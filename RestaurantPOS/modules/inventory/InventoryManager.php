<?php
/**
 * Inventory Management Module
 */
class InventoryManager {

    /** Get all inventory items with stock status */
    public static function getItems(int $locationId, array $filters = []): array {
        $where  = ['i.location_id = ?', 'i.is_active = 1'];
        $params = [$locationId];

        if (!empty($filters['category'])) {
            $where[]  = 'i.inv_cat_id = ?';
            $params[] = $filters['category'];
        }
        if (!empty($filters['low_stock'])) {
            $where[] = 'i.quantity_on_hand <= i.reorder_level';
        }
        if (!empty($filters['search'])) {
            $where[]  = '(i.item_name LIKE ? OR i.sku LIKE ?)';
            $t        = '%' . $filters['search'] . '%';
            $params   = array_merge($params, [$t, $t]);
        }

        $whereStr = implode(' AND ', $where);
        return Database::fetchAll(
            "SELECT i.*, c.cat_name AS category_name, s.supplier_name,
                    CASE WHEN i.quantity_on_hand <= 0               THEN 'out'
                         WHEN i.quantity_on_hand <= i.reorder_level THEN 'low'
                         ELSE 'ok'
                    END AS stock_status
             FROM inventory_items i
             LEFT JOIN inventory_categories c ON c.inv_cat_id = i.inv_cat_id
             LEFT JOIN suppliers s ON s.supplier_id = i.supplier_id
             WHERE $whereStr
             ORDER BY stock_status, i.item_name",
            $params
        );
    }

    /** Adjust stock quantity (in/out/waste/adjustment) */
    public static function adjustStock(
        int    $invItemId,
        float  $quantity,
        string $type,
        string $notes     = '',
        int    $refId     = 0,
        string $refType   = ''
    ): void {
        $item = Database::fetchOne(
            "SELECT quantity_on_hand, location_id FROM inventory_items WHERE inv_item_id = ?",
            [$invItemId]
        );
        if (!$item) throw new RuntimeException('Inventory item not found');

        $newBalance = $type === 'out' || $type === 'waste'
            ? (float)$item['quantity_on_hand'] - $quantity
            : (float)$item['quantity_on_hand'] + $quantity;

        Database::beginTransaction();
        try {
            Database::query(
                "UPDATE inventory_items
                 SET quantity_on_hand = ?, updated_at = datetime('now')
                 WHERE inv_item_id = ?",
                [$newBalance, $invItemId]
            );
            Database::query(
                "INSERT INTO inventory_movements
                    (inv_item_id, location_id, movement_type, quantity, balance_after,
                     reference_id, reference_type, notes, created_by)
                 VALUES (?,?,?,?,?,?,?,?,?)",
                [
                    $invItemId, $item['location_id'], $type, $quantity, $newBalance,
                    $refId ?: null, $refType ?: null, $notes, Auth::id(),
                ]
            );
            // Check if reorder needed → create notification
            $updated = Database::fetchOne(
                "SELECT item_name, quantity_on_hand, reorder_level, tenant_id FROM inventory_items
                 WHERE inv_item_id = ?",
                [$invItemId]
            );
            if ((float)$updated['quantity_on_hand'] <= (float)$updated['reorder_level']) {
                self::sendLowStockAlert($updated);
            }
            Database::commit();
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Deduct inventory based on recipe when an order item is sold */
    public static function deductFromOrder(int $orderItemId): void {
        $orderItem = Database::fetchOne(
            "SELECT oi.item_id, oi.quantity, o.location_id
             FROM order_items oi
             JOIN orders o ON o.order_id = oi.order_id
             WHERE oi.order_item_id = ?",
            [$orderItemId]
        );
        if (!$orderItem) return;

        $recipes = Database::fetchAll(
            "SELECT r.inv_item_id, r.quantity_used * ? AS total_qty
             FROM recipes r
             WHERE r.item_id = ?",
            [$orderItem['quantity'], $orderItem['item_id']]
        );

        foreach ($recipes as $recipe) {
            // Find inventory item at this location
            $inv = Database::fetchOne(
                "SELECT inv_item_id FROM inventory_items
                 WHERE inv_item_id = ? AND location_id = ?",
                [$recipe['inv_item_id'], $orderItem['location_id']]
            );
            if ($inv) {
                self::adjustStock(
                    $inv['inv_item_id'],
                    (float) $recipe['total_qty'],
                    'out',
                    'Auto-deducted from order',
                    $orderItemId,
                    'order_item'
                );
            }
        }
    }

    /** Create purchase order */
    public static function createPurchaseOrder(array $data): int {
        $poNumber = 'PO-' . date('Ymd') . '-' . str_pad(mt_rand(0, 999), 3, '0', STR_PAD_LEFT);

        Database::beginTransaction();
        try {
            $poId = Database::insert(
                "INSERT INTO purchase_orders
                    (tenant_id, location_id, supplier_id, po_number, status,
                     order_date, expected_date, notes, created_by)
                 VALUES (?,?,?,?,?,?,?,?,?)",
                [
                    $data['tenant_id'],
                    $data['location_id'],
                    $data['supplier_id'] ?? null,
                    $poNumber,
                    'draft',
                    $data['order_date']    ?? date('Y-m-d'),
                    $data['expected_date'] ?? null,
                    $data['notes']         ?? null,
                    Auth::id(),
                ]
            );

            $subtotal = 0;
            foreach ($data['items'] as $item) {
                $lineTotal = (float)$item['qty_ordered'] * (float)$item['unit_cost'];
                $subtotal += $lineTotal;
                Database::query(
                    "INSERT INTO purchase_order_items
                        (po_id, inv_item_id, qty_ordered, unit_cost, line_total)
                     VALUES (?,?,?,?,?)",
                    [$poId, $item['inv_item_id'], $item['qty_ordered'], $item['unit_cost'], $lineTotal]
                );
            }

            Database::query(
                "UPDATE purchase_orders SET subtotal = ?, total_amount = ? WHERE po_id = ?",
                [$subtotal, $subtotal, $poId]
            );

            Database::commit();
            return $poId;
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Receive a purchase order — updates stock */
    public static function receivePurchaseOrder(int $poId, array $received): void {
        Database::beginTransaction();
        try {
            foreach ($received as $poItemId => $qtyReceived) {
                $poItem = Database::fetchOne(
                    "SELECT poi.*, po.location_id FROM purchase_order_items poi
                     JOIN purchase_orders po ON po.po_id = poi.po_id
                     WHERE poi.po_item_id = ?",
                    [$poItemId]
                );
                if (!$poItem) continue;

                Database::query(
                    "UPDATE purchase_order_items SET qty_received = ? WHERE po_item_id = ?",
                    [$qtyReceived, $poItemId]
                );
                self::adjustStock(
                    $poItem['inv_item_id'],
                    (float)$qtyReceived,
                    'in',
                    "Received via PO #$poId",
                    $poId,
                    'purchase_order'
                );
                // Update last restocked date
                Database::query(
                    "UPDATE inventory_items SET last_restocked = datetime('now') WHERE inv_item_id = ?",
                    [$poItem['inv_item_id']]
                );
            }

            // Update PO status
            $allReceived = Database::fetchValue(
                "SELECT COUNT(*) FROM purchase_order_items
                 WHERE po_id = ? AND qty_received < qty_ordered",
                [$poId]
            );
            $status = (int)$allReceived === 0 ? 'received' : 'partial';
            Database::query(
                "UPDATE purchase_orders SET status = ?, received_date = CAST(datetime('now') AS DATE)
                 WHERE po_id = ?",
                [$status, $poId]
            );

            Database::commit();
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Get low-stock items */
    public static function getLowStockItems(int $tenantId): array {
        return Database::fetchAll(
            "SELECT i.*, l.location_name, s.supplier_name, s.email AS supplier_email
             FROM inventory_items i
             JOIN locations l ON l.location_id = i.location_id
             LEFT JOIN suppliers s ON s.supplier_id = i.supplier_id
             WHERE l.tenant_id = ? AND i.quantity_on_hand <= i.reorder_level AND i.is_active = 1
             ORDER BY (i.quantity_on_hand - i.reorder_level)",
            [$tenantId]
        );
    }

    /** Get movement history */
    public static function getMovements(int $invItemId, int $page = 1): array {
        $perPage = 30;
        $offset  = ($page - 1) * $perPage;
        $rows = Database::fetchAll(
            "SELECT m.*, u.first_name || ' ' || u.last_name AS created_by_name
             FROM inventory_movements m
             LEFT JOIN users u ON u.user_id = m.created_by
             WHERE m.inv_item_id = ?
             ORDER BY m.created_at DESC
             OFFSET ? ROWS FETCH NEXT ? ROWS ONLY",
            [$invItemId, $offset, $perPage]
        );
        return $rows;
    }

    private static function sendLowStockAlert(array $item): void {
        Database::query(
            "INSERT INTO notifications (tenant_id, type, title, message)
             VALUES (?, 'low_stock', ?, ?)",
            [
                $item['tenant_id'],
                'Low Stock Alert: ' . $item['item_name'],
                $item['item_name'] . ' is running low ('. $item['quantity_on_hand'] .' remaining)',
            ]
        );
    }
}
