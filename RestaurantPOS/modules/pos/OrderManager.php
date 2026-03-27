<?php
/**
 * POS Order Manager — handles order lifecycle
 */
class OrderManager {

    /** Create a new order */
    public static function create(array $data): int {
        $orderNumber = generateOrderNumber($data['location_id']);
        Database::beginTransaction();
        try {
            $orderId = Database::insert(
                "INSERT INTO orders
                    (tenant_id, location_id, table_id, customer_id, user_id,
                     order_number, order_type, status, subtotal, tax_amount,
                     discount_amount, tip_amount, delivery_fee, total_amount, notes, source)
                 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                [
                    $data['tenant_id'],
                    $data['location_id'],
                    $data['table_id']     ?? null,
                    $data['customer_id']  ?? null,
                    $data['user_id'],
                    $orderNumber,
                    $data['order_type']   ?? 'dine-in',
                    'pending',
                    0, 0, 0, 0, 0, 0,
                    $data['notes']        ?? null,
                    $data['source']       ?? 'pos',
                ]
            );
            Database::commit();
            auditLog('order_created', 'orders', $orderId, [], ['order_number' => $orderNumber]);
            return $orderId;
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Add item to an open order */
    public static function addItem(int $orderId, array $item): int {
        $menuItem = Database::fetchOne(
            "SELECT * FROM menu_items WHERE item_id = ?",
            [$item['item_id']]
        );
        if (!$menuItem) throw new RuntimeException('Menu item not found');

        $unitPrice    = (float) $menuItem['price'];
        $modifierTotal = 0.00;
        $modifiers     = $item['modifiers'] ?? [];

        // Calculate modifier totals
        if (!empty($modifiers)) {
            foreach ($modifiers as $modId) {
                $mod = Database::fetchOne(
                    "SELECT price_add FROM modifiers WHERE modifier_id = ?",
                    [$modId]
                );
                if ($mod) $modifierTotal += (float) $mod['price_add'];
            }
        }

        $lineTotal = ($unitPrice + $modifierTotal) * (int) ($item['quantity'] ?? 1);

        Database::beginTransaction();
        try {
            $orderItemId = Database::insert(
                "INSERT INTO order_items
                    (order_id, item_id, variant_id, quantity, unit_price, modifier_total, line_total, notes, status)
                 VALUES (?,?,?,?,?,?,?,?,?)",
                [
                    $orderId,
                    $item['item_id'],
                    $item['variant_id'] ?? null,
                    $item['quantity']   ?? 1,
                    $unitPrice,
                    $modifierTotal,
                    $lineTotal,
                    $item['notes']      ?? null,
                    'pending',
                ]
            );

            // Save modifier selections
            foreach ($modifiers as $modId) {
                $mod = Database::fetchOne(
                    "SELECT modifier_name, price_add FROM modifiers WHERE modifier_id = ?",
                    [$modId]
                );
                if ($mod) {
                    Database::query(
                        "INSERT INTO order_item_modifiers (order_item_id, modifier_id, modifier_name, price_add)
                         VALUES (?,?,?,?)",
                        [$orderItemId, $modId, $mod['modifier_name'], $mod['price_add']]
                    );
                }
            }

            self::recalculate($orderId);
            Database::commit();
            return $orderItemId;
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Remove item from order */
    public static function removeItem(int $orderItemId): void {
        $item = Database::fetchOne("SELECT order_id FROM order_items WHERE order_item_id = ?", [$orderItemId]);
        Database::query("DELETE FROM order_item_modifiers WHERE order_item_id = ?", [$orderItemId]);
        Database::query("DELETE FROM order_items WHERE order_item_id = ?", [$orderItemId]);
        if ($item) self::recalculate($item['order_id']);
    }

    /** Recalculate order totals */
    public static function recalculate(int $orderId): void {
        $order = Database::fetchOne(
            "SELECT o.*, l.tax_rate FROM orders o
             JOIN locations l ON l.location_id = o.location_id
             WHERE o.order_id = ?",
            [$orderId]
        );
        if (!$order) return;

        $subtotal = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(line_total), 0) FROM order_items WHERE order_id = ? AND status != 'void'",
            [$orderId]
        );
        $discount = (float) ($order['discount_amount'] ?? 0);
        $taxable  = max(0, $subtotal - $discount);
        $tax      = round($taxable * (float) $order['tax_rate'], 2);
        $delivery = (float) ($order['delivery_fee'] ?? 0);
        $tip      = (float) ($order['tip_amount']   ?? 0);
        $total    = $taxable + $tax + $delivery + $tip;

        Database::query(
            "UPDATE orders SET subtotal=?, tax_amount=?, total_amount=?, updated_at=datetime('now')
             WHERE order_id=?",
            [$subtotal, $tax, $total, $orderId]
        );
    }

    /** Apply discount to order */
    public static function applyDiscount(int $orderId, int $discountId): bool {
        $discount = Database::fetchOne(
            "SELECT * FROM discounts WHERE discount_id = ? AND is_active = 1
             AND (start_date IS NULL OR start_date <= datetime('now'))
             AND (end_date   IS NULL OR end_date   >= datetime('now'))",
            [$discountId]
        );
        if (!$discount) return false;

        $order = Database::fetchOne("SELECT subtotal FROM orders WHERE order_id = ?", [$orderId]);
        $amount = $discount['discount_type'] === 'percentage'
            ? round((float)$order['subtotal'] * (float)$discount['value'] / 100, 2)
            : (float) $discount['value'];

        Database::query(
            "UPDATE orders SET discount_amount = ?, updated_at = datetime('now') WHERE order_id = ?",
            [$amount, $orderId]
        );
        Database::query(
            "UPDATE discounts SET uses_count = uses_count + 1 WHERE discount_id = ?",
            [$discountId]
        );
        self::recalculate($orderId);
        return true;
    }

    /** Update order status */
    public static function updateStatus(int $orderId, string $status): void {
        $completedAt = in_array($status, ['completed', 'served']) ? 'datetime('now')' : 'NULL';
        Database::query(
            "UPDATE orders SET status = ?, completed_at = $completedAt, updated_at = datetime('now')
             WHERE order_id = ?",
            [$status, $orderId]
        );
        // If table dine-in completed, free the table
        if ($status === 'completed') {
            Database::query(
                "UPDATE restaurant_tables SET status = 'available'
                 WHERE table_id = (SELECT table_id FROM orders WHERE order_id = ?)",
                [$orderId]
            );
        }
        auditLog("order_status_$status", 'orders', $orderId);
    }

    /** Get full order with items */
    public static function getOrder(int $orderId): ?array {
        $order = Database::fetchOne(
            "SELECT o.*, u.first_name || ' ' || u.last_name AS cashier_name,
                    t.table_number,
                    c.first_name || ' ' || COALESCE(c.last_name,'') AS customer_name
             FROM orders o
             LEFT JOIN users u ON u.user_id = o.user_id
             LEFT JOIN restaurant_tables t ON t.table_id = o.table_id
             LEFT JOIN customers c ON c.customer_id = o.customer_id
             WHERE o.order_id = ?",
            [$orderId]
        );
        if (!$order) return null;

        $order['items'] = Database::fetchAll(
            "SELECT oi.*, mi.item_name, mi.image_url
             FROM order_items oi
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE oi.order_id = ?",
            [$orderId]
        );
        foreach ($order['items'] as &$item) {
            $item['modifiers'] = Database::fetchAll(
                "SELECT * FROM order_item_modifiers WHERE order_item_id = ?",
                [$item['order_item_id']]
            );
        }
        $order['payments'] = Database::fetchAll(
            "SELECT * FROM payments WHERE order_id = ?",
            [$orderId]
        );
        return $order;
    }

    /** List orders for a location with optional filters */
    public static function listOrders(int $locationId, array $filters = [], int $page = 1): array {
        $where  = ['o.location_id = ?'];
        $params = [$locationId];

        if (!empty($filters['status'])) {
            $where[]  = 'o.status = ?';
            $params[] = $filters['status'];
        }
        if (!empty($filters['order_type'])) {
            $where[]  = 'o.order_type = ?';
            $params[] = $filters['order_type'];
        }
        if (!empty($filters['date'])) {
            $where[]  = 'date(o.created_at) = ?';
            $params[] = $filters['date'];
        }
        if (!empty($filters['search'])) {
            $where[]  = "(o.order_number LIKE ? OR c.first_name LIKE ? OR c.phone LIKE ?)";
            $term     = '%' . $filters['search'] . '%';
            $params   = array_merge($params, [$term, $term, $term]);
        }

        $whereStr = implode(' AND ', $where);
        $sql = "SELECT o.order_id, o.order_number, o.order_type, o.status,
                       o.total_amount, o.created_at, o.source,
                       t.table_number,
                       c.first_name || ' ' || COALESCE(c.last_name,'') AS customer_name,
                       u.first_name || ' ' || u.last_name AS cashier_name
                FROM orders o
                LEFT JOIN restaurant_tables t ON t.table_id = o.table_id
                LEFT JOIN customers c ON c.customer_id = o.customer_id
                LEFT JOIN users u ON u.user_id = o.user_id
                WHERE $whereStr";

        $perPage = 25;
        $offset  = ($page - 1) * $perPage;
        $total   = (int) Database::fetchValue(
            "SELECT COUNT(*) FROM orders o
             LEFT JOIN customers c ON c.customer_id = o.customer_id
             WHERE $whereStr",
            $params
        );
        $rows = Database::fetchAll(
            "$sql ORDER BY o.created_at DESC LIMIT $perPage OFFSET $offset",
            $params
        );

        return [
            'data'         => $rows,
            'total'        => $total,
            'per_page'     => $perPage,
            'current_page' => $page,
            'last_page'    => (int) ceil($total / $perPage),
        ];
    }

    /** Get open orders for KDS */
    public static function getKDSTickets(int $locationId): array {
        return Database::fetchAll(
            "SELECT o.order_id, o.order_number, o.order_type, o.created_at,
                    o.notes, t.table_number,
                    CAST((julianday(datetime('now')) - julianday(o.created_at)) * 1440 AS INTEGER) AS elapsed_minutes
             FROM orders o
             LEFT JOIN restaurant_tables t ON t.table_id = o.table_id
             WHERE o.location_id = ? AND o.status IN ('confirmed','preparing')
             ORDER BY o.created_at ASC",
            [$locationId]
        );
    }
}
