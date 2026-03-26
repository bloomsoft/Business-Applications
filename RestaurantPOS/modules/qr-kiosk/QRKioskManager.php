<?php
/**
 * QR Code Self-Ordering & Kiosk Module
 */
class QRKioskManager {

    /** Generate QR code URL for a table */
    public static function getTableQRUrl(int $tableId): string {
        $token = TableManager::getQRToken($tableId);
        if (!$token) {
            // Generate token if missing
            $token = bin2hex(random_bytes(16));
            Database::query(
                "UPDATE restaurant_tables SET qr_code_token = ? WHERE table_id = ?",
                [$token, $tableId]
            );
        }
        $orderUrl = APP_URL . '/order.php?t=' . urlencode($token);
        return qrCodeUrl($orderUrl, 300);
    }

    /** Get menu for QR/Kiosk ordering (public — no auth) */
    public static function getPublicMenu(int $tenantId): array {
        $categories = Database::fetchAll(
            "SELECT * FROM menu_categories
             WHERE tenant_id = ? AND is_active = 1 AND parent_id IS NULL
             ORDER BY sort_order",
            [$tenantId]
        );

        foreach ($categories as &$cat) {
            $cat['items'] = Database::fetchAll(
                "SELECT item_id, item_name, description, price, image_url, calories, prep_time_min
                 FROM menu_items
                 WHERE category_id = ? AND is_available = 1
                 ORDER BY sort_order, item_name",
                [$cat['category_id']]
            );
            foreach ($cat['items'] as &$item) {
                $item['modifier_groups'] = self::getItemModifiers($item['item_id']);
            }
        }
        return $categories;
    }

    /** Get modifier groups for an item */
    public static function getItemModifiers(int $itemId): array {
        $groups = Database::fetchAll(
            "SELECT mg.* FROM modifier_groups mg
             JOIN item_modifier_groups img ON img.group_id = mg.group_id
             WHERE img.item_id = ?",
            [$itemId]
        );
        foreach ($groups as &$group) {
            $group['options'] = Database::fetchAll(
                "SELECT * FROM modifiers WHERE group_id = ? AND is_available = 1 ORDER BY sort_order",
                [$group['group_id']]
            );
        }
        return $groups;
    }

    /** Place order from QR/Kiosk */
    public static function placeOrder(string $tableToken, array $cartItems, array $customerInfo = []): array {
        $table = TableManager::getByQRToken($tableToken);
        if (!$table) return ['success' => false, 'message' => 'Invalid table QR code'];

        // Optionally find/create customer
        $customerId = null;
        if (!empty($customerInfo['phone']) || !empty($customerInfo['email'])) {
            $customerId = self::findOrCreateCustomer($customerInfo, $table['tenant_id']);
        }

        Database::beginTransaction();
        try {
            $orderId = OrderManager::create([
                'tenant_id'   => $table['tenant_id'],
                'location_id' => $table['location_id'],
                'table_id'    => $table['table_id'],
                'customer_id' => $customerId,
                'user_id'     => 0,  // system order
                'order_type'  => 'qr-order',
                'source'      => 'qr',
                'notes'       => $customerInfo['notes'] ?? null,
            ]);

            foreach ($cartItems as $cartItem) {
                OrderManager::addItem($orderId, $cartItem);
            }

            // Mark table as occupied
            TableManager::updateStatus($table['table_id'], 'occupied');

            // Notify kitchen
            Database::query(
                "INSERT INTO notifications (tenant_id, type, title, message)
                 VALUES (?, 'new_order', ?, ?)",
                [
                    $table['tenant_id'],
                    'New QR Order — Table ' . $table['table_number'],
                    'Self-service order placed for table ' . $table['table_number'],
                ]
            );

            Database::commit();

            $order = OrderManager::getOrder($orderId);
            return ['success' => true, 'order_id' => $orderId, 'order' => $order];
        } catch (Throwable $e) {
            Database::rollback();
            return ['success' => false, 'message' => $e->getMessage()];
        }
    }

    /** Kiosk self-checkout order (takeout / dine-in without table) */
    public static function kioskOrder(int $locationId, int $tenantId, array $cartItems, string $orderType = 'takeout', array $customerInfo = []): array {
        $customerId = null;
        if (!empty($customerInfo)) {
            $customerId = self::findOrCreateCustomer($customerInfo, $tenantId);
        }

        Database::beginTransaction();
        try {
            $orderId = OrderManager::create([
                'tenant_id'   => $tenantId,
                'location_id' => $locationId,
                'customer_id' => $customerId,
                'user_id'     => 0,
                'order_type'  => $orderType,
                'source'      => 'kiosk',
            ]);

            foreach ($cartItems as $item) {
                OrderManager::addItem($orderId, $item);
            }

            Database::commit();
            return ['success' => true, 'order_id' => $orderId, 'order' => OrderManager::getOrder($orderId)];
        } catch (Throwable $e) {
            Database::rollback();
            return ['success' => false, 'message' => $e->getMessage()];
        }
    }

    /** Get order status for customer-facing display */
    public static function getOrderStatus(int $orderId, int $tenantId): ?array {
        $order = Database::fetchOne(
            "SELECT order_id, order_number, order_type, status, total_amount,
                    created_at, estimated_ready_time
             FROM orders
             WHERE order_id = ? AND tenant_id = ?",
            [$orderId, $tenantId]
        );
        if (!$order) return null;

        $order['items'] = Database::fetchAll(
            "SELECT oi.quantity, mi.item_name, oi.status
             FROM order_items oi
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE oi.order_id = ?",
            [$orderId]
        );
        return $order;
    }

    /** Generate feedback QR after meal */
    public static function getFeedbackQRUrl(int $orderId, int $tenantId): string {
        $token = base64_encode("$tenantId:$orderId:" . substr(md5(APP_KEY . $orderId), 0, 8));
        $url   = APP_URL . '/feedback.php?token=' . urlencode($token);
        return qrCodeUrl($url, 250);
    }

    /** Print-ready QR code HTML for a table */
    public static function printTableQR(int $tableId): string {
        $table  = Database::fetchOne("SELECT * FROM restaurant_tables WHERE table_id = ?", [$tableId]);
        $qrUrl  = self::getTableQRUrl($tableId);
        $orderUrl = APP_URL . '/order.php?t=' . urlencode($table['qr_code_token'] ?? '');

        return <<<HTML
        <!DOCTYPE html>
        <html>
        <head><meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 40px; }
            .qr-card { border: 2px solid #333; border-radius: 12px; padding: 30px; display: inline-block; max-width: 300px; }
            .qr-card h2 { margin: 0 0 5px; font-size: 24px; }
            .qr-card p  { color: #666; font-size: 14px; margin: 0 0 20px; }
            .qr-card img { width: 200px; height: 200px; }
            .qr-card small { display: block; margin-top: 15px; color: #999; font-size: 11px; word-break: break-all; }
        </style>
        </head>
        <body>
            <div class="qr-card">
                <h2>Table {$table['table_number']}</h2>
                <p>Scan to order & pay</p>
                <img src="{$qrUrl}" alt="QR Code">
                <small>{$orderUrl}</small>
            </div>
        </body>
        </html>
        HTML;
    }

    private static function findOrCreateCustomer(array $info, int $tenantId): int {
        $existing = null;
        if (!empty($info['phone'])) {
            $existing = Database::fetchValue(
                "SELECT customer_id FROM customers WHERE tenant_id = ? AND phone = ?",
                [$tenantId, $info['phone']]
            );
        }
        if (!$existing && !empty($info['email'])) {
            $existing = Database::fetchValue(
                "SELECT customer_id FROM customers WHERE tenant_id = ? AND email = ?",
                [$tenantId, $info['email']]
            );
        }
        if ($existing) return (int)$existing;

        return Database::insert(
            "INSERT INTO customers (tenant_id, first_name, last_name, email, phone)
             VALUES (?,?,?,?,?)",
            [$tenantId, $info['first_name'] ?? null, $info['last_name'] ?? null,
             $info['email'] ?? null, $info['phone'] ?? null]
        );
    }
}
