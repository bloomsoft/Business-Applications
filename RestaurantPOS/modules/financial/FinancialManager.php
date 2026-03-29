<?php
/**
 * Financial Management Module
 * Cash drawers, expenses, tax reports, end-of-day closing
 */
class FinancialManager {

    // ── Cash Drawer ──────────────────────────────────────────────────────────

    public static function openDrawer(int $locationId, float $openingFloat): int {
        // Close any previously open drawer for this location
        Database::query(
            "UPDATE cash_drawers SET status = 'closed', closed_at = datetime('now')
             WHERE location_id = ? AND status = 'open'",
            [$locationId]
        );
        return Database::insert(
            "INSERT INTO cash_drawers (location_id, user_id, opening_float, status)
             VALUES (?,?,?,'open')",
            [$locationId, Auth::id(), $openingFloat]
        );
    }

    public static function closeDrawer(int $drawerId, float $closingAmount): array {
        $drawer = Database::fetchOne(
            "SELECT d.*, l.location_id FROM cash_drawers d
             JOIN locations l ON l.location_id = d.location_id
             WHERE d.drawer_id = ?",
            [$drawerId]
        );

        // Calculate expected cash (opening float + all cash payments - refunds)
        $cashIn = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(p.amount),0)
             FROM payments p
             JOIN orders o ON o.order_id = p.order_id
             WHERE o.location_id = ? AND p.payment_method = 'cash'
               AND p.status = 'completed'
               AND p.processed_at >= ?",
            [$drawer['location_id'], $drawer['opened_at']]
        );
        $expected = (float)$drawer['opening_float'] + $cashIn;
        $variance = $closingAmount - $expected;

        Database::query(
            "UPDATE cash_drawers
             SET closing_amount = ?, expected_amount = ?, variance = ?,
                 status = 'closed', closed_at = datetime('now')
             WHERE drawer_id = ?",
            [$closingAmount, $expected, $variance, $drawerId]
        );

        auditLog('drawer_closed', 'cash_drawers', $drawerId);
        return ['expected' => $expected, 'actual' => $closingAmount, 'variance' => $variance];
    }

    // ── Expenses ─────────────────────────────────────────────────────────────

    public static function addExpense(array $data): int {
        return Database::insert(
            "INSERT INTO expenses
                (tenant_id, location_id, category, description, amount, expense_date, receipt_url, created_by)
             VALUES (?,?,?,?,?,?,?,?)",
            [
                $data['tenant_id'],
                $data['location_id'],
                $data['category'],
                $data['description'] ?? null,
                $data['amount'],
                $data['expense_date'],
                $data['receipt_url'] ?? null,
                Auth::id(),
            ]
        );
    }

    public static function getExpenses(int $locationId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT e.*, u.first_name || ' ' || u.last_name AS created_by_name
             FROM expenses e
             LEFT JOIN users u ON u.user_id = e.created_by
             WHERE e.location_id = ? AND e.expense_date BETWEEN ? AND ?
             ORDER BY e.expense_date DESC",
            [$locationId, $startDate, $endDate]
        );
    }

    public static function getExpenseSummary(int $locationId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT category,
                    COUNT(*)         AS item_count,
                    SUM(amount)      AS total_amount
             FROM expenses
             WHERE location_id = ? AND expense_date BETWEEN ? AND ?
             GROUP BY category
             ORDER BY total_amount DESC",
            [$locationId, $startDate, $endDate]
        );
    }

    // ── End of Day Report ────────────────────────────────────────────────────

    public static function getEODReport(int $locationId, string $date = ''): array {
        $date = $date ?: date('Y-m-d');

        $sales = Database::fetchOne(
            "SELECT
                COUNT(*)                              AS total_orders,
                COALESCE(SUM(subtotal),0)               AS subtotal,
                COALESCE(SUM(discount_amount),0)        AS total_discounts,
                COALESCE(SUM(tax_amount),0)             AS total_tax,
                COALESCE(SUM(tip_amount),0)             AS total_tips,
                COALESCE(SUM(delivery_fee),0)           AS total_delivery_fees,
                COALESCE(SUM(total_amount),0)           AS gross_revenue,
                SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) AS cancelled_orders,
                SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END) AS completed_orders
             FROM orders
             WHERE location_id = ? AND date(created_at) = ?",
            [$locationId, $date]
        );

        $paymentBreakdown = Database::fetchAll(
            "SELECT p.payment_method,
                    COUNT(*)         AS txn_count,
                    SUM(p.amount)    AS total
             FROM payments p
             JOIN orders o ON o.order_id = p.order_id
             WHERE o.location_id = ? AND date(p.processed_at) = ? AND p.status='completed'
             GROUP BY p.payment_method",
            [$locationId, $date]
        );

        $topItems = Database::fetchAll(
            "SELECT mi.item_name, SUM(oi.quantity) AS qty, SUM(oi.line_total) AS revenue
             FROM order_items oi
             JOIN orders o ON o.order_id = oi.order_id
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE o.location_id = ? AND date(o.created_at) = ? AND o.status='completed'
             GROUP BY mi.item_name
             ORDER BY qty DESC
             LIMIT 5",
            [$locationId, $date]
        );

        $voids = Database::fetchAll(
            "SELECT oi.*, mi.item_name FROM order_items oi
             JOIN orders o ON o.order_id = oi.order_id
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE o.location_id = ? AND date(oi.created_at) = ? AND oi.status='void'",
            [$locationId, $date]
        );

        $refunds = Database::fetchAll(
            "SELECT r.*, p.payment_method FROM refunds r
             JOIN payments p ON p.payment_id = r.payment_id
             JOIN orders o ON o.order_id = r.order_id
             WHERE o.location_id = ? AND date(r.created_at) = ?",
            [$locationId, $date]
        );

        return [
            'date'              => $date,
            'sales'             => $sales,
            'payment_breakdown' => $paymentBreakdown,
            'top_items'         => $topItems,
            'voids'             => $voids,
            'refunds'           => $refunds,
        ];
    }

    // ── Tax Report ───────────────────────────────────────────────────────────

    public static function getTaxReport(int $locationId, string $startDate, string $endDate): array {
        return Database::fetchOne(
            "SELECT
                COALESCE(SUM(subtotal),0)      AS net_sales,
                COALESCE(SUM(tax_amount),0)    AS tax_collected,
                COALESCE(SUM(total_amount),0)  AS gross_sales,
                COUNT(*)                     AS transaction_count
             FROM orders
             WHERE location_id = ? AND status = 'completed'
               AND date(created_at) BETWEEN ? AND ?",
            [$locationId, $startDate, $endDate]
        );
    }

    // ── Invoice Generation ────────────────────────────────────────────────────

    public static function generateInvoiceHTML(int $orderId): string {
        $order = OrderManager::getOrder($orderId);
        if (!$order) return '';

        $location = LocationManager::get($order['location_id']);
        $tenant   = Database::fetchOne("SELECT * FROM tenants WHERE tenant_id = ?", [Auth::tenantId()]);

        ob_start();
        ?>
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; font-size: 12px; }
                .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; }
                .items table { width: 100%; border-collapse: collapse; }
                .items th, .items td { border-bottom: 1px solid #ddd; padding: 4px 8px; }
                .totals { text-align: right; margin-top: 10px; }
                .footer { text-align: center; margin-top: 20px; font-size: 11px; }
            </style>
        </head>
        <body>
            <div class="header">
                <h2><?= sanitize($tenant['company_name'] ?? '') ?></h2>
                <p><?= sanitize($location['address'] ?? '') ?><br>
                   <?= sanitize($location['phone'] ?? '') ?></p>
                <p><strong>Order #<?= sanitize($order['order_number']) ?></strong><br>
                   <?= fmtDateTime($order['created_at']) ?></p>
            </div>
            <?php if ($order['table_number']): ?>
            <p>Table: <?= sanitize($order['table_number']) ?></p>
            <?php endif; ?>
            <div class="items">
                <table>
                    <tr><th>Item</th><th>Qty</th><th>Price</th><th>Total</th></tr>
                    <?php foreach ($order['items'] as $item): ?>
                    <tr>
                        <td><?= sanitize($item['item_name']) ?></td>
                        <td><?= (int)$item['quantity'] ?></td>
                        <td><?= money($item['unit_price']) ?></td>
                        <td><?= money($item['line_total']) ?></td>
                    </tr>
                    <?php foreach ($item['modifiers'] as $mod): ?>
                    <tr><td colspan="3" style="padding-left:20px">+ <?= sanitize($mod['modifier_name']) ?></td>
                        <td><?= money($mod['price_add']) ?></td></tr>
                    <?php endforeach; ?>
                    <?php endforeach; ?>
                </table>
            </div>
            <div class="totals">
                <p>Subtotal: <?= money($order['subtotal']) ?></p>
                <?php if ($order['discount_amount'] > 0): ?>
                <p>Discount: -<?= money($order['discount_amount']) ?></p>
                <?php endif; ?>
                <p>Tax: <?= money($order['tax_amount']) ?></p>
                <?php if ($order['tip_amount'] > 0): ?>
                <p>Tip: <?= money($order['tip_amount']) ?></p>
                <?php endif; ?>
                <p><strong>Total: <?= money($order['total_amount']) ?></strong></p>
            </div>
            <div class="footer">
                <p>Thank you for dining with us!</p>
                <p>Powered by RestaurantPOS SaaS</p>
            </div>
        </body>
        </html>
        <?php
        return ob_get_clean();
    }
}
