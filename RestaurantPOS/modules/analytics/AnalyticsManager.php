<?php
/**
 * Analytics & Reporting Engine
 */
class AnalyticsManager {

    /** Dashboard KPI summary */
    public static function getDashboardKPIs(int $locationId, string $date = ''): array {
        $date = $date ?: date('Y-m-d');
        $prevDate = date('Y-m-d', strtotime($date . ' -1 day'));

        $today = Database::fetchOne(
            "SELECT
                COUNT(*)                                     AS total_orders,
                COALESCE(SUM(total_amount),0)                  AS total_revenue,
                COALESCE(AVG(total_amount),0)                  AS avg_order_value,
                COUNT(DISTINCT customer_id)                  AS unique_customers,
                SUM(CASE WHEN order_type='dine-in'  THEN 1 ELSE 0 END) AS dine_in,
                SUM(CASE WHEN order_type='takeout'  THEN 1 ELSE 0 END) AS takeout,
                SUM(CASE WHEN order_type='delivery' THEN 1 ELSE 0 END) AS delivery,
                SUM(CASE WHEN order_type IN ('qr-order','kiosk') THEN 1 ELSE 0 END) AS self_service
             FROM orders
             WHERE location_id = ? AND date(created_at) = ? AND status = 'completed'",
            [$locationId, $date]
        );

        $prev = Database::fetchOne(
            "SELECT COALESCE(SUM(total_amount),0) AS total_revenue,
                    COUNT(*) AS total_orders
             FROM orders
             WHERE location_id = ? AND date(created_at) = ? AND status = 'completed'",
            [$locationId, $prevDate]
        );

        $today['revenue_change'] = $prev['total_revenue'] > 0
            ? round((($today['total_revenue'] - $prev['total_revenue']) / $prev['total_revenue']) * 100, 1)
            : 0;
        $today['orders_change']  = $prev['total_orders'] > 0
            ? round((($today['total_orders'] - $prev['total_orders']) / $prev['total_orders']) * 100, 1)
            : 0;

        return $today;
    }

    /** Revenue by period (daily/weekly/monthly) */
    public static function getRevenueChart(int $locationId, string $period = 'daily', int $days = 30): array {
        $groupBy = match($period) {
            'weekly'  => "DATEPART(YEAR, created_at), DATEPART(WEEK, created_at)",
            'monthly' => "YEAR(created_at), MONTH(created_at)",
            default   => "date(created_at)",
        };
        $label = match($period) {
            'weekly'  => "CAST(DATEPART(YEAR,created_at) AS VARCHAR) + '-W' + CAST(DATEPART(WEEK,created_at) AS VARCHAR)",
            'monthly' => "FORMAT(created_at,'yyyy-MM')",
            default   => "date(created_at)",
        };

        return Database::fetchAll(
            "SELECT $label AS period,
                    COALESCE(SUM(total_amount),0)  AS revenue,
                    COUNT(*)                      AS order_count,
                    COALESCE(AVG(total_amount),0)   AS avg_order_value
             FROM orders
             WHERE location_id = ?
               AND created_at >= DATEADD(DAY,-?,datetime('now'))
               AND status = 'completed'
             GROUP BY $groupBy
             ORDER BY MIN(created_at)",
            [$locationId, $days]
        );
    }

    /** Top-selling items */
    public static function getTopItems(int $tenantId, string $startDate = '', string $endDate = '', int $limit = 10): array {
        $startDate = $startDate ?: date('Y-m-d', strtotime('-30 days'));
        $endDate   = $endDate   ?: date('Y-m-d');

        return Database::fetchAll(
            "SELECT mi.item_id, mi.item_name, mi.category_id,
                    SUM(oi.quantity)     AS qty_sold,
                    SUM(oi.line_total)   AS total_revenue,
                    AVG(oi.unit_price)   AS avg_price,
                    COUNT(DISTINCT o.order_id) AS order_count
             FROM order_items oi
             JOIN orders o    ON o.order_id  = oi.order_id
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE o.tenant_id = ?
               AND date(o.created_at) BETWEEN ? AND ?
               AND o.status = 'completed'
               AND oi.status != 'void'
             GROUP BY mi.item_id, mi.item_name, mi.category_id
             ORDER BY qty_sold DESC
             LIMIT ?",
            [$tenantId, $startDate, $endDate, $limit]
        );
    }

    /** Hourly sales heatmap */
    public static function getHourlySales(int $locationId, int $days = 7): array {
        return Database::fetchAll(
            "SELECT DATEPART(WEEKDAY, created_at) AS day_of_week,
                    DATEPART(HOUR, created_at)    AS hour_of_day,
                    COUNT(*)                       AS order_count,
                    COALESCE(SUM(total_amount),0)   AS revenue
             FROM orders
             WHERE location_id = ?
               AND created_at >= DATEADD(DAY,-?,datetime('now'))
               AND status = 'completed'
             GROUP BY DATEPART(WEEKDAY,created_at), DATEPART(HOUR,created_at)
             ORDER BY day_of_week, hour_of_day",
            [$locationId, $days]
        );
    }

    /** Staff performance */
    public static function getStaffPerformance(int $locationId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT u.user_id, u.first_name || ' ' || u.last_name AS staff_name,
                    COUNT(o.order_id)             AS orders_handled,
                    COALESCE(SUM(o.total_amount),0) AS total_revenue,
                    COALESCE(AVG(o.total_amount),0) AS avg_order_value,
                    COALESCE(SUM(o.tip_amount),0)   AS total_tips
             FROM users u
             LEFT JOIN orders o ON o.user_id = u.user_id
                 AND o.status = 'completed'
                 AND date(o.created_at) BETWEEN ? AND ?
             WHERE u.location_id = ? AND u.is_active = 1
             GROUP BY u.user_id, u.first_name, u.last_name
             ORDER BY total_revenue DESC",
            [$startDate, $endDate, $locationId]
        );
    }

    /** Menu category performance */
    public static function getCategoryPerformance(int $tenantId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT mc.category_id, mc.category_name,
                    SUM(oi.quantity)   AS qty_sold,
                    SUM(oi.line_total) AS revenue,
                    COUNT(DISTINCT mi.item_id) AS unique_items
             FROM order_items oi
             JOIN orders o ON o.order_id = oi.order_id
             JOIN menu_items mi ON mi.item_id = oi.item_id
             JOIN menu_categories mc ON mc.category_id = mi.category_id
             WHERE o.tenant_id = ?
               AND date(o.created_at) BETWEEN ? AND ?
               AND o.status = 'completed'
             GROUP BY mc.category_id, mc.category_name
             ORDER BY revenue DESC",
            [$tenantId, $startDate, $endDate]
        );
    }

    /** Order channel breakdown */
    public static function getChannelBreakdown(int $locationId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT source,
                    COUNT(*)                     AS order_count,
                    COALESCE(SUM(total_amount),0)  AS revenue,
                    ROUND(
                        100.0 * COUNT(*) / NULLIF((SELECT COUNT(*) FROM orders
                            WHERE location_id = ? AND status = 'completed'
                            AND date(created_at) BETWEEN ? AND ?), 0), 1
                    ) AS percentage
             FROM orders
             WHERE location_id = ?
               AND status = 'completed'
               AND date(created_at) BETWEEN ? AND ?
             GROUP BY source
             ORDER BY order_count DESC",
            [$locationId, $startDate, $endDate, $locationId, $startDate, $endDate]
        );
    }

    /** Financial P&L summary */
    public static function getPLSummary(int $locationId, string $month): array {
        $start = $month . '-01';
        $end   = date('Y-m-t', strtotime($start));

        $revenue = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(total_amount),0) FROM orders
             WHERE location_id = ? AND status='completed'
               AND date(created_at) BETWEEN ? AND ?",
            [$locationId, $start, $end]
        );
        $cogs = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(oi.quantity * mi.cost_price),0)
             FROM order_items oi
             JOIN orders o ON o.order_id = oi.order_id
             JOIN menu_items mi ON mi.item_id = oi.item_id
             WHERE o.location_id = ? AND o.status='completed'
               AND date(o.created_at) BETWEEN ? AND ?",
            [$locationId, $start, $end]
        );
        $expenses = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(amount),0) FROM expenses
             WHERE location_id = ? AND expense_date BETWEEN ? AND ?",
            [$locationId, $start, $end]
        );
        $payroll = (float) Database::fetchValue(
            "SELECT COALESCE(SUM(gross_pay),0) FROM payroll
             WHERE location_id = ? AND period_start >= ? AND period_end <= ?",
            [$locationId, $start, $end]
        );

        $grossProfit    = $revenue - $cogs;
        $operatingCost  = $expenses + $payroll;
        $netProfit      = $grossProfit - $operatingCost;
        $grossMargin    = $revenue > 0 ? round(($grossProfit / $revenue) * 100, 1) : 0;

        return compact('revenue','cogs','grossProfit','grossMargin','expenses','payroll','operatingCost','netProfit');
    }

    /** Export report data as CSV string */
    public static function exportCSV(array $rows, array $headers): string {
        $lines = [implode(',', array_map(fn($h) => '"' . $h . '"', $headers))];
        foreach ($rows as $row) {
            $lines[] = implode(',', array_map(fn($v) => '"' . str_replace('"', '""', $v ?? '') . '"', array_values($row)));
        }
        return implode("\r\n", $lines);
    }
}
