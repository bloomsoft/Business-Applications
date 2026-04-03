<?php
/**
 * One-time patch script — fixes all SQL Server syntax for SQLite
 * Run: php patch.php
 */

$base = __DIR__;
$fixed = 0;
$errors = [];

function patchFile(string $path, array $replacements): bool {
    if (!file_exists($path)) return false;
    $content = file_get_contents($path);
    $original = $content;
    foreach ($replacements as $from => $to) {
        $content = str_replace($from, $to, $content);
    }
    if ($content !== $original) {
        file_put_contents($path, $content);
        return true;
    }
    return false;
}

// ── 1. Rewrite AnalyticsManager.php getRevenueChart & getHourlySales ─────────
$analyticsFile = "$base/modules/analytics/AnalyticsManager.php";
if (file_exists($analyticsFile)) {
    $content = file_get_contents($analyticsFile);

    // Fix getRevenueChart — replace old broken version
    $oldChart = <<<'OLD'
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
OLD;

    $newChart = <<<'NEW'
    /** Revenue by period (daily/weekly/monthly) */
    public static function getRevenueChart(int $locationId, string $period = 'daily', int $days = 30): array {
        $since   = date('Y-m-d', strtotime("-$days days"));
        $groupBy = match($period) {
            'weekly'  => "strftime('%Y-%W', created_at)",
            'monthly' => "strftime('%Y-%m', created_at)",
            default   => "date(created_at)",
        };
        $label = match($period) {
            'weekly'  => "strftime('%Y-W%W', created_at)",
            'monthly' => "strftime('%Y-%m', created_at)",
            default   => "date(created_at)",
        };

        return Database::fetchAll(
            "SELECT $label AS period,
                    COALESCE(SUM(total_amount),0)  AS revenue,
                    COUNT(*)                       AS order_count,
                    COALESCE(AVG(total_amount),0)  AS avg_order_value
             FROM orders
             WHERE location_id = ?
               AND date(created_at) >= ?
               AND status = 'completed'
             GROUP BY $groupBy
             ORDER BY MIN(created_at)",
            [$locationId, $since]
        );
    }
NEW;

    // Fix getHourlySales
    $oldHourly = <<<'OLD'
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
OLD;

    $newHourly = <<<'NEW'
    /** Hourly sales heatmap */
    public static function getHourlySales(int $locationId, int $days = 7): array {
        $since = date('Y-m-d', strtotime("-$days days"));
        return Database::fetchAll(
            "SELECT CAST(strftime('%w', created_at) AS INTEGER) AS day_of_week,
                    CAST(strftime('%H', created_at) AS INTEGER) AS hour_of_day,
                    COUNT(*)                       AS order_count,
                    COALESCE(SUM(total_amount),0)  AS revenue
             FROM orders
             WHERE location_id = ?
               AND date(created_at) >= ?
               AND status = 'completed'
             GROUP BY strftime('%w', created_at), strftime('%H', created_at)
             ORDER BY day_of_week, hour_of_day",
            [$locationId, $since]
        );
    }
NEW;

    $new = str_replace($oldChart, $newChart, $content);
    $new = str_replace($oldHourly, $newHourly, $new);
    if ($new !== $content) {
        file_put_contents($analyticsFile, $new);
        echo "FIXED: modules/analytics/AnalyticsManager.php\n";
        $fixed++;
    } else {
        echo "OK   : modules/analytics/AnalyticsManager.php (already correct or pattern not found)\n";
    }
}

// ── 2. Patch all files with simple string replacements ────────────────────────
$patches = [
    // Auth.php
    "$base/core/Auth.php" => [
        "GETDATE()" => "datetime('now')",
    ],
    // helpers.php
    "$base/core/helpers.php" => [
        "ORDER BY (SELECT NULL) OFFSET \$offset ROWS FETCH NEXT \$perPage ROWS ONLY" =>
            "LIMIT \$perPage OFFSET \$offset",
    ],
    // settings.php
    "$base/settings.php" => [
        "updated_at = GETDATE()" => "updated_at = datetime('now')",
    ],
    // reservations.php
    "$base/reservations.php" => [
        "c.first_name + ' ' + ISNULL(c.last_name,'')" =>
            "c.first_name || ' ' || COALESCE(c.last_name,'')",
        "c.first_name + ' ' + COALESCE(c.last_name,'')" =>
            "c.first_name || ' ' || COALESCE(c.last_name,'')",
    ],
    // CustomerManager.php
    "$base/modules/crm/CustomerManager.php" => [
        "OFFSET 0 ROWS FETCH NEXT 10 ROWS ONLY" => "LIMIT 10",
        "OFFSET 0 ROWS FETCH NEXT 5 ROWS ONLY"  => "LIMIT 5",
        "SELECT TOP 20 customer_id"              => "SELECT customer_id",
        "ORDER BY total_spent DESC\","            => "ORDER BY total_spent DESC LIMIT 20\",",
        "DATEADD(DAY,-90,datetime('now'))"       => "date('now','-90 days')",
        "DATEADD(DAY,-45,datetime('now'))"       => "date('now','-45 days')",
        "DATEADD(DAY,-?,datetime('now'))"        => "date('now','-' || ? || ' days')",
    ],
    // DeliveryManager.php
    "$base/modules/delivery/DeliveryManager.php" => [
        "DATEDIFF(MINUTE, o.created_at, datetime('now'))" =>
            "CAST((julianday(datetime('now')) - julianday(o.created_at)) * 1440 AS INTEGER)",
    ],
    // InventoryManager.php
    "$base/modules/inventory/InventoryManager.php" => [
        "CAST(datetime('now') AS DATE)" => "date('now')",
        "OFFSET ? ROWS FETCH NEXT ? ROWS ONLY" => "LIMIT ? OFFSET ?",
    ],
    // TableManager.php
    "$base/modules/pos/TableManager.php" => [
        "DATEDIFF(MINUTE, o.created_at, datetime('now'))" =>
            "CAST((julianday(datetime('now')) - julianday(o.created_at)) * 1440 AS INTEGER)",
    ],
    // staff.php
    "$base/staff.php" => [
        "u.first_name + ' ' + u.last_name AS name"      => "u.first_name || ' ' || u.last_name AS name",
        "u.first_name + ' ' + u.last_name AS staff_name"=> "u.first_name || ' ' || u.last_name AS staff_name",
        "CAST(tc.clock_in AS DATE)"                      => "date(tc.clock_in)",
        "OFFSET 0 ROWS FETCH NEXT 20 ROWS ONLY"          => "LIMIT 20",
    ],
    // layout.php
    "$base/templates/layout.php" => [
        "SELECT TOP 5 * FROM notifications" =>
            "SELECT * FROM notifications",
        "ORDER BY created_at DESC\"," =>
            "ORDER BY created_at DESC LIMIT 5\",",
    ],
];

foreach ($patches as $file => $replacements) {
    $short = str_replace($base . '/', '', $file);
    if (!file_exists($file)) {
        echo "SKIP : $short (not found)\n";
        continue;
    }
    $content  = file_get_contents($file);
    $original = $content;
    foreach ($replacements as $from => $to) {
        $content = str_replace($from, $to, $content);
    }
    if ($content !== $original) {
        file_put_contents($file, $content);
        echo "FIXED: $short\n";
        $fixed++;
    } else {
        echo "OK   : $short\n";
    }
}

// ── 3. Fix InventoryManager LIMIT/OFFSET order (params were swapped) ──────────
$invFile = "$base/modules/inventory/InventoryManager.php";
if (file_exists($invFile)) {
    $content = file_get_contents($invFile);
    // After replacing OFFSET ? ROWS FETCH NEXT ? ROWS ONLY → LIMIT ? OFFSET ?
    // params were [$invItemId, $offset, $perPage] — now should be [$invItemId, $perPage, $offset]
    $old = '[$invItemId, $offset, $perPage]';
    $new = '[$invItemId, $perPage, $offset]';
    $patched = str_replace($old, $new, $content);
    if ($patched !== $content) {
        file_put_contents($invFile, $patched);
        echo "FIXED: modules/inventory/InventoryManager.php (param order)\n";
        $fixed++;
    }
}

echo "\n==========================================\n";
echo " Patch complete! $fixed file(s) updated.\n";
echo "==========================================\n";
echo "\nNow run:\n";
echo "  php database\\setup.php\n";
echo "  php -S localhost:8000\n\n";
