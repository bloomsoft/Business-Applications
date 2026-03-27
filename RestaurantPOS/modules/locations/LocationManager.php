<?php
/**
 * Multi-Location Management Module
 */
class LocationManager {

    public static function getAll(int $tenantId): array {
        return Database::fetchAll(
            "SELECT l.*,
                    (SELECT COUNT(*) FROM users u WHERE u.location_id = l.location_id AND u.is_active = 1) AS staff_count,
                    (SELECT COUNT(*) FROM orders o
                     WHERE o.location_id = l.location_id
                       AND date(o.created_at) = date('now')
                       AND o.status = 'completed') AS todays_orders,
                    (SELECT COALESCE(SUM(total_amount),0) FROM orders o
                     WHERE o.location_id = l.location_id
                       AND date(o.created_at) = date('now')
                       AND o.status = 'completed') AS todays_revenue
             FROM locations l
             WHERE l.tenant_id = ?
             ORDER BY l.location_name",
            [$tenantId]
        );
    }

    public static function create(array $data, int $tenantId): int {
        return Database::insert(
            "INSERT INTO locations
                (tenant_id, location_name, address, city, state, zip, country,
                 phone, email, timezone, currency, tax_rate)
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            [
                $tenantId,
                $data['location_name'],
                $data['address']   ?? null,
                $data['city']      ?? null,
                $data['state']     ?? null,
                $data['zip']       ?? null,
                $data['country']   ?? 'US',
                $data['phone']     ?? null,
                $data['email']     ?? null,
                $data['timezone']  ?? 'UTC',
                $data['currency']  ?? 'USD',
                $data['tax_rate']  ?? 0.08,
            ]
        );
    }

    public static function update(int $locationId, array $data): void {
        $old = Database::fetchOne("SELECT * FROM locations WHERE location_id = ?", [$locationId]);
        Database::query(
            "UPDATE locations
             SET location_name = ?, address = ?, city = ?, state = ?, zip = ?,
                 country = ?, phone = ?, email = ?, timezone = ?, currency = ?, tax_rate = ?
             WHERE location_id = ?",
            [
                $data['location_name'] ?? $old['location_name'],
                $data['address']       ?? $old['address'],
                $data['city']          ?? $old['city'],
                $data['state']         ?? $old['state'],
                $data['zip']           ?? $old['zip'],
                $data['country']       ?? $old['country'],
                $data['phone']         ?? $old['phone'],
                $data['email']         ?? $old['email'],
                $data['timezone']      ?? $old['timezone'],
                $data['currency']      ?? $old['currency'],
                $data['tax_rate']      ?? $old['tax_rate'],
                $locationId,
            ]
        );
        auditLog('location_updated', 'locations', $locationId, $old, $data);
    }

    /** Get comparative stats across all locations for a tenant */
    public static function getComparisonStats(int $tenantId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT l.location_id, l.location_name,
                    COUNT(o.order_id)              AS order_count,
                    ISNULL(SUM(o.total_amount),0)  AS revenue,
                    ISNULL(AVG(o.total_amount),0)  AS avg_order_value,
                    COUNT(DISTINCT o.customer_id)  AS unique_customers
             FROM locations l
             LEFT JOIN orders o ON o.location_id = l.location_id
                 AND o.status = 'completed'
                 AND o.created_at BETWEEN ? AND ?
             WHERE l.tenant_id = ? AND l.is_active = 1
             GROUP BY l.location_id, l.location_name
             ORDER BY revenue DESC",
            [$startDate, $endDate, $tenantId]
        );
    }

    /** Transfer menu items from one location to another */
    public static function copyMenuToLocation(int $fromLocationId, int $toLocationId): int {
        // Copy inventory items (not menu items — menu is tenant-wide)
        $items = Database::fetchAll(
            "SELECT * FROM inventory_items WHERE location_id = ? AND is_active = 1",
            [$fromLocationId]
        );
        $count = 0;
        foreach ($items as $item) {
            $exists = Database::fetchValue(
                "SELECT COUNT(*) FROM inventory_items WHERE location_id = ? AND item_name = ?",
                [$toLocationId, $item['item_name']]
            );
            if (!$exists) {
                Database::query(
                    "INSERT INTO inventory_items
                        (tenant_id, location_id, inv_cat_id, item_name, sku, unit,
                         reorder_level, reorder_qty, cost_per_unit, supplier_id)
                     VALUES (?,?,?,?,?,?,?,?,?,?)",
                    [
                        $item['tenant_id'], $toLocationId, $item['inv_cat_id'],
                        $item['item_name'], $item['sku'], $item['unit'],
                        $item['reorder_level'], $item['reorder_qty'],
                        $item['cost_per_unit'], $item['supplier_id'],
                    ]
                );
                $count++;
            }
        }
        return $count;
    }

    public static function get(int $locationId): ?array {
        return Database::fetchOne(
            "SELECT * FROM locations WHERE location_id = ?",
            [$locationId]
        ) ?: null;
    }

    public static function deactivate(int $locationId): void {
        Database::query(
            "UPDATE locations SET is_active = 0 WHERE location_id = ?",
            [$locationId]
        );
        auditLog('location_deactivated', 'locations', $locationId);
    }
}
