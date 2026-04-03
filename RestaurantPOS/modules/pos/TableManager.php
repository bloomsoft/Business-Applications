<?php
/**
 * Table / Floor Plan Manager
 */
class TableManager {

    public static function getTables(int $locationId): array {
        return Database::fetchAll(
            "SELECT t.*, a.area_name,
                    o.order_id, o.order_number, o.total_amount,
                    CAST((julianday(datetime('now')) - julianday(o.created_at)) * 1440 AS INTEGER) AS occupied_minutes
             FROM restaurant_tables t
             LEFT JOIN dining_areas a ON a.area_id = t.area_id
             LEFT JOIN orders o ON o.table_id = t.table_id
                AND o.status NOT IN ('completed','cancelled')
             WHERE t.location_id = ?
             ORDER BY t.area_id, t.table_number",
            [$locationId]
        );
    }

    public static function getAreas(int $locationId): array {
        return Database::fetchAll(
            "SELECT * FROM dining_areas WHERE location_id = ? ORDER BY sort_order",
            [$locationId]
        );
    }

    public static function updateStatus(int $tableId, string $status): void {
        Database::query(
            "UPDATE restaurant_tables SET status = ? WHERE table_id = ?",
            [$status, $tableId]
        );
    }

    public static function createTable(array $data): int {
        $token = bin2hex(random_bytes(16));
        return Database::insert(
            "INSERT INTO restaurant_tables
                (location_id, area_id, table_number, capacity, qr_code_token, pos_x, pos_y, shape)
             VALUES (?,?,?,?,?,?,?,?)",
            [
                $data['location_id'],
                $data['area_id']      ?? null,
                $data['table_number'],
                $data['capacity']     ?? 4,
                $token,
                $data['pos_x']        ?? 0,
                $data['pos_y']        ?? 0,
                $data['shape']        ?? 'rectangle',
            ]
        );
    }

    public static function getQRToken(int $tableId): ?string {
        return Database::fetchValue(
            "SELECT qr_code_token FROM restaurant_tables WHERE table_id = ?",
            [$tableId]
        ) ?: null;
    }

    public static function getByQRToken(string $token): ?array {
        return Database::fetchOne(
            "SELECT t.*, l.tenant_id, l.tax_rate, l.currency
             FROM restaurant_tables t
             JOIN locations l ON l.location_id = t.location_id
             WHERE t.qr_code_token = ?",
            [$token]
        ) ?: null;
    }

    /** Move table on floor plan */
    public static function updatePosition(int $tableId, int $x, int $y): void {
        Database::query(
            "UPDATE restaurant_tables SET pos_x = ?, pos_y = ? WHERE table_id = ?",
            [$x, $y, $tableId]
        );
    }
}
