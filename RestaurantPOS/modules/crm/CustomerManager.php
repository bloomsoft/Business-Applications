<?php
/**
 * CRM — Customer Relationship Manager
 */
class CustomerManager {

    public static function create(array $data, int $tenantId): int {
        return Database::insert(
            "INSERT INTO customers
                (tenant_id, first_name, last_name, email, phone,
                 date_of_birth, gender, address, city, notes)
             VALUES (?,?,?,?,?,?,?,?,?,?)",
            [
                $tenantId,
                $data['first_name']   ?? null,
                $data['last_name']    ?? null,
                $data['email']        ?? null,
                $data['phone']        ?? null,
                $data['date_of_birth']?? null,
                $data['gender']       ?? null,
                $data['address']      ?? null,
                $data['city']         ?? null,
                $data['notes']        ?? null,
            ]
        );
    }

    public static function update(int $customerId, array $data): void {
        $old = Database::fetchOne("SELECT * FROM customers WHERE customer_id = ?", [$customerId]);
        Database::query(
            "UPDATE customers
             SET first_name = ?, last_name = ?, email = ?, phone = ?,
                 date_of_birth = ?, gender = ?, address = ?, city = ?,
                 notes = ?, updated_at = datetime('now')
             WHERE customer_id = ?",
            [
                $data['first_name']    ?? $old['first_name'],
                $data['last_name']     ?? $old['last_name'],
                $data['email']         ?? $old['email'],
                $data['phone']         ?? $old['phone'],
                $data['date_of_birth'] ?? $old['date_of_birth'],
                $data['gender']        ?? $old['gender'],
                $data['address']       ?? $old['address'],
                $data['city']          ?? $old['city'],
                $data['notes']         ?? $old['notes'],
                $customerId,
            ]
        );
        auditLog('customer_updated', 'customers', $customerId, $old, $data);
    }

    public static function find(int $tenantId, string $search): array {
        $term = '%' . $search . '%';
        return Database::fetchAll(
            "SELECT customer_id, first_name, last_name, email, phone,
                    loyalty_points, total_visits, total_spent, segment
             FROM customers
             WHERE tenant_id = ? AND is_active = 1
               AND (first_name LIKE ? OR last_name LIKE ? OR email LIKE ? OR phone LIKE ?)
             ORDER BY total_spent DESC LIMIT 20",
            [$tenantId, $term, $term, $term, $term]
        );
    }

    public static function getProfile(int $customerId): ?array {
        $customer = Database::fetchOne(
            "SELECT * FROM customers WHERE customer_id = ?",
            [$customerId]
        );
        if (!$customer) return null;

        $customer['recent_orders'] = Database::fetchAll(
            "SELECT order_id, order_number, order_type, total_amount, status, created_at
             FROM orders WHERE customer_id = ? ORDER BY created_at DESC LIMIT 10",
            [$customerId]
        );
        $customer['loyalty_history'] = Database::fetchAll(
            "SELECT * FROM loyalty_transactions WHERE customer_id = ?
             ORDER BY created_at DESC LIMIT 10",
            [$customerId]
        );
        $customer['feedbacks'] = Database::fetchAll(
            "SELECT * FROM customer_feedback WHERE customer_id = ?
             ORDER BY created_at DESC LIMIT 5",
            [$customerId]
        );
        return $customer;
    }

    /** List all customers with filtering, pagination */
    public static function list(int $tenantId, array $filters = [], int $page = 1): array {
        $where  = ['c.tenant_id = ?', 'c.is_active = 1'];
        $params = [$tenantId];

        if (!empty($filters['segment'])) {
            $where[]  = 'c.segment = ?';
            $params[] = $filters['segment'];
        }
        if (!empty($filters['search'])) {
            $t        = '%' . $filters['search'] . '%';
            $where[]  = '(c.first_name LIKE ? OR c.last_name LIKE ? OR c.email LIKE ? OR c.phone LIKE ?)';
            $params   = array_merge($params, [$t, $t, $t, $t]);
        }

        $whereStr = implode(' AND ', $where);
        $perPage  = 25;
        $offset   = ($page - 1) * $perPage;
        $total    = (int) Database::fetchValue("SELECT COUNT(*) FROM customers c WHERE $whereStr", $params);
        $data     = Database::fetchAll(
            "SELECT c.customer_id, c.first_name, c.last_name, c.email, c.phone,
                    c.loyalty_points, c.total_visits, c.total_spent,
                    c.segment, c.last_visit, c.created_at
             FROM customers c
             WHERE $whereStr
             ORDER BY c.total_spent DESC
             LIMIT $perPage OFFSET $offset",
            $params
        );

        return ['data' => $data, 'total' => $total, 'per_page' => $perPage,
                'current_page' => $page, 'last_page' => (int)ceil($total / $perPage)];
    }

    /** Redeem loyalty points */
    public static function redeemPoints(int $customerId, int $orderId, int $points): float {
        $customer = Database::fetchOne(
            "SELECT loyalty_points, tenant_id FROM customers WHERE customer_id = ?",
            [$customerId]
        );
        if (!$customer || $customer['loyalty_points'] < $points) {
            throw new RuntimeException('Insufficient loyalty points');
        }

        $program = Database::fetchOne(
            "SELECT redeem_rate, min_redeem FROM loyalty_programs WHERE tenant_id = ? AND is_active = 1",
            [$customer['tenant_id']]
        );
        if (!$program || $points < (int)$program['min_redeem']) {
            throw new RuntimeException('Below minimum redemption threshold');
        }

        $discount     = round($points * (float)$program['redeem_rate'], 2);
        $newBalance   = (int)$customer['loyalty_points'] - $points;

        Database::query(
            "UPDATE customers SET loyalty_points = ? WHERE customer_id = ?",
            [$newBalance, $customerId]
        );
        Database::query(
            "INSERT INTO loyalty_transactions (customer_id, order_id, points_redeemed, balance_after, txn_type)
             VALUES (?,?,?,?,'redeem')",
            [$customerId, $orderId, $points, $newBalance]
        );

        // Apply as discount to order
        Database::query(
            "UPDATE orders SET discount_amount = discount_amount + ?, updated_at = datetime('now')
             WHERE order_id = ?",
            [$discount, $orderId]
        );
        OrderManager::recalculate($orderId);
        return $discount;
    }

    /** Auto-segment customers based on visit/spend patterns */
    public static function resegmentAll(int $tenantId): void {
        Database::query(
            "UPDATE customers SET segment =
                CASE
                    WHEN total_visits = 0                             THEN 'new'
                    WHEN total_spent >= 1000 AND total_visits >= 20   THEN 'vip'
                    WHEN last_visit < date('now','-90 days')      THEN 'lost'
                    WHEN last_visit < date('now','-45 days')      THEN 'at-risk'
                    ELSE 'regular'
                END
             WHERE tenant_id = ?",
            [$tenantId]
        );
    }

    /** Submit customer feedback */
    public static function submitFeedback(array $data, int $tenantId): int {
        return Database::insert(
            "INSERT INTO customer_feedback
                (tenant_id, customer_id, order_id, rating, food_rating,
                 service_rating, ambiance_rating, comment, source)
             VALUES (?,?,?,?,?,?,?,?,?)",
            [
                $tenantId,
                $data['customer_id']     ?? null,
                $data['order_id']        ?? null,
                $data['rating'],
                $data['food_rating']     ?? null,
                $data['service_rating']  ?? null,
                $data['ambiance_rating'] ?? null,
                $data['comment']         ?? null,
                $data['source']          ?? 'qr',
            ]
        );
    }

    /** Get feedback summary */
    public static function getFeedbackSummary(int $tenantId, string $period = '30'): array {
        return Database::fetchOne(
            "SELECT
                COUNT(*)                        AS total_reviews,
                ROUND(AVG(CAST(rating AS FLOAT)),2) AS avg_rating,
                ROUND(AVG(CAST(food_rating AS FLOAT)),2) AS avg_food,
                ROUND(AVG(CAST(service_rating AS FLOAT)),2) AS avg_service,
                ROUND(AVG(CAST(ambiance_rating AS FLOAT)),2) AS avg_ambiance,
                SUM(CASE WHEN rating >= 4 THEN 1 ELSE 0 END) AS positive,
                SUM(CASE WHEN rating <= 2 THEN 1 ELSE 0 END) AS negative
             FROM customer_feedback
             WHERE tenant_id = ?
               AND created_at >= date('now','-' || ? || ' days')",
            [$tenantId, (int)$period]
        );
    }
}
