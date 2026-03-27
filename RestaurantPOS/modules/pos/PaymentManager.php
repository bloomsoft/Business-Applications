<?php
/**
 * Payment Processing Manager
 */
class PaymentManager {

    /** Process a payment for an order */
    public static function process(int $orderId, array $paymentData): array {
        $order = Database::fetchOne(
            "SELECT * FROM orders WHERE order_id = ?",
            [$orderId]
        );
        if (!$order) return ['success' => false, 'message' => 'Order not found'];
        if ($order['status'] === 'completed') return ['success' => false, 'message' => 'Order already paid'];

        $method  = $paymentData['method'];
        $amount  = (float) $paymentData['amount'];
        $tip     = (float) ($paymentData['tip']  ?? 0);

        // Gateway processing (stubbed — swap for real SDK calls)
        $gatewayResult = self::processGateway($method, $amount, $paymentData);
        if (!$gatewayResult['success']) {
            return $gatewayResult;
        }

        Database::beginTransaction();
        try {
            $paymentId = Database::insert(
                "INSERT INTO payments
                    (order_id, payment_method, amount, tip_amount, status,
                     reference_no, gateway, gateway_txn_id, processed_by, processed_at)
                 VALUES (?,?,?,?,?,?,?,?,?,datetime('now'))",
                [
                    $orderId, $method, $amount, $tip,
                    'completed',
                    $gatewayResult['reference'] ?? null,
                    $gatewayResult['gateway']   ?? 'manual',
                    $gatewayResult['txn_id']    ?? null,
                    Auth::id(),
                ]
            );

            // Update order tip and status
            Database::query(
                "UPDATE orders SET tip_amount = tip_amount + ?,
                                   total_amount = total_amount + ?,
                                   status = 'completed',
                                   completed_at = datetime('now'),
                                   updated_at   = datetime('now')
                 WHERE order_id = ?",
                [$tip, $tip, $orderId]
            );

            // Update loyalty points if customer linked
            if ($order['customer_id']) {
                self::awardLoyaltyPoints($order['customer_id'], $orderId, $amount);
            }

            Database::commit();
            auditLog('payment_processed', 'payments', $paymentId);
            return ['success' => true, 'payment_id' => $paymentId, 'change' => max(0, $amount - (float)$order['total_amount'])];
        } catch (Throwable $e) {
            Database::rollback();
            throw $e;
        }
    }

    /** Process split payment */
    public static function splitPayment(int $orderId, array $splits): array {
        $results = [];
        foreach ($splits as $split) {
            $results[] = self::process($orderId, $split);
        }
        return $results;
    }

    /** Void/refund a payment */
    public static function refund(int $paymentId, float $amount, string $reason): array {
        $payment = Database::fetchOne(
            "SELECT * FROM payments WHERE payment_id = ?",
            [$paymentId]
        );
        if (!$payment) return ['success' => false, 'message' => 'Payment not found'];

        $refundId = Database::insert(
            "INSERT INTO refunds (payment_id, order_id, amount, reason, status, processed_by)
             VALUES (?,?,?,?,'completed',?)",
            [$paymentId, $payment['order_id'], $amount, $reason, Auth::id()]
        );
        Database::query(
            "UPDATE payments SET status = 'refunded' WHERE payment_id = ?",
            [$paymentId]
        );
        auditLog('payment_refunded', 'refunds', $refundId);
        return ['success' => true, 'refund_id' => $refundId];
    }

    /** Get daily cash summary for a location */
    public static function getDailySummary(int $locationId, string $date): array {
        return Database::fetchAll(
            "SELECT p.payment_method,
                    COUNT(*) AS txn_count,
                    SUM(p.amount) AS total_amount,
                    SUM(p.tip_amount) AS total_tips
             FROM payments p
             JOIN orders o ON o.order_id = p.order_id
             WHERE o.location_id = ?
               AND CAST(p.processed_at AS DATE) = ?
               AND p.status = 'completed'
             GROUP BY p.payment_method",
            [$locationId, $date]
        );
    }

    /** Award loyalty points */
    private static function awardLoyaltyPoints(int $customerId, int $orderId, float $amount): void {
        $program = Database::fetchOne(
            "SELECT lp.* FROM loyalty_programs lp
             JOIN customers c ON c.tenant_id = (
                 SELECT tenant_id FROM customers WHERE customer_id = ?
             )
             WHERE lp.is_active = 1",
            [$customerId]
        );
        if (!$program) return;

        $points = (int) floor($amount * (float)$program['points_per_dollar']);
        if ($points <= 0) return;

        $current = (int) Database::fetchValue(
            "SELECT loyalty_points FROM customers WHERE customer_id = ?",
            [$customerId]
        );
        $newBalance = $current + $points;

        Database::query(
            "UPDATE customers SET loyalty_points = ?, total_visits = total_visits + 1,
                    total_spent = total_spent + ?, last_visit = datetime('now')
             WHERE customer_id = ?",
            [$newBalance, $amount, $customerId]
        );
        Database::query(
            "INSERT INTO loyalty_transactions (customer_id, order_id, points_earned, balance_after, txn_type)
             VALUES (?,?,?,?,'earn')",
            [$customerId, $orderId, $points, $newBalance]
        );
    }

    /** Gateway dispatcher — integrate real SDKs here */
    private static function processGateway(string $method, float $amount, array $data): array {
        return match ($method) {
            'cash'   => ['success' => true, 'gateway' => 'cash',   'reference' => 'CASH-' . time()],
            'card'   => self::processStripe($amount, $data),
            'paypal' => self::processPaypal($amount, $data),
            default  => ['success' => true, 'gateway' => 'manual', 'reference' => 'MAN-' . time()],
        };
    }

    private static function processStripe(float $amount, array $data): array {
        // Placeholder: integrate with Stripe PHP SDK
        // \Stripe\Stripe::setApiKey(STRIPE_SECRET_KEY);
        // $charge = \Stripe\PaymentIntent::create([...]);
        return ['success' => true, 'gateway' => 'stripe', 'txn_id' => 'pi_stub_' . uniqid(), 'reference' => uniqid('STR-')];
    }

    private static function processPaypal(float $amount, array $data): array {
        return ['success' => true, 'gateway' => 'paypal', 'txn_id' => 'pp_stub_' . uniqid(), 'reference' => uniqid('PP-')];
    }
}
