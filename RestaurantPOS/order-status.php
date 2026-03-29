<?php
/**
 * Order Status Page — shown after QR/Kiosk order is placed
 */
require_once __DIR__ . '/core/bootstrap.php';

$orderId  = (int) get('order_id');
$token    = get('t');
$order    = null;

if ($orderId) {
    $order = QRKioskManager::getOrderStatus($orderId, 0 /* public */);
    // If token provided, get tenant from table
    if ($token) {
        $table = TableManager::getByQRToken($token);
        if ($table) {
            $order = QRKioskManager::getOrderStatus($orderId, $table['tenant_id']);
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Order Status — RestaurantPOS</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        body { background: #f8fafc; }
        .status-circle { width:80px;height:80px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:36px;margin:0 auto; }
        @keyframes pulse { 0%,100%{transform:scale(1)} 50%{transform:scale(1.08)} }
        .pulse { animation: pulse 1.5s infinite; }
    </style>
</head>
<body>
<div class="container py-5">
    <div class="row justify-content-center">
        <div class="col-sm-10 col-md-6 col-lg-4">
            <?php if (!$order): ?>
            <div class="card shadow text-center p-4">
                <i class="bi bi-exclamation-circle text-warning fs-1"></i>
                <h4 class="mt-3">Order not found</h4>
                <a href="javascript:history.back()" class="btn btn-outline-secondary mt-3">Go Back</a>
            </div>
            <?php else: ?>
            <div class="card shadow text-center p-4">
                <?php
                $icon = match($order['status']) {
                    'completed','served','ready' => ['bi-check-circle-fill text-success', 'Your order is ready!', false],
                    'preparing'                  => ['bi-fire text-warning pulse', 'Preparing your order...', true],
                    'confirmed'                  => ['bi-clock text-info pulse', 'Order confirmed!', true],
                    'cancelled'                  => ['bi-x-circle-fill text-danger', 'Order cancelled', false],
                    default                      => ['bi-hourglass-split text-secondary pulse', 'Processing...', true],
                };
                [$iconClass, $statusMsg, $refresh] = $icon;
                ?>
                <div class="status-circle bg-light mx-auto mb-3">
                    <i class="bi <?= $iconClass ?>"></i>
                </div>
                <h4><?= $statusMsg ?></h4>
                <p class="text-muted">Order #<?= sanitize($order['order_number']) ?></p>

                <div class="list-group text-start mb-3">
                    <?php foreach ($order['items'] as $item): ?>
                    <div class="list-group-item d-flex justify-content-between">
                        <span><?= (int)$item['quantity'] ?>× <?= sanitize($item['item_name']) ?></span>
                        <?= statusBadge($item['status']) ?>
                    </div>
                    <?php endforeach; ?>
                </div>

                <div class="d-flex justify-content-between fw-bold border-top pt-2">
                    <span>Total</span>
                    <span><?= money($order['total_amount']) ?></span>
                </div>

                <?php if ($token): ?>
                <div class="mt-4 d-flex gap-2">
                    <a href="/order.php?t=<?= urlencode($token) ?>" class="btn btn-outline-primary flex-grow-1">
                        <i class="bi bi-plus me-1"></i>Add More
                    </a>
                    <a href="/feedback.php?order_id=<?= $orderId ?>" class="btn btn-outline-warning flex-grow-1">
                        <i class="bi bi-star me-1"></i>Rate Us
                    </a>
                </div>
                <?php endif; ?>
            </div>
            <?php if ($refresh): ?>
            <script>setTimeout(()=>location.reload(), 20000);</script>
            <?php endif; ?>
            <?php endif; ?>
        </div>
    </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
