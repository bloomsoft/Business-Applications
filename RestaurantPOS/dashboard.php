<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$today      = date('Y-m-d');
$kpis       = AnalyticsManager::getDashboardKPIs($locationId, $today);
$revenue7d  = AnalyticsManager::getRevenueChart($locationId, 'daily', 7);
$topItems   = AnalyticsManager::getTopItems($tenantId, date('Y-m-d', strtotime('-30 days')), $today, 5);
$channels   = AnalyticsManager::getChannelBreakdown($locationId, date('Y-m-d', strtotime('-7 days')), $today);
$lowStock   = InventoryManager::getLowStockItems($tenantId);
$liveOrders = OrderManager::listOrders($locationId, ['status' => 'preparing'], 1)['data'];
$pageTitle  = 'Dashboard';
$activeMenu = 'dashboard';

ob_start();
?>
<!-- KPI Row -->
<div class="row g-3 mb-4">
    <?php
    $kpis_display = [
        ['label'=>"Today's Revenue",  'value'=>money($kpis['total_revenue']),   'icon'=>'bi-currency-dollar', 'color'=>'bg-success bg-opacity-10 text-success', 'change'=>$kpis['revenue_change']],
        ['label'=>"Total Orders",     'value'=>number_format($kpis['total_orders']),  'icon'=>'bi-receipt',        'color'=>'bg-primary bg-opacity-10 text-primary', 'change'=>$kpis['orders_change']],
        ['label'=>"Avg Order Value",  'value'=>money($kpis['avg_order_value']),  'icon'=>'bi-graph-up',       'color'=>'bg-warning bg-opacity-10 text-warning', 'change'=>null],
        ['label'=>"Unique Customers", 'value'=>number_format($kpis['unique_customers']), 'icon'=>'bi-people', 'color'=>'bg-info bg-opacity-10 text-info',    'change'=>null],
    ];
    foreach ($kpis_display as $k):
    ?>
    <div class="col-6 col-xl-3">
        <div class="card kpi-card shadow-sm h-100">
            <div class="card-body d-flex align-items-center gap-3">
                <div class="kpi-icon <?= $k['color'] ?>">
                    <i class="bi <?= $k['icon'] ?>"></i>
                </div>
                <div>
                    <div class="text-muted small"><?= $k['label'] ?></div>
                    <div class="fs-4 fw-bold"><?= $k['value'] ?></div>
                    <?php if ($k['change'] !== null): ?>
                    <small class="<?= $k['change'] >= 0 ? 'text-success' : 'text-danger' ?>">
                        <i class="bi bi-arrow-<?= $k['change'] >= 0 ? 'up' : 'down' ?>-short"></i>
                        <?= abs($k['change']) ?>% vs yesterday
                    </small>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>
    <?php endforeach; ?>
</div>

<!-- Order Type Breakdown -->
<div class="row g-3 mb-4">
    <div class="col-md-3">
        <div class="card text-center border-0 shadow-sm">
            <div class="card-body py-3">
                <i class="bi bi-table fs-3 text-primary"></i>
                <div class="fw-bold fs-5 mt-1"><?= number_format($kpis['dine_in']) ?></div>
                <div class="text-muted small">Dine-In</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center border-0 shadow-sm">
            <div class="card-body py-3">
                <i class="bi bi-bag fs-3 text-warning"></i>
                <div class="fw-bold fs-5 mt-1"><?= number_format($kpis['takeout']) ?></div>
                <div class="text-muted small">Takeout</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center border-0 shadow-sm">
            <div class="card-body py-3">
                <i class="bi bi-truck fs-3 text-success"></i>
                <div class="fw-bold fs-5 mt-1"><?= number_format($kpis['delivery']) ?></div>
                <div class="text-muted small">Delivery</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center border-0 shadow-sm">
            <div class="card-body py-3">
                <i class="bi bi-qr-code fs-3 text-info"></i>
                <div class="fw-bold fs-5 mt-1"><?= number_format($kpis['self_service']) ?></div>
                <div class="text-muted small">QR / Kiosk</div>
            </div>
        </div>
    </div>
</div>

<!-- Charts Row -->
<div class="row g-3 mb-4">
    <!-- Revenue Chart -->
    <div class="col-xl-8">
        <div class="card chart-card shadow-sm h-100">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-bar-chart-line me-2 text-primary"></i>Revenue (Last 7 Days)
            </div>
            <div class="card-body">
                <canvas id="revenueChart" height="100"></canvas>
            </div>
        </div>
    </div>
    <!-- Channel Doughnut -->
    <div class="col-xl-4">
        <div class="card chart-card shadow-sm h-100">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-pie-chart me-2 text-success"></i>Order Channels
            </div>
            <div class="card-body d-flex align-items-center justify-content-center">
                <canvas id="channelChart" height="200"></canvas>
            </div>
        </div>
    </div>
</div>

<!-- Live Orders + Top Items + Low Stock -->
<div class="row g-3">
    <!-- Live Orders -->
    <div class="col-xl-4">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3 d-flex justify-content-between">
                <span><i class="bi bi-fire me-2 text-danger"></i>Live Orders</span>
                <a href="/orders.php" class="btn btn-sm btn-outline-secondary">View All</a>
            </div>
            <div class="card-body p-0">
                <?php if (empty($liveOrders)): ?>
                <div class="text-center text-muted py-4">No active orders</div>
                <?php else: ?>
                <div class="list-group list-group-flush">
                    <?php foreach ($liveOrders as $order): ?>
                    <a href="/orders.php?id=<?= $order['order_id'] ?>"
                       class="list-group-item list-group-item-action d-flex justify-content-between align-items-center">
                        <div>
                            <div class="fw-600">#<?= sanitize($order['order_number']) ?></div>
                            <small class="text-muted"><?= sanitize($order['order_type']) ?>
                                <?= $order['table_number'] ? '· Table ' . sanitize($order['table_number']) : '' ?>
                            </small>
                        </div>
                        <div class="text-end">
                            <?= statusBadge($order['status']) ?>
                            <div class="fw-bold mt-1"><?= money($order['total_amount']) ?></div>
                        </div>
                    </a>
                    <?php endforeach; ?>
                </div>
                <?php endif; ?>
            </div>
        </div>
    </div>

    <!-- Top Items -->
    <div class="col-xl-4">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-trophy me-2 text-warning"></i>Top Items (30 Days)
            </div>
            <div class="card-body p-0">
                <div class="list-group list-group-flush">
                    <?php foreach ($topItems as $i => $item): ?>
                    <div class="list-group-item d-flex justify-content-between align-items-center">
                        <div class="d-flex align-items-center gap-2">
                            <span class="badge bg-secondary"><?= $i + 1 ?></span>
                            <span><?= sanitize($item['item_name']) ?></span>
                        </div>
                        <div class="text-end">
                            <div class="fw-bold"><?= money($item['total_revenue']) ?></div>
                            <small class="text-muted"><?= number_format($item['qty_sold']) ?> sold</small>
                        </div>
                    </div>
                    <?php endforeach; ?>
                    <?php if (empty($topItems)): ?>
                    <div class="list-group-item text-muted text-center">No data yet</div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>

    <!-- Low Stock Alerts -->
    <div class="col-xl-4">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3 d-flex justify-content-between">
                <span>
                    <i class="bi bi-exclamation-triangle me-2 text-warning"></i>Low Stock Alerts
                    <?php if (count($lowStock)): ?>
                    <span class="badge bg-danger"><?= count($lowStock) ?></span>
                    <?php endif; ?>
                </span>
                <a href="/inventory.php?filter=low_stock" class="btn btn-sm btn-outline-secondary">Manage</a>
            </div>
            <div class="card-body p-0">
                <?php if (empty($lowStock)): ?>
                <div class="text-center text-muted py-4">
                    <i class="bi bi-check-circle text-success fs-3"></i>
                    <p class="mt-2 mb-0">Stock levels OK</p>
                </div>
                <?php else: ?>
                <div class="list-group list-group-flush">
                    <?php foreach (array_slice($lowStock, 0, 6) as $item): ?>
                    <div class="list-group-item d-flex justify-content-between align-items-center">
                        <div>
                            <div class="fw-600"><?= sanitize($item['item_name']) ?></div>
                            <small class="text-muted"><?= sanitize($item['location_name']) ?></small>
                        </div>
                        <div class="text-end">
                            <span class="badge <?= $item['quantity_on_hand'] <= 0 ? 'bg-danger' : 'bg-warning text-dark' ?>">
                                <?= number_format($item['quantity_on_hand'], 1) ?> <?= sanitize($item['unit']) ?>
                            </span>
                        </div>
                    </div>
                    <?php endforeach; ?>
                </div>
                <?php endif; ?>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();

// Prepare chart data
$chartLabels   = array_column($revenue7d, 'period');
$chartRevenue  = array_column($revenue7d, 'revenue');
$channelLabels = array_column($channels, 'source');
$channelData   = array_column($channels, 'order_count');

$scripts = <<<JS
<script>
const rev7Labels = <?= json_encode($chartLabels) ?>;
const rev7Data   = <?= json_encode(array_map('floatval', $chartRevenue)) ?>;

renderLineChart('revenueChart', rev7Labels, [{
    label: 'Revenue',
    data: rev7Data,
    borderColor: '#f97316',
    backgroundColor: 'rgba(249,115,22,.1)',
}]);

renderDoughnut('channelChart',
    <?= json_encode($channelLabels) ?>,
    <?= json_encode(array_map('intval', $channelData)) ?>,
    ['#3b82f6','#f97316','#22c55e','#a855f7','#f59e0b']
);
</script>
JS;

require_once __DIR__ . '/templates/layout.php';
