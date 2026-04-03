<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$startDate  = get('start', date('Y-m-d', strtotime('-30 days')));
$endDate    = get('end',   date('Y-m-d'));
$period     = get('period', 'daily');

$revenue    = AnalyticsManager::getRevenueChart($locationId, $period, 30);
$topItems   = AnalyticsManager::getTopItems($tenantId, $startDate, $endDate, 10);
$categories = AnalyticsManager::getCategoryPerformance($tenantId, $startDate, $endDate);
$channels   = AnalyticsManager::getChannelBreakdown($locationId, $startDate, $endDate);
$staffPerf  = AnalyticsManager::getStaffPerformance($locationId, $startDate, $endDate);
$hourly     = AnalyticsManager::getHourlySales($locationId, 30);
$pl         = AnalyticsManager::getPLSummary($locationId, date('Y-m', strtotime($startDate)));

$pageTitle  = 'Analytics & Reports';
$activeMenu = 'analytics';
ob_start();
?>
<!-- Date Range Filter -->
<div class="card shadow-sm mb-4">
    <div class="card-body py-2">
        <form method="GET" class="row g-2 align-items-end">
            <div class="col-auto">
                <label class="form-label mb-1 small">Start Date</label>
                <input type="date" name="start" class="form-control form-control-sm"
                       value="<?= sanitize($startDate) ?>">
            </div>
            <div class="col-auto">
                <label class="form-label mb-1 small">End Date</label>
                <input type="date" name="end" class="form-control form-control-sm"
                       value="<?= sanitize($endDate) ?>">
            </div>
            <div class="col-auto">
                <label class="form-label mb-1 small">Period</label>
                <select name="period" class="form-select form-select-sm">
                    <option value="daily"   <?= $period === 'daily'   ? 'selected' : '' ?>>Daily</option>
                    <option value="weekly"  <?= $period === 'weekly'  ? 'selected' : '' ?>>Weekly</option>
                    <option value="monthly" <?= $period === 'monthly' ? 'selected' : '' ?>>Monthly</option>
                </select>
            </div>
            <div class="col-auto">
                <button type="submit" class="btn btn-sm btn-primary">Apply</button>
                <a href="?start=<?= date('Y-m-d', strtotime('-7 days')) ?>&end=<?= date('Y-m-d') ?>"
                   class="btn btn-sm btn-outline-secondary">Last 7D</a>
                <a href="?start=<?= date('Y-m-01') ?>&end=<?= date('Y-m-d') ?>"
                   class="btn btn-sm btn-outline-secondary">This Month</a>
            </div>
            <div class="col-auto ms-auto">
                <a href="?<?= http_build_query(array_merge($_GET, ['export' => 'csv'])) ?>"
                   class="btn btn-sm btn-outline-success">
                    <i class="bi bi-download me-1"></i>Export CSV
                </a>
            </div>
        </form>
    </div>
</div>

<!-- P&L Summary -->
<div class="row g-3 mb-4">
    <?php
    $plItems = [
        ['label'=>'Gross Revenue',  'value'=>$pl['revenue'],      'color'=>'text-success', 'icon'=>'bi-currency-dollar'],
        ['label'=>'Cost of Goods',  'value'=>$pl['cogs'],         'color'=>'text-danger',  'icon'=>'bi-box-seam'],
        ['label'=>'Gross Profit',   'value'=>$pl['grossProfit'],  'color'=>'text-primary', 'icon'=>'bi-graph-up'],
        ['label'=>'Net Profit',     'value'=>$pl['netProfit'],    'color'=>$pl['netProfit']>=0?'text-success':'text-danger', 'icon'=>'bi-wallet2'],
    ];
    foreach ($plItems as $p):
    ?>
    <div class="col-6 col-md-3">
        <div class="card shadow-sm text-center">
            <div class="card-body py-3">
                <i class="bi <?= $p['icon'] ?> fs-3 <?= $p['color'] ?>"></i>
                <div class="fw-bold fs-5 mt-1 <?= $p['color'] ?>"><?= money($p['value']) ?></div>
                <div class="text-muted small"><?= $p['label'] ?></div>
            </div>
        </div>
    </div>
    <?php endforeach; ?>
</div>

<!-- Revenue + Category Charts -->
<div class="row g-3 mb-4">
    <div class="col-xl-8">
        <div class="card chart-card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-bar-chart me-2 text-primary"></i>Revenue Trend
            </div>
            <div class="card-body">
                <canvas id="revTrendChart" height="100"></canvas>
            </div>
        </div>
    </div>
    <div class="col-xl-4">
        <div class="card chart-card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-pie-chart me-2 text-success"></i>By Category
            </div>
            <div class="card-body d-flex align-items-center justify-content-center">
                <canvas id="catChart" height="220"></canvas>
            </div>
        </div>
    </div>
</div>

<!-- Channels + Hourly Heatmap -->
<div class="row g-3 mb-4">
    <div class="col-md-5">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-layers me-2 text-info"></i>Order Channels
            </div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>Channel</th><th>Orders</th><th>Revenue</th><th>%</th></tr></thead>
                    <tbody>
                        <?php foreach ($channels as $ch): ?>
                        <tr>
                            <td><?= sanitize(ucfirst($ch['source'])) ?></td>
                            <td><?= number_format($ch['order_count']) ?></td>
                            <td><?= money($ch['revenue']) ?></td>
                            <td>
                                <div class="d-flex align-items-center gap-2">
                                    <div class="progress flex-grow-1" style="height:6px">
                                        <div class="progress-bar bg-primary" style="width:<?= min(100,$ch['percentage']) ?>%"></div>
                                    </div>
                                    <?= $ch['percentage'] ?>%
                                </div>
                            </td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <div class="col-md-7">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-clock me-2 text-warning"></i>Peak Hours Heatmap
            </div>
            <div class="card-body">
                <div id="heatmapContainer" style="overflow-x:auto"></div>
            </div>
        </div>
    </div>
</div>

<!-- Top Items Table -->
<div class="row g-3 mb-4">
    <div class="col-md-7">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-trophy me-2 text-warning"></i>Top Selling Items
            </div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>#</th><th>Item</th><th>Qty Sold</th><th>Revenue</th><th>Avg Price</th></tr></thead>
                    <tbody>
                        <?php foreach ($topItems as $i => $item): ?>
                        <tr>
                            <td><span class="badge bg-secondary"><?= $i + 1 ?></span></td>
                            <td><?= sanitize($item['item_name']) ?></td>
                            <td><?= number_format($item['qty_sold']) ?></td>
                            <td><?= money($item['total_revenue']) ?></td>
                            <td><?= money($item['avg_price']) ?></td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Staff Performance -->
    <div class="col-md-5">
        <div class="card shadow-sm">
            <div class="card-header bg-transparent fw-600 border-0 pt-3">
                <i class="bi bi-person-check me-2 text-info"></i>Staff Performance
            </div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>Staff</th><th>Orders</th><th>Revenue</th><th>Tips</th></tr></thead>
                    <tbody>
                        <?php foreach ($staffPerf as $s): ?>
                        <tr>
                            <td><?= sanitize($s['staff_name']) ?></td>
                            <td><?= number_format($s['orders_handled']) ?></td>
                            <td><?= money($s['total_revenue']) ?></td>
                            <td><?= money($s['total_tips']) ?></td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$revLabelsJson  = json_encode(array_column($revenue, 'period'));
$revDataJson    = json_encode(array_map('floatval', array_column($revenue, 'revenue')));
$catLabelsJson  = json_encode(array_column($categories, 'category_name'));
$catDataJson    = json_encode(array_map('floatval', array_column($categories, 'revenue')));
$hourlyJson     = json_encode($hourly);
$scripts = <<<JS
<script>
const revLabels = $revLabelsJson;
const revData   = $revDataJson;
const catLabels = $catLabelsJson;
const catData   = $catDataJson;

renderLineChart('revTrendChart', revLabels, [{
    label: 'Revenue',
    data: revData,
    borderColor: '#f97316',
    backgroundColor: 'rgba(249,115,22,.1)',
}]);

renderDoughnut('catChart', catLabels, catData,
    ['#3b82f6','#f97316','#22c55e','#a855f7','#f59e0b','#14b8a6','#ec4899','#64748b']);

// Heatmap
const hourlyData = $hourlyJson;
const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const hours = Array.from({length:24}, (_,i) => i === 0 ? '12am' : i < 12 ? i+'am' : i === 12 ? '12pm' : (i-12)+'pm');
const maxOrders = Math.max(...hourlyData.map(d=>d.order_count), 1);

let html = '<table class="table table-sm" style="font-size:11px"><thead><tr><th></th>';
hours.forEach(h => html += `<th style="writing-mode:vertical-lr;transform:rotate(180deg);white-space:nowrap;padding:2px">\${h}</th>`);
html += '</tr></thead><tbody>';
days.forEach((day, di) => {
    html += `<tr><td class="fw-600">\${day}</td>`;
    hours.forEach((_, hi) => {
        const cell = hourlyData.find(d => d.day_of_week == di+1 && d.hour_of_day == hi);
        const count = cell ? cell.order_count : 0;
        const intensity = Math.round((count / maxOrders) * 100);
        const bg = count === 0 ? '#f8fafc' : `hsl(24,95%,\${90-intensity*0.5}%)`;
        html += `<td style="background:\${bg};padding:4px;text-align:center" title="\${day} \${hours[hi]}: \${count} orders">
            \${count || ''}</td>`;
    });
    html += '</tr>';
});
html += '</tbody></table>';
document.getElementById('heatmapContainer').innerHTML = html;
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
