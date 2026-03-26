<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('financial.view');

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$activeTab  = get('tab', 'eod');
$date       = get('date', date('Y-m-d'));
$startDate  = get('start', date('Y-m-01'));
$endDate    = get('end',   date('Y-m-d'));

$eod      = FinancialManager::getEODReport($locationId, $date);
$tax      = FinancialManager::getTaxReport($locationId, $startDate, $endDate);
$expenses = FinancialManager::getExpenses($locationId, $startDate, $endDate);
$expSum   = FinancialManager::getExpenseSummary($locationId, $startDate, $endDate);
$pl       = AnalyticsManager::getPLSummary($locationId, substr($startDate, 0, 7));
$payments = PaymentManager::getDailySummary($locationId, $date);

$pageTitle = 'Financial Management';
$activeMenu = 'financial';

if ($_SERVER['REQUEST_METHOD'] === 'POST' && post('action') === 'add_expense') {
    FinancialManager::addExpense(array_merge($_POST, ['tenant_id' => $tenantId, 'location_id' => $locationId]));
    flash('success', 'Expense recorded');
    redirect('/financial.php?tab=expenses');
}

ob_start();
?>
<ul class="nav nav-tabs mb-4">
    <li class="nav-item"><a class="nav-link <?= $activeTab==='eod'     ?'active':'' ?>" href="?tab=eod">End of Day</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab==='expenses'?'active':'' ?>" href="?tab=expenses">Expenses</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab==='tax'     ?'active':'' ?>" href="?tab=tax">Tax Report</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab==='pl'      ?'active':'' ?>" href="?tab=pl">P&amp;L</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab==='cash'    ?'active':'' ?>" href="?tab=cash">Cash Drawer</a></li>
</ul>

<?php if ($activeTab === 'eod'): ?>
<!-- EOD Report -->
<div class="d-flex justify-content-between align-items-center mb-3">
    <div class="d-flex gap-2 align-items-center">
        <label class="mb-0">Date:</label>
        <input type="date" class="form-control form-control-sm" value="<?= $date ?>"
               onchange="location.href='?tab=eod&date='+this.value">
    </div>
    <button class="btn btn-sm btn-outline-secondary" onclick="window.print()">
        <i class="bi bi-printer me-1"></i>Print Report
    </button>
</div>

<div class="row g-3 mb-4">
    <?php
    $eodCards = [
        ['label'=>'Completed Orders', 'value'=>number_format($eod['sales']['completed_orders']),   'cls'=>'text-success'],
        ['label'=>'Gross Revenue',    'value'=>money($eod['sales']['gross_revenue']),               'cls'=>'text-primary'],
        ['label'=>'Tax Collected',    'value'=>money($eod['sales']['total_tax']),                   'cls'=>'text-warning'],
        ['label'=>'Total Tips',       'value'=>money($eod['sales']['total_tips']),                  'cls'=>'text-info'],
        ['label'=>'Discounts Given',  'value'=>money($eod['sales']['total_discounts']),             'cls'=>'text-danger'],
        ['label'=>'Cancelled Orders', 'value'=>number_format($eod['sales']['cancelled_orders']),   'cls'=>'text-danger'],
    ];
    foreach ($eodCards as $card):
    ?>
    <div class="col-6 col-md-4 col-lg-2">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold <?= $card['cls'] ?>"><?= $card['value'] ?></div>
                <div class="text-muted small"><?= $card['label'] ?></div>
            </div>
        </div>
    </div>
    <?php endforeach; ?>
</div>

<div class="row g-3">
    <!-- Payment Breakdown -->
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-header fw-600">Payment Breakdown</div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>Method</th><th>Txns</th><th>Amount</th><th>Tips</th></tr></thead>
                    <tbody>
                        <?php foreach ($eod['payment_breakdown'] as $p): ?>
                        <tr>
                            <td><?= ucfirst(sanitize($p['payment_method'])) ?></td>
                            <td><?= $p['txn_count'] ?></td>
                            <td><?= money($p['total']) ?></td>
                            <td><?= money($p['total_tips']) ?></td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <!-- Top Items -->
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-header fw-600">Top 5 Items Today</div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>Item</th><th>Qty</th><th>Revenue</th></tr></thead>
                    <tbody>
                        <?php foreach ($eod['top_items'] as $item): ?>
                        <tr>
                            <td><?= sanitize($item['item_name']) ?></td>
                            <td><?= $item['qty'] ?></td>
                            <td><?= money($item['revenue']) ?></td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <!-- Voids & Refunds -->
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-header fw-600">Voids & Refunds</div>
            <div class="card-body p-0">
                <?php if (empty($eod['voids']) && empty($eod['refunds'])): ?>
                <div class="text-center text-muted py-3">None today</div>
                <?php else: ?>
                <table class="table table-sm mb-0">
                    <tbody>
                        <?php foreach ($eod['voids'] as $v): ?>
                        <tr>
                            <td><?= sanitize($v['item_name']) ?></td>
                            <td class="text-danger">Void</td>
                        </tr>
                        <?php endforeach; ?>
                        <?php foreach ($eod['refunds'] as $r): ?>
                        <tr>
                            <td><?= money($r['amount']) ?></td>
                            <td class="text-warning">Refund</td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
                <?php endif; ?>
            </div>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'expenses'): ?>
<div class="d-flex justify-content-between mb-3">
    <form method="GET" class="d-flex gap-2 align-items-end">
        <input type="hidden" name="tab" value="expenses">
        <div><label class="form-label mb-1 small">From</label>
            <input type="date" name="start" class="form-control form-control-sm" value="<?= $startDate ?>"></div>
        <div><label class="form-label mb-1 small">To</label>
            <input type="date" name="end" class="form-control form-control-sm" value="<?= $endDate ?>"></div>
        <button type="submit" class="btn btn-sm btn-primary align-self-end">Filter</button>
    </form>
    <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addExpenseModal">
        <i class="bi bi-plus me-1"></i>Add Expense
    </button>
</div>
<div class="row g-3">
    <div class="col-md-8">
        <div class="card shadow-sm">
            <div class="card-body p-0">
                <table class="table table-sm align-middle mb-0">
                    <thead class="table-light"><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>By</th></tr></thead>
                    <tbody>
                        <?php foreach ($expenses as $e): ?>
                        <tr>
                            <td><?= fmtDate($e['expense_date']) ?></td>
                            <td><span class="badge bg-secondary"><?= sanitize($e['category']) ?></span></td>
                            <td><?= sanitize($e['description'] ?? '—') ?></td>
                            <td class="fw-600 text-danger"><?= money($e['amount']) ?></td>
                            <td><small><?= sanitize($e['created_by_name'] ?? '') ?></small></td>
                        </tr>
                        <?php endforeach; ?>
                        <?php if (empty($expenses)): ?>
                        <tr><td colspan="5" class="text-center text-muted py-4">No expenses recorded</td></tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-header fw-600">By Category</div>
            <div class="card-body p-0">
                <table class="table table-sm mb-0">
                    <?php foreach ($expSum as $cat): ?>
                    <tr>
                        <td><?= sanitize($cat['category']) ?></td>
                        <td class="text-end fw-600"><?= money($cat['total_amount']) ?></td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (!empty($expSum)): ?>
                    <tr class="table-light">
                        <td class="fw-bold">Total</td>
                        <td class="text-end fw-bold"><?= money(array_sum(array_column($expSum,'total_amount'))) ?></td>
                    </tr>
                    <?php endif; ?>
                </table>
            </div>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'tax'): ?>
<div class="card shadow-sm">
    <div class="card-header fw-600">Tax Report — <?= fmtDate($startDate) ?> to <?= fmtDate($endDate) ?></div>
    <div class="card-body">
        <div class="row g-3">
            <div class="col-md-3"><div class="card border text-center py-3">
                <div class="fw-bold fs-4"><?= money($tax['net_sales']) ?></div>
                <div class="text-muted">Net Sales</div>
            </div></div>
            <div class="col-md-3"><div class="card border text-center py-3">
                <div class="fw-bold fs-4 text-warning"><?= money($tax['tax_collected']) ?></div>
                <div class="text-muted">Tax Collected</div>
            </div></div>
            <div class="col-md-3"><div class="card border text-center py-3">
                <div class="fw-bold fs-4"><?= money($tax['gross_sales']) ?></div>
                <div class="text-muted">Gross Sales</div>
            </div></div>
            <div class="col-md-3"><div class="card border text-center py-3">
                <div class="fw-bold fs-4"><?= number_format($tax['transaction_count']) ?></div>
                <div class="text-muted">Transactions</div>
            </div></div>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'pl'): ?>
<div class="card shadow-sm">
    <div class="card-header fw-600">Profit & Loss — <?= date('F Y', strtotime($startDate)) ?></div>
    <div class="card-body">
        <table class="table table-sm" style="max-width:500px">
            <tr><td>Gross Revenue</td><td class="text-end fw-bold text-success"><?= money($pl['revenue']) ?></td></tr>
            <tr><td>Cost of Goods (COGS)</td><td class="text-end text-danger">-<?= money($pl['cogs']) ?></td></tr>
            <tr class="table-light"><td class="fw-bold">Gross Profit</td><td class="text-end fw-bold"><?= money($pl['grossProfit']) ?></td></tr>
            <tr><td class="text-muted ps-3">Margin</td><td class="text-end text-muted"><?= $pl['grossMargin'] ?>%</td></tr>
            <tr><td>Operating Expenses</td><td class="text-end text-danger">-<?= money($pl['expenses']) ?></td></tr>
            <tr><td>Payroll</td><td class="text-end text-danger">-<?= money($pl['payroll']) ?></td></tr>
            <tr class="table-<?= $pl['netProfit'] >= 0 ? 'success' : 'danger' ?>">
                <td class="fw-bold fs-5">Net Profit</td>
                <td class="text-end fw-bold fs-5"><?= money($pl['netProfit']) ?></td>
            </tr>
        </table>
    </div>
</div>
<?php endif; ?>

<!-- Add Expense Modal -->
<div class="modal fade" id="addExpenseModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Record Expense</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="add_expense">
                    <div class="col-6"><label class="form-label">Category</label>
                        <select name="category" class="form-select" required>
                            <?php foreach (['Supplies','Utilities','Marketing','Maintenance','Food','Packaging','Other'] as $cat): ?>
                            <option><?= $cat ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-6"><label class="form-label">Date</label>
                        <input type="date" name="expense_date" class="form-control" value="<?= date('Y-m-d') ?>" required></div>
                    <div class="col-12"><label class="form-label">Description</label>
                        <input type="text" name="description" class="form-control"></div>
                    <div class="col-6"><label class="form-label">Amount</label>
                        <input type="number" name="amount" class="form-control" step="0.01" min="0.01" required></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Save Expense</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
require_once __DIR__ . '/templates/layout.php';
