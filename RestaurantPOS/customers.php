<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$tenantId  = Auth::tenantId();
$page      = (int) get('page', 1);
$filters   = ['search' => get('search'), 'segment' => get('segment')];
$customers = CustomerManager::list($tenantId, $filters, $page);
$feedback  = CustomerManager::getFeedbackSummary($tenantId, 30);
$pageTitle = 'Customers & CRM';
$activeMenu= 'crm';

// Handle create customer
if ($_SERVER['REQUEST_METHOD'] === 'POST' && post('action') === 'create_customer') {
    CustomerManager::create($_POST, $tenantId);
    flash('success', 'Customer added successfully');
    redirect('/customers.php');
}

ob_start();
?>
<!-- Summary Cards -->
<div class="row g-2 mb-4">
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <i class="bi bi-star-fill text-warning fs-4"></i>
                <div class="fw-bold"><?= number_format($feedback['avg_rating'] ?? 0, 1) ?>/5</div>
                <div class="text-muted small">Avg Rating</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <i class="bi bi-chat-square-text text-info fs-4"></i>
                <div class="fw-bold"><?= number_format($feedback['total_reviews'] ?? 0) ?></div>
                <div class="text-muted small">Reviews (30d)</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <i class="bi bi-hand-thumbs-up text-success fs-4"></i>
                <div class="fw-bold"><?= number_format($feedback['positive'] ?? 0) ?></div>
                <div class="text-muted small">Positive</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <i class="bi bi-people text-primary fs-4"></i>
                <div class="fw-bold"><?= number_format($customers['total']) ?></div>
                <div class="text-muted small">Total Customers</div>
            </div>
        </div>
    </div>
</div>

<!-- Filters + Actions -->
<div class="d-flex flex-wrap gap-2 justify-content-between mb-3">
    <form method="GET" class="d-flex gap-2 flex-wrap">
        <input type="text" name="search" class="form-control form-control-sm" style="width:220px"
               placeholder="Name, phone, email..." value="<?= sanitize(get('search')) ?>">
        <select name="segment" class="form-select form-select-sm" style="width:140px">
            <option value="">All Segments</option>
            <?php foreach (['new','regular','vip','at-risk','lost'] as $seg): ?>
            <option value="<?= $seg ?>" <?= get('segment') === $seg ? 'selected' : '' ?>><?= ucfirst($seg) ?></option>
            <?php endforeach; ?>
        </select>
        <button type="submit" class="btn btn-sm btn-primary">Filter</button>
        <a href="/customers.php" class="btn btn-sm btn-outline-secondary">Reset</a>
    </form>
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addCustomerModal">
            <i class="bi bi-person-plus me-1"></i>Add Customer
        </button>
        <button class="btn btn-sm btn-outline-secondary" onclick="resegment()">
            <i class="bi bi-funnel me-1"></i>Auto-Segment
        </button>
    </div>
</div>

<!-- Customers Table -->
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover table-sm align-middle mb-0">
                <thead class="table-light">
                    <tr>
                        <th>Customer</th><th>Phone/Email</th><th>Visits</th>
                        <th>Total Spent</th><th>Loyalty Pts</th><th>Segment</th>
                        <th>Last Visit</th><th></th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($customers['data'] as $c): ?>
                    <tr>
                        <td>
                            <div class="fw-600"><?= sanitize($c['first_name'] . ' ' . $c['last_name']) ?></div>
                        </td>
                        <td><small class="text-muted"><?= sanitize($c['phone'] ?: $c['email'] ?: '—') ?></small></td>
                        <td><?= number_format($c['total_visits']) ?></td>
                        <td class="fw-600"><?= money($c['total_spent']) ?></td>
                        <td>
                            <span class="badge bg-warning text-dark">
                                <i class="bi bi-star-fill me-1"></i><?= number_format($c['loyalty_points']) ?>
                            </span>
                        </td>
                        <td>
                            <?php
                            $segColor = match($c['segment']) {
                                'vip'     => 'bg-purple text-white',
                                'new'     => 'bg-info',
                                'at-risk' => 'bg-warning text-dark',
                                'lost'    => 'bg-danger',
                                default   => 'bg-secondary',
                            };
                            ?>
                            <span class="badge <?= $segColor ?>"><?= ucfirst($c['segment']) ?></span>
                        </td>
                        <td><small><?= $c['last_visit'] ? fmtDate($c['last_visit']) : 'Never' ?></small></td>
                        <td>
                            <a href="/customer-profile.php?id=<?= $c['customer_id'] ?>"
                               class="btn btn-sm btn-outline-primary">
                                <i class="bi bi-person-circle"></i>
                            </a>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($customers['data'])): ?>
                    <tr><td colspan="8" class="text-center text-muted py-4">No customers found</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
    <!-- Pagination -->
    <?php if ($customers['last_page'] > 1): ?>
    <div class="card-footer bg-transparent">
        <nav>
            <ul class="pagination pagination-sm mb-0 justify-content-end">
                <?php for ($i = 1; $i <= $customers['last_page']; $i++): ?>
                <li class="page-item <?= $i === $customers['current_page'] ? 'active' : '' ?>">
                    <a class="page-link" href="?<?= http_build_query(array_merge($_GET, ['page' => $i])) ?>"><?= $i ?></a>
                </li>
                <?php endfor; ?>
            </ul>
        </nav>
    </div>
    <?php endif; ?>
</div>

<!-- Add Customer Modal -->
<div class="modal fade" id="addCustomerModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-person-plus me-2"></i>Add Customer</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="create_customer">
                    <div class="col-6">
                        <label class="form-label">First Name</label>
                        <input type="text" name="first_name" class="form-control" required>
                    </div>
                    <div class="col-6">
                        <label class="form-label">Last Name</label>
                        <input type="text" name="last_name" class="form-control">
                    </div>
                    <div class="col-6">
                        <label class="form-label">Phone</label>
                        <input type="tel" name="phone" class="form-control">
                    </div>
                    <div class="col-6">
                        <label class="form-label">Email</label>
                        <input type="email" name="email" class="form-control">
                    </div>
                    <div class="col-6">
                        <label class="form-label">Date of Birth</label>
                        <input type="date" name="date_of_birth" class="form-control">
                    </div>
                    <div class="col-6">
                        <label class="form-label">City</label>
                        <input type="text" name="city" class="form-control">
                    </div>
                    <div class="col-12">
                        <label class="form-label">Notes</label>
                        <textarea name="notes" class="form-control" rows="2"></textarea>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Add Customer</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
async function resegment() {
    if (!confirm('Auto-segment all customers based on their activity?')) return;
    await api('/api/customers/resegment.php','POST',{tenant_id: {$tenantId}});
    showToast('Customers re-segmented','success');
    setTimeout(()=>location.reload(),1200);
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
