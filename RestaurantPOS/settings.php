<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('settings.manage');

$tenantId = Auth::tenantId();
$tenant   = Database::fetchOne("SELECT * FROM tenants WHERE tenant_id = ?", [$tenantId]);
$location = LocationManager::get(Auth::locationId());
$taxRates = Database::fetchAll("SELECT * FROM tax_rates WHERE tenant_id = ? ORDER BY tax_name", [$tenantId]);
$loyalty  = Database::fetchOne("SELECT * FROM loyalty_programs WHERE tenant_id = ? AND is_active = 1", [$tenantId]);

$pageTitle  = 'Settings';
$activeMenu = 'settings';

// Handle updates
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $section = post('section');

    if ($section === 'company') {
        Database::query(
            "UPDATE tenants SET company_name = ?, email = ?, phone = ?, address = ?, updated_at = GETDATE()
             WHERE tenant_id = ?",
            [post('company_name'), post('email'), post('phone'), post('address'), $tenantId]
        );
        flash('success', 'Company settings updated');
        redirect('/settings.php');
    }

    if ($section === 'tax') {
        Database::insert(
            "INSERT INTO tax_rates (tenant_id, tax_name, rate, applies_to, is_inclusive)
             VALUES (?,?,?,?,?)",
            [$tenantId, post('tax_name'), (float)post('rate') / 100, post('applies_to'), post('is_inclusive') ? 1 : 0]
        );
        flash('success', 'Tax rate added');
        redirect('/settings.php#tax');
    }

    if ($section === 'loyalty') {
        if ($loyalty) {
            Database::query(
                "UPDATE loyalty_programs SET program_name = ?, points_per_dollar = ?, redeem_rate = ?, min_redeem = ?
                 WHERE program_id = ?",
                [post('program_name'), (float)post('points_per_dollar'), (float)post('redeem_rate'), (int)post('min_redeem'), $loyalty['program_id']]
            );
        } else {
            Database::insert(
                "INSERT INTO loyalty_programs (tenant_id, program_name, points_per_dollar, redeem_rate, min_redeem)
                 VALUES (?,?,?,?,?)",
                [$tenantId, post('program_name'), (float)post('points_per_dollar'), (float)post('redeem_rate'), (int)post('min_redeem')]
            );
        }
        flash('success', 'Loyalty program updated');
        redirect('/settings.php#loyalty');
    }
}

ob_start();
?>
<div class="row g-4">
    <!-- Company Info -->
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-building me-2"></i>Company Information
            </div>
            <div class="card-body">
                <form method="POST">
                    <input type="hidden" name="section" value="company">
                    <div class="mb-3">
                        <label class="form-label">Company Name</label>
                        <input type="text" name="company_name" class="form-control"
                               value="<?= sanitize($tenant['company_name']) ?>" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Email</label>
                        <input type="email" name="email" class="form-control"
                               value="<?= sanitize($tenant['email']) ?>">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Phone</label>
                        <input type="tel" name="phone" class="form-control"
                               value="<?= sanitize($tenant['phone'] ?? '') ?>">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Address</label>
                        <textarea name="address" class="form-control" rows="2"><?= sanitize($tenant['address'] ?? '') ?></textarea>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Organization Slug</label>
                        <input type="text" class="form-control" value="<?= sanitize($tenant['slug']) ?>" disabled>
                        <small class="text-muted">Used for login: <code><?= sanitize($tenant['slug']) ?></code></small>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Subscription Plan</label>
                        <div>
                            <span class="badge <?= $tenant['plan'] === 'enterprise' ? 'bg-purple' : ($tenant['plan'] === 'professional' ? 'bg-primary' : 'bg-secondary') ?> fs-6">
                                <?= ucfirst(sanitize($tenant['plan'])) ?>
                            </span>
                        </div>
                    </div>
                    <button type="submit" class="btn btn-primary">Save Changes</button>
                </form>
            </div>
        </div>
    </div>

    <!-- Current Location Settings -->
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-geo-alt me-2"></i>Current Location Settings
            </div>
            <div class="card-body">
                <table class="table table-sm mb-0">
                    <tr><td class="text-muted">Location</td><td class="fw-600"><?= sanitize($location['location_name']) ?></td></tr>
                    <tr><td class="text-muted">Address</td><td><?= sanitize($location['address'] ?? '—') ?></td></tr>
                    <tr><td class="text-muted">City/State</td><td><?= sanitize($location['city'] . ', ' . $location['state']) ?></td></tr>
                    <tr><td class="text-muted">Phone</td><td><?= sanitize($location['phone'] ?? '—') ?></td></tr>
                    <tr><td class="text-muted">Timezone</td><td><?= sanitize($location['timezone']) ?></td></tr>
                    <tr><td class="text-muted">Currency</td><td><?= sanitize($location['currency']) ?></td></tr>
                    <tr><td class="text-muted">Tax Rate</td><td><?= number_format($location['tax_rate'] * 100, 2) ?>%</td></tr>
                </table>
                <a href="/locations.php" class="btn btn-outline-primary btn-sm mt-3">
                    <i class="bi bi-pencil me-1"></i>Edit Location
                </a>
            </div>
        </div>

        <!-- Tax Rates -->
        <div class="card shadow-sm mt-3" id="tax">
            <div class="card-header fw-600 bg-transparent d-flex justify-content-between">
                <span><i class="bi bi-percent me-2"></i>Tax Rates</span>
                <button class="btn btn-sm btn-outline-primary" data-bs-toggle="collapse" data-bs-target="#taxForm">
                    <i class="bi bi-plus"></i>
                </button>
            </div>
            <div class="card-body p-0">
                <div class="collapse p-3 border-bottom" id="taxForm">
                    <form method="POST" class="row g-2 align-items-end">
                        <input type="hidden" name="section" value="tax">
                        <div class="col-3"><label class="form-label mb-1 small">Name</label>
                            <input type="text" name="tax_name" class="form-control form-control-sm" required></div>
                        <div class="col-2"><label class="form-label mb-1 small">Rate (%)</label>
                            <input type="number" name="rate" class="form-control form-control-sm" step="0.01" required></div>
                        <div class="col-3"><label class="form-label mb-1 small">Applies To</label>
                            <select name="applies_to" class="form-select form-select-sm">
                                <option value="all">All</option><option value="food">Food</option><option value="beverage">Beverage</option>
                            </select></div>
                        <div class="col-2"><label class="form-label mb-1 small">Inclusive</label>
                            <select name="is_inclusive" class="form-select form-select-sm">
                                <option value="0">No</option><option value="1">Yes</option>
                            </select></div>
                        <div class="col-2"><button type="submit" class="btn btn-sm btn-success w-100">Add</button></div>
                    </form>
                </div>
                <table class="table table-sm mb-0">
                    <?php foreach ($taxRates as $tr): ?>
                    <tr>
                        <td><?= sanitize($tr['tax_name']) ?></td>
                        <td><?= number_format($tr['rate'] * 100, 2) ?>%</td>
                        <td><?= ucfirst($tr['applies_to']) ?></td>
                        <td><?= $tr['is_inclusive'] ? 'Inclusive' : 'Exclusive' ?></td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($taxRates)): ?>
                    <tr><td colspan="4" class="text-muted text-center">No custom tax rates</td></tr>
                    <?php endif; ?>
                </table>
            </div>
        </div>
    </div>

    <!-- Loyalty Program -->
    <div class="col-md-6" id="loyalty">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-star me-2"></i>Loyalty Program
            </div>
            <div class="card-body">
                <form method="POST">
                    <input type="hidden" name="section" value="loyalty">
                    <div class="mb-3">
                        <label class="form-label">Program Name</label>
                        <input type="text" name="program_name" class="form-control"
                               value="<?= sanitize($loyalty['program_name'] ?? 'Rewards') ?>">
                    </div>
                    <div class="row g-3 mb-3">
                        <div class="col-4">
                            <label class="form-label">Points per $1</label>
                            <input type="number" name="points_per_dollar" class="form-control" step="0.1" min="0"
                                   value="<?= $loyalty['points_per_dollar'] ?? 1 ?>">
                        </div>
                        <div class="col-4">
                            <label class="form-label">Redeem Rate ($/pt)</label>
                            <input type="number" name="redeem_rate" class="form-control" step="0.001" min="0"
                                   value="<?= $loyalty['redeem_rate'] ?? 0.01 ?>">
                        </div>
                        <div class="col-4">
                            <label class="form-label">Min Redeem Pts</label>
                            <input type="number" name="min_redeem" class="form-control" min="1"
                                   value="<?= $loyalty['min_redeem'] ?? 100 ?>">
                        </div>
                    </div>
                    <button type="submit" class="btn btn-primary">Save Loyalty Settings</button>
                </form>
            </div>
        </div>
    </div>

    <!-- Roles & Permissions -->
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-shield-lock me-2"></i>Roles & Permissions
            </div>
            <div class="card-body p-0">
                <?php $roles = StaffManager::getRoles($tenantId); ?>
                <table class="table table-sm mb-0">
                    <thead class="table-light"><tr><th>Role</th><th>Permissions</th></tr></thead>
                    <tbody>
                        <?php foreach ($roles as $role): ?>
                        <tr>
                            <td class="fw-600"><?= sanitize($role['role_name']) ?></td>
                            <td>
                                <?php
                                $perms = json_decode($role['permissions'] ?? '[]', true);
                                if (in_array('*', $perms ?: [])):
                                ?>
                                <span class="badge bg-danger">Full Access</span>
                                <?php else: ?>
                                <?php foreach (array_slice($perms ?: [], 0, 5) as $p): ?>
                                <span class="badge bg-light text-dark border"><?= sanitize($p) ?></span>
                                <?php endforeach; ?>
                                <?php if (count($perms ?: []) > 5): ?>
                                <span class="badge bg-secondary">+<?= count($perms) - 5 ?> more</span>
                                <?php endif; ?>
                                <?php endif; ?>
                            </td>
                        </tr>
                        <?php endforeach; ?>
                        <?php if (empty($roles)): ?>
                        <tr><td colspan="2" class="text-muted text-center py-3">No roles configured</td></tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
require_once __DIR__ . '/templates/layout.php';
