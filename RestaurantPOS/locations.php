<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('locations.manage');

$tenantId  = Auth::tenantId();
$locations = LocationManager::getAll($tenantId);
$startDate = date('Y-m-01');
$endDate   = date('Y-m-d');
$comparison= LocationManager::getComparisonStats($tenantId, $startDate, $endDate);

// Handle create
if ($_SERVER['REQUEST_METHOD'] === 'POST' && post('action') === 'create') {
    LocationManager::create($_POST, $tenantId);
    flash('success', 'Location added');
    redirect('/locations.php');
}

$pageTitle = 'Locations Management';
$activeMenu= 'locations';
ob_start();
?>
<!-- Comparison Chart -->
<div class="card shadow-sm mb-4">
    <div class="card-header fw-600 bg-transparent border-0 pt-3">
        <i class="bi bi-geo-alt me-2 text-primary"></i>Location Performance — <?= date('F Y') ?>
    </div>
    <div class="card-body">
        <canvas id="locationChart" height="80"></canvas>
    </div>
</div>

<!-- Location Cards -->
<div class="row g-3 mb-4">
    <?php foreach ($locations as $loc): ?>
    <div class="col-md-6 col-xl-4">
        <div class="card shadow-sm h-100">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-start mb-3">
                    <div>
                        <h5 class="mb-0 fw-bold"><?= sanitize($loc['location_name']) ?></h5>
                        <small class="text-muted">
                            <i class="bi bi-geo-alt me-1"></i><?= sanitize($loc['city'] . ', ' . $loc['state']) ?>
                        </small>
                    </div>
                    <div>
                        <?php if ($loc['location_id'] == Auth::locationId()): ?>
                        <span class="badge bg-success">Current</span>
                        <?php endif; ?>
                        <?= statusBadge($loc['is_active'] ? 'active' : 'inactive') ?>
                    </div>
                </div>
                <div class="row g-2 mb-3">
                    <div class="col-6">
                        <div class="text-muted small">Today's Revenue</div>
                        <div class="fw-bold"><?= money($loc['todays_revenue']) ?></div>
                    </div>
                    <div class="col-6">
                        <div class="text-muted small">Today's Orders</div>
                        <div class="fw-bold"><?= number_format($loc['todays_orders']) ?></div>
                    </div>
                    <div class="col-6">
                        <div class="text-muted small">Staff</div>
                        <div class="fw-bold"><?= $loc['staff_count'] ?></div>
                    </div>
                    <div class="col-6">
                        <div class="text-muted small">Tax Rate</div>
                        <div class="fw-bold"><?= number_format($loc['tax_rate'] * 100, 1) ?>%</div>
                    </div>
                </div>
                <div class="d-flex gap-2">
                    <button class="btn btn-sm btn-outline-primary flex-grow-1"
                            onclick="switchLocation(<?= $loc['location_id'] ?>)">
                        <i class="bi bi-arrow-right-circle me-1"></i>Switch To
                    </button>
                    <button class="btn btn-sm btn-outline-secondary"
                            onclick="editLocation(<?= $loc['location_id'] ?>)">
                        <i class="bi bi-pencil"></i>
                    </button>
                </div>
            </div>
        </div>
    </div>
    <?php endforeach; ?>

    <!-- Add New Location Card -->
    <div class="col-md-6 col-xl-4">
        <div class="card shadow-sm h-100 border-dashed" style="border-style:dashed">
            <div class="card-body d-flex flex-column align-items-center justify-content-center text-center text-muted py-5">
                <i class="bi bi-plus-circle fs-1 mb-2"></i>
                <h6>Add New Location</h6>
                <button class="btn btn-outline-primary btn-sm mt-2"
                        data-bs-toggle="modal" data-bs-target="#addLocationModal">
                    Add Location
                </button>
            </div>
        </div>
    </div>
</div>

<!-- Comparison Table -->
<div class="card shadow-sm">
    <div class="card-header fw-600 bg-transparent border-0 pt-3">
        <i class="bi bi-table me-2"></i>Month-to-Date Comparison
    </div>
    <div class="card-body p-0">
        <table class="table table-sm align-middle mb-0">
            <thead class="table-light">
                <tr><th>Location</th><th>Orders</th><th>Revenue</th><th>Avg Order</th><th>Customers</th></tr>
            </thead>
            <tbody>
                <?php foreach ($comparison as $row): ?>
                <tr>
                    <td class="fw-600"><?= sanitize($row['location_name']) ?></td>
                    <td><?= number_format($row['order_count']) ?></td>
                    <td><?= money($row['revenue']) ?></td>
                    <td><?= money($row['avg_order_value']) ?></td>
                    <td><?= number_format($row['unique_customers']) ?></td>
                </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
    </div>
</div>

<!-- Add Location Modal -->
<div class="modal fade" id="addLocationModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-geo-alt me-2"></i>Add Location</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="create">
                    <div class="col-12"><label class="form-label">Location Name *</label>
                        <input type="text" name="location_name" class="form-control" required></div>
                    <div class="col-12"><label class="form-label">Address</label>
                        <input type="text" name="address" class="form-control"></div>
                    <div class="col-4"><label class="form-label">City</label>
                        <input type="text" name="city" class="form-control"></div>
                    <div class="col-4"><label class="form-label">State</label>
                        <input type="text" name="state" class="form-control"></div>
                    <div class="col-4"><label class="form-label">ZIP</label>
                        <input type="text" name="zip" class="form-control"></div>
                    <div class="col-4"><label class="form-label">Phone</label>
                        <input type="tel" name="phone" class="form-control"></div>
                    <div class="col-4"><label class="form-label">Timezone</label>
                        <select name="timezone" class="form-select">
                            <?php foreach (['UTC','America/New_York','America/Chicago','America/Denver','America/Los_Angeles'] as $tz): ?>
                            <option value="<?= $tz ?>"><?= $tz ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-4"><label class="form-label">Tax Rate (%)</label>
                        <input type="number" name="tax_rate" class="form-control" step="0.001" min="0" max="1" value="0.08">
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-primary">Add Location</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$locNamesJson    = json_encode(array_column($comparison, 'location_name'));
$locRevenuesJson = json_encode(array_map('floatval', array_column($comparison, 'revenue')));
$scripts = <<<JS
<script>
renderBarChart('locationChart',
    $locNamesJson,
    [{label:'Revenue', data: $locRevenuesJson,
      backgroundColor:'rgba(249,115,22,.7)'}]
);

async function switchLocation(locationId) {
    await fetch('/api/switch-location.php',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({location_id:locationId})
    });
    showToast('Switched location','success');
    setTimeout(()=>location.reload(),800);
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
