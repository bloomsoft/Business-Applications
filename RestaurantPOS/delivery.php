<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$deliveries = DeliveryManager::getActiveDeliveries($locationId);
$drivers    = DeliveryManager::getAvailableDrivers($locationId);
$zones      = Database::fetchAll(
    "SELECT * FROM delivery_zones WHERE location_id = ? AND is_active = 1 ORDER BY zone_name",
    [$locationId]
);

$pageTitle  = 'Delivery Management';
$activeMenu = 'delivery';

// Handle actions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');
    if ($action === 'assign_driver') {
        DeliveryManager::assignDriver((int)post('delivery_id'), (int)post('driver_id'));
        flash('success', 'Driver assigned');
        redirect('/delivery.php');
    }
    if ($action === 'update_status') {
        DeliveryManager::updateStatus((int)post('delivery_id'), post('status'));
        flash('success', 'Delivery status updated');
        redirect('/delivery.php');
    }
    if ($action === 'add_zone') {
        Database::insert(
            "INSERT INTO delivery_zones (location_id, zone_name, min_order, delivery_fee, free_delivery_above, estimated_time)
             VALUES (?,?,?,?,?,?)",
            [$locationId, post('zone_name'), post('min_order'), post('delivery_fee'), post('free_delivery_above') ?: null, post('estimated_time')]
        );
        flash('success', 'Delivery zone added');
        redirect('/delivery.php?tab=zones');
    }
}

$activeTab = get('tab', 'active');
ob_start();
?>
<!-- Tabs -->
<ul class="nav nav-tabs mb-4">
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'active' ? 'active' : '' ?>" href="?tab=active">Active Deliveries</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'zones'  ? 'active' : '' ?>" href="?tab=zones">Delivery Zones</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'drivers'? 'active' : '' ?>" href="?tab=drivers">Drivers</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'integrations' ? 'active' : '' ?>" href="?tab=integrations">Platform Integrations</a></li>
</ul>

<?php if ($activeTab === 'active'): ?>
<!-- Active Deliveries -->
<div class="row g-3 mb-4">
    <div class="col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= count($deliveries) ?></div>
                <div class="text-muted small">Active Deliveries</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= count(array_filter($deliveries, fn($d) => $d['status'] === 'pending')) ?></div>
                <div class="text-muted small">Awaiting Pickup</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= count(array_filter($deliveries, fn($d) => $d['status'] === 'in_transit')) ?></div>
                <div class="text-muted small">In Transit</div>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= count($drivers) ?></div>
                <div class="text-muted small">Available Drivers</div>
            </div>
        </div>
    </div>
</div>

<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover align-middle mb-0">
                <thead class="table-light">
                    <tr>
                        <th>Order #</th><th>Platform</th><th>Address</th><th>Driver</th>
                        <th>Status</th><th>Time</th><th>Amount</th><th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($deliveries as $d): ?>
                    <tr>
                        <td><span class="fw-600">#<?= sanitize($d['order_number']) ?></span></td>
                        <td>
                            <?php
                            $platformIcon = match($d['platform'] ?? 'in-house') {
                                'ubereats'  => 'bg-success',
                                'doordash'  => 'bg-danger',
                                'grubhub'   => 'bg-warning text-dark',
                                default     => 'bg-primary',
                            };
                            ?>
                            <span class="badge <?= $platformIcon ?>"><?= ucfirst(sanitize($d['platform'] ?? 'In-House')) ?></span>
                        </td>
                        <td>
                            <div class="small"><?= sanitize($d['delivery_address'] ?? '—') ?></div>
                        </td>
                        <td>
                            <?php if ($d['driver_name']): ?>
                            <div class="fw-600"><?= sanitize($d['driver_name']) ?></div>
                            <small class="text-muted"><?= sanitize($d['driver_phone'] ?? '') ?></small>
                            <?php else: ?>
                            <form method="POST" class="d-flex gap-1">
                                <input type="hidden" name="action" value="assign_driver">
                                <input type="hidden" name="delivery_id" value="<?= $d['delivery_id'] ?>">
                                <select name="driver_id" class="form-select form-select-sm" style="width:130px">
                                    <option value="">— Driver —</option>
                                    <?php foreach ($drivers as $dr): ?>
                                    <option value="<?= $dr['user_id'] ?>"><?= sanitize($dr['full_name']) ?> (<?= $dr['active_deliveries'] ?>)</option>
                                    <?php endforeach; ?>
                                </select>
                                <button type="submit" class="btn btn-sm btn-primary">Assign</button>
                            </form>
                            <?php endif; ?>
                        </td>
                        <td><?= statusBadge($d['status']) ?></td>
                        <td>
                            <span class="<?= $d['elapsed_min'] > 45 ? 'text-danger fw-bold' : 'text-muted' ?>">
                                <?= $d['elapsed_min'] ?>m
                            </span>
                        </td>
                        <td class="fw-600"><?= money($d['total_amount']) ?></td>
                        <td>
                            <div class="dropdown">
                                <button class="btn btn-sm btn-outline-secondary dropdown-toggle" data-bs-toggle="dropdown">
                                    Update
                                </button>
                                <ul class="dropdown-menu">
                                    <?php foreach (['assigned','picked_up','in_transit','delivered','failed'] as $st): ?>
                                    <li>
                                        <form method="POST" class="d-inline">
                                            <input type="hidden" name="action" value="update_status">
                                            <input type="hidden" name="delivery_id" value="<?= $d['delivery_id'] ?>">
                                            <input type="hidden" name="status" value="<?= $st ?>">
                                            <button type="submit" class="dropdown-item"><?= ucfirst(str_replace('_', ' ', $st)) ?></button>
                                        </form>
                                    </li>
                                    <?php endforeach; ?>
                                </ul>
                            </div>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($deliveries)): ?>
                    <tr><td colspan="8" class="text-center text-muted py-4">No active deliveries</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'zones'): ?>
<!-- Delivery Zones -->
<div class="d-flex justify-content-between mb-3">
    <h5 class="mb-0">Delivery Zones</h5>
    <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addZoneModal">
        <i class="bi bi-plus me-1"></i>Add Zone
    </button>
</div>
<div class="card shadow-sm">
    <div class="card-body p-0">
        <table class="table table-sm align-middle mb-0">
            <thead class="table-light">
                <tr><th>Zone Name</th><th>Min Order</th><th>Delivery Fee</th><th>Free Above</th><th>Est. Time</th><th>Status</th></tr>
            </thead>
            <tbody>
                <?php foreach ($zones as $z): ?>
                <tr>
                    <td class="fw-600"><?= sanitize($z['zone_name']) ?></td>
                    <td><?= money($z['min_order']) ?></td>
                    <td><?= money($z['delivery_fee']) ?></td>
                    <td><?= $z['free_delivery_above'] ? money($z['free_delivery_above']) : '—' ?></td>
                    <td><?= $z['estimated_time'] ?> min</td>
                    <td><?= statusBadge($z['is_active'] ? 'active' : 'inactive') ?></td>
                </tr>
                <?php endforeach; ?>
                <?php if (empty($zones)): ?>
                <tr><td colspan="6" class="text-center text-muted py-4">No delivery zones configured</td></tr>
                <?php endif; ?>
            </tbody>
        </table>
    </div>
</div>

<!-- Add Zone Modal -->
<div class="modal fade" id="addZoneModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Add Delivery Zone</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="add_zone">
                    <div class="col-12"><label class="form-label">Zone Name *</label>
                        <input type="text" name="zone_name" class="form-control" required placeholder="e.g. Downtown 5km"></div>
                    <div class="col-6"><label class="form-label">Min Order ($)</label>
                        <input type="number" name="min_order" class="form-control" step="0.01" value="15.00"></div>
                    <div class="col-6"><label class="form-label">Delivery Fee ($)</label>
                        <input type="number" name="delivery_fee" class="form-control" step="0.01" value="3.99"></div>
                    <div class="col-6"><label class="form-label">Free Above ($)</label>
                        <input type="number" name="free_delivery_above" class="form-control" step="0.01" placeholder="Optional"></div>
                    <div class="col-6"><label class="form-label">Est. Time (min)</label>
                        <input type="number" name="estimated_time" class="form-control" value="30"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Add Zone</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'drivers'): ?>
<!-- Drivers -->
<div class="card shadow-sm">
    <div class="card-header fw-600 bg-transparent border-0 pt-3">Active Drivers</div>
    <div class="card-body p-0">
        <table class="table table-sm align-middle mb-0">
            <thead class="table-light">
                <tr><th>Driver</th><th>Phone</th><th>Active Deliveries</th><th>Status</th></tr>
            </thead>
            <tbody>
                <?php foreach ($drivers as $dr): ?>
                <tr>
                    <td class="fw-600"><?= sanitize($dr['full_name']) ?></td>
                    <td><?= sanitize($dr['phone'] ?? '—') ?></td>
                    <td><span class="badge bg-primary"><?= $dr['active_deliveries'] ?></span></td>
                    <td><?= $dr['active_deliveries'] > 0 ? statusBadge('occupied') : statusBadge('available') ?></td>
                </tr>
                <?php endforeach; ?>
                <?php if (empty($drivers)): ?>
                <tr><td colspan="4" class="text-center text-muted py-4">No drivers found. Add staff with the "Driver" role.</td></tr>
                <?php endif; ?>
            </tbody>
        </table>
    </div>
</div>

<?php elseif ($activeTab === 'integrations'): ?>
<!-- Third-Party Integrations -->
<div class="row g-3">
    <?php
    $platforms = [
        ['name'=>'UberEats',  'icon'=>'bi-bag',     'color'=>'success', 'status'=> UBEREATS_CLIENT_ID ? 'Connected' : 'Not configured',  'configured' => (bool)UBEREATS_CLIENT_ID],
        ['name'=>'DoorDash',  'icon'=>'bi-door-open','color'=>'danger',  'status'=> DOORDASH_DEVELOPER_ID ? 'Connected' : 'Not configured','configured' => (bool)DOORDASH_DEVELOPER_ID],
        ['name'=>'Grubhub',   'icon'=>'bi-shop',    'color'=>'warning', 'status'=> GRUBHUB_API_KEY ? 'Connected' : 'Not configured',      'configured' => (bool)GRUBHUB_API_KEY],
    ];
    foreach ($platforms as $p):
    ?>
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-body text-center py-4">
                <i class="bi <?= $p['icon'] ?> fs-1 text-<?= $p['color'] ?>"></i>
                <h5 class="mt-2"><?= $p['name'] ?></h5>
                <p class="text-muted small">Receive orders from <?= $p['name'] ?> directly into your POS</p>
                <?php if ($p['configured']): ?>
                <span class="badge bg-success"><i class="bi bi-check me-1"></i><?= $p['status'] ?></span>
                <?php else: ?>
                <span class="badge bg-secondary"><?= $p['status'] ?></span>
                <p class="text-muted small mt-2">Set API keys in environment variables to enable.</p>
                <?php endif; ?>
            </div>
        </div>
    </div>
    <?php endforeach; ?>
</div>
<div class="alert alert-info mt-3">
    <i class="bi bi-info-circle me-2"></i>
    <strong>Webhook URLs:</strong> Configure these endpoints in your delivery platform dashboards:
    <ul class="mb-0 mt-1">
        <li>UberEats: <code><?= APP_URL ?>/api/webhooks/ubereats.php</code></li>
        <li>DoorDash: <code><?= APP_URL ?>/api/webhooks/doordash.php</code></li>
        <li>Grubhub: <code><?= APP_URL ?>/api/webhooks/grubhub.php</code></li>
    </ul>
</div>
<?php endif; ?>

<?php
$content = ob_get_clean();
require_once __DIR__ . '/templates/layout.php';
