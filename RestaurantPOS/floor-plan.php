<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tables     = TableManager::getTables($locationId);
$areas      = TableManager::getAreas($locationId);
$pageTitle  = 'Floor Plan';
$activeMenu = 'pos';
ob_start();
?>
<div class="d-flex justify-content-between align-items-center mb-3">
    <div class="d-flex gap-2">
        <span class="badge bg-success py-2"><span class="rounded-circle bg-white d-inline-block" style="width:10px;height:10px"></span> Available</span>
        <span class="badge bg-danger py-2"><span class="rounded-circle bg-white d-inline-block" style="width:10px;height:10px"></span> Occupied</span>
        <span class="badge bg-info py-2"><span class="rounded-circle bg-white d-inline-block" style="width:10px;height:10px"></span> Reserved</span>
        <span class="badge bg-warning py-2"><span class="rounded-circle bg-white d-inline-block" style="width:10px;height:10px"></span> Cleaning</span>
    </div>
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-outline-primary" data-bs-toggle="modal" data-bs-target="#addTableModal">
            <i class="bi bi-plus me-1"></i>Add Table
        </button>
        <a href="/pos.php" class="btn btn-sm btn-outline-secondary">Back to POS</a>
    </div>
</div>

<!-- Floor Plan Canvas -->
<div id="floorPlan" style="min-height:600px;position:relative">
    <?php foreach ($tables as $table):
        $w = ($table['shape'] === 'circle') ? 80 : 100;
        $h = 80;
    ?>
    <div class="table-shape <?= $table['status'] ?> <?= $table['shape'] ?>"
         data-table-id="<?= $table['table_id'] ?>"
         style="left:<?= $table['pos_x'] ?>px; top:<?= $table['pos_y'] ?>px; width:<?= $w ?>px; height:<?= $h ?>px"
         onclick="selectTable(<?= $table['table_id'] ?>, '<?= $table['status'] ?>', <?= $table['order_id'] ?? 'null' ?>)"
         title="Table <?= sanitize($table['table_number']) ?> — <?= ucfirst($table['status']) ?>&#10;<?= $table['capacity'] ?> seats<?= $table['order_id'] ? '&#10;Order #' . sanitize($table['order_number']) . ' (' . money($table['total_amount']) . ') ' . $table['occupied_minutes'] . 'm' : '' ?>">
        <div class="text-center">
            <div class="fw-bold"><?= sanitize($table['table_number']) ?></div>
            <small><?= $table['capacity'] ?>p</small>
        </div>
    </div>
    <?php endforeach; ?>
</div>

<!-- Table Action Panel -->
<div class="offcanvas offcanvas-end" id="tablePanel" tabindex="-1">
    <div class="offcanvas-header border-bottom">
        <h5 class="offcanvas-title" id="tablePanelTitle">Table Details</h5>
        <button class="btn-close" data-bs-dismiss="offcanvas"></button>
    </div>
    <div class="offcanvas-body" id="tablePanelBody"></div>
</div>

<!-- Add Table Modal -->
<div class="modal fade" id="addTableModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Add Table</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST" action="/api/tables/create.php">
                <div class="modal-body row g-3">
                    <div class="col-6"><label class="form-label">Table Number *</label>
                        <input type="text" name="table_number" class="form-control" required></div>
                    <div class="col-6"><label class="form-label">Capacity</label>
                        <input type="number" name="capacity" class="form-control" value="4" min="1"></div>
                    <div class="col-6"><label class="form-label">Shape</label>
                        <select name="shape" class="form-select">
                            <option value="rectangle">Rectangle</option>
                            <option value="circle">Circle</option>
                        </select></div>
                    <div class="col-6"><label class="form-label">Area</label>
                        <select name="area_id" class="form-select">
                            <option value="">None</option>
                            <?php foreach ($areas as $area): ?>
                            <option value="<?= $area['area_id'] ?>"><?= sanitize($area['area_name']) ?></option>
                            <?php endforeach; ?>
                        </select></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-primary">Add Table</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
initFloorPlan();

function selectTable(tableId, status, orderId) {
    const panel = new bootstrap.Offcanvas('#tablePanel');
    const body  = document.getElementById('tablePanelBody');

    let actions = '';
    if (status === 'available') {
        actions = `<a href="/pos.php?table_id=${tableId}" class="btn btn-accent w-100 mb-2">
            <i class="bi bi-plus me-1"></i>New Order</a>`;
    }
    if (status === 'occupied' && orderId) {
        actions = `
            <a href="/pos.php?order_id=${orderId}" class="btn btn-primary w-100 mb-2">
                <i class="bi bi-pencil me-1"></i>View Order #${orderId}</a>
            <button class="btn btn-outline-success w-100 mb-2" onclick="setTableStatus(${tableId},'available')">
                <i class="bi bi-check me-1"></i>Mark Available</button>`;
    }
    if (status === 'reserved') {
        actions = `<button class="btn btn-success w-100 mb-2" onclick="setTableStatus(${tableId},'occupied')">
            <i class="bi bi-person-check me-1"></i>Seat Guests</button>`;
    }
    actions += `
        <hr>
        <div class="d-flex gap-2">
            <button class="btn btn-outline-secondary flex-grow-1" onclick="setTableStatus(${tableId},'cleaning')">Cleaning</button>
            <button class="btn btn-outline-secondary flex-grow-1" onclick="setTableStatus(${tableId},'reserved')">Reserve</button>
            <button class="btn btn-outline-secondary flex-grow-1" onclick="setTableStatus(${tableId},'available')">Free</button>
        </div>`;

    body.innerHTML = actions;
    panel.show();
}

async function setTableStatus(tableId, status) {
    await api('/api/tables/update-status.php','POST',{table_id: tableId, status: status});
    showToast('Table updated','success');
    setTimeout(()=>location.reload(),600);
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
