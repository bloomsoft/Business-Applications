<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('inventory.view');

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$filters    = ['search' => get('search'), 'category' => get('category'), 'low_stock' => get('filter') === 'low_stock'];
$items      = InventoryManager::getItems($locationId, $filters);
$categories = Database::fetchAll("SELECT * FROM inventory_categories WHERE tenant_id = ?", [$tenantId]);
$suppliers  = Database::fetchAll("SELECT * FROM suppliers WHERE tenant_id = ? AND is_active = 1 ORDER BY supplier_name", [$tenantId]);
$pageTitle  = 'Inventory Management';
$activeMenu = 'inventory';

// Handle actions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');
    if ($action === 'adjust') {
        InventoryManager::adjustStock(
            (int)post('inv_item_id'),
            (float)post('quantity'),
            post('type'),
            post('notes')
        );
        flash('success', 'Stock adjusted successfully');
        redirect('/inventory.php');
    }
}

ob_start();
?>
<!-- Filters & Actions -->
<div class="d-flex flex-wrap gap-2 justify-content-between align-items-end mb-3">
    <form method="GET" class="d-flex gap-2 flex-wrap">
        <input type="text" name="search" class="form-control form-control-sm" style="width:200px"
               placeholder="Search..." value="<?= sanitize(get('search')) ?>">
        <select name="category" class="form-select form-select-sm" style="width:160px">
            <option value="">All Categories</option>
            <?php foreach ($categories as $cat): ?>
            <option value="<?= $cat['inv_cat_id'] ?>" <?= get('category') == $cat['inv_cat_id'] ? 'selected' : '' ?>>
                <?= sanitize($cat['cat_name']) ?>
            </option>
            <?php endforeach; ?>
        </select>
        <a href="?filter=low_stock" class="btn btn-sm <?= get('filter') === 'low_stock' ? 'btn-warning' : 'btn-outline-warning' ?>">
            <i class="bi bi-exclamation-triangle me-1"></i>Low Stock
        </a>
        <button type="submit" class="btn btn-sm btn-primary">Filter</button>
    </form>
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addItemModal">
            <i class="bi bi-plus-lg me-1"></i>Add Item
        </button>
        <button class="btn btn-sm btn-outline-primary" data-bs-toggle="modal" data-bs-target="#poModal">
            <i class="bi bi-cart-plus me-1"></i>Purchase Order
        </button>
    </div>
</div>

<!-- Summary Cards -->
<div class="row g-2 mb-3">
    <?php
    $totalItems = count($items);
    $lowItems   = count(array_filter($items, fn($i) => $i['stock_status'] === 'low'));
    $outItems   = count(array_filter($items, fn($i) => $i['stock_status'] === 'out'));
    $invValue   = array_sum(array_map(fn($i) => $i['quantity_on_hand'] * $i['cost_per_unit'], $items));
    ?>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= $totalItems ?></div>
                <div class="text-muted small">Total Items</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm border-warning">
            <div class="card-body py-2">
                <div class="fw-bold fs-5 text-warning"><?= $lowItems ?></div>
                <div class="text-muted small">Low Stock</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm border-danger">
            <div class="card-body py-2">
                <div class="fw-bold fs-5 text-danger"><?= $outItems ?></div>
                <div class="text-muted small">Out of Stock</div>
            </div>
        </div>
    </div>
    <div class="col-6 col-md-3">
        <div class="card text-center shadow-sm">
            <div class="card-body py-2">
                <div class="fw-bold fs-5"><?= money($invValue) ?></div>
                <div class="text-muted small">Inventory Value</div>
            </div>
        </div>
    </div>
</div>

<!-- Inventory Table -->
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover table-sm align-middle mb-0">
                <thead class="table-light">
                    <tr>
                        <th>Item</th><th>Category</th><th>Supplier</th>
                        <th>On Hand</th><th>Reorder Level</th><th>Cost/Unit</th>
                        <th>Value</th><th>Status</th><th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($items as $item): ?>
                    <tr>
                        <td>
                            <div class="fw-600"><?= sanitize($item['item_name']) ?></div>
                            <?php if ($item['sku']): ?>
                            <small class="text-muted">SKU: <?= sanitize($item['sku']) ?></small>
                            <?php endif; ?>
                        </td>
                        <td><small><?= sanitize($item['category_name'] ?? '—') ?></small></td>
                        <td><small><?= sanitize($item['supplier_name'] ?? '—') ?></small></td>
                        <td class="fw-600">
                            <?= number_format($item['quantity_on_hand'], 2) ?> <?= sanitize($item['unit']) ?>
                        </td>
                        <td><?= number_format($item['reorder_level'], 2) ?></td>
                        <td><?= money($item['cost_per_unit']) ?></td>
                        <td><?= money($item['quantity_on_hand'] * $item['cost_per_unit']) ?></td>
                        <td>
                            <?php
                            $stockClass = match($item['stock_status']) {
                                'out' => 'bg-danger',
                                'low' => 'bg-warning text-dark',
                                default => 'bg-success'
                            };
                            ?>
                            <span class="badge <?= $stockClass ?>">
                                <?= ucfirst($item['stock_status']) ?>
                            </span>
                        </td>
                        <td>
                            <button class="btn btn-sm btn-outline-primary"
                                    onclick="adjustStock(<?= $item['inv_item_id'] ?>, '<?= addslashes($item['item_name']) ?>', <?= $item['quantity_on_hand'] ?>)">
                                <i class="bi bi-arrows-vertical"></i>
                            </button>
                            <a href="/inventory-detail.php?id=<?= $item['inv_item_id'] ?>"
                               class="btn btn-sm btn-outline-secondary">
                                <i class="bi bi-clock-history"></i>
                            </a>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($items)): ?>
                    <tr><td colspan="9" class="text-center text-muted py-4">No inventory items found</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Stock Adjustment Modal -->
<div class="modal fade" id="adjustModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Adjust Stock — <span id="adjustItemName"></span></h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body">
                    <input type="hidden" name="action" value="adjust">
                    <input type="hidden" name="inv_item_id" id="adjustItemId">
                    <div class="mb-3">
                        <label class="form-label">Current Stock</label>
                        <div class="fw-bold" id="currentStock"></div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Adjustment Type</label>
                        <select name="type" class="form-select" required>
                            <option value="in">Stock In (Receiving)</option>
                            <option value="out">Stock Out (Usage)</option>
                            <option value="waste">Waste/Spoilage</option>
                            <option value="adjustment">Manual Adjustment</option>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Quantity</label>
                        <input type="number" name="quantity" class="form-control" step="0.001" min="0.001" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Notes</label>
                        <textarea name="notes" class="form-control" rows="2"></textarea>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-primary">Save Adjustment</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
function adjustStock(id, name, current) {
    document.getElementById('adjustItemId').value  = id;
    document.getElementById('adjustItemName').textContent = name;
    document.getElementById('currentStock').textContent   = current;
    new bootstrap.Modal('#adjustModal').show();
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
