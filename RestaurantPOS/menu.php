<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('menu.manage');

$tenantId   = Auth::tenantId();
$categories = Database::fetchAll(
    "SELECT mc.*, COUNT(mi.item_id) AS item_count
     FROM menu_categories mc
     LEFT JOIN menu_items mi ON mi.category_id = mc.category_id AND mi.is_available = 1
     WHERE mc.tenant_id = ? AND mc.is_active = 1 AND mc.parent_id IS NULL
     GROUP BY mc.category_id, mc.tenant_id, mc.parent_id, mc.category_name,
              mc.description, mc.image_url, mc.sort_order, mc.is_active, mc.created_at
     ORDER BY mc.sort_order",
    [$tenantId]
);
$activeCat = get('category');
$items = Database::fetchAll(
    "SELECT mi.*, mc.category_name
     FROM menu_items mi
     LEFT JOIN menu_categories mc ON mc.category_id = mi.category_id
     WHERE mi.tenant_id = ?" . ($activeCat ? " AND mi.category_id = ?" : "") . "
     ORDER BY mi.sort_order, mi.item_name",
    $activeCat ? [$tenantId, (int)$activeCat] : [$tenantId]
);

$pageTitle  = 'Menu Manager';
$activeMenu = 'menu';

// Handle actions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');

    if ($action === 'add_category') {
        Database::insert(
            "INSERT INTO menu_categories (tenant_id, category_name, description, sort_order)
             VALUES (?,?,?,?)",
            [$tenantId, post('category_name'), post('description'), (int)post('sort_order')]
        );
        flash('success', 'Category added');
        redirect('/menu.php');
    }

    if ($action === 'add_item') {
        Database::insert(
            "INSERT INTO menu_items
                (tenant_id, category_id, item_name, description, price, cost_price,
                 sku, item_type, calories, prep_time_min, is_taxable, track_inventory, sort_order)
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [
                $tenantId,
                post('category_id') ?: null,
                post('item_name'),
                post('description'),
                (float) post('price'),
                (float) post('cost_price'),
                post('sku'),
                post('item_type'),
                post('calories') ?: null,
                (int) post('prep_time_min'),
                post('is_taxable') ? 1 : 0,
                post('track_inventory') ? 1 : 0,
                (int) post('sort_order'),
            ]
        );
        flash('success', 'Menu item added');
        redirect('/menu.php');
    }

    if ($action === 'toggle_item') {
        $itemId = (int)post('item_id');
        Database::query(
            "UPDATE menu_items SET is_available = CASE WHEN is_available=1 THEN 0 ELSE 1 END WHERE item_id = ?",
            [$itemId]
        );
        flash('info', 'Item availability updated');
        redirect('/menu.php');
    }

    if ($action === 'delete_item') {
        Database::query("UPDATE menu_items SET is_available = 0 WHERE item_id = ?", [(int)post('item_id')]);
        flash('warning', 'Menu item deactivated');
        redirect('/menu.php');
    }
}

ob_start();
?>
<!-- Actions Bar -->
<div class="d-flex flex-wrap gap-2 justify-content-between align-items-center mb-3">
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addItemModal">
            <i class="bi bi-plus-lg me-1"></i>Add Item
        </button>
        <button class="btn btn-sm btn-outline-primary" data-bs-toggle="modal" data-bs-target="#addCategoryModal">
            <i class="bi bi-folder-plus me-1"></i>Add Category
        </button>
    </div>
    <form method="GET" class="d-flex gap-2">
        <select name="category" class="form-select form-select-sm" style="width:180px" onchange="this.form.submit()">
            <option value="">All Categories</option>
            <?php foreach ($categories as $cat): ?>
            <option value="<?= $cat['category_id'] ?>" <?= $activeCat == $cat['category_id'] ? 'selected' : '' ?>>
                <?= sanitize($cat['category_name']) ?> (<?= $cat['item_count'] ?>)
            </option>
            <?php endforeach; ?>
        </select>
    </form>
</div>

<!-- Categories Overview -->
<div class="row g-2 mb-4">
    <?php foreach ($categories as $cat): ?>
    <div class="col-6 col-md-3 col-lg-2">
        <a href="?category=<?= $cat['category_id'] ?>"
           class="card shadow-sm text-decoration-none h-100 <?= $activeCat == $cat['category_id'] ? 'border-primary' : '' ?>">
            <div class="card-body text-center py-3">
                <div class="fw-bold"><?= sanitize($cat['category_name']) ?></div>
                <small class="text-muted"><?= $cat['item_count'] ?> items</small>
            </div>
        </a>
    </div>
    <?php endforeach; ?>
</div>

<!-- Items Table -->
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover align-middle mb-0">
                <thead class="table-light">
                    <tr>
                        <th>Item</th><th>Category</th><th>Type</th>
                        <th>Price</th><th>Cost</th><th>Margin</th>
                        <th>Prep Time</th><th>Status</th><th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($items as $item):
                        $margin = $item['price'] > 0
                            ? round((($item['price'] - $item['cost_price']) / $item['price']) * 100, 1)
                            : 0;
                    ?>
                    <tr class="<?= !$item['is_available'] ? 'table-light text-muted' : '' ?>">
                        <td>
                            <div class="d-flex align-items-center gap-2">
                                <?php if ($item['image_url']): ?>
                                <img src="<?= sanitize($item['image_url']) ?>" class="rounded"
                                     style="width:40px;height:40px;object-fit:cover" alt="">
                                <?php else: ?>
                                <div class="rounded bg-light d-flex align-items-center justify-content-center"
                                     style="width:40px;height:40px">
                                    <i class="bi bi-cup-hot text-secondary"></i>
                                </div>
                                <?php endif; ?>
                                <div>
                                    <div class="fw-600"><?= sanitize($item['item_name']) ?></div>
                                    <?php if ($item['sku']): ?>
                                    <small class="text-muted">SKU: <?= sanitize($item['sku']) ?></small>
                                    <?php endif; ?>
                                </div>
                            </div>
                        </td>
                        <td><small><?= sanitize($item['category_name'] ?? '—') ?></small></td>
                        <td><span class="badge bg-light text-dark border"><?= ucfirst(sanitize($item['item_type'])) ?></span></td>
                        <td class="fw-bold text-accent"><?= money($item['price']) ?></td>
                        <td><?= money($item['cost_price']) ?></td>
                        <td>
                            <span class="<?= $margin >= 60 ? 'text-success' : ($margin >= 40 ? 'text-warning' : 'text-danger') ?>">
                                <?= $margin ?>%
                            </span>
                        </td>
                        <td><?= $item['prep_time_min'] ?>m</td>
                        <td>
                            <?= statusBadge($item['is_available'] ? 'available' : 'inactive') ?>
                            <?php if ($item['track_inventory']): ?>
                            <span class="badge bg-info">Tracked</span>
                            <?php endif; ?>
                        </td>
                        <td>
                            <form method="POST" class="d-inline">
                                <input type="hidden" name="action" value="toggle_item">
                                <input type="hidden" name="item_id" value="<?= $item['item_id'] ?>">
                                <button type="submit" class="btn btn-sm btn-outline-<?= $item['is_available'] ? 'warning' : 'success' ?>"
                                        title="<?= $item['is_available'] ? 'Disable' : 'Enable' ?>">
                                    <i class="bi bi-<?= $item['is_available'] ? 'pause' : 'play' ?>"></i>
                                </button>
                            </form>
                            <button class="btn btn-sm btn-outline-primary"
                                    onclick="editItem(<?= $item['item_id'] ?>)">
                                <i class="bi bi-pencil"></i>
                            </button>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($items)): ?>
                    <tr><td colspan="9" class="text-center text-muted py-4">No menu items found</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Add Category Modal -->
<div class="modal fade" id="addCategoryModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-folder-plus me-2"></i>Add Category</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="add_category">
                    <div class="col-12"><label class="form-label">Category Name *</label>
                        <input type="text" name="category_name" class="form-control" required></div>
                    <div class="col-12"><label class="form-label">Description</label>
                        <textarea name="description" class="form-control" rows="2"></textarea></div>
                    <div class="col-6"><label class="form-label">Sort Order</label>
                        <input type="number" name="sort_order" class="form-control" value="0"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-primary">Add Category</button>
                </div>
            </form>
        </div>
    </div>
</div>

<!-- Add Item Modal -->
<div class="modal fade" id="addItemModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-plus-circle me-2"></i>Add Menu Item</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="add_item">
                    <div class="col-8"><label class="form-label">Item Name *</label>
                        <input type="text" name="item_name" class="form-control" required></div>
                    <div class="col-4"><label class="form-label">Category</label>
                        <select name="category_id" class="form-select">
                            <option value="">— None —</option>
                            <?php foreach ($categories as $c): ?>
                            <option value="<?= $c['category_id'] ?>"><?= sanitize($c['category_name']) ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-12"><label class="form-label">Description</label>
                        <textarea name="description" class="form-control" rows="2"></textarea></div>
                    <div class="col-3"><label class="form-label">Price *</label>
                        <input type="number" name="price" class="form-control" step="0.01" min="0" required></div>
                    <div class="col-3"><label class="form-label">Cost Price</label>
                        <input type="number" name="cost_price" class="form-control" step="0.01" min="0" value="0"></div>
                    <div class="col-3"><label class="form-label">SKU</label>
                        <input type="text" name="sku" class="form-control"></div>
                    <div class="col-3"><label class="form-label">Type</label>
                        <select name="item_type" class="form-select">
                            <option value="food">Food</option>
                            <option value="beverage">Beverage</option>
                            <option value="combo">Combo</option>
                            <option value="modifier">Modifier</option>
                        </select>
                    </div>
                    <div class="col-3"><label class="form-label">Calories</label>
                        <input type="number" name="calories" class="form-control" min="0"></div>
                    <div class="col-3"><label class="form-label">Prep Time (min)</label>
                        <input type="number" name="prep_time_min" class="form-control" min="0" value="10"></div>
                    <div class="col-3"><label class="form-label">Sort Order</label>
                        <input type="number" name="sort_order" class="form-control" value="0"></div>
                    <div class="col-3 d-flex align-items-end gap-3">
                        <div class="form-check">
                            <input type="checkbox" name="is_taxable" value="1" class="form-check-input" checked>
                            <label class="form-check-label">Taxable</label>
                        </div>
                        <div class="form-check">
                            <input type="checkbox" name="track_inventory" value="1" class="form-check-input">
                            <label class="form-check-label">Track Stock</label>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Add Item</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
function editItem(itemId) {
    // Open edit modal (extend with AJAX load)
    window.location.href = '/menu-edit.php?id=' + itemId;
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
