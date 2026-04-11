<?php
/**
 * QR Code Self-Ordering Page (public — no auth required)
 */
require_once __DIR__ . '/core/bootstrap.php';

$token = get('t');
if (!$token) {
    http_response_code(404);
    die('<h3>Invalid QR code</h3>');
}

$table = TableManager::getByQRToken($token);
if (!$table) {
    http_response_code(404);
    die('<h3>Table not found</h3>');
}

$tenantId  = $table['tenant_id'];
$menu      = QRKioskManager::getPublicMenu($tenantId);
$tenant    = Database::fetchOne("SELECT * FROM tenants WHERE tenant_id = ?", [$tenantId]);
$pageTitle = sanitize($tenant['company_name']) . ' — Order';
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= $pageTitle ?></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link rel="stylesheet" href="/public/css/app.css">
    <style>
        body { background: #f8fafc; }
        .sticky-header { position: sticky; top: 0; z-index: 100; }
        .category-pill.active { background: var(--pos-accent); color: #fff; border-color: var(--pos-accent); }
        .kiosk-cart-btn {
            position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%);
            z-index: 200; min-width: 220px;
        }
    </style>
</head>
<body>

<!-- Header -->
<div class="sticky-header bg-white shadow-sm py-3 px-4">
    <div class="d-flex align-items-center justify-content-between">
        <div>
            <?php if ($tenant['logo_url']): ?>
            <img src="<?= sanitize($tenant['logo_url']) ?>" height="36" alt="">
            <?php endif; ?>
            <span class="fw-bold ms-2"><?= sanitize($tenant['company_name']) ?></span>
        </div>
        <div class="text-muted">
            <i class="bi bi-table me-1"></i>Table <?= sanitize($table['table_number']) ?>
        </div>
    </div>
    <!-- Category Tabs -->
    <div class="d-flex gap-2 mt-2 overflow-auto pb-1">
        <button class="btn btn-sm rounded-pill border category-pill active" data-category="">All</button>
        <?php foreach ($menu as $cat): ?>
        <button class="btn btn-sm rounded-pill border category-pill"
                data-category="<?= $cat['category_id'] ?>">
            <?= sanitize($cat['category_name']) ?>
        </button>
        <?php endforeach; ?>
    </div>
</div>

<!-- Menu Items -->
<div class="container-fluid px-3 py-3 pb-5 mb-5">
    <?php foreach ($menu as $cat): if (empty($cat['items'])) continue; ?>
    <h5 class="fw-bold mt-3 mb-2" id="cat<?= $cat['category_id'] ?>">
        <?= sanitize($cat['category_name']) ?>
    </h5>
    <div class="row g-2" data-category-section="<?= $cat['category_id'] ?>">
        <?php foreach ($cat['items'] as $item): ?>
        <div class="col-6 col-md-4 col-lg-3 kiosk-item" data-category="<?= $cat['category_id'] ?>">
            <div class="kiosk-item-card card h-100" onclick="openItem(<?= htmlspecialchars(json_encode($item)) ?>)">
                <?php if ($item['image_url']): ?>
                <img src="<?= sanitize($item['image_url']) ?>"
                     class="card-img-top" style="height:120px;object-fit:cover" alt="">
                <?php endif; ?>
                <div class="card-body p-2">
                    <div class="fw-600"><?= sanitize($item['item_name']) ?></div>
                    <?php if ($item['description']): ?>
                    <div class="text-muted small mt-1" style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">
                        <?= sanitize($item['description']) ?>
                    </div>
                    <?php endif; ?>
                    <div class="mt-2 d-flex justify-content-between align-items-center">
                        <strong class="text-accent"><?= money($item['price']) ?></strong>
                        <?php if ($item['calories']): ?>
                        <small class="text-muted"><?= $item['calories'] ?> cal</small>
                        <?php endif; ?>
                    </div>
                </div>
            </div>
        </div>
        <?php endforeach; ?>
    </div>
    <?php endforeach; ?>
</div>

<!-- Cart FAB -->
<button class="btn btn-accent btn-lg shadow kiosk-cart-btn d-none" id="cartFab"
        data-bs-toggle="offcanvas" data-bs-target="#cartPanel">
    <i class="bi bi-cart3 me-2"></i>View Cart
    <span class="badge bg-white text-accent ms-2" id="kioskCartCount">0</span>
</button>

<!-- Cart Offcanvas -->
<div class="offcanvas offcanvas-end" tabindex="-1" id="cartPanel" style="max-width:400px">
    <div class="offcanvas-header border-bottom">
        <h5 class="offcanvas-title"><i class="bi bi-cart3 me-2"></i>Your Order</h5>
        <button type="button" class="btn-close" data-bs-dismiss="offcanvas"></button>
    </div>
    <div class="offcanvas-body" id="kioskCart">
        <div class="text-center text-muted py-4">Your cart is empty</div>
    </div>
</div>

<!-- Item Detail Modal -->
<div class="modal fade" id="itemModal" tabindex="-1">
    <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
            <div class="modal-header border-0 pb-0">
                <h5 class="modal-title" id="itemModalTitle"></h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body" id="itemModalBody"></div>
            <div class="modal-footer">
                <div class="d-flex align-items-center gap-3 me-auto">
                    <button class="btn btn-outline-secondary" onclick="changeModalQty(-1)">−</button>
                    <span id="modalQty" class="fw-bold fs-5">1</span>
                    <button class="btn btn-outline-secondary" onclick="changeModalQty(1)">+</button>
                </div>
                <button class="btn btn-accent" onclick="addModalItem()">
                    Add — <span id="modalItemTotal"></span>
                </button>
            </div>
        </div>
    </div>
</div>

<!-- Customer Info Modal -->
<div class="modal fade" id="customerModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5>Your Details (Optional)</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="mb-3">
                    <label class="form-label">Name</label>
                    <input type="text" class="form-control" id="custName" placeholder="Your name">
                </div>
                <div class="mb-3">
                    <label class="form-label">Phone</label>
                    <input type="tel" class="form-control" id="custPhone" placeholder="For order updates">
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" onclick="placeOrderNow(null)">Skip</button>
                <button class="btn btn-accent" onclick="placeOrderWithCustomer()">Place Order</button>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="/public/js/app.js"></script>
<script>
const TABLE_TOKEN = <?= json_encode($token) ?>;
let modalItem = null;
let modalQty  = 1;

// Category filter
document.querySelectorAll('.category-pill').forEach(pill => {
    pill.addEventListener('click', function() {
        document.querySelectorAll('.category-pill').forEach(p => p.classList.remove('active'));
        this.classList.add('active');
        const cat = this.dataset.category;
        document.querySelectorAll('.kiosk-item').forEach(el => {
            el.style.display = (!cat || el.dataset.category === cat) ? '' : 'none';
        });
        if (cat) {
            document.getElementById('cat' + cat)?.scrollIntoView({ behavior: 'smooth' });
        }
    });
});

function openItem(item) {
    modalItem = item;
    modalQty  = 1;
    document.getElementById('itemModalTitle').textContent = item.item_name;
    document.getElementById('modalQty').textContent = 1;
    document.getElementById('modalItemTotal').textContent = money(item.price);

    let html = '';
    if (item.description) html += `<p class="text-muted">${item.description}</p>`;
    if (item.modifier_groups?.length) {
        html += item.modifier_groups.map(grp => `
            <div class="mb-3">
                <div class="fw-600">${grp.group_name}
                    ${grp.is_required ? '<span class="badge bg-danger">Required</span>' : ''}
                </div>
                ${grp.options.map(opt => `
                    <div class="form-check">
                        <input class="form-check-input mod-opt"
                               type="${grp.selection_type==='single'?'radio':'checkbox'}"
                               name="grp_${grp.group_id}" value="${opt.modifier_id}"
                               data-price="${opt.price_add}">
                        <label class="form-check-label">
                            ${opt.modifier_name}
                            ${opt.price_add>0?`<span class="text-accent">+${money(opt.price_add)}</span>`:''}
                        </label>
                    </div>`).join('')}
            </div>`).join('');
    }
    html += `<div class="mt-2">
        <label class="form-label text-muted small">Special instructions</label>
        <input type="text" class="form-control form-control-sm" id="itemNotes" placeholder="e.g. no onions">
    </div>`;

    document.getElementById('itemModalBody').innerHTML = html;

    document.querySelectorAll('.mod-opt').forEach(el => {
        el.addEventListener('change', updateModalTotal);
    });

    new bootstrap.Modal('#itemModal').show();
}

function changeModalQty(d) {
    modalQty = Math.max(1, modalQty + d);
    document.getElementById('modalQty').textContent = modalQty;
    updateModalTotal();
}

function updateModalTotal() {
    let extra = 0;
    document.querySelectorAll('.mod-opt:checked').forEach(el => extra += parseFloat(el.dataset.price)||0);
    document.getElementById('modalItemTotal').textContent = money((modalItem.price + extra) * modalQty);
}

function addModalItem() {
    const mods  = [...document.querySelectorAll('.mod-opt:checked')].map(el => parseInt(el.value));
    const notes = document.getElementById('itemNotes')?.value || '';
    KioskCart.add({ ...modalItem, unit_price: modalItem.price, modifiers: mods, notes, quantity: 1 });
    document.getElementById('cartFab').classList.remove('d-none');
    bootstrap.Modal.getInstance('#itemModal')?.hide();
    showToast(modalItem.item_name + ' added to cart', 'success');
}

// Override checkout to ask for customer info
KioskCart.checkout = function() {
    if (!this.items.length) { showToast('Cart is empty','warning'); return; }
    new bootstrap.Modal('#customerModal').show();
};

function placeOrderWithCustomer() {
    const name  = document.getElementById('custName').value.trim();
    const phone = document.getElementById('custPhone').value.trim();
    bootstrap.Modal.getInstance('#customerModal')?.hide();
    placeOrderNow({ name, phone });
}

async function placeOrderNow(customer) {
    bootstrap.Modal.getInstance('#customerModal')?.hide();
    try {
        const res = await api('/api/qr/place-order.php', 'POST', {
            table_token: TABLE_TOKEN,
            cart: KioskCart.items,
            customer: customer || {},
        });
        if (res.success) {
            KioskCart.items = [];
            KioskCart.render();
            window.location.href = `/order-status.php?order_id=\${res.order_id}&t=\${TABLE_TOKEN}`;
        } else {
            showToast(res.message || 'Order failed', 'error');
        }
    } catch(e) { showToast(e.message, 'error'); }
}
</script>
</body>
</html>
