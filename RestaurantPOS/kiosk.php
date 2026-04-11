<?php
/**
 * Self-Service Kiosk Mode — full-screen ordering terminal
 */
require_once __DIR__ . '/core/bootstrap.php';

$locationId = (int) get('location', 0);
if (!$locationId) {
    Auth::requireAuth();
    $locationId = Auth::locationId();
}
$location   = LocationManager::get($locationId);
$tenantId   = $location['tenant_id'] ?? 0;
$tenant     = Database::fetchOne("SELECT * FROM tenants WHERE tenant_id = ?", [$tenantId]);
$menu       = QRKioskManager::getPublicMenu($tenantId);
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Self-Order Kiosk — <?= sanitize($tenant['company_name'] ?? '') ?></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link rel="stylesheet" href="/public/css/app.css">
    <style>
        body { background: #1a1f36; color: #fff; min-height: 100vh; user-select: none; }
        .kiosk-header { background: rgba(255,255,255,.05); border-bottom: 1px solid rgba(255,255,255,.1); }
        .kiosk-cat-btn { background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.15);
            color: #fff; border-radius: 50px; padding: 8px 20px; white-space: nowrap; cursor: pointer; transition: all .2s; }
        .kiosk-cat-btn.active { background: var(--pos-accent); border-color: var(--pos-accent); }
        .kiosk-item { background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.1);
            border-radius: 12px; cursor: pointer; transition: all .2s; overflow: hidden; }
        .kiosk-item:hover { background: rgba(249,115,22,.2); border-color: var(--pos-accent); transform: scale(1.02); }
        .kiosk-item .price { color: var(--pos-accent); font-weight: 700; font-size: 1.1rem; }
        .cart-panel { background: rgba(255,255,255,.05); border-left: 1px solid rgba(255,255,255,.1);
            width: 380px; flex-shrink: 0; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
        .cart-items { flex: 1; overflow-y: auto; }
        .idle-overlay { position: fixed; inset: 0; background: #1a1f36;
            z-index: 9999; display: flex; align-items: center; justify-content: center; cursor: pointer; }
        .idle-overlay h1 { font-size: 4rem; font-weight: 900; }
    </style>
</head>
<body>

<!-- Idle Screen -->
<div class="idle-overlay" id="idleOverlay" style="cursor:pointer">
    <div class="text-center">
        <?php if ($tenant['logo_url']): ?>
        <img src="<?= sanitize($tenant['logo_url']) ?>" height="80" class="mb-4" alt="">
        <?php endif; ?>
        <h1 class="text-white"><?= sanitize($tenant['company_name']) ?></h1>
        <p class="text-white-50 fs-4">Tap anywhere to start ordering</p>
        <div class="mt-4 fs-1">
            <i class="bi bi-hand-index-thumb text-accent"></i>
        </div>
    </div>
</div>

<div id="kioskApp" style="height:100vh; display:none; flex-direction:row">

    <!-- Menu Panel -->
    <div class="flex-grow-1 d-flex flex-column overflow-hidden">
        <!-- Header -->
        <div class="kiosk-header d-flex align-items-center gap-3 px-4 py-3">
            <div>
                <h4 class="mb-0 fw-bold text-white"><?= sanitize($tenant['company_name']) ?></h4>
                <small class="text-white-50"><?= sanitize($location['location_name']) ?></small>
            </div>
            <div class="ms-auto d-flex align-items-center gap-3">
                <select class="form-select form-select-sm bg-dark text-white border-secondary"
                        style="width:160px" id="kioskOrderType">
                    <option value="dine-in">Dine-In</option>
                    <option value="takeout">Takeout</option>
                </select>
                <span class="text-white-50" id="kioskClock" style="font-size:1.1rem"></span>
            </div>
        </div>

        <!-- Categories -->
        <div class="px-4 py-3 d-flex gap-2 overflow-auto" style="border-bottom:1px solid rgba(255,255,255,.1)">
            <button class="kiosk-cat-btn active" data-category="">All</button>
            <?php foreach ($menu as $cat): ?>
            <button class="kiosk-cat-btn" data-category="<?= $cat['category_id'] ?>">
                <?= sanitize($cat['category_name']) ?>
            </button>
            <?php endforeach; ?>
        </div>

        <!-- Menu Grid -->
        <div class="flex-grow-1 overflow-auto p-4">
            <div class="row g-3" id="kioskMenuGrid">
                <?php foreach ($menu as $cat): foreach ($cat['items'] as $item): ?>
                <div class="col-6 col-lg-4 col-xl-3 kiosk-item-wrap" data-category="<?= $cat['category_id'] ?>">
                    <div class="kiosk-item"
                         onclick="kioskSelectItem(<?= htmlspecialchars(json_encode($item)) ?>)">
                        <?php if ($item['image_url']): ?>
                        <img src="<?= sanitize($item['image_url']) ?>"
                             style="width:100%;height:140px;object-fit:cover" alt="">
                        <?php else: ?>
                        <div class="d-flex align-items-center justify-content-center"
                             style="height:100px;background:rgba(255,255,255,.04)">
                            <i class="bi bi-cup-hot fs-1 text-secondary"></i>
                        </div>
                        <?php endif; ?>
                        <div class="p-3">
                            <div class="fw-600 fs-6 text-white"><?= sanitize($item['item_name']) ?></div>
                            <?php if ($item['description']): ?>
                            <div class="text-white-50 small mt-1" style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">
                                <?= sanitize($item['description']) ?>
                            </div>
                            <?php endif; ?>
                            <div class="price mt-2"><?= money($item['price']) ?></div>
                        </div>
                    </div>
                </div>
                <?php endforeach; endforeach; ?>
            </div>
        </div>
    </div>

    <!-- Cart Panel -->
    <div class="cart-panel">
        <div class="p-4 border-bottom" style="border-color:rgba(255,255,255,.1) !important">
            <h5 class="mb-0 text-white fw-bold">
                <i class="bi bi-cart3 me-2"></i>Your Order
                <span class="badge bg-accent ms-2" id="kioskCartCount">0</span>
            </h5>
        </div>
        <div class="cart-items p-3" id="kioskCart">
            <div class="text-center text-white-50 py-5">
                <i class="bi bi-cart fs-1"></i>
                <p class="mt-2">Select items to add</p>
            </div>
        </div>
        <div class="p-4 border-top" style="border-color:rgba(255,255,255,.1) !important">
            <div class="d-flex justify-content-between text-white-50 mb-1">
                <span>Subtotal</span><span id="kioskSubtotal">$0.00</span>
            </div>
            <div class="d-flex justify-content-between text-white-50 mb-3">
                <span>Tax</span><span id="kioskTax">$0.00</span>
            </div>
            <div class="d-flex justify-content-between text-white fw-bold fs-4 mb-4">
                <span>Total</span><span id="kioskTotal">$0.00</span>
            </div>
            <button class="btn btn-accent btn-lg w-100 fw-bold"
                    id="kioskPlaceBtn" onclick="kioskCheckout()" disabled>
                <i class="bi bi-bag-check me-2"></i>Place Order
            </button>
            <button class="btn btn-outline-secondary btn-sm w-100 mt-2" onclick="kioskClearCart()">
                <i class="bi bi-x me-1"></i>Clear Cart
            </button>
        </div>
    </div>
</div>

<!-- Item Modal -->
<div class="modal fade" id="kioskItemModal" tabindex="-1" data-bs-theme="dark">
    <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content bg-dark text-white">
            <div class="modal-header border-0">
                <h5 class="modal-title" id="kioskItemTitle"></h5>
                <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body" id="kioskItemBody"></div>
            <div class="modal-footer border-0">
                <div class="d-flex align-items-center gap-3 me-auto">
                    <button class="btn btn-outline-secondary btn-lg" onclick="kioskChangeQty(-1)">−</button>
                    <span id="kioskModalQty" class="fw-bold fs-4">1</span>
                    <button class="btn btn-outline-secondary btn-lg" onclick="kioskChangeQty(1)">+</button>
                </div>
                <button class="btn btn-accent btn-lg px-4" onclick="kioskAddToCart()">
                    Add — <span id="kioskItemTotal"></span>
                </button>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="/public/js/app.js"></script>
<script>
const KIOSK_LOCATION_ID = <?= (int)$locationId ?>;
const KIOSK_TENANT_ID   = <?= (int)$tenantId ?>;
const TAX_RATE          = <?= floatval($location['tax_rate'] ?? 0.08) ?>;
let kioskItem = null, kioskQty = 1;
let idleTimer = null;

function wakeUp() {
    document.getElementById('idleOverlay').style.display = 'none';
    document.getElementById('kioskApp').style.display   = 'flex';
    resetIdleTimer();
}

function goIdle() {
    document.getElementById('kioskApp').style.display   = 'none';
    document.getElementById('idleOverlay').style.display = 'flex';
    kioskClearCart();
}

function resetIdleTimer() {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(goIdle, 120000); // 2 min idle
}

// Script is at bottom of body — DOM is already ready, attach directly
document.getElementById('idleOverlay').addEventListener('click', wakeUp);
document.getElementById('idleOverlay').addEventListener('touchstart', function(e) {
    e.preventDefault();
    wakeUp();
}, { passive: false });
document.addEventListener('touchstart', resetIdleTimer, { passive: true });
document.addEventListener('mousemove',  resetIdleTimer);

// Clock
setInterval(() => {
    const el = document.getElementById('kioskClock');
    if (el) el.textContent = new Date().toLocaleTimeString();
}, 1000);

// Category filter
document.querySelectorAll('.kiosk-cat-btn').forEach(btn => {
    btn.addEventListener('click', function() {
        document.querySelectorAll('.kiosk-cat-btn').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        const cat = this.dataset.category;
        document.querySelectorAll('.kiosk-item-wrap').forEach(el => {
            el.style.display = (!cat || el.dataset.category === cat) ? '' : 'none';
        });
    });
});

const kioskCartItems = [];

function kioskSelectItem(item) {
    kioskItem = item;
    kioskQty  = 1;
    document.getElementById('kioskItemTitle').textContent = item.item_name;
    document.getElementById('kioskModalQty').textContent = 1;
    document.getElementById('kioskItemTotal').textContent = money(item.price);

    let html = '';
    if (item.description) html += `<p class="text-white-50">\${item.description}</p>`;
    if (item.modifier_groups?.length) {
        html += item.modifier_groups.map(grp => `
            <div class="mb-3">
                <div class="fw-600 text-white">\${grp.group_name}
                    \${grp.is_required ? '<span class="badge bg-danger">Required</span>' : ''}
                </div>
                \${grp.options.map(opt => `
                    <div class="form-check">
                        <input class="form-check-input kmod" type="\${grp.selection_type==='single'?'radio':'checkbox'}"
                               name="grp_\${grp.group_id}" value="\${opt.modifier_id}" data-price="\${opt.price_add}">
                        <label class="form-check-label text-white">
                            \${opt.modifier_name}
                            \${opt.price_add>0?`<span class="text-accent">+\${money(opt.price_add)}</span>`:''}
                        </label>
                    </div>`).join('')}
            </div>`).join('');
    }
    document.getElementById('kioskItemBody').innerHTML = html;
    document.querySelectorAll('.kmod').forEach(el => el.addEventListener('change', updateKioskItemTotal));
    new bootstrap.Modal('#kioskItemModal').show();
}

function kioskChangeQty(d) {
    kioskQty = Math.max(1, kioskQty + d);
    document.getElementById('kioskModalQty').textContent = kioskQty;
    updateKioskItemTotal();
}

function updateKioskItemTotal() {
    let extra = 0;
    document.querySelectorAll('.kmod:checked').forEach(el => extra += parseFloat(el.dataset.price)||0);
    document.getElementById('kioskItemTotal').textContent = money((kioskItem.price + extra) * kioskQty);
}

function kioskAddToCart() {
    const mods = [...document.querySelectorAll('.kmod:checked')].map(el => parseInt(el.value));
    const existing = kioskCartItems.find(i => i.item_id === kioskItem.item_id);
    if (existing) existing.quantity += kioskQty;
    else kioskCartItems.push({ ...kioskItem, quantity: kioskQty, modifiers: mods, line_total: kioskItem.price * kioskQty });
    renderKioskCart();
    bootstrap.Modal.getInstance('#kioskItemModal')?.hide();
}

function renderKioskCart() {
    const total    = kioskCartItems.reduce((s,i) => s + i.price * i.quantity, 0);
    const tax      = total * TAX_RATE;
    const count    = kioskCartItems.reduce((s,i) => s+i.quantity, 0);

    const el = id => document.getElementById(id);
    if (el('kioskCartCount')) el('kioskCartCount').textContent = count;
    if (el('kioskSubtotal'))  el('kioskSubtotal').textContent  = money(total);
    if (el('kioskTax'))       el('kioskTax').textContent       = money(tax);
    if (el('kioskTotal'))     el('kioskTotal').textContent     = money(total + tax);
    if (el('kioskPlaceBtn'))  el('kioskPlaceBtn').disabled     = !kioskCartItems.length;

    document.getElementById('kioskCart').innerHTML = kioskCartItems.length ?
        kioskCartItems.map(i => `
            <div class="d-flex justify-content-between align-items-center mb-3">
                <div>
                    <div class="text-white fw-600">\${i.item_name}</div>
                    <div class="d-flex align-items-center gap-2 mt-1">
                        <button class="btn btn-sm btn-outline-secondary py-0"
                                onclick="kioskQtyChange(\${i.item_id}, -1)">−</button>
                        <span class="text-white-50">×\${i.quantity}</span>
                        <button class="btn btn-sm btn-outline-secondary py-0"
                                onclick="kioskQtyChange(\${i.item_id}, 1)">+</button>
                    </div>
                </div>
                <div class="text-end">
                    <div class="text-white">\${money(i.price * i.quantity)}</div>
                    <button class="btn btn-sm text-danger p-0" onclick="kioskRemove(\${i.item_id})">
                        <i class="bi bi-trash"></i>
                    </button>
                </div>
            </div>`).join('') :
        '<div class="text-center text-white-50 py-5"><i class="bi bi-cart fs-1"></i><p class="mt-2">Empty</p></div>';
}

function kioskQtyChange(itemId, delta) {
    const item = kioskCartItems.find(i => i.item_id === itemId);
    if (item) {
        item.quantity = Math.max(0, item.quantity + delta);
        if (item.quantity === 0) kioskRemove(itemId);
        else renderKioskCart();
    }
}

function kioskRemove(itemId) {
    const idx = kioskCartItems.findIndex(i => i.item_id === itemId);
    if (idx !== -1) kioskCartItems.splice(idx, 1);
    renderKioskCart();
}

function kioskClearCart() {
    kioskCartItems.length = 0;
    renderKioskCart();
}

async function kioskCheckout() {
    if (!kioskCartItems.length) return;
    const orderType = document.getElementById('kioskOrderType').value;
    try {
        const res = await api('/api/qr/kiosk-order.php', 'POST', {
            location_id: KIOSK_LOCATION_ID,
            tenant_id:   KIOSK_TENANT_ID,
            order_type: orderType,
            cart: kioskCartItems,
        });
        if (res.success) {
            kioskClearCart();
            // Show confirmation screen
            document.getElementById('kioskApp').innerHTML = `
                <div style="width:100%;display:flex;align-items:center;justify-content:center;height:100vh">
                    <div class="text-center text-white">
                        <i class="bi bi-check-circle-fill text-success" style="font-size:6rem"></i>
                        <h2 class="mt-4">Order Placed!</h2>
                        <p class="text-white-50 fs-5">Order #\${res.order.order_number}</p>
                        <p class="text-white-50">We'll have it ready shortly.</p>
                        <div class="mt-4 fw-bold fs-3">\${money(res.order.total_amount)}</div>
                        <button class="btn btn-accent btn-lg mt-5" onclick="location.reload()">
                            New Order
                        </button>
                    </div>
                </div>`;
        }
    } catch(e) { showToast(e.message, 'error'); }
}
</script>
</body>
</html>
