<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('pos.access');

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$categories = Database::fetchAll(
    "SELECT * FROM menu_categories WHERE tenant_id = ? AND is_active = 1 AND parent_id IS NULL ORDER BY sort_order",
    [$tenantId]
);
$tables     = TableManager::getTables($locationId);
$pageTitle  = 'Point of Sale';
$activeMenu = 'pos';

ob_start();
?>
<div class="pos-screen d-flex gap-0" style="margin:-1.5rem; height:calc(100vh - 56px)">

    <!-- LEFT: Menu Panel -->
    <div class="flex-grow-1 d-flex flex-column overflow-hidden">

        <!-- Category Tabs -->
        <div class="bg-white border-bottom px-3 py-2 d-flex gap-2 overflow-auto flex-shrink-0">
            <button class="btn btn-accent btn-sm rounded-pill px-3 category-btn active" data-category="">
                All
            </button>
            <?php foreach ($categories as $cat): ?>
            <button class="btn btn-outline-secondary btn-sm rounded-pill px-3 category-btn"
                    data-category="<?= $cat['category_id'] ?>">
                <?= sanitize($cat['category_name']) ?>
            </button>
            <?php endforeach; ?>
        </div>

        <!-- Search + Table Select -->
        <div class="bg-white border-bottom px-3 py-2 d-flex gap-2 align-items-center">
            <div class="input-group input-group-sm" style="max-width:280px">
                <span class="input-group-text"><i class="bi bi-search"></i></span>
                <input type="text" class="form-control" id="menuSearch" placeholder="Search items...">
            </div>
            <select class="form-select form-select-sm" style="max-width:180px" id="tableSelect">
                <option value="">— Takeout / No Table —</option>
                <?php foreach ($tables as $t):
                    if ($t['status'] === 'available'): ?>
                <option value="<?= $t['table_id'] ?>">
                    Table <?= sanitize($t['table_number']) ?> (<?= $t['capacity'] ?> seats)
                </option>
                <?php endif; endforeach; ?>
            </select>
            <select class="form-select form-select-sm" style="max-width:140px" id="orderTypeSelect">
                <option value="dine-in">Dine-In</option>
                <option value="takeout">Takeout</option>
                <option value="delivery">Delivery</option>
            </select>
            <button class="btn btn-sm btn-primary" onclick="newOrder()">
                <i class="bi bi-plus-lg me-1"></i>New Order
            </button>
            <div class="ms-auto d-flex gap-2">
                <a href="/kds.php" class="btn btn-sm btn-outline-secondary">
                    <i class="bi bi-display me-1"></i>KDS
                </a>
                <button class="btn btn-sm btn-outline-secondary" onclick="showFloorPlan()">
                    <i class="bi bi-layout-text-window me-1"></i>Floor Plan
                </button>
            </div>
        </div>

        <!-- Active Order Indicator -->
        <div id="activeOrderBar" class="d-none bg-warning-subtle border-bottom px-3 py-2 d-flex align-items-center gap-3">
            <span><i class="bi bi-receipt me-2"></i>Order: <strong id="activeOrderNumber">—</strong></span>
            <span class="text-muted small" id="activeOrderType"></span>
            <button class="btn btn-sm btn-outline-secondary ms-auto" onclick="clearOrder()">
                <i class="bi bi-x me-1"></i>Clear
            </button>
        </div>

        <!-- Menu Grid -->
        <div class="flex-grow-1 overflow-auto p-3">
            <div class="menu-grid" id="menuGrid">
                <?php
                $items = Database::fetchAll(
                    "SELECT * FROM menu_items WHERE tenant_id = ? AND is_available = 1 ORDER BY sort_order, item_name",
                    [$tenantId]
                );
                foreach ($items as $item):
                ?>
                <div class="menu-item-card card p-0 shadow-sm"
                     data-item-id="<?= $item['item_id'] ?>"
                     data-category="<?= $item['category_id'] ?>"
                     data-name="<?= strtolower(sanitize($item['item_name'])) ?>"
                     onclick="selectItem(<?= $item['item_id'] ?>, '<?= addslashes($item['item_name']) ?>', <?= $item['price'] ?>)">
                    <?php if ($item['image_url']): ?>
                    <img src="<?= sanitize($item['image_url']) ?>"
                         class="card-img-top" style="height:80px;object-fit:cover" alt="">
                    <?php else: ?>
                    <div class="d-flex align-items-center justify-content-center bg-light" style="height:60px">
                        <i class="bi bi-cup-hot fs-3 text-secondary"></i>
                    </div>
                    <?php endif; ?>
                    <div class="card-body p-2">
                        <div class="fw-600 small lh-sm"><?= sanitize($item['item_name']) ?></div>
                        <div class="price small mt-1"><?= money($item['price']) ?></div>
                    </div>
                </div>
                <?php endforeach; ?>
            </div>
        </div>
    </div>

    <!-- RIGHT: Order Panel -->
    <div class="order-panel">
        <div class="p-3 border-bottom">
            <h6 class="mb-0 fw-bold"><i class="bi bi-cart3 me-2"></i>Current Order</h6>
        </div>

        <!-- Customer Search -->
        <div class="p-2 border-bottom">
            <div class="input-group input-group-sm">
                <span class="input-group-text"><i class="bi bi-person-search"></i></span>
                <input type="text" class="form-control" id="customerSearch"
                       placeholder="Search customer (phone/email)..."
                       oninput="searchCustomer(this.value)">
            </div>
            <div id="customerResults" class="list-group mt-1 position-absolute z-1" style="width:310px"></div>
            <div id="selectedCustomer" class="d-none mt-1">
                <span class="badge bg-success" id="customerBadge"></span>
                <button class="btn btn-sm btn-link text-danger p-0 ms-1" onclick="clearCustomer()">×</button>
            </div>
        </div>

        <!-- Cart Items -->
        <div class="order-items-list p-2" id="orderItemsList">
            <div class="text-center text-muted py-5">
                <i class="bi bi-cart fs-1"></i>
                <p class="mt-2">Start a new order</p>
            </div>
        </div>

        <!-- Totals -->
        <div class="p-3 border-top">
            <div class="d-flex justify-content-between mb-1">
                <span class="text-muted">Subtotal</span>
                <span id="cartSubtotal">$0.00</span>
            </div>
            <div class="d-flex justify-content-between mb-1">
                <span class="text-muted">Tax</span>
                <span id="cartTax">$0.00</span>
            </div>
            <div class="d-flex justify-content-between mb-1">
                <span class="text-muted small">
                    <a href="#" data-bs-toggle="modal" data-bs-target="#discountModal">+ Discount</a>
                </span>
                <span id="cartDiscount" class="text-danger"></span>
            </div>
            <hr class="my-2">
            <div class="d-flex justify-content-between fw-bold fs-5 mb-3">
                <span>Total</span>
                <span id="cartTotal">$0.00</span>
            </div>

            <!-- Send to Kitchen -->
            <div class="d-grid mb-2">
                <button class="btn btn-warning btn-sm" onclick="sendToKitchen()">
                    <i class="bi bi-fire me-1"></i>Send to Kitchen
                </button>
            </div>

            <!-- Payment Buttons -->
            <div class="d-grid gap-2">
                <button class="btn btn-accent btn-lg" onclick="openPaymentModal('card')">
                    <i class="bi bi-credit-card me-2"></i>Card
                </button>
                <div class="row g-2">
                    <div class="col-6">
                        <button class="btn btn-outline-secondary w-100" onclick="openPaymentModal('cash')">
                            <i class="bi bi-cash me-1"></i>Cash
                        </button>
                    </div>
                    <div class="col-6">
                        <button class="btn btn-outline-secondary w-100" onclick="openPaymentModal('split')">
                            <i class="bi bi-distribute-vertical me-1"></i>Split
                        </button>
                    </div>
                </div>
                <button class="btn btn-outline-danger btn-sm" onclick="voidOrder()">
                    <i class="bi bi-x-circle me-1"></i>Void Order
                </button>
            </div>
        </div>
    </div>
</div>

<!-- Modifier Modal -->
<div class="modal fade" id="modifierModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="modifierModalTitle">Customize Item</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body" id="modifierModalBody"></div>
            <div class="modal-footer">
                <div class="d-flex align-items-center gap-2 me-auto">
                    <button class="btn btn-outline-secondary btn-sm" onclick="changeQty(-1)">−</button>
                    <span id="itemQty" class="fw-bold px-2">1</span>
                    <button class="btn btn-outline-secondary btn-sm" onclick="changeQty(1)">+</button>
                </div>
                <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button class="btn btn-accent" onclick="addToOrder()">Add to Order</button>
            </div>
        </div>
    </div>
</div>

<!-- Payment Modal -->
<div class="modal fade" id="paymentModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Process Payment</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="row g-3">
                    <div class="col-md-6">
                        <label class="form-label">Amount Due</label>
                        <div class="fs-3 fw-bold text-accent" id="payDue">$0.00</div>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Tip Amount</label>
                        <div class="d-flex gap-2 mb-2">
                            <button class="btn btn-outline-secondary btn-sm" onclick="setTipPercent(15)">15%</button>
                            <button class="btn btn-outline-secondary btn-sm" onclick="setTipPercent(18)">18%</button>
                            <button class="btn btn-outline-secondary btn-sm" onclick="setTipPercent(20)">20%</button>
                            <button class="btn btn-outline-secondary btn-sm" onclick="setTipPercent(0)">No Tip</button>
                        </div>
                        <input type="number" class="form-control" id="tipAmount" placeholder="Custom tip" step="0.01" min="0">
                    </div>
                    <div class="col-12" id="cashInputGroup" style="display:none">
                        <label class="form-label">Cash Tendered</label>
                        <input type="number" class="form-control form-control-lg" id="cashTendered"
                               placeholder="Enter amount" step="0.01" oninput="calcChange()">
                        <div class="mt-2 fs-5">Change: <strong id="changeAmount" class="text-success">$0.00</strong></div>
                    </div>
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button class="btn btn-accent btn-lg" id="confirmPayBtn" onclick="confirmPayment()">
                    <i class="bi bi-check-lg me-2"></i>Confirm Payment
                </button>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
window.LOCATION_ID = {$locationId};
window.TENANT_ID   = {$tenantId};

let currentItem       = null;
let currentItemQty    = 1;
let currentPayMethod  = 'card';
let currentCustomerId = null;

// Category filter
document.querySelectorAll('.category-btn').forEach(btn => {
    btn.addEventListener('click', function() {
        document.querySelectorAll('.category-btn').forEach(b => b.classList.remove('active','btn-accent'));
        this.classList.add('active','btn-accent');
        this.classList.remove('btn-outline-secondary');
        filterMenu(this.dataset.category);
    });
});

// Search filter
document.getElementById('menuSearch').addEventListener('input', e => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('.menu-item-card').forEach(card => {
        card.style.display = card.dataset.name.includes(q) ? '' : 'none';
    });
});

function filterMenu(catId) {
    document.querySelectorAll('.menu-item-card').forEach(card => {
        card.style.display = (!catId || card.dataset.category === catId) ? '' : 'none';
    });
}

async function newOrder() {
    const tableId   = document.getElementById('tableSelect').value;
    const orderType = document.getElementById('orderTypeSelect').value;
    try {
        const res = await api('/api/orders/create.php', 'POST', {
            location_id: window.LOCATION_ID,
            table_id: tableId || null,
            order_type: orderType,
            customer_id: currentCustomerId,
        });
        POSCart.orderId = res.order_id;
        document.getElementById('activeOrderBar').classList.remove('d-none');
        document.getElementById('activeOrderNumber').textContent = res.order_number;
        document.getElementById('activeOrderType').textContent = orderType;
        showToast('New order started: #' + res.order_number, 'success');
    } catch(e) { showToast(e.message, 'error'); }
}

async function selectItem(itemId, itemName, price) {
    if (!POSCart.orderId) {
        showToast('Start a new order first', 'warning'); return;
    }
    // Load modifiers
    const mods = await api('/api/menu/modifiers.php?item_id=' + itemId);
    currentItem = { itemId, itemName, price };
    currentItemQty = 1;

    document.getElementById('modifierModalTitle').textContent = itemName + ' — ' + money(price);
    const body = document.getElementById('modifierModalBody');

    if (!mods.length) {
        // No modifiers — add directly
        POSCart.addItem(itemId);
        return;
    }

    body.innerHTML = mods.map(grp => `
        <div class="mb-3">
            <label class="fw-600">\${grp.group_name}
                \${grp.is_required ? '<span class="badge bg-danger ms-1">Required</span>' : ''}
            </label>
            \${grp.options.map(opt => `
                <div class="form-check">
                    <input class="form-check-input modifier-check" type="\${grp.selection_type==='single'?'radio':'checkbox'}"
                           name="grp_\${grp.group_id}" value="\${opt.modifier_id}"
                           id="mod_\${opt.modifier_id}">
                    <label class="form-check-label" for="mod_\${opt.modifier_id}">
                        \${opt.modifier_name}
                        \${opt.price_add > 0 ? '<span class="text-accent">+' + money(opt.price_add) + '</span>' : ''}
                    </label>
                </div>`).join('')}
        </div>`).join('');

    new bootstrap.Modal('#modifierModal').show();
}

function changeQty(delta) {
    currentItemQty = Math.max(1, currentItemQty + delta);
    document.getElementById('itemQty').textContent = currentItemQty;
}

async function addToOrder() {
    const checked  = [...document.querySelectorAll('.modifier-check:checked')].map(el => parseInt(el.value));
    const notesEl  = document.getElementById('itemNotes');
    await POSCart.addItem(currentItem.itemId, null, checked, currentItemQty, notesEl?.value || '');
    bootstrap.Modal.getInstance('#modifierModal')?.hide();
}

function openPaymentModal(method) {
    if (!POSCart.orderId) { showToast('No active order', 'warning'); return; }
    currentPayMethod = method;
    const total = document.getElementById('cartTotal').textContent;
    document.getElementById('payDue').textContent = total;
    document.getElementById('cashInputGroup').style.display = method === 'cash' ? '' : 'none';
    new bootstrap.Modal('#paymentModal').show();
}

function setTipPercent(pct) {
    const total = parseFloat(document.getElementById('payDue').textContent.replace('$','')) || 0;
    document.getElementById('tipAmount').value = (total * pct / 100).toFixed(2);
}

function calcChange() {
    const due     = parseFloat(document.getElementById('payDue').textContent.replace('$','')) || 0;
    const tip     = parseFloat(document.getElementById('tipAmount').value) || 0;
    const cash    = parseFloat(document.getElementById('cashTendered').value) || 0;
    const change  = Math.max(0, cash - due - tip);
    document.getElementById('changeAmount').textContent = money(change);
}

async function confirmPayment() {
    const total = parseFloat(document.getElementById('payDue').textContent.replace('$',''));
    const tip   = parseFloat(document.getElementById('tipAmount').value) || 0;
    bootstrap.Modal.getInstance('#paymentModal')?.hide();
    await POSCart.processPayment(currentPayMethod, total + tip, tip);
}

async function searchCustomer(q) {
    if (q.length < 2) { document.getElementById('customerResults').innerHTML = ''; return; }
    const res = await api('/api/customers/search.php?q=' + encodeURIComponent(q) + '&tenant_id=' + window.TENANT_ID);
    const div = document.getElementById('customerResults');
    div.innerHTML = res.map(c => `
        <a href="#" class="list-group-item list-group-item-action py-1"
           onclick="selectCustomer(\${c.customer_id},'\${c.first_name} \${c.last_name}');return false">
            \${c.first_name} \${c.last_name} — \${c.phone||c.email||''}
            <span class="badge bg-warning text-dark ms-1">\${c.loyalty_points} pts</span>
        </a>`).join('');
}

function selectCustomer(id, name) {
    currentCustomerId = id;
    document.getElementById('selectedCustomer').classList.remove('d-none');
    document.getElementById('customerBadge').textContent = name;
    document.getElementById('customerResults').innerHTML = '';
    document.getElementById('customerSearch').value = '';
}

function clearCustomer() {
    currentCustomerId = null;
    document.getElementById('selectedCustomer').classList.add('d-none');
}

function clearOrder() {
    POSCart.orderId = null;
    document.getElementById('activeOrderBar').classList.add('d-none');
    document.getElementById('orderItemsList').innerHTML = '<div class="text-center text-muted py-5"><i class="bi bi-cart fs-1"></i><p class="mt-2">Start a new order</p></div>';
    ['cartSubtotal','cartTax','cartTotal'].forEach(id => document.getElementById(id).textContent = '$0.00');
}

async function sendToKitchen() {
    if (!POSCart.orderId) { showToast('No active order', 'warning'); return; }
    try {
        await api('/api/orders/update-status.php','POST',{order_id:POSCart.orderId,status:'confirmed'});
        showToast('Order sent to kitchen!','success');
    } catch(e) { showToast(e.message,'error'); }
}

async function voidOrder() {
    if (!POSCart.orderId) return;
    if (!confirm('Void this order?')) return;
    await api('/api/orders/update-status.php','POST',{order_id:POSCart.orderId,status:'cancelled'});
    clearOrder();
    showToast('Order voided','warning');
}

function showFloorPlan() { window.open('/floor-plan.php','_blank','width=900,height=600'); }
</script>
JS;

require_once __DIR__ . '/templates/layout.php';
