/**
 * RestaurantPOS SaaS — Core JavaScript
 */

// ── CSRF Token helper ─────────────────────────────────────────
const csrf = () => document.querySelector('meta[name="csrf-token"]')?.content ?? '';

// ── API helper ────────────────────────────────────────────────
async function api(url, method = 'GET', data = null) {
    const opts = {
        method,
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    };
    if (data) opts.body = JSON.stringify(data);
    const res  = await fetch(url, opts);
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || json.message || 'Request failed');
    return json;
}

// ── Toast notification ────────────────────────────────────────
function showToast(message, type = 'success') {
    const id      = 'toast_' + Date.now();
    const colors  = { success:'bg-success', error:'bg-danger', warning:'bg-warning text-dark', info:'bg-info text-dark' };
    const html    = `
        <div id="${id}" class="toast align-items-center text-white ${colors[type]||colors.info} border-0"
             role="alert" data-bs-delay="3500">
            <div class="d-flex">
                <div class="toast-body fw-500">${message}</div>
                <button type="button" class="btn-close btn-close-white me-2 m-auto"
                        data-bs-dismiss="toast"></button>
            </div>
        </div>`;
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'toast-container position-fixed bottom-0 end-0 p-3';
        container.style.zIndex = 9999;
        document.body.appendChild(container);
    }
    container.insertAdjacentHTML('beforeend', html);
    new bootstrap.Toast(document.getElementById(id)).show();
}

// ── Confirm dialog ────────────────────────────────────────────
function confirmAction(message, callback) {
    if (window.confirm(message)) callback();
}

// ── Money formatter ───────────────────────────────────────────
function money(amount, symbol = '$') {
    return symbol + parseFloat(amount || 0).toFixed(2);
}

// ── POS Cart State ────────────────────────────────────────────
const POSCart = {
    orderId: null,
    items: [],
    taxRate: 0.08,

    async addItem(itemId, variantId = null, modifiers = [], qty = 1, notes = '') {
        if (!this.orderId) {
            showToast('No active order. Start a new order first.', 'warning');
            return;
        }
        try {
            const res = await api('/api/orders/add-item.php', 'POST', {
                order_id: this.orderId, item_id: itemId,
                variant_id: variantId, modifiers, quantity: qty, notes
            });
            this.refresh();
            showToast('Item added', 'success');
        } catch(e) { showToast(e.message, 'error'); }
    },

    async removeItem(orderItemId) {
        try {
            await api('/api/orders/remove-item.php', 'POST', { order_item_id: orderItemId });
            this.refresh();
        } catch(e) { showToast(e.message, 'error'); }
    },

    async refresh() {
        if (!this.orderId) return;
        try {
            const order = await api(`/api/orders/get.php?order_id=${this.orderId}`);
            this.renderCart(order);
        } catch(e) { console.error(e); }
    },

    renderCart(order) {
        const list = document.getElementById('orderItemsList');
        if (!list) return;

        if (!order.items?.length) {
            list.innerHTML = '<div class="text-center text-muted py-4"><i class="bi bi-cart fs-1"></i><p class="mt-2">Cart is empty</p></div>';
        } else {
            list.innerHTML = order.items.map(item => `
                <div class="d-flex align-items-center gap-2 py-2 border-bottom">
                    <div class="flex-grow-1">
                        <div class="fw-600">${item.item_name}</div>
                        ${item.modifiers?.map(m => `<small class="text-muted">+${m.modifier_name}</small>`).join(', ') || ''}
                        ${item.notes ? `<small class="text-muted d-block">${item.notes}</small>` : ''}
                    </div>
                    <div class="text-end">
                        <div>${money(item.line_total)}</div>
                        <small class="text-muted">×${item.quantity}</small>
                    </div>
                    <button class="btn btn-sm btn-outline-danger"
                            onclick="POSCart.removeItem(${item.order_item_id})">
                        <i class="bi bi-trash"></i>
                    </button>
                </div>
            `).join('');
        }

        // Update totals
        document.getElementById('cartSubtotal')?.querySelector('span')?.textContent !== undefined &&
            (document.getElementById('cartSubtotal').textContent = money(order.subtotal));
        document.getElementById('cartTax') &&
            (document.getElementById('cartTax').textContent = money(order.tax_amount));
        document.getElementById('cartTotal') &&
            (document.getElementById('cartTotal').textContent = money(order.total_amount));
    },

    async processPayment(method, amount, tip = 0) {
        if (!this.orderId) return;
        try {
            const res = await api('/api/payments/process.php', 'POST', {
                order_id: this.orderId, method, amount, tip
            });
            showToast('Payment successful! Change: ' + money(res.change || 0), 'success');
            setTimeout(() => { window.location.href = '/pos.php'; }, 1500);
        } catch(e) { showToast(e.message, 'error'); }
    }
};

// ── KDS Auto-Refresh ──────────────────────────────────────────
function initKDSRefresh(intervalSec = 30) {
    setInterval(() => {
        fetch('/api/kds/tickets.php?location_id=' + (window.LOCATION_ID || ''))
            .then(r => r.json())
            .then(data => renderKDSTickets(data))
            .catch(console.error);
    }, intervalSec * 1000);
}

function renderKDSTickets(tickets) {
    const grid = document.getElementById('kdsGrid');
    if (!grid) return;
    if (!tickets.length) {
        grid.innerHTML = '<div class="col-12 text-center text-muted py-5"><i class="bi bi-check-circle fs-1 text-success"></i><h4 class="mt-3">All orders ready</h4></div>';
        return;
    }
    grid.innerHTML = tickets.map(t => {
        const mins    = parseInt(t.elapsed_minutes) || 0;
        const urgency = mins >= 20 ? 'urgent' : mins >= 10 ? 'warning' : '';
        return `
            <div class="col-md-4 col-lg-3">
                <div class="kds-ticket card p-3 ${urgency}">
                    <div class="d-flex justify-content-between align-items-start mb-2">
                        <div>
                            <h5 class="mb-0">#${t.order_number}</h5>
                            <small class="text-muted">${t.order_type}${t.table_number ? ' · Table ' + t.table_number : ''}</small>
                        </div>
                        <span class="kds-timer ${urgency}">${mins}m</span>
                    </div>
                    <ul class="list-unstyled mb-3" id="kds_items_${t.order_id}">
                        <!-- items loaded separately -->
                    </ul>
                    ${t.notes ? `<small class="text-muted fst-italic">${t.notes}</small>` : ''}
                    <button class="btn btn-success btn-sm mt-2 w-100"
                            onclick="bumpKDSOrder(${t.order_id})">
                        <i class="bi bi-check2 me-1"></i>Ready
                    </button>
                </div>
            </div>`;
    }).join('');
}

async function bumpKDSOrder(orderId) {
    try {
        await api('/api/kds/bump.php', 'POST', { order_id: orderId });
        showToast('Order marked ready', 'success');
        document.querySelector(`[data-order-id="${orderId}"]`)?.remove();
    } catch(e) { showToast(e.message, 'error'); }
}

// ── Floor Plan Drag ───────────────────────────────────────────
function initFloorPlan() {
    document.querySelectorAll('.table-shape').forEach(el => {
        let startX, startY, isDragging = false;

        el.addEventListener('mousedown', e => {
            isDragging = true;
            startX = e.clientX - el.offsetLeft;
            startY = e.clientY - el.offsetTop;
            el.style.zIndex = 100;
        });

        document.addEventListener('mousemove', e => {
            if (!isDragging) return;
            el.style.left = (e.clientX - startX) + 'px';
            el.style.top  = (e.clientY - startY) + 'px';
        });

        document.addEventListener('mouseup', () => {
            if (!isDragging) return;
            isDragging = false;
            el.style.zIndex = '';
            api('/api/tables/update-position.php', 'POST', {
                table_id: el.dataset.tableId,
                x: parseInt(el.style.left),
                y: parseInt(el.style.top)
            });
        });
    });
}

// ── Chart helpers ─────────────────────────────────────────────
function renderLineChart(canvasId, labels, datasets) {
    const ctx = document.getElementById(canvasId)?.getContext('2d');
    if (!ctx) return;
    return new Chart(ctx, {
        type: 'line',
        data: { labels, datasets: datasets.map(d => ({ tension: .4, fill: true, ...d })) },
        options: { responsive: true, plugins: { legend: { position: 'top' } } }
    });
}

function renderBarChart(canvasId, labels, datasets, horizontal = false) {
    const ctx = document.getElementById(canvasId)?.getContext('2d');
    if (!ctx) return;
    return new Chart(ctx, {
        type: horizontal ? 'bar' : 'bar',
        data: { labels, datasets },
        options: {
            indexAxis: horizontal ? 'y' : 'x',
            responsive: true,
            plugins: { legend: { display: false } }
        }
    });
}

function renderDoughnut(canvasId, labels, data, colors) {
    const ctx = document.getElementById(canvasId)?.getContext('2d');
    if (!ctx) return;
    return new Chart(ctx, {
        type: 'doughnut',
        data: { labels, datasets: [{ data, backgroundColor: colors }] },
        options: { responsive: true, plugins: { legend: { position: 'right' } } }
    });
}

// ── QR/Kiosk Cart ─────────────────────────────────────────────
const KioskCart = {
    items: [],

    add(item) {
        const existing = this.items.find(i => i.item_id === item.item_id);
        if (existing) { existing.quantity++; existing.line_total += item.unit_price; }
        else          { this.items.push({ ...item, quantity: 1, line_total: item.unit_price }); }
        this.render();
    },

    remove(itemId) {
        this.items = this.items.filter(i => i.item_id !== itemId);
        this.render();
    },

    total() { return this.items.reduce((s, i) => s + i.line_total, 0); },

    render() {
        const panel = document.getElementById('kioskCart');
        if (!panel) return;
        const count = this.items.reduce((s, i) => s + i.quantity, 0);
        document.getElementById('kioskCartCount') &&
            (document.getElementById('kioskCartCount').textContent = count);

        panel.innerHTML = this.items.length ? `
            <div class="p-3">
                ${this.items.map(i => `
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <div><strong>${i.item_name}</strong> ×${i.quantity}</div>
                        <div>${money(i.line_total)}
                            <button class="btn btn-sm btn-outline-danger ms-1"
                                    onclick="KioskCart.remove(${i.item_id})">
                                <i class="bi bi-x"></i>
                            </button>
                        </div>
                    </div>`).join('')}
                <hr>
                <div class="d-flex justify-content-between fw-bold fs-5">
                    <span>Total</span><span>${money(this.total())}</span>
                </div>
                <button class="btn btn-accent btn-lg w-100 mt-3" onclick="KioskCart.checkout()">
                    Proceed to Payment
                </button>
            </div>` : '<div class="text-center text-muted p-4">Your cart is empty</div>';
    },

    async checkout() {
        if (!this.items.length) { showToast('Cart is empty', 'warning'); return; }
        const tableToken = new URLSearchParams(location.search).get('t') || '';
        try {
            const res = await api('/api/qr/place-order.php', 'POST', {
                table_token: tableToken,
                cart: this.items,
            });
            if (res.success) {
                this.items = [];
                this.render();
                window.location.href = `/order-status.php?order_id=${res.order_id}`;
            }
        } catch(e) { showToast(e.message, 'error'); }
    }
};
