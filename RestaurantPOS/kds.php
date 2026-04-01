<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tickets    = OrderManager::getKDSTickets($locationId);
$pageTitle  = 'Kitchen Display';
$activeMenu = 'kds';
ob_start();
?>
<div class="d-flex justify-content-between align-items-center mb-3">
    <div class="d-flex gap-2">
        <span class="badge bg-success fs-6"><i class="bi bi-circle-fill me-1"></i>Live</span>
        <span class="text-muted" id="kdsLastRefresh">Last refresh: now</span>
    </div>
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-outline-secondary" onclick="location.reload()">
            <i class="bi bi-arrow-clockwise me-1"></i>Refresh
        </button>
        <button class="btn btn-sm btn-outline-secondary" onclick="toggleFullscreen()">
            <i class="bi bi-fullscreen me-1"></i>Fullscreen
        </button>
    </div>
</div>

<div class="row g-3" id="kdsGrid">
    <?php if (empty($tickets)): ?>
    <div class="col-12 text-center text-muted py-5">
        <i class="bi bi-check-circle fs-1 text-success"></i>
        <h4 class="mt-3">All clear — no pending orders</h4>
    </div>
    <?php else: ?>
    <?php foreach ($tickets as $ticket):
        $mins    = (int)$ticket['elapsed_minutes'];
        $urgency = $mins >= 20 ? 'urgent' : ($mins >= 10 ? 'warning' : '');
        $items   = $ticket['items'] ?? [];
    ?>
    <div class="col-md-4 col-lg-3" data-order-id="<?= $ticket['order_id'] ?>">
        <div class="kds-ticket card p-0 shadow <?= $urgency ?>">
            <div class="card-header d-flex justify-content-between align-items-center
                        <?= $urgency === 'urgent' ? 'bg-danger text-white' : ($urgency === 'warning' ? 'bg-warning' : 'bg-light') ?>">
                <div>
                    <h5 class="mb-0">#<?= sanitize($ticket['order_number']) ?></h5>
                    <small><?= ucfirst($ticket['order_type']) ?>
                        <?= $ticket['table_number'] ? ' · Table ' . sanitize($ticket['table_number']) : '' ?>
                    </small>
                </div>
                <div class="text-end">
                    <div class="kds-timer <?= $urgency ?>"><?= $mins ?>m</div>
                    <small><?= fmtDateTime($ticket['created_at'], 'g:i A') ?></small>
                </div>
            </div>
            <div class="card-body p-3">
                <ul class="list-unstyled mb-0">
                    <?php foreach ($items as $item): ?>
                    <li class="d-flex gap-2 mb-2">
                        <span class="badge bg-dark"><?= (int)$item['quantity'] ?>×</span>
                        <div>
                            <strong><?= sanitize($item['item_name']) ?></strong>
                            <?php if ($item['modifiers']): ?>
                            <div class="text-muted small"><?= sanitize($item['modifiers']) ?></div>
                            <?php endif; ?>
                            <?php if ($item['notes']): ?>
                            <div class="text-danger small fst-italic"><?= sanitize($item['notes']) ?></div>
                            <?php endif; ?>
                        </div>
                    </li>
                    <?php endforeach; ?>
                </ul>
                <?php if ($ticket['notes']): ?>
                <div class="alert alert-warning py-1 px-2 mt-2 mb-0 small">
                    <i class="bi bi-info-circle me-1"></i><?= sanitize($ticket['notes']) ?>
                </div>
                <?php endif; ?>
            </div>
            <div class="card-footer bg-transparent d-flex gap-2">
                <button class="btn btn-warning btn-sm flex-grow-1"
                        onclick="setOrderStatus(<?= $ticket['order_id'] ?>,'preparing',this)">
                    <i class="bi bi-fire me-1"></i>Preparing
                </button>
                <button class="btn btn-success btn-sm flex-grow-1"
                        onclick="bumpOrder(<?= $ticket['order_id'] ?>,this)">
                    <i class="bi bi-check2-all me-1"></i>Ready
                </button>
            </div>
        </div>
    </div>
    <?php endforeach; ?>
    <?php endif; ?>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
window.LOCATION_ID = {$locationId};

async function bumpOrder(orderId, btn) {
    btn.disabled = true;
    try {
        await api('/api/orders/update-status.php','POST',{order_id:orderId,status:'ready'});
        document.querySelector(`[data-order-id="\${orderId}"]`).remove();
        checkEmpty();
        showToast('Order marked ready!','success');
    } catch(e) { showToast(e.message,'error'); btn.disabled=false; }
}

async function setOrderStatus(orderId, status, btn) {
    btn.disabled = true;
    try {
        await api('/api/orders/update-status.php','POST',{order_id:orderId,status});
        showToast('Status updated','info');
    } catch(e) { showToast(e.message,'error'); }
    btn.disabled = false;
}

function checkEmpty() {
    if (!document.querySelectorAll('[data-order-id]').length) {
        document.getElementById('kdsGrid').innerHTML = `
            <div class="col-12 text-center text-muted py-5">
                <i class="bi bi-check-circle fs-1 text-success"></i>
                <h4 class="mt-3">All clear!</h4>
            </div>`;
    }
}

function toggleFullscreen() {
    document.fullscreenElement
        ? document.exitFullscreen()
        : document.documentElement.requestFullscreen();
}

function urgencyClass(mins) {
    return mins >= 20 ? 'urgent' : mins >= 10 ? 'warning' : '';
}
function headerClass(mins) {
    return mins >= 20 ? 'bg-danger text-white' : mins >= 10 ? 'bg-warning' : 'bg-light';
}

function renderKDSTickets(tickets) {
    const grid = document.getElementById('kdsGrid');
    if (!tickets.length) {
        grid.innerHTML = `<div class="col-12 text-center text-muted py-5">
            <i class="bi bi-check-circle fs-1 text-success"></i>
            <h4 class="mt-3">All clear — no pending orders</h4></div>`;
        return;
    }
    grid.innerHTML = tickets.map(t => {
        const mins = t.elapsed_minutes || 0;
        const urg  = urgencyClass(mins);
        const items = (t.items||[]).map(i => `
            <li class="d-flex gap-2 mb-2">
                <span class="badge bg-dark">\${i.quantity}×</span>
                <div><strong>\${i.item_name}</strong>
                    \${i.modifiers ? `<div class="text-muted small">\${i.modifiers}</div>` : ''}
                    \${i.notes    ? `<div class="text-danger small fst-italic">\${i.notes}</div>` : ''}
                </div>
            </li>`).join('');
        return `
        <div class="col-md-4 col-lg-3" data-order-id="\${t.order_id}">
            <div class="kds-ticket card p-0 shadow \${urg}">
                <div class="card-header d-flex justify-content-between align-items-center \${headerClass(mins)}">
                    <div>
                        <h5 class="mb-0">#\${t.order_number}</h5>
                        <small>\${t.order_type}\${t.table_number?' · Table '+t.table_number:''}</small>
                    </div>
                    <div class="text-end">
                        <div class="kds-timer \${urg}">\${mins}m</div>
                    </div>
                </div>
                <div class="card-body p-3"><ul class="list-unstyled mb-0">\${items}</ul>
                    \${t.notes?`<div class="alert alert-warning py-1 px-2 mt-2 mb-0 small">\${t.notes}</div>`:''}
                </div>
                <div class="card-footer bg-transparent d-flex gap-2">
                    <button class="btn btn-warning btn-sm flex-grow-1"
                            onclick="setOrderStatus(\${t.order_id},'preparing',this)">
                        <i class="bi bi-fire me-1"></i>Preparing</button>
                    <button class="btn btn-success btn-sm flex-grow-1"
                            onclick="bumpOrder(\${t.order_id},this)">
                        <i class="bi bi-check2-all me-1"></i>Ready</button>
                </div>
            </div>
        </div>`;
    }).join('');
}

// Auto-refresh every 30 seconds
setInterval(async () => {
    try {
        const res     = await fetch('/api/kds/tickets.php?location_id=' + window.LOCATION_ID);
        const tickets = await res.json();
        renderKDSTickets(tickets);
        document.getElementById('kdsLastRefresh').textContent = 'Last refresh: ' + new Date().toLocaleTimeString();
    } catch(e) { console.error('KDS refresh failed', e); }
}, 30000);
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
