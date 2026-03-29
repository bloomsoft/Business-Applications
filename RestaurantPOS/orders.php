<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$page       = max(1, (int) get('page', 1));
$filters    = [
    'status'     => get('status'),
    'order_type' => get('type'),
    'date'       => get('date', date('Y-m-d')),
    'search'     => get('search'),
];
$orders     = OrderManager::listOrders($locationId, $filters, $page);
$pageTitle  = 'Orders';
$activeMenu = 'orders';
ob_start();
?>
<!-- Filters -->
<div class="card shadow-sm mb-3">
    <div class="card-body py-2">
        <form method="GET" class="row g-2 align-items-end">
            <div class="col-auto">
                <input type="date" name="date" class="form-control form-control-sm" value="<?= sanitize($filters['date']) ?>">
            </div>
            <div class="col-auto">
                <select name="status" class="form-select form-select-sm">
                    <option value="">All Statuses</option>
                    <?php foreach (['pending','confirmed','preparing','ready','served','completed','cancelled'] as $s): ?>
                    <option value="<?= $s ?>" <?= $filters['status'] === $s ? 'selected' : '' ?>><?= ucfirst($s) ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="col-auto">
                <select name="type" class="form-select form-select-sm">
                    <option value="">All Types</option>
                    <?php foreach (['dine-in','takeout','delivery','qr-order','kiosk'] as $t): ?>
                    <option value="<?= $t ?>" <?= $filters['order_type'] === $t ? 'selected' : '' ?>><?= ucfirst($t) ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="col-auto">
                <input type="text" name="search" class="form-control form-control-sm"
                       placeholder="Order # or customer..." value="<?= sanitize(get('search')) ?>">
            </div>
            <div class="col-auto">
                <button type="submit" class="btn btn-sm btn-primary">Filter</button>
                <a href="/orders.php" class="btn btn-sm btn-outline-secondary">Reset</a>
            </div>
            <div class="col-auto ms-auto">
                <a href="/pos.php" class="btn btn-sm btn-accent">
                    <i class="bi bi-plus me-1"></i>New Order
                </a>
            </div>
        </form>
    </div>
</div>

<!-- Orders Table -->
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover align-middle mb-0">
                <thead class="table-light">
                    <tr>
                        <th>Order #</th><th>Type</th><th>Table/Customer</th>
                        <th>Items</th><th>Total</th><th>Source</th>
                        <th>Status</th><th>Time</th><th></th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($orders['data'] as $order): ?>
                    <tr>
                        <td><span class="fw-600">#<?= sanitize($order['order_number']) ?></span></td>
                        <td><?= ucfirst(sanitize($order['order_type'])) ?></td>
                        <td>
                            <?php if ($order['table_number']): ?>
                            <span class="badge bg-light text-dark border">
                                <i class="bi bi-table me-1"></i>Table <?= sanitize($order['table_number']) ?>
                            </span>
                            <?php endif; ?>
                            <?php if ($order['customer_name'] && trim($order['customer_name'])): ?>
                            <div class="small text-muted mt-1"><i class="bi bi-person me-1"></i><?= sanitize($order['customer_name']) ?></div>
                            <?php endif; ?>
                        </td>
                        <td>
                            <?php
                            $itemCount = Database::fetchValue(
                                "SELECT COUNT(*) FROM order_items WHERE order_id = ?",
                                [$order['order_id']]
                            );
                            ?>
                            <span class="badge bg-secondary"><?= $itemCount ?></span>
                        </td>
                        <td class="fw-600"><?= money($order['total_amount']) ?></td>
                        <td><span class="badge bg-light text-dark border"><?= ucfirst(sanitize($order['source'])) ?></span></td>
                        <td><?= statusBadge($order['status']) ?></td>
                        <td><small class="text-muted"><?= timeAgo($order['created_at']) ?></small></td>
                        <td>
                            <button class="btn btn-sm btn-outline-primary"
                                    onclick="viewOrder(<?= $order['order_id'] ?>)">
                                <i class="bi bi-eye"></i>
                            </button>
                            <?php if (in_array($order['status'], ['pending','confirmed','preparing'])): ?>
                            <a href="/pos.php?order_id=<?= $order['order_id'] ?>"
                               class="btn btn-sm btn-outline-success">
                                <i class="bi bi-pencil"></i>
                            </a>
                            <?php endif; ?>
                            <button class="btn btn-sm btn-outline-secondary"
                                    onclick="printReceipt(<?= $order['order_id'] ?>)">
                                <i class="bi bi-printer"></i>
                            </button>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($orders['data'])): ?>
                    <tr><td colspan="9" class="text-center text-muted py-4">No orders found</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
    <!-- Pagination -->
    <?php if ($orders['last_page'] > 1): ?>
    <div class="card-footer bg-transparent">
        <nav><ul class="pagination pagination-sm mb-0 justify-content-end">
            <?php for ($i = 1; $i <= $orders['last_page']; $i++): ?>
            <li class="page-item <?= $i === $orders['current_page'] ? 'active' : '' ?>">
                <a class="page-link" href="?<?= http_build_query(array_merge($_GET, ['page' => $i])) ?>"><?= $i ?></a>
            </li>
            <?php endfor; ?>
        </ul></nav>
    </div>
    <?php endif; ?>
</div>

<!-- Order Detail Modal -->
<div class="modal fade" id="orderModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Order Details</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body" id="orderModalBody">Loading...</div>
            <div class="modal-footer">
                <div id="orderModalActions"></div>
                <button class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
async function viewOrder(orderId) {
    const modal = new bootstrap.Modal('#orderModal');
    document.getElementById('orderModalBody').innerHTML = '<div class="text-center py-4"><div class="spinner-border"></div></div>';
    modal.show();
    const order = await api('/api/orders/get.php?order_id=' + orderId);
    document.getElementById('orderModalBody').innerHTML = renderOrderDetail(order);
    document.getElementById('orderModalActions').innerHTML =
        `<button class="btn btn-outline-secondary btn-sm" onclick="printReceipt(${orderId})">
            <i class="bi bi-printer me-1"></i>Print Receipt
        </button>`;
}

function renderOrderDetail(order) {
    const items = (order.items||[]).map(i => `
        <tr>
            <td>${i.item_name}</td>
            <td>${i.quantity}</td>
            <td>${money(i.unit_price)}</td>
            <td>${money(i.line_total)}</td>
        </tr>`).join('');
    return `
        <div class="row mb-3">
            <div class="col-6">
                <strong>Order #:</strong> ${order.order_number}<br>
                <strong>Type:</strong> ${order.order_type}<br>
                <strong>Status:</strong> ${order.status}<br>
                <strong>Source:</strong> ${order.source}
            </div>
            <div class="col-6">
                <strong>Cashier:</strong> ${order.cashier_name||'—'}<br>
                <strong>Table:</strong> ${order.table_number||'—'}<br>
                <strong>Customer:</strong> ${order.customer_name||'—'}<br>
                <strong>Date:</strong> ${order.created_at}
            </div>
        </div>
        <table class="table table-sm">
            <thead class="table-light"><tr><th>Item</th><th>Qty</th><th>Price</th><th>Total</th></tr></thead>
            <tbody>${items}</tbody>
        </table>
        <div class="text-end">
            <div>Subtotal: ${money(order.subtotal)}</div>
            <div>Tax: ${money(order.tax_amount)}</div>
            ${order.discount_amount>0?`<div class="text-danger">Discount: -${money(order.discount_amount)}</div>`:''}
            ${order.tip_amount>0?`<div>Tip: ${money(order.tip_amount)}</div>`:''}
            <div class="fw-bold fs-5">Total: ${money(order.total_amount)}</div>
        </div>`;
}

async function printReceipt(orderId) {
    window.open('/api/orders/receipt.php?order_id='+orderId,'_blank','width=400,height=600');
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
