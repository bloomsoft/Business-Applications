<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$tables     = TableManager::getTables($locationId);
$pageTitle  = 'QR Code & Kiosk Manager';
$activeMenu = 'qr';
ob_start();
?>
<div class="row g-3 mb-4">
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-body text-center">
                <i class="bi bi-qr-code fs-1 text-primary"></i>
                <h5 class="mt-2">Table QR Ordering</h5>
                <p class="text-muted small">Generate QR codes for each table. Customers scan and order directly from their phone.</p>
                <a href="#tablesTab" data-bs-toggle="tab" class="btn btn-primary btn-sm">View Tables</a>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-body text-center">
                <i class="bi bi-display fs-1 text-success"></i>
                <h5 class="mt-2">Self-Service Kiosk</h5>
                <p class="text-muted small">Launch a full-screen kiosk ordering mode for self-checkout counters.</p>
                <a href="/kiosk.php?location=<?= $locationId ?>" target="_blank" class="btn btn-success btn-sm">
                    <i class="bi bi-fullscreen me-1"></i>Launch Kiosk
                </a>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card shadow-sm">
            <div class="card-body text-center">
                <i class="bi bi-star fs-1 text-warning"></i>
                <h5 class="mt-2">Feedback QR</h5>
                <p class="text-muted small">Place feedback QR codes on tables to collect ratings after meals.</p>
                <a href="/feedback.php" target="_blank" class="btn btn-warning btn-sm text-white">
                    <i class="bi bi-chat-square-text me-1"></i>Feedback Page
                </a>
            </div>
        </div>
    </div>
</div>

<!-- Table QR Grid -->
<div class="card shadow-sm">
    <div class="card-header bg-transparent fw-600 border-0 pt-3 d-flex justify-content-between">
        <span><i class="bi bi-qr-code me-2"></i>Table QR Codes</span>
        <div class="d-flex gap-2">
            <button class="btn btn-sm btn-outline-secondary" onclick="printAll()">
                <i class="bi bi-printer me-1"></i>Print All
            </button>
        </div>
    </div>
    <div class="card-body">
        <div class="row g-3">
            <?php foreach ($tables as $table): ?>
            <div class="col-6 col-md-4 col-lg-3 col-xl-2">
                <div class="card text-center border shadow-sm h-100">
                    <div class="card-body p-2">
                        <div class="fw-bold mb-1">Table <?= sanitize($table['table_number']) ?></div>
                        <div class="mb-2"><?= statusBadge($table['status']) ?></div>
                        <?php
                        $qrUrl = QRKioskManager::getTableQRUrl($table['table_id']);
                        $orderUrl = APP_URL . '/order.php?t=' . urlencode($table['qr_code_token'] ?? '');
                        ?>
                        <img src="<?= $qrUrl ?>" alt="QR" class="img-fluid" style="max-width:120px">
                        <div class="mt-2 d-flex flex-column gap-1">
                            <a href="<?= $orderUrl ?>" target="_blank"
                               class="btn btn-sm btn-outline-primary btn-sm">
                                <i class="bi bi-box-arrow-up-right me-1"></i>Test
                            </a>
                            <button class="btn btn-sm btn-outline-secondary"
                                    onclick="printQR(<?= $table['table_id'] ?>)">
                                <i class="bi bi-printer me-1"></i>Print
                            </button>
                            <button class="btn btn-sm btn-outline-danger"
                                    onclick="regenerateQR(<?= $table['table_id'] ?>)">
                                <i class="bi bi-arrow-repeat me-1"></i>Regenerate
                            </button>
                        </div>
                    </div>
                </div>
            </div>
            <?php endforeach; ?>
        </div>
    </div>
</div>

<!-- Kiosk Settings Card -->
<div class="card shadow-sm mt-4">
    <div class="card-header bg-transparent fw-600 border-0 pt-3">
        <i class="bi bi-gear me-2"></i>Kiosk Settings
    </div>
    <div class="card-body">
        <div class="row g-3">
            <div class="col-md-4">
                <label class="form-label">Kiosk Mode</label>
                <select class="form-select" id="kioskMode">
                    <option value="dine-in">Dine-In</option>
                    <option value="takeout">Takeout / Walk-in</option>
                    <option value="both">Both Options</option>
                </select>
            </div>
            <div class="col-md-4">
                <label class="form-label">Payment Methods</label>
                <div class="form-check"><input class="form-check-input" type="checkbox" checked> Cash</div>
                <div class="form-check"><input class="form-check-input" type="checkbox" checked> Card</div>
                <div class="form-check"><input class="form-check-input" type="checkbox" checked> QR Pay</div>
            </div>
            <div class="col-md-4">
                <label class="form-label">Idle Timeout (seconds)</label>
                <input type="number" class="form-control" value="120" min="30" max="600">
            </div>
        </div>
        <div class="mt-3">
            <a href="/kiosk.php?location=<?= $locationId ?>" target="_blank"
               class="btn btn-success">
                <i class="bi bi-display me-2"></i>Launch Kiosk Mode
            </a>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
async function printQR(tableId) {
    const win = window.open('/api/qr/print.php?table_id=' + tableId, '_blank');
    win.onload = () => { win.print(); };
}

async function regenerateQR(tableId) {
    if (!confirm('Regenerate QR code? The old QR will no longer work.')) return;
    try {
        await api('/api/qr/regenerate.php','POST',{table_id:tableId});
        showToast('QR code regenerated','success');
        setTimeout(()=>location.reload(),1000);
    } catch(e) { showToast(e.message,'error'); }
}

function printAll() {
    window.print();
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
