<?php
require_once __DIR__ . '/core/bootstrap.php';

if (Auth::check()) redirect('/dashboard.php');

$error = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $pin        = post('pin');
    $locationId = (int) post('location_id');

    if (!$pin || !$locationId) {
        $error = 'Please enter a PIN and select a location.';
    } else {
        $result = Auth::loginByPin($pin, $locationId);
        if ($result['success']) {
            redirect('/pos.php');
        } else {
            $error = $result['message'];
        }
    }
}

$locations = Database::fetchAll(
    "SELECT l.location_id, l.location_name, t.company_name
     FROM locations l
     JOIN tenants t ON t.tenant_id = l.tenant_id
     WHERE l.is_active = 1
     ORDER BY t.company_name, l.location_name"
);
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Quick PIN Login — RestaurantPOS</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        :root { --pos-accent: #f97316; }
        body { background: linear-gradient(135deg, #1a1f36 0%, #2d3561 100%); min-height: 100vh; }
        .pin-pad button { width: 70px; height: 70px; font-size: 28px; font-weight: 700; border-radius: 50%; }
        .btn-accent { background: var(--pos-accent); border-color: var(--pos-accent); color: #fff; }
        .btn-accent:hover { background: #ea6c0c; color: #fff; }
        .pin-display { letter-spacing: 16px; font-size: 36px; font-weight: 800; }
    </style>
</head>
<body class="d-flex align-items-center justify-content-center py-5">
<div class="container">
    <div class="row justify-content-center">
        <div class="col-sm-8 col-md-6 col-lg-4">
            <div class="text-center mb-4">
                <i class="bi bi-shop text-white" style="font-size:48px"></i>
                <h4 class="text-white mt-2">Quick PIN Login</h4>
            </div>
            <div class="card border-0" style="border-radius:16px">
                <div class="card-body p-4">
                    <?php if ($error): ?>
                    <div class="alert alert-danger py-2"><?= sanitize($error) ?></div>
                    <?php endif; ?>

                    <form method="POST" id="pinForm">
                        <div class="mb-3">
                            <select name="location_id" class="form-select" required>
                                <option value="">Select Location...</option>
                                <?php foreach ($locations as $loc): ?>
                                <option value="<?= $loc['location_id'] ?>">
                                    <?= sanitize($loc['company_name'] . ' — ' . $loc['location_name']) ?>
                                </option>
                                <?php endforeach; ?>
                            </select>
                        </div>

                        <div class="text-center mb-4">
                            <input type="password" name="pin" id="pinInput"
                                   class="form-control form-control-lg text-center pin-display border-0 bg-light"
                                   maxlength="6" readonly>
                        </div>

                        <div class="pin-pad text-center">
                            <div class="d-flex justify-content-center gap-2 mb-2">
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(1)">1</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(2)">2</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(3)">3</button>
                            </div>
                            <div class="d-flex justify-content-center gap-2 mb-2">
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(4)">4</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(5)">5</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(6)">6</button>
                            </div>
                            <div class="d-flex justify-content-center gap-2 mb-2">
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(7)">7</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(8)">8</button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(9)">9</button>
                            </div>
                            <div class="d-flex justify-content-center gap-2">
                                <button type="button" class="btn btn-outline-danger" onclick="clearPin()">
                                    <i class="bi bi-x-lg"></i>
                                </button>
                                <button type="button" class="btn btn-outline-secondary" onclick="addDigit(0)">0</button>
                                <button type="submit" class="btn btn-accent">
                                    <i class="bi bi-check-lg"></i>
                                </button>
                            </div>
                        </div>
                    </form>

                    <div class="text-center mt-3">
                        <a href="/login.php" class="text-muted small">
                            <i class="bi bi-envelope me-1"></i>Sign in with email
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
<script>
const pinInput = document.getElementById('pinInput');
function addDigit(d) { if (pinInput.value.length < 6) pinInput.value += d; }
function clearPin()   { pinInput.value = ''; }
</script>
</body>
</html>
