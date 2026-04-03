<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><?= sanitize($pageTitle ?? 'Dashboard') ?> — RestaurantPOS</title>
    <!-- Bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
    <link rel="stylesheet" href="/public/css/app.css">
</head>
<body class="bg-light">

<!-- Sidebar -->
<div class="d-flex">
<nav id="sidebar" class="d-flex flex-column flex-shrink-0 p-3 text-white sidebar-nav">
    <a href="/dashboard.php" class="d-flex align-items-center mb-3 mb-md-0 me-md-auto text-white text-decoration-none">
        <i class="bi bi-shop fs-4 me-2"></i>
        <span class="fs-5 fw-bold">RestaurantPOS</span>
    </a>
    <hr>

    <!-- Location Switcher -->
    <div class="mb-3">
        <select class="form-select form-select-sm bg-dark text-white border-secondary" id="locationSwitcher"
                onchange="switchLocation(this.value)">
            <?php
            $locations = LocationManager::getAll(Auth::tenantId());
            foreach ($locations as $loc):
            ?>
            <option value="<?= $loc['location_id'] ?>"
                <?= $loc['location_id'] == Auth::locationId() ? 'selected' : '' ?>>
                <?= sanitize($loc['location_name']) ?>
            </option>
            <?php endforeach; ?>
        </select>
    </div>

    <ul class="nav nav-pills flex-column mb-auto">
        <li class="nav-item">
            <a href="/dashboard.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'dashboard' ? 'active' : '' ?>">
                <i class="bi bi-speedometer2 me-2"></i>Dashboard
            </a>
        </li>
        <li>
            <a href="/pos.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'pos' ? 'active' : '' ?>">
                <i class="bi bi-cash-register me-2"></i>Point of Sale
            </a>
        </li>
        <li>
            <a href="/orders.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'orders' ? 'active' : '' ?>">
                <i class="bi bi-receipt me-2"></i>Orders
            </a>
        </li>
        <li>
            <a href="/kds.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'kds' ? 'active' : '' ?>">
                <i class="bi bi-display me-2"></i>Kitchen Display
            </a>
        </li>
        <li><hr class="text-secondary"></li>
        <li>
            <a href="/menu.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'menu' ? 'active' : '' ?>">
                <i class="bi bi-menu-button-wide me-2"></i>Menu Manager
            </a>
        </li>
        <li>
            <a href="/inventory.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'inventory' ? 'active' : '' ?>">
                <i class="bi bi-box-seam me-2"></i>Inventory
            </a>
        </li>
        <li>
            <a href="/customers.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'crm' ? 'active' : '' ?>">
                <i class="bi bi-people me-2"></i>Customers / CRM
            </a>
        </li>
        <li>
            <a href="/delivery.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'delivery' ? 'active' : '' ?>">
                <i class="bi bi-truck me-2"></i>Delivery
            </a>
        </li>
        <li>
            <a href="/reservations.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'reservations' ? 'active' : '' ?>">
                <i class="bi bi-calendar-check me-2"></i>Reservations
            </a>
        </li>
        <li><hr class="text-secondary"></li>
        <li>
            <a href="/staff.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'staff' ? 'active' : '' ?>">
                <i class="bi bi-person-badge me-2"></i>Staff
            </a>
        </li>
        <li>
            <a href="/financial.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'financial' ? 'active' : '' ?>">
                <i class="bi bi-cash-coin me-2"></i>Financial
            </a>
        </li>
        <li>
            <a href="/analytics.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'analytics' ? 'active' : '' ?>">
                <i class="bi bi-bar-chart-line me-2"></i>Analytics
            </a>
        </li>
        <li>
            <a href="/qr-manager.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'qr' ? 'active' : '' ?>">
                <i class="bi bi-qr-code me-2"></i>QR & Kiosk
            </a>
        </li>
        <li><hr class="text-secondary"></li>
        <li>
            <a href="/locations.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'locations' ? 'active' : '' ?>">
                <i class="bi bi-geo-alt me-2"></i>Locations
            </a>
        </li>
        <li>
            <a href="/settings.php" class="nav-link text-white <?= ($activeMenu ?? '') === 'settings' ? 'active' : '' ?>">
                <i class="bi bi-gear me-2"></i>Settings
            </a>
        </li>
    </ul>
    <hr>
    <div class="dropdown">
        <a href="#" class="d-flex align-items-center text-white text-decoration-none dropdown-toggle"
           data-bs-toggle="dropdown">
            <i class="bi bi-person-circle me-2 fs-5"></i>
            <strong><?= sanitize((Auth::user()['first_name'] ?? '') . ' ' . (Auth::user()['last_name'] ?? '')) ?></strong>
        </a>
        <ul class="dropdown-menu dropdown-menu-dark text-small shadow">
            <li><a class="dropdown-item" href="/profile.php"><i class="bi bi-person me-2"></i>Profile</a></li>
            <li><hr class="dropdown-divider"></li>
            <li><a class="dropdown-item" href="/logout.php"><i class="bi bi-box-arrow-right me-2"></i>Sign out</a></li>
        </ul>
    </div>
</nav>

<!-- Main Content -->
<main class="flex-grow-1">
    <!-- Top bar -->
    <div class="topbar d-flex align-items-center justify-content-between px-4 py-2 bg-white border-bottom shadow-sm">
        <div class="d-flex align-items-center gap-2">
            <button class="btn btn-sm btn-outline-secondary d-lg-none" id="sidebarToggle">
                <i class="bi bi-list"></i>
            </button>
            <h6 class="mb-0 text-muted"><?= sanitize($pageTitle ?? '') ?></h6>
        </div>
        <div class="d-flex align-items-center gap-3">
            <!-- Notifications -->
            <div class="dropdown">
                <a href="#" class="position-relative text-secondary" data-bs-toggle="dropdown">
                    <i class="bi bi-bell fs-5"></i>
                    <?php
                    $unreadCount = (int) Database::fetchValue(
                        "SELECT COUNT(*) FROM notifications WHERE tenant_id = ? AND is_read = 0",
                        [Auth::tenantId()]
                    );
                    if ($unreadCount > 0):
                    ?>
                    <span class="position-absolute top-0 start-100 translate-middle badge rounded-pill bg-danger" style="font-size:10px">
                        <?= $unreadCount ?>
                    </span>
                    <?php endif; ?>
                </a>
                <ul class="dropdown-menu dropdown-menu-end shadow" style="min-width:300px">
                    <li><h6 class="dropdown-header">Notifications</h6></li>
                    <?php
                    $notifs = Database::fetchAll(
                        "SELECT * FROM notifications WHERE tenant_id = ? ORDER BY created_at DESC LIMIT 5",
                        [Auth::tenantId()]
                    );
                    foreach ($notifs as $n):
                    ?>
                    <li>
                        <a class="dropdown-item <?= !$n['is_read'] ? 'fw-bold' : '' ?>" href="#">
                            <small class="text-muted d-block"><?= timeAgo($n['created_at']) ?></small>
                            <?= sanitize($n['title']) ?>
                        </a>
                    </li>
                    <?php endforeach; ?>
                    <?php if (empty($notifs)): ?>
                    <li><span class="dropdown-item text-muted">No notifications</span></li>
                    <?php endif; ?>
                </ul>
            </div>
            <span class="text-muted small" id="liveClock"></span>
        </div>
    </div>

    <!-- Flash Message -->
    <?php $flash = getFlash(); if ($flash): ?>
    <div class="alert <?= alertClass($flash['type']) ?> alert-dismissible m-3 mb-0" role="alert">
        <?= sanitize($flash['message']) ?>
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    </div>
    <?php endif; ?>

    <!-- Page Body -->
    <div class="p-4">
        <?= $content ?? '' ?>
    </div>
</main>
</div><!-- end d-flex -->

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="/public/js/app.js"></script>
<script>
    // Live clock
    function updateClock() {
        const now = new Date();
        document.getElementById('liveClock').textContent = now.toLocaleTimeString();
    }
    setInterval(updateClock, 1000); updateClock();

    // Location switcher
    function switchLocation(locationId) {
        fetch('/api/switch-location.php', {
            method: 'POST',
            headers: {'Content-Type':'application/json'},
            body: JSON.stringify({location_id: locationId})
        }).then(() => location.reload());
    }

    // Sidebar toggle (mobile)
    document.getElementById('sidebarToggle')?.addEventListener('click', () => {
        document.getElementById('sidebar').classList.toggle('sidebar-open');
    });
</script>
<?= $scripts ?? '' ?>
</body>
</html>
