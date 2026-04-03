<?php
require_once __DIR__ . '/core/bootstrap.php';

// Already logged in
if (Auth::check()) redirect('/dashboard.php');

$error = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $email    = trim(post('email'));
    $password = post('password');
    $slug     = trim(post('slug'));

    $tenant = Database::fetchOne(
        "SELECT * FROM tenants WHERE slug = ? AND is_active = 1",
        [$slug]
    );
    if (!$tenant) {
        $error = 'Organization not found.';
    } else {
        $result = Auth::login($email, $password, $tenant['tenant_id']);
        if ($result['success']) {
            redirect('/dashboard.php');
        } else {
            $error = $result['message'];
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Sign In — RestaurantPOS</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        :root { --pos-accent: #f97316; }
        body { background: linear-gradient(135deg, #1a1f36 0%, #2d3561 100%); min-height: 100vh; }
        .login-card { border: none; border-radius: 16px; box-shadow: 0 20px 60px rgba(0,0,0,.3); }
        .btn-accent { background: var(--pos-accent); border-color: var(--pos-accent); color: #fff; }
        .btn-accent:hover { background: #ea6c0c; color: #fff; }
        .text-accent { color: var(--pos-accent); }
    </style>
</head>
<body class="d-flex align-items-center justify-content-center py-5">
<div class="container">
    <div class="row justify-content-center">
        <div class="col-sm-10 col-md-7 col-lg-5 col-xl-4">
            <!-- Logo -->
            <div class="text-center mb-4">
                <i class="bi bi-shop text-white" style="font-size:48px"></i>
                <h3 class="text-white fw-bold mt-2">RestaurantPOS</h3>
                <p class="text-white-50">Cloud Restaurant Management</p>
            </div>

            <div class="card login-card">
                <div class="card-body p-4">
                    <h5 class="fw-bold mb-4">Sign in to your account</h5>

                    <?php if ($error): ?>
                    <div class="alert alert-danger d-flex align-items-center gap-2">
                        <i class="bi bi-exclamation-circle"></i><?= sanitize($error) ?>
                    </div>
                    <?php endif; ?>

                    <form method="POST" novalidate>
                        <div class="mb-3">
                            <label class="form-label">Organization</label>
                            <div class="input-group">
                                <span class="input-group-text"><i class="bi bi-building"></i></span>
                                <input type="text" name="slug" class="form-control"
                                       placeholder="your-restaurant-slug"
                                       value="<?= sanitize(post('slug')) ?>" required>
                            </div>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Email</label>
                            <div class="input-group">
                                <span class="input-group-text"><i class="bi bi-envelope"></i></span>
                                <input type="email" name="email" class="form-control"
                                       value="<?= sanitize(post('email')) ?>" required autofocus>
                            </div>
                        </div>
                        <div class="mb-4">
                            <label class="form-label">Password</label>
                            <div class="input-group">
                                <span class="input-group-text"><i class="bi bi-lock"></i></span>
                                <input type="password" name="password" class="form-control" required>
                                <button type="button" class="btn btn-outline-secondary"
                                        onclick="this.previousElementSibling.type = this.previousElementSibling.type === 'password' ? 'text' : 'password'">
                                    <i class="bi bi-eye"></i>
                                </button>
                            </div>
                        </div>
                        <div class="d-grid">
                            <button type="submit" class="btn btn-accent btn-lg">
                                <i class="bi bi-box-arrow-in-right me-2"></i>Sign In
                            </button>
                        </div>
                    </form>

                    <hr class="my-4">

                    <!-- Demo Credentials -->
                    <div class="alert alert-info py-2 px-3 small mb-3">
                        <strong><i class="bi bi-info-circle me-1"></i>Demo Login:</strong><br>
                        Organization: <code>demo</code><br>
                        Email: <code>admin@demo.com</code><br>
                        Password: <code>password123</code>
                    </div>

                    <!-- PIN Login -->
                    <div class="text-center">
                        <a href="/pin-login.php" class="text-muted small">
                            <i class="bi bi-grid-3x3-gap me-1"></i>Quick PIN Login
                        </a>
                    </div>
                </div>
            </div>

            <div class="text-center mt-3">
                <small class="text-white-50">
                    &copy; <?= date('Y') ?> RestaurantPOS SaaS. All rights reserved.
                </small>
            </div>
        </div>
    </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
