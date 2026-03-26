<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$user      = Database::fetchOne("SELECT * FROM users WHERE user_id = ?", [Auth::id()]);
$pageTitle = 'My Profile';
$activeMenu= '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');

    if ($action === 'update_profile') {
        Database::query(
            "UPDATE users SET first_name = ?, last_name = ?, phone = ?, email = ? WHERE user_id = ?",
            [post('first_name'), post('last_name'), post('phone'), post('email'), Auth::id()]
        );
        $_SESSION['user']['first_name'] = post('first_name');
        $_SESSION['user']['last_name']  = post('last_name');
        flash('success', 'Profile updated');
        redirect('/profile.php');
    }

    if ($action === 'change_password') {
        if (!password_verify(post('current_password'), $user['password_hash'])) {
            flash('error', 'Current password is incorrect');
            redirect('/profile.php');
        }
        if (post('new_password') !== post('confirm_password')) {
            flash('error', 'Passwords do not match');
            redirect('/profile.php');
        }
        Database::query(
            "UPDATE users SET password_hash = ? WHERE user_id = ?",
            [Auth::hashPassword(post('new_password')), Auth::id()]
        );
        flash('success', 'Password changed');
        redirect('/profile.php');
    }
}

ob_start();
?>
<div class="row g-4 justify-content-center">
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-person me-2"></i>Profile Information
            </div>
            <div class="card-body">
                <form method="POST">
                    <input type="hidden" name="action" value="update_profile">
                    <div class="row g-3">
                        <div class="col-6"><label class="form-label">First Name</label>
                            <input type="text" name="first_name" class="form-control" value="<?= sanitize($user['first_name']) ?>" required></div>
                        <div class="col-6"><label class="form-label">Last Name</label>
                            <input type="text" name="last_name" class="form-control" value="<?= sanitize($user['last_name']) ?>" required></div>
                        <div class="col-12"><label class="form-label">Email</label>
                            <input type="email" name="email" class="form-control" value="<?= sanitize($user['email']) ?>" required></div>
                        <div class="col-12"><label class="form-label">Phone</label>
                            <input type="tel" name="phone" class="form-control" value="<?= sanitize($user['phone'] ?? '') ?>"></div>
                    </div>
                    <button type="submit" class="btn btn-primary mt-3">Update Profile</button>
                </form>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600 bg-transparent">
                <i class="bi bi-lock me-2"></i>Change Password
            </div>
            <div class="card-body">
                <form method="POST">
                    <input type="hidden" name="action" value="change_password">
                    <div class="mb-3"><label class="form-label">Current Password</label>
                        <input type="password" name="current_password" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label">New Password</label>
                        <input type="password" name="new_password" class="form-control" required minlength="8"></div>
                    <div class="mb-3"><label class="form-label">Confirm New Password</label>
                        <input type="password" name="confirm_password" class="form-control" required minlength="8"></div>
                    <button type="submit" class="btn btn-warning">Change Password</button>
                </form>
            </div>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
require_once __DIR__ . '/templates/layout.php';
