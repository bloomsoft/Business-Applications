<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('staff.manage');

$data = $_POST ?: (json_decode(file_get_contents('php://input'), true) ?? []);
$data['tenant_id']   = Auth::tenantId();
$data['location_id'] = Auth::locationId();

if (empty($data['first_name']) || empty($data['last_name']) || empty($data['email']) || empty($data['password'])) {
    flash('error', 'All required fields must be filled');
    header('Location: /staff.php');
    exit;
}

// Check email uniqueness within tenant
$exists = Database::fetchValue(
    "SELECT COUNT(*) FROM users WHERE email = ? AND tenant_id = ?",
    [$data['email'], $data['tenant_id']]
);
if ($exists) {
    flash('error', 'Email already in use');
    header('Location: /staff.php');
    exit;
}

StaffManager::createStaff($data);
flash('success', 'Staff member added');
header('Location: /staff.php');
