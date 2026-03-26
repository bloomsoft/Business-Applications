<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('staff.manage');

$data   = json_decode(file_get_contents('php://input'), true) ?? [];
$userId = (int)($data['user_id'] ?? 0);
if (!$userId) jsonResponse(['error' => 'Missing user_id'], 422);

Database::query("UPDATE users SET is_active = 0 WHERE user_id = ? AND tenant_id = ?", [$userId, Auth::tenantId()]);
auditLog('staff_deactivated', 'users', $userId);
jsonResponse(['success' => true]);
