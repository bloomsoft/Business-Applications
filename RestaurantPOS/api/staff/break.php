<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data   = json_decode(file_get_contents('php://input'), true) ?? [];
$userId = (int)($data['user_id'] ?? 0);

// Check if currently on break
$clock = Database::fetchOne(
    "SELECT * FROM time_clocks WHERE user_id = ? AND clock_out IS NULL ORDER BY clock_in DESC",
    [$userId]
);
if (!$clock) jsonResponse(['error' => 'Not clocked in'], 422);

if ($clock['break_start'] && !$clock['break_end']) {
    StaffManager::endBreak($userId);
    jsonResponse(['success' => true, 'action' => 'break_ended']);
} else {
    StaffManager::startBreak($userId);
    jsonResponse(['success' => true, 'action' => 'break_started']);
}
