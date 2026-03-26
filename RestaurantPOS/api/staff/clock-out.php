<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
$data   = json_decode(file_get_contents('php://input'), true) ?? [];
$result = StaffManager::clockOut((int)$data['user_id']);
jsonResponse(['success' => true, ...$result]);
