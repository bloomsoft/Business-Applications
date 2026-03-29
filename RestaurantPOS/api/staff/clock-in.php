<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
$data = json_decode(file_get_contents('php://input'), true) ?? [];
$clockId = StaffManager::clockIn((int)$data['user_id'], (int)$data['location_id']);
jsonResponse(['success' => true, 'clock_id' => $clockId]);
