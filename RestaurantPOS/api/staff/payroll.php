<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
$data = json_decode(file_get_contents('php://input'), true) ?? [];
$results = StaffManager::generatePayroll(
    (int)$data['location_id'],
    $data['period_start'],
    $data['period_end']
);
jsonResponse($results);
