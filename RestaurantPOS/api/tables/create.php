<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = $_POST ?: (json_decode(file_get_contents('php://input'), true) ?? []);
$data['location_id'] = Auth::locationId();

if (empty($data['table_number'])) {
    flash('error', 'Table number is required');
    header('Location: /floor-plan.php');
    exit;
}

TableManager::createTable($data);
flash('success', 'Table added successfully');
header('Location: /floor-plan.php');
