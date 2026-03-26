<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$tableId = (int) ($_GET['table_id'] ?? 0);
if (!$tableId) die('Invalid table');

echo QRKioskManager::printTableQR($tableId);
