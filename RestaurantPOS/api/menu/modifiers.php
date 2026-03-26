<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$itemId = (int) ($_GET['item_id'] ?? 0);
if (!$itemId) jsonResponse([]);

jsonResponse(QRKioskManager::getItemModifiers($itemId));
