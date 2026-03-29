<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data    = json_decode(file_get_contents('php://input'), true) ?? [];
$orderId = (int)($data['order_id'] ?? 0);
if (!$orderId) jsonResponse(['error' => 'Missing order_id'], 422);

OrderManager::updateStatus($orderId, 'ready');
jsonResponse(['success' => true]);
