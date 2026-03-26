<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['order_item_id'])) jsonResponse(['error' => 'Missing order_item_id'], 422);

OrderManager::removeItem((int)$data['order_item_id']);
jsonResponse(['success' => true]);
