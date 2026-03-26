<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['order_id']) || empty($data['item_id'])) jsonResponse(['error' => 'Missing fields'], 422);

$orderItemId = OrderManager::addItem((int)$data['order_id'], $data);
$order       = OrderManager::getOrder((int)$data['order_id']);
jsonResponse(['order_item_id' => $orderItemId, 'order' => $order]);
