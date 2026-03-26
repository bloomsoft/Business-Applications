<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$orderId = (int) ($_GET['order_id'] ?? 0);
if (!$orderId) jsonResponse(['error' => 'Missing order_id'], 422);

$order = OrderManager::getOrder($orderId);
if (!$order) jsonResponse(['error' => 'Not found'], 404);
jsonResponse($order);
