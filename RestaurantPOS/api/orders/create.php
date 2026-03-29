<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

if ($_SERVER['REQUEST_METHOD'] !== 'POST') { jsonResponse(['error' => 'Method not allowed'], 405); }

$data = json_decode(file_get_contents('php://input'), true) ?? [];
$data['tenant_id'] = Auth::tenantId();
$data['user_id']   = Auth::id();

if (empty($data['location_id'])) jsonResponse(['error' => 'location_id required'], 422);

$orderId = OrderManager::create($data);
$order   = Database::fetchOne("SELECT * FROM orders WHERE order_id = ?", [$orderId]);
jsonResponse(['order_id' => $orderId, 'order_number' => $order['order_number']]);
