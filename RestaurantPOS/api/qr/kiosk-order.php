<?php
/**
 * Kiosk self-checkout order placement
 */
require_once __DIR__ . '/../../core/bootstrap.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') jsonResponse(['error' => 'Method not allowed'], 405);

$data       = json_decode(file_get_contents('php://input'), true) ?? [];
$locationId = (int)($data['location_id'] ?? 0);
$tenantId   = (int)($data['tenant_id']   ?? 0);
$orderType  = $data['order_type']         ?? 'takeout';
$cart       = $data['cart']               ?? [];

if (!$locationId || !$tenantId || empty($cart)) {
    jsonResponse(['success' => false, 'message' => 'Missing required data'], 422);
}

$result = QRKioskManager::kioskOrder($locationId, $tenantId, $cart, $orderType);
jsonResponse($result, $result['success'] ? 200 : 422);
