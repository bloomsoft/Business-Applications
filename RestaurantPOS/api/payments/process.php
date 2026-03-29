<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['order_id']) || empty($data['method']) || !isset($data['amount'])) {
    jsonResponse(['error' => 'Missing required fields'], 422);
}

$result = PaymentManager::process((int)$data['order_id'], $data);
jsonResponse($result, $result['success'] ? 200 : 422);
