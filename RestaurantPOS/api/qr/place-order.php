<?php
require_once __DIR__ . '/../../core/bootstrap.php';
// Public endpoint — no auth required

if ($_SERVER['REQUEST_METHOD'] !== 'POST') jsonResponse(['error' => 'Method not allowed'], 405);

$data        = json_decode(file_get_contents('php://input'), true) ?? [];
$tableToken  = $data['table_token'] ?? '';
$cart        = $data['cart']        ?? [];
$customerInfo= $data['customer']    ?? [];

if (!$tableToken || empty($cart)) {
    jsonResponse(['success' => false, 'message' => 'Missing required data'], 422);
}

// Rate limit: max 3 orders per IP per minute
$ip       = $_SERVER['REMOTE_ADDR'] ?? '';
$cacheKey = 'qr_rate_' . md5($ip);
$count    = (int) ($_SESSION[$cacheKey] ?? 0);
if ($count >= 3) jsonResponse(['success' => false, 'message' => 'Too many requests'], 429);
$_SESSION[$cacheKey] = $count + 1;

$result = QRKioskManager::placeOrder($tableToken, $cart, $customerInfo);
jsonResponse($result, $result['success'] ? 200 : 422);
