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

// Rate limit: max 10 orders per IP per minute (time-based reset)
$ip       = $_SERVER['REMOTE_ADDR'] ?? '';
$cacheKey = 'qr_rate_' . md5($ip);
$tsKey    = 'qr_rate_ts_' . md5($ip);
$now      = time();
if (isset($_SESSION[$tsKey]) && ($now - $_SESSION[$tsKey]) > 60) {
    $_SESSION[$cacheKey] = 0; // reset counter after 60 seconds
}
$count = (int) ($_SESSION[$cacheKey] ?? 0);
if ($count >= 10) jsonResponse(['success' => false, 'message' => 'Too many requests, please wait a minute'], 429);
$_SESSION[$cacheKey] = $count + 1;
$_SESSION[$tsKey]    = $_SESSION[$tsKey] ?? $now;

$result = QRKioskManager::placeOrder($tableToken, $cart, $customerInfo);
jsonResponse($result, $result['success'] ? 200 : 422);
