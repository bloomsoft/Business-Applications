<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$orderId = (int) ($_GET['order_id'] ?? 0);
if (!$orderId) die('Invalid order');

echo FinancialManager::generateInvoiceHTML($orderId);
