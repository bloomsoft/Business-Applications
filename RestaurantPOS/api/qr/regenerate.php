<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
$data    = json_decode(file_get_contents('php://input'), true) ?? [];
$tableId = (int)($data['table_id'] ?? 0);
if (!$tableId) jsonResponse(['error' => 'Missing table_id'], 422);

$token = bin2hex(random_bytes(16));
Database::query(
    "UPDATE restaurant_tables SET qr_code_token = ? WHERE table_id = ?",
    [$token, $tableId]
);
jsonResponse(['success' => true, 'token' => $token]);
