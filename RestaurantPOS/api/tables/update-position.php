<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['table_id'])) jsonResponse(['error' => 'Missing table_id'], 422);

TableManager::updatePosition((int)$data['table_id'], (int)($data['x'] ?? 0), (int)($data['y'] ?? 0));
jsonResponse(['success' => true]);
