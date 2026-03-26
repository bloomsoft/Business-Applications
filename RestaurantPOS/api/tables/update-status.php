<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['table_id']) || empty($data['status'])) jsonResponse(['error' => 'Missing fields'], 422);

$allowed = ['available', 'occupied', 'reserved', 'cleaning'];
if (!in_array($data['status'], $allowed)) jsonResponse(['error' => 'Invalid status'], 422);

TableManager::updateStatus((int)$data['table_id'], $data['status']);
jsonResponse(['success' => true]);
