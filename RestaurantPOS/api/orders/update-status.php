<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
if (empty($data['order_id']) || empty($data['status'])) jsonResponse(['error' => 'Missing fields'], 422);

$allowed = ['confirmed','preparing','ready','served','completed','cancelled'];
if (!in_array($data['status'], $allowed)) jsonResponse(['error' => 'Invalid status'], 422);

OrderManager::updateStatus((int)$data['order_id'], $data['status']);
jsonResponse(['success' => true]);
