<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();
$data     = json_decode(file_get_contents('php://input'), true) ?? [];
$tenantId = (int)($data['tenant_id'] ?? Auth::tenantId());
CustomerManager::resegmentAll($tenantId);
jsonResponse(['success' => true]);
