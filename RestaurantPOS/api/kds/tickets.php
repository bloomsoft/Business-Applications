<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$locationId = (int) ($_GET['location_id'] ?? Auth::locationId() ?? 0);
if (!$locationId) jsonResponse(['error' => 'location_id required'], 422);
jsonResponse(OrderManager::getKDSTickets($locationId));
