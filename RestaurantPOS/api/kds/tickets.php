<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$locationId = (int) ($_GET['location_id'] ?? Auth::locationId());
jsonResponse(OrderManager::getKDSTickets($locationId));
