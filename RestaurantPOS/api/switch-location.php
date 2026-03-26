<?php
require_once __DIR__ . '/../core/bootstrap.php';
Auth::requireAuth();

$data = json_decode(file_get_contents('php://input'), true) ?? [];
$locationId = (int)($data['location_id'] ?? 0);

// Verify this location belongs to the tenant
$loc = Database::fetchOne(
    "SELECT location_id FROM locations WHERE location_id = ? AND tenant_id = ?",
    [$locationId, Auth::tenantId()]
);
if (!$loc) jsonResponse(['error' => 'Location not found'], 404);

$_SESSION['location_id'] = $locationId;
jsonResponse(['success' => true]);
