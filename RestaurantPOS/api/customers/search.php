<?php
require_once __DIR__ . '/../../core/bootstrap.php';
Auth::requireAuth();

$q        = get('q');
$tenantId = Auth::tenantId();
if (strlen($q) < 2) jsonResponse([]);

jsonResponse(CustomerManager::find($tenantId, $q));
