<?php
/**
 * Application Bootstrap — include this at the top of every page
 */

require_once __DIR__ . '/../config/config.php';
require_once __DIR__ . '/Database.php';
require_once __DIR__ . '/Auth.php';
require_once __DIR__ . '/helpers.php';

// Auto-load module classes
spl_autoload_register(function (string $class): void {
    $dirs = [
        __DIR__ . '/../modules/pos/',
        __DIR__ . '/../modules/inventory/',
        __DIR__ . '/../modules/crm/',
        __DIR__ . '/../modules/delivery/',
        __DIR__ . '/../modules/locations/',
        __DIR__ . '/../modules/analytics/',
        __DIR__ . '/../modules/financial/',
        __DIR__ . '/../modules/staff/',
        __DIR__ . '/../modules/qr-kiosk/',
        __DIR__ . '/',
    ];
    foreach ($dirs as $dir) {
        $file = $dir . $class . '.php';
        if (file_exists($file)) {
            require_once $file;
            return;
        }
    }
});

Auth::startSession();
