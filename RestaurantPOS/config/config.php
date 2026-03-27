<?php
/**
 * RestaurantPOS SaaS Platform
 * Core Configuration
 */

define('APP_NAME',    'RestaurantPOS');
define('APP_VERSION', '1.0.0');
define('APP_URL',     getenv('APP_URL') ?: 'http://localhost');
define('APP_ENV',     getenv('APP_ENV') ?: 'production'); // development | production

// ── Database (SQL Server via PDO/SQLSRV) ─────────────────────────────────────
define('DB_HOST',     getenv('DB_HOST')     ?: 'localhost');
define('DB_NAME',     getenv('DB_NAME')     ?: 'RestaurantPOS');
define('DB_USER',     getenv('DB_USER')     ?: 'sa');
define('DB_PASS',     getenv('DB_PASS')     ?: '');
define('DB_PORT',     getenv('DB_PORT')     ?: '1433');

// ── Security ─────────────────────────────────────────────────────────────────
define('APP_KEY',          getenv('APP_KEY')     ?: 'change-this-32-char-secret-key!!');
define('SESSION_LIFETIME', 480);   // minutes
define('BCRYPT_COST',      12);

// ── Uploads ───────────────────────────────────────────────────────────────────
define('UPLOAD_DIR',      __DIR__ . '/../uploads/');
define('MAX_UPLOAD_SIZE', 5 * 1024 * 1024); // 5MB

// ── Timezone ─────────────────────────────────────────────────────────────────
date_default_timezone_set('UTC');

// ── Delivery Integrations ─────────────────────────────────────────────────────
define('UBEREATS_CLIENT_ID',     getenv('UBEREATS_CLIENT_ID')     ?: '');
define('UBEREATS_CLIENT_SECRET', getenv('UBEREATS_CLIENT_SECRET') ?: '');
define('DOORDASH_DEVELOPER_ID',  getenv('DOORDASH_DEVELOPER_ID')  ?: '');
define('DOORDASH_KEY_ID',        getenv('DOORDASH_KEY_ID')        ?: '');
define('DOORDASH_SIGNING_SECRET',getenv('DOORDASH_SIGNING_SECRET')?:'');
define('GRUBHUB_API_KEY',        getenv('GRUBHUB_API_KEY')        ?: '');

// ── Payment Gateways ─────────────────────────────────────────────────────────
define('STRIPE_SECRET_KEY',      getenv('STRIPE_SECRET_KEY')      ?: '');
define('STRIPE_PUBLIC_KEY',      getenv('STRIPE_PUBLIC_KEY')      ?: '');
define('PAYPAL_CLIENT_ID',       getenv('PAYPAL_CLIENT_ID')       ?: '');
define('PAYPAL_CLIENT_SECRET',   getenv('PAYPAL_CLIENT_SECRET')   ?: '');

// ── Email (SMTP) ─────────────────────────────────────────────────────────────
define('SMTP_HOST',     getenv('SMTP_HOST')     ?: 'smtp.mailtrap.io');
define('SMTP_PORT',     getenv('SMTP_PORT')     ?: 587);
define('SMTP_USER',     getenv('SMTP_USER')     ?: '');
define('SMTP_PASS',     getenv('SMTP_PASS')     ?: '');
define('SMTP_FROM',     getenv('SMTP_FROM')     ?: 'noreply@restaurantpos.app');
define('SMTP_FROM_NAME',getenv('SMTP_FROM_NAME')?:'RestaurantPOS');

// ── Error Reporting ───────────────────────────────────────────────────────────
if (APP_ENV === 'development') {
    error_reporting(E_ALL);
    ini_set('display_errors', 1);
} else {
    error_reporting(0);
    ini_set('display_errors', 0);
}
