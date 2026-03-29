<?php
/**
 * RestaurantPOS - SQLite Database Setup
 * Run once: php setup.php
 */

$dbPath = __DIR__ . '/restaurant_pos.db';

if (file_exists($dbPath)) {
    echo "Database already exists. Deleting and recreating...\n";
    unlink($dbPath);
}

try {
    $db = new PDO('sqlite:' . $dbPath);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("PRAGMA foreign_keys = ON");
    $db->exec("PRAGMA journal_mode = WAL");

    echo "Loading schema...\n";
    $schema = file_get_contents(__DIR__ . '/sqlite_schema.sql');
    $db->exec($schema);
    echo "Schema loaded OK\n";

    echo "Loading seed data...\n";
    $seed = file_get_contents(__DIR__ . '/sqlite_seed.sql');
    $db->exec($seed);
    echo "Seed data loaded OK\n";

    $tables = $db->query("SELECT COUNT(*) FROM sqlite_master WHERE type='table'")->fetchColumn();
    $users  = $db->query("SELECT COUNT(*) FROM users")->fetchColumn();
    $items  = $db->query("SELECT COUNT(*) FROM menu_items")->fetchColumn();

    echo "\n=============================================\n";
    echo " Database created successfully!\n";
    echo "=============================================\n";
    echo " Tables   : $tables\n";
    echo " Users    : $users\n";
    echo " Menu items: $items\n";
    echo "=============================================\n";
    echo "\n Login credentials:\n";
    echo "  Email   : admin@demo.com\n";
    echo "  Password: password123\n";
    echo "\n Now run: php -S localhost:8000\n";
    echo " Open   : http://localhost:8000/login.php\n\n";

} catch (Exception $e) {
    echo "ERROR: " . $e->getMessage() . "\n";
    exit(1);
}
