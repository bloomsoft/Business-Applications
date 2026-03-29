<?php
// Check if PDO SQLite is available
if (!in_array('sqlite', PDO::getAvailableDrivers())) {
    echo "ERROR: PDO SQLite extension is not enabled!\n\n";
    echo "Fix karo:\n";
    echo "1. php --ini   (php.ini ka path pata karo)\n";
    echo "2. php.ini mein ye lines uncomment karo:\n";
    echo "   extension=pdo_sqlite\n";
    echo "   extension=sqlite3\n";
    echo "3. CMD band karke dobara kholو aur phir chalao\n";
    exit(1);
}

$dbPath = __DIR__ . '/restaurant_pos.db';
if (file_exists($dbPath)) {
    unlink($dbPath);
    echo "Old database deleted.\n";
}

try {
    $db = new PDO('sqlite:' . $dbPath);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("PRAGMA foreign_keys = ON");
    $db->exec("PRAGMA journal_mode = WAL");

    echo "Loading schema...\n";
    $db->exec(file_get_contents(__DIR__ . '/sqlite_schema.sql'));
    echo "Schema OK\n";

    echo "Loading seed data...\n";
    $db->exec(file_get_contents(__DIR__ . '/sqlite_seed.sql'));
    echo "Seed OK\n";

    $tables = $db->query("SELECT COUNT(*) FROM sqlite_master WHERE type='table'")->fetchColumn();
    $users  = $db->query("SELECT COUNT(*) FROM users")->fetchColumn();
    $items  = $db->query("SELECT COUNT(*) FROM menu_items")->fetchColumn();

    echo "\n=============================================\n";
    echo " Database Ready!\n";
    echo "=============================================\n";
    echo " Tables    : $tables\n";
    echo " Users     : $users\n";
    echo " Menu Items: $items\n";
    echo "=============================================\n";
    echo " Email   : admin@demo.com\n";
    echo " Password: password123\n";
    echo "=============================================\n";
    echo "\nServer start karo:\n";
    echo "  php -S localhost:8000\n\n";
    echo "Browser mein kholو:\n";
    echo "  http://localhost:8000/login.php\n\n";

} catch (Exception $e) {
    echo "ERROR: " . $e->getMessage() . "\n";
    exit(1);
}
