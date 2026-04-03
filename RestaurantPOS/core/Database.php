<?php
/**
 * Database Connection — SQLite via PDO (no installation required)
 */
class Database {
    private static ?PDO $instance = null;

    public static function getInstance(): PDO {
        if (self::$instance === null) {
            $dsn = 'sqlite:' . DB_PATH;
            try {
                self::$instance = new PDO($dsn, null, null, [
                    PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
                    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
                    PDO::ATTR_EMULATE_PREPARES   => false,
                ]);
                self::$instance->exec("PRAGMA foreign_keys = ON");
                self::$instance->exec("PRAGMA journal_mode = WAL");
            } catch (PDOException $e) {
                error_log('DB Connection failed: ' . $e->getMessage());
                http_response_code(503);
                $msg = htmlspecialchars($e->getMessage());
                $db  = htmlspecialchars(DB_PATH);
                die("<!DOCTYPE html><html><head><title>Database Error</title>
                <style>body{font-family:Arial;max-width:700px;margin:60px auto;padding:20px}
                .box{background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:20px}
                code{background:#f8f9fa;padding:2px 6px;border-radius:4px;font-size:13px}</style></head>
                <body><div class='box'>
                <h2>&#9888; Database Not Ready</h2>
                <p><strong>Error:</strong> $msg</p>
                <p><strong>Database file:</strong> <code>$db</code></p>
                <hr>
                <h3>Fix: Run setup first</h3>
                <p>Open <strong>Command Prompt</strong> in the RestaurantPOS folder and run:</p>
                <p><code>php database\\setup.php</code></p>
                <p>Then start the server: <code>php -S localhost:8000</code></p>
                <p>Then open: <a href='http://localhost:8000/login.php'>http://localhost:8000/login.php</a></p>
                </div></body></html>");
            }
        }
        return self::$instance;
    }

    /** Execute a prepared query and return the statement */
    public static function query(string $sql, array $params = []): PDOStatement {
        $stmt = self::getInstance()->prepare($sql);
        $stmt->execute($params);
        return $stmt;
    }

    /** Fetch all rows */
    public static function fetchAll(string $sql, array $params = []): array {
        return self::query($sql, $params)->fetchAll();
    }

    /** Fetch a single row */
    public static function fetchOne(string $sql, array $params = []): array|false {
        return self::query($sql, $params)->fetch();
    }

    /** Fetch a single column value */
    public static function fetchValue(string $sql, array $params = []): mixed {
        return self::query($sql, $params)->fetchColumn();
    }

    /** Insert and return new id */
    public static function insert(string $sql, array $params = []): int {
        self::query($sql, $params);
        return (int) self::getInstance()->lastInsertId();
    }

    /** Begin transaction */
    public static function beginTransaction(): void {
        self::getInstance()->beginTransaction();
    }

    /** Commit transaction */
    public static function commit(): void {
        self::getInstance()->commit();
    }

    /** Rollback transaction */
    public static function rollback(): void {
        self::getInstance()->rollBack();
    }
}
