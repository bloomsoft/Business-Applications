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
                error_reporting(E_ALL);
                ini_set('display_errors', 1);
                error_log('DB Connection failed: ' . $e->getMessage());
                http_response_code(503);
                die(json_encode(['error' => 'Database unavailable: ' . $e->getMessage()]));
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
