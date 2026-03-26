<?php
/**
 * Database Connection — SQL Server via PDO (sqlsrv driver)
 */
class Database {
    private static ?PDO $instance = null;

    public static function getInstance(): PDO {
        if (self::$instance === null) {
            $dsn = 'sqlsrv:Server=' . DB_HOST . ',' . DB_PORT . ';Database=' . DB_NAME . ';TrustServerCertificate=1';
            try {
                self::$instance = new PDO($dsn, DB_USER, DB_PASS, [
                    PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
                    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
                    PDO::ATTR_EMULATE_PREPARES   => false,
                ]);
            } catch (PDOException $e) {
                error_log('DB Connection failed: ' . $e->getMessage());
                http_response_code(503);
                die(json_encode(['error' => 'Database unavailable']));
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

    /** Insert and return new identity */
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
