<?php
/**
 * Authentication & Session Management
 */
class Auth {
    public static function startSession(): void {
        if (session_status() === PHP_SESSION_NONE) {
            ini_set('session.cookie_httponly', 1);
            ini_set('session.cookie_secure',   APP_ENV === 'production' ? 1 : 0);
            ini_set('session.cookie_samesite', 'Strict');
            ini_set('session.gc_maxlifetime',  SESSION_LIFETIME * 60);
            session_start();
        }
    }

    public static function login(string $email, string $password, int $tenantId): array {
        $user = Database::fetchOne(
            "SELECT u.*, r.permissions FROM users u
             LEFT JOIN roles r ON r.role_id = u.role_id
             WHERE u.email = ? AND u.tenant_id = ? AND u.is_active = 1",
            [$email, $tenantId]
        );
        if (!$user || !password_verify($password, $user['password_hash'])) {
            return ['success' => false, 'message' => 'Invalid credentials'];
        }
        self::setSession($user);
        Database::query(
            "UPDATE users SET last_login = datetime('now') WHERE user_id = ?",
            [$user['user_id']]
        );
        return ['success' => true, 'user' => self::safeUser($user)];
    }

    public static function loginByPin(string $pin, int $locationId): array {
        $user = Database::fetchOne(
            "SELECT u.*, r.permissions FROM users u
             LEFT JOIN roles r ON r.role_id = u.role_id
             LEFT JOIN locations l ON l.location_id = u.location_id
             WHERE u.pin_code = ? AND u.location_id = ? AND u.is_active = 1",
            [$pin, $locationId]
        );
        if (!$user) {
            return ['success' => false, 'message' => 'Invalid PIN'];
        }
        self::setSession($user);
        return ['success' => true, 'user' => self::safeUser($user)];
    }

    public static function logout(): void {
        self::startSession();
        session_destroy();
    }

    public static function check(): bool {
        self::startSession();
        return isset($_SESSION['user_id']);
    }

    public static function user(): ?array {
        self::startSession();
        if (!isset($_SESSION['user_id'])) return null;
        return $_SESSION['user'] ?? null;
    }

    public static function id(): ?int {
        self::startSession();
        return $_SESSION['user_id'] ?? null;
    }

    public static function tenantId(): ?int {
        self::startSession();
        return $_SESSION['tenant_id'] ?? null;
    }

    public static function locationId(): ?int {
        self::startSession();
        return $_SESSION['location_id'] ?? null;
    }

    public static function can(string $permission): bool {
        self::startSession();
        $perms = $_SESSION['permissions'] ?? [];
        return in_array('*', $perms) || in_array($permission, $perms);
    }

    public static function requireAuth(): void {
        if (!self::check()) {
            if (self::isApiRequest()) {
                http_response_code(401);
                die(json_encode(['error' => 'Unauthorized']));
            }
            header('Location: /login.php');
            exit;
        }
    }

    public static function requirePermission(string $permission): void {
        self::requireAuth();
        if (!self::can($permission)) {
            if (self::isApiRequest()) {
                http_response_code(403);
                die(json_encode(['error' => 'Forbidden']));
            }
            header('Location: /unauthorized.php');
            exit;
        }
    }

    public static function hashPassword(string $password): string {
        return password_hash($password, PASSWORD_BCRYPT, ['cost' => BCRYPT_COST]);
    }

    public static function generateToken(int $length = 32): string {
        return bin2hex(random_bytes($length));
    }

    private static function setSession(array $user): void {
        self::startSession();
        session_regenerate_id(true);
        $_SESSION['user_id']    = $user['user_id'];
        $_SESSION['tenant_id']  = $user['tenant_id'];
        $_SESSION['location_id']= $user['location_id'];
        $_SESSION['permissions']= json_decode($user['permissions'] ?? '[]', true);
        $_SESSION['user']       = self::safeUser($user);
    }

    private static function safeUser(array $user): array {
        unset($user['password_hash'], $user['pin_code'], $user['permissions']);
        return $user;
    }

    private static function isApiRequest(): bool {
        return (
            isset($_SERVER['HTTP_ACCEPT']) &&
            str_contains($_SERVER['HTTP_ACCEPT'], 'application/json')
        ) || str_starts_with($_SERVER['REQUEST_URI'] ?? '', '/api/');
    }
}
