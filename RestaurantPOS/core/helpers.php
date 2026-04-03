<?php
/**
 * Global Helper Functions
 */

/** Format currency */
function money(float $amount, string $currency = 'USD'): string {
    return '$' . number_format($amount, 2);
}

/** Generate unique order number */
function generateOrderNumber(int $locationId): string {
    return strtoupper(substr(base_convert($locationId, 10, 36), 0, 2))
         . date('ymd')
         . str_pad(mt_rand(0, 9999), 4, '0', STR_PAD_LEFT);
}

/** Sanitize string input */
function sanitize(string $input): string {
    return htmlspecialchars(trim($input), ENT_QUOTES, 'UTF-8');
}

/** Return JSON response */
function jsonResponse(mixed $data, int $status = 200): never {
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

/** Redirect */
function redirect(string $url): never {
    header("Location: $url");
    exit;
}

/** Get POST value */
function post(string $key, mixed $default = ''): mixed {
    return $_POST[$key] ?? $default;
}

/** Get GET value */
function get(string $key, mixed $default = ''): mixed {
    return $_GET[$key] ?? $default;
}

/** Flash message */
function flash(string $type, string $message): void {
    Auth::startSession();
    $_SESSION['flash'] = ['type' => $type, 'message' => $message];
}

/** Get and clear flash */
function getFlash(): ?array {
    Auth::startSession();
    if (isset($_SESSION['flash'])) {
        $flash = $_SESSION['flash'];
        unset($_SESSION['flash']);
        return $flash;
    }
    return null;
}

/** Format date */
function fmtDate(string|null $date, string $format = 'M j, Y'): string {
    if (!$date) return '—';
    return date($format, strtotime($date));
}

/** Format datetime */
function fmtDateTime(string|null $dt, string $format = 'M j, Y g:i A'): string {
    if (!$dt) return '—';
    return date($format, strtotime($dt));
}

/** Time ago */
function timeAgo(string $datetime): string {
    $diff = time() - strtotime($datetime);
    if ($diff < 60)    return $diff . 's ago';
    if ($diff < 3600)  return floor($diff / 60) . 'm ago';
    if ($diff < 86400) return floor($diff / 3600) . 'h ago';
    return floor($diff / 86400) . 'd ago';
}

/** Generate QR code URL using Google Charts API as fallback */
function qrCodeUrl(string $data, int $size = 200): string {
    return 'https://api.qrserver.com/v1/create-qr-code/?size=' . $size . 'x' . $size
         . '&data=' . urlencode($data);
}

/** Paginate results */
function paginate(string $sql, array $params, int $page, int $perPage = 20): array {
    $offset = ($page - 1) * $perPage;
    $countSql = preg_replace('/SELECT .+? FROM/is', 'SELECT COUNT(*) FROM', $sql);
    $total = (int) Database::fetchValue($countSql, $params);
    $rows  = Database::fetchAll("$sql LIMIT $perPage OFFSET $offset", $params);
    return [
        'data'        => $rows,
        'total'       => $total,
        'per_page'    => $perPage,
        'current_page'=> $page,
        'last_page'   => (int) ceil($total / $perPage),
    ];
}

/** Log audit event */
function auditLog(string $action, string $table = '', int $recordId = 0, array $old = [], array $new = []): void {
    try {
        Database::query(
            "INSERT INTO audit_logs (tenant_id, user_id, action, table_name, record_id, old_values, new_values, ip_address)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            [
                Auth::tenantId() ?? 0,
                Auth::id() ?? 0,
                $action, $table, $recordId ?: null,
                $old ? json_encode($old) : null,
                $new  ? json_encode($new)  : null,
                $_SERVER['REMOTE_ADDR'] ?? null,
            ]
        );
    } catch (Throwable) { /* non-blocking */ }
}

/** Bootstrap alert class */
function alertClass(string $type): string {
    return match($type) {
        'success' => 'alert-success',
        'error', 'danger' => 'alert-danger',
        'warning' => 'alert-warning',
        default   => 'alert-info',
    };
}

/** Status badge HTML */
function statusBadge(string $status): string {
    $map = [
        'active'     => 'bg-success',
        'available'  => 'bg-success',
        'completed'  => 'bg-success',
        'delivered'  => 'bg-success',
        'paid'       => 'bg-success',
        'pending'    => 'bg-warning text-dark',
        'preparing'  => 'bg-warning text-dark',
        'in_progress'=> 'bg-warning text-dark',
        'draft'      => 'bg-secondary',
        'scheduled'  => 'bg-info',
        'cancelled'  => 'bg-danger',
        'failed'     => 'bg-danger',
        'occupied'   => 'bg-primary',
        'reserved'   => 'bg-info',
    ];
    $cls = $map[strtolower($status)] ?? 'bg-secondary';
    return '<span class="badge ' . $cls . '">' . ucfirst(str_replace('_', ' ', $status)) . '</span>';
}
