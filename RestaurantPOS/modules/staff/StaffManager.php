<?php
/**
 * Staff Management Module
 * Scheduling, time clocks, payroll
 */
class StaffManager {

    // ── Users / Staff ─────────────────────────────────────────────────────────

    public static function createStaff(array $data): int {
        return Database::insert(
            "INSERT INTO users
                (tenant_id, location_id, role_id, first_name, last_name,
                 email, phone, password_hash, pin_code)
             VALUES (?,?,?,?,?,?,?,?,?)",
            [
                $data['tenant_id'],
                $data['location_id'] ?? null,
                $data['role_id']     ?? null,
                $data['first_name'],
                $data['last_name'],
                $data['email'],
                $data['phone']       ?? null,
                Auth::hashPassword($data['password']),
                $data['pin_code']    ?? null,
            ]
        );
    }

    public static function getStaff(int $locationId, array $filters = []): array {
        $where  = ['u.location_id = ?', 'u.is_active = 1'];
        $params = [$locationId];

        if (!empty($filters['role_id'])) {
            $where[]  = 'u.role_id = ?';
            $params[] = $filters['role_id'];
        }
        if (!empty($filters['search'])) {
            $t        = '%' . $filters['search'] . '%';
            $where[]  = "(u.first_name LIKE ? OR u.last_name LIKE ? OR u.email LIKE ?)";
            $params   = array_merge($params, [$t, $t, $t]);
        }

        $whereStr = implode(' AND ', $where);
        return Database::fetchAll(
            "SELECT u.user_id, u.first_name, u.last_name, u.email, u.phone,
                    u.last_login, u.created_at,
                    r.role_name
             FROM users u
             LEFT JOIN roles r ON r.role_id = u.role_id
             WHERE $whereStr
             ORDER BY u.first_name",
            $params
        );
    }

    // ── Scheduling ────────────────────────────────────────────────────────────

    public static function createShift(array $data): int {
        return Database::insert(
            "INSERT INTO shifts
                (location_id, user_id, shift_date, start_time, end_time, break_minutes, notes, created_by)
             VALUES (?,?,?,?,?,?,?,?)",
            [
                $data['location_id'],
                $data['user_id'],
                $data['shift_date'],
                $data['start_time'],
                $data['end_time'],
                $data['break_minutes'] ?? 0,
                $data['notes']         ?? null,
                Auth::id(),
            ]
        );
    }

    public static function getSchedule(int $locationId, string $weekStart): array {
        $weekEnd = date('Y-m-d', strtotime($weekStart . ' +6 days'));
        $shifts  = Database::fetchAll(
            "SELECT s.*,
                    u.first_name || ' ' || u.last_name AS staff_name,
                    r.role_name
             FROM shifts s
             JOIN users u ON u.user_id = s.user_id
             LEFT JOIN roles r ON r.role_id = u.role_id
             WHERE s.location_id = ?
               AND s.shift_date BETWEEN ? AND ?
             ORDER BY s.shift_date, s.start_time",
            [$locationId, $weekStart, $weekEnd]
        );

        // Group by staff member
        $schedule = [];
        foreach ($shifts as $shift) {
            $schedule[$shift['user_id']]['staff_name'] = $shift['staff_name'];
            $schedule[$shift['user_id']]['role_name']  = $shift['role_name'];
            $schedule[$shift['user_id']]['shifts'][]   = $shift;
        }
        return array_values($schedule);
    }

    public static function deleteShift(int $shiftId): void {
        Database::query("DELETE FROM shifts WHERE shift_id = ?", [$shiftId]);
    }

    // ── Time Clock ────────────────────────────────────────────────────────────

    public static function clockIn(int $userId, int $locationId): int {
        // Ensure not already clocked in
        $open = Database::fetchOne(
            "SELECT clock_id FROM time_clocks WHERE user_id = ? AND clock_out IS NULL",
            [$userId]
        );
        if ($open) throw new RuntimeException('Already clocked in');

        return Database::insert(
            "INSERT INTO time_clocks (user_id, location_id, clock_in) VALUES (?,?,datetime('now'))",
            [$userId, $locationId]
        );
    }

    public static function clockOut(int $userId): array {
        $clock = Database::fetchOne(
            "SELECT * FROM time_clocks WHERE user_id = ? AND clock_out IS NULL ORDER BY clock_in DESC",
            [$userId]
        );
        if (!$clock) throw new RuntimeException('Not clocked in');

        $totalHours    = self::calcHours($clock['clock_in'], null, (int)$clock['break_start'] ? $clock : null);
        $regularHours  = min($totalHours, 8);
        $overtimeHours = max(0, $totalHours - 8);

        Database::query(
            "UPDATE time_clocks
             SET clock_out = datetime('now'), total_hours = ?, overtime_hours = ?
             WHERE clock_id = ?",
            [$totalHours, $overtimeHours, $clock['clock_id']]
        );

        return ['total_hours' => $totalHours, 'overtime_hours' => $overtimeHours];
    }

    public static function startBreak(int $userId): void {
        Database::query(
            "UPDATE time_clocks SET break_start = datetime('now')
             WHERE user_id = ? AND clock_out IS NULL",
            [$userId]
        );
    }

    public static function endBreak(int $userId): void {
        Database::query(
            "UPDATE time_clocks SET break_end = datetime('now')
             WHERE user_id = ? AND clock_out IS NULL",
            [$userId]
        );
    }

    public static function getTimeLog(int $userId, string $startDate, string $endDate): array {
        return Database::fetchAll(
            "SELECT * FROM time_clocks
             WHERE user_id = ? AND date(clock_in) BETWEEN ? AND ?
             ORDER BY clock_in",
            [$userId, $startDate, $endDate]
        );
    }

    // ── Payroll ───────────────────────────────────────────────────────────────

    public static function generatePayroll(int $locationId, string $periodStart, string $periodEnd): array {
        $staff = Database::fetchAll(
            "SELECT u.user_id, u.first_name || ' ' || u.last_name AS full_name
             FROM users u WHERE u.location_id = ? AND u.is_active = 1",
            [$locationId]
        );

        $results = [];
        foreach ($staff as $employee) {
            $hours = Database::fetchOne(
                "SELECT COALESCE(SUM(total_hours),0)    AS regular_hours,
                        COALESCE(SUM(overtime_hours),0) AS overtime_hours
                 FROM time_clocks
                 WHERE user_id = ? AND date(clock_in) BETWEEN ? AND ?",
                [$employee['user_id'], $periodStart, $periodEnd]
            );

            // Fetch hourly rate from user settings (stub — add rate column or separate table in prod)
            $hourlyRate = 15.00;

            $tips = (float) Database::fetchValue(
                "SELECT COALESCE(SUM(tip_amount),0) FROM orders
                 WHERE user_id = ? AND status='completed'
                   AND date(created_at) BETWEEN ? AND ?",
                [$employee['user_id'], $periodStart, $periodEnd]
            );

            $grossPay = ((float)$hours['regular_hours'] * $hourlyRate)
                      + ((float)$hours['overtime_hours'] * $hourlyRate * 1.5)
                      + $tips;

            $payrollId = Database::insert(
                "INSERT INTO payroll
                    (user_id, location_id, period_start, period_end,
                     regular_hours, overtime_hours, hourly_rate, tips_amount, gross_pay, net_pay)
                 VALUES (?,?,?,?,?,?,?,?,?,?)",
                [
                    $employee['user_id'], $locationId, $periodStart, $periodEnd,
                    $hours['regular_hours'], $hours['overtime_hours'],
                    $hourlyRate, $tips, $grossPay, $grossPay * 0.85, // 15% deduction placeholder
                ]
            );

            $results[] = [
                'payroll_id'     => $payrollId,
                'staff_name'     => $employee['full_name'],
                'regular_hours'  => $hours['regular_hours'],
                'overtime_hours' => $hours['overtime_hours'],
                'gross_pay'      => $grossPay,
            ];
        }
        return $results;
    }

    public static function getRoles(int $tenantId): array {
        return Database::fetchAll(
            "SELECT * FROM roles WHERE tenant_id = ? ORDER BY role_name",
            [$tenantId]
        );
    }

    private static function calcHours(string $clockIn, ?string $clockOut, ?array $breaks): float {
        $in   = strtotime($clockIn);
        $out  = $clockOut ? strtotime($clockOut) : time();
        $mins = ($out - $in) / 60;

        if ($breaks && $breaks['break_start'] && $breaks['break_end']) {
            $breakMins = (strtotime($breaks['break_end']) - strtotime($breaks['break_start'])) / 60;
            $mins -= max(0, $breakMins);
        }

        return round($mins / 60, 2);
    }
}
