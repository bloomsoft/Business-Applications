<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();
Auth::requirePermission('staff.view');

$locationId  = Auth::locationId();
$tenantId    = Auth::tenantId();
$activeTab   = get('tab', 'staff');

$staff       = StaffManager::getStaff($locationId);
$roles       = StaffManager::getRoles($tenantId);
$today       = date('Y-m-d');
$weekStart   = date('Y-m-d', strtotime('last Monday', strtotime($today . ' +1 day')));
$schedule    = StaffManager::getSchedule($locationId, $weekStart);
$weekEnd     = date('Y-m-d', strtotime($weekStart . ' +6 days'));

$pageTitle   = 'Staff Management';
$activeMenu  = 'staff';
ob_start();
?>
<!-- Tabs -->
<ul class="nav nav-tabs mb-4">
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'staff'   ? 'active' : '' ?>" href="?tab=staff">Staff List</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'schedule'? 'active' : '' ?>" href="?tab=schedule">Schedule</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'clock'   ? 'active' : '' ?>" href="?tab=clock">Time Clock</a></li>
    <li class="nav-item"><a class="nav-link <?= $activeTab === 'payroll' ? 'active' : '' ?>" href="?tab=payroll">Payroll</a></li>
</ul>

<?php if ($activeTab === 'staff'): ?>
<!-- Staff List -->
<div class="d-flex justify-content-between mb-3">
    <h5 class="mb-0">Team Members</h5>
    <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addStaffModal">
        <i class="bi bi-person-plus me-1"></i>Add Staff
    </button>
</div>
<div class="card shadow-sm">
    <div class="card-body p-0">
        <table class="table table-hover align-middle mb-0">
            <thead class="table-light">
                <tr><th>Name</th><th>Role</th><th>Email</th><th>Phone</th><th>Last Login</th><th></th></tr>
            </thead>
            <tbody>
                <?php foreach ($staff as $s): ?>
                <tr>
                    <td>
                        <div class="d-flex align-items-center gap-2">
                            <div class="rounded-circle bg-primary text-white d-flex align-items-center justify-content-center"
                                 style="width:36px;height:36px;font-size:14px">
                                <?= strtoupper(substr($s['first_name'],0,1) . substr($s['last_name'],0,1)) ?>
                            </div>
                            <div>
                                <div class="fw-600"><?= sanitize($s['first_name'] . ' ' . $s['last_name']) ?></div>
                            </div>
                        </div>
                    </td>
                    <td><span class="badge bg-secondary"><?= sanitize($s['role_name'] ?? 'N/A') ?></span></td>
                    <td><small><?= sanitize($s['email']) ?></small></td>
                    <td><small><?= sanitize($s['phone'] ?? '—') ?></small></td>
                    <td><small class="text-muted"><?= $s['last_login'] ? timeAgo($s['last_login']) : 'Never' ?></small></td>
                    <td>
                        <button class="btn btn-sm btn-outline-secondary">
                            <i class="bi bi-pencil"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deactivateStaff(<?= $s['user_id'] ?>)">
                            <i class="bi bi-person-x"></i>
                        </button>
                    </td>
                </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
    </div>
</div>

<?php elseif ($activeTab === 'schedule'): ?>
<!-- Weekly Schedule -->
<div class="d-flex justify-content-between align-items-center mb-3">
    <div>
        <h5 class="mb-0">Week of <?= fmtDate($weekStart) ?> — <?= fmtDate($weekEnd) ?></h5>
    </div>
    <div class="d-flex gap-2">
        <a href="?tab=schedule&week=<?= date('Y-m-d', strtotime($weekStart . ' -7 days')) ?>"
           class="btn btn-sm btn-outline-secondary"><i class="bi bi-chevron-left"></i></a>
        <a href="?tab=schedule&week=<?= date('Y-m-d', strtotime($weekStart . ' +7 days')) ?>"
           class="btn btn-sm btn-outline-secondary"><i class="bi bi-chevron-right"></i></a>
        <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addShiftModal">
            <i class="bi bi-plus me-1"></i>Add Shift
        </button>
    </div>
</div>
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-bordered mb-0" style="min-width:700px">
                <thead class="table-light">
                    <tr>
                        <th>Staff</th>
                        <?php
                        for ($i = 0; $i < 7; $i++) {
                            $d = date('Y-m-d', strtotime($weekStart . " +$i days"));
                            echo '<th class="text-center ' . ($d === $today ? 'table-primary' : '') . '">'
                               . date('D', strtotime($d)) . '<br><small>' . date('M j', strtotime($d)) . '</small>'
                               . '</th>';
                        }
                        ?>
                    </tr>
                </thead>
                <tbody>
                    <?php if (empty($schedule)): ?>
                    <tr><td colspan="8" class="text-center text-muted py-4">No shifts scheduled this week</td></tr>
                    <?php else: ?>
                    <?php foreach ($schedule as $row): ?>
                    <tr>
                        <td>
                            <div class="fw-600"><?= sanitize($row['staff_name']) ?></div>
                            <small class="text-muted"><?= sanitize($row['role_name'] ?? '') ?></small>
                        </td>
                        <?php
                        for ($i = 0; $i < 7; $i++) {
                            $d = date('Y-m-d', strtotime($weekStart . " +$i days"));
                            $dayShifts = array_filter($row['shifts'], fn($s) => $s['shift_date'] === $d);
                            echo '<td class="text-center">';
                            foreach ($dayShifts as $shift) {
                                $start = date('g:iA', strtotime($shift['start_time']));
                                $end   = date('g:iA', strtotime($shift['end_time']));
                                echo '<span class="badge bg-primary d-block mb-1">' . $start . '–' . $end . '</span>';
                            }
                            echo '</td>';
                        }
                        ?>
                    </tr>
                    <?php endforeach; ?>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'clock'): ?>
<!-- Time Clock -->
<div class="row g-3">
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600">Clock In / Out</div>
            <div class="card-body">
                <div class="mb-3">
                    <label class="form-label">Select Staff Member</label>
                    <select class="form-select" id="clockStaffSelect">
                        <?php foreach ($staff as $s): ?>
                        <option value="<?= $s['user_id'] ?>"><?= sanitize($s['first_name'] . ' ' . $s['last_name']) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                <div class="d-flex gap-2">
                    <button class="btn btn-success flex-grow-1" onclick="clockAction('in')">
                        <i class="bi bi-box-arrow-in-right me-1"></i>Clock In
                    </button>
                    <button class="btn btn-warning flex-grow-1" onclick="clockAction('break')">
                        <i class="bi bi-cup-hot me-1"></i>Break
                    </button>
                    <button class="btn btn-danger flex-grow-1" onclick="clockAction('out')">
                        <i class="bi bi-box-arrow-right me-1"></i>Clock Out
                    </button>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-header fw-600">Today's Log</div>
            <div class="card-body p-0">
                <?php
                $todayLog = Database::fetchAll(
                    "SELECT tc.*, u.first_name + ' ' + u.last_name AS name
                     FROM time_clocks tc
                     JOIN users u ON u.user_id = tc.user_id
                     WHERE tc.location_id = ? AND CAST(tc.clock_in AS DATE) = ?
                     ORDER BY tc.clock_in DESC",
                    [$locationId, $today]
                );
                ?>
                <div class="list-group list-group-flush">
                    <?php foreach ($todayLog as $log): ?>
                    <div class="list-group-item d-flex justify-content-between align-items-center">
                        <div>
                            <div class="fw-600"><?= sanitize($log['name']) ?></div>
                            <small class="text-muted">
                                In: <?= fmtDateTime($log['clock_in'], 'g:i A') ?>
                                <?= $log['clock_out'] ? '· Out: ' . fmtDateTime($log['clock_out'], 'g:i A') : '' ?>
                            </small>
                        </div>
                        <div>
                            <?php if ($log['clock_out']): ?>
                            <span class="badge bg-secondary"><?= $log['total_hours'] ?>h</span>
                            <?php else: ?>
                            <span class="badge bg-success">Active</span>
                            <?php endif; ?>
                        </div>
                    </div>
                    <?php endforeach; ?>
                    <?php if (empty($todayLog)): ?>
                    <div class="list-group-item text-muted text-center">No clock-ins today</div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>
</div>

<?php elseif ($activeTab === 'payroll'): ?>
<!-- Payroll -->
<div class="card shadow-sm">
    <div class="card-header fw-600 d-flex justify-content-between">
        <span>Payroll Generation</span>
        <button class="btn btn-sm btn-primary" onclick="generatePayroll()">
            <i class="bi bi-calculator me-1"></i>Generate Payroll
        </button>
    </div>
    <div class="card-body row g-3">
        <div class="col-md-4">
            <label class="form-label">Period Start</label>
            <input type="date" class="form-control" id="payrollStart"
                   value="<?= date('Y-m-01') ?>">
        </div>
        <div class="col-md-4">
            <label class="form-label">Period End</label>
            <input type="date" class="form-control" id="payrollEnd"
                   value="<?= date('Y-m-t') ?>">
        </div>
    </div>
    <div class="card-body p-0">
        <?php
        $payrolls = Database::fetchAll(
            "SELECT p.*, u.first_name + ' ' + u.last_name AS staff_name
             FROM payroll p
             JOIN users u ON u.user_id = p.user_id
             WHERE p.location_id = ?
             ORDER BY p.created_at DESC
             OFFSET 0 ROWS FETCH NEXT 20 ROWS ONLY",
            [$locationId]
        );
        ?>
        <table class="table table-sm mb-0">
            <thead class="table-light">
                <tr><th>Staff</th><th>Period</th><th>Reg Hours</th><th>OT Hours</th><th>Tips</th><th>Gross Pay</th><th>Net Pay</th><th>Status</th></tr>
            </thead>
            <tbody>
                <?php foreach ($payrolls as $pr): ?>
                <tr>
                    <td><?= sanitize($pr['staff_name']) ?></td>
                    <td><small><?= fmtDate($pr['period_start'], 'M j') ?> – <?= fmtDate($pr['period_end'], 'M j, Y') ?></small></td>
                    <td><?= number_format($pr['regular_hours'], 1) ?></td>
                    <td><?= number_format($pr['overtime_hours'], 1) ?></td>
                    <td><?= money($pr['tips_amount']) ?></td>
                    <td class="fw-600"><?= money($pr['gross_pay']) ?></td>
                    <td><?= money($pr['net_pay']) ?></td>
                    <td><?= statusBadge($pr['status']) ?></td>
                </tr>
                <?php endforeach; ?>
                <?php if (empty($payrolls)): ?>
                <tr><td colspan="8" class="text-center text-muted py-4">No payroll records</td></tr>
                <?php endif; ?>
            </tbody>
        </table>
    </div>
</div>
<?php endif; ?>

<!-- Add Staff Modal -->
<div class="modal fade" id="addStaffModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Add Team Member</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST" action="/api/staff/create.php">
                <div class="modal-body row g-3">
                    <div class="col-6"><label class="form-label">First Name</label>
                        <input type="text" name="first_name" class="form-control" required></div>
                    <div class="col-6"><label class="form-label">Last Name</label>
                        <input type="text" name="last_name" class="form-control" required></div>
                    <div class="col-12"><label class="form-label">Email</label>
                        <input type="email" name="email" class="form-control" required></div>
                    <div class="col-6"><label class="form-label">Phone</label>
                        <input type="tel" name="phone" class="form-control"></div>
                    <div class="col-6"><label class="form-label">PIN (4-6 digits)</label>
                        <input type="text" name="pin_code" class="form-control" pattern="[0-9]{4,6}" maxlength="6"></div>
                    <div class="col-6"><label class="form-label">Role</label>
                        <select name="role_id" class="form-select">
                            <?php foreach ($roles as $role): ?>
                            <option value="<?= $role['role_id'] ?>"><?= sanitize($role['role_name']) ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-6"><label class="form-label">Password</label>
                        <input type="password" name="password" class="form-control" required></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Add Member</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
$scripts = <<<JS
<script>
async function clockAction(type) {
    const userId = parseInt(document.getElementById('clockStaffSelect').value);
    const endpoints = { in:'/api/staff/clock-in.php', out:'/api/staff/clock-out.php', break:'/api/staff/break.php' };
    try {
        await api(endpoints[type],'POST',{user_id:userId,location_id:<?= (int)$locationId ?>});
        showToast('Clock ' + type + ' recorded','success');
        setTimeout(()=>location.reload(),1000);
    } catch(e) { showToast(e.message,'error'); }
}

async function generatePayroll() {
    const start = document.getElementById('payrollStart')?.value;
    const end   = document.getElementById('payrollEnd')?.value;
    if (!start||!end) { showToast('Select period dates','warning'); return; }
    try {
        const res = await api('/api/staff/payroll.php','POST',{
            location_id:<?= (int)$locationId ?>, period_start:start, period_end:end
        });
        showToast('Payroll generated for ' + res.length + ' employees','success');
        setTimeout(()=>location.reload(),1200);
    } catch(e) { showToast(e.message,'error'); }
}

async function deactivateStaff(userId) {
    if (!confirm('Deactivate this staff member?')) return;
    await api('/api/staff/deactivate.php','POST',{user_id:userId});
    showToast('Staff member deactivated','warning');
    setTimeout(()=>location.reload(),1000);
}
</script>
JS;
require_once __DIR__ . '/templates/layout.php';
