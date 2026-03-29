<?php
require_once __DIR__ . '/core/bootstrap.php';
Auth::requireAuth();

$locationId = Auth::locationId();
$tenantId   = Auth::tenantId();
$date       = get('date', date('Y-m-d'));

$reservations = Database::fetchAll(
    "SELECT r.*, t.table_number, t.capacity,
            c.first_name || ' ' || COALESCE(c.last_name,'') AS customer_name,
            c.phone AS customer_phone
     FROM reservations r
     LEFT JOIN restaurant_tables t ON t.table_id = r.table_id
     LEFT JOIN customers c ON c.customer_id = r.customer_id
     WHERE r.location_id = ? AND r.reservation_date = ?
     ORDER BY r.reservation_time",
    [$locationId, $date]
);

$tables = Database::fetchAll(
    "SELECT * FROM restaurant_tables WHERE location_id = ? ORDER BY table_number",
    [$locationId]
);

$pageTitle  = 'Reservations';
$activeMenu = 'reservations';

// Handle actions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');

    if ($action === 'create_reservation') {
        $confCode = strtoupper(substr(uniqid(), -6));
        Database::insert(
            "INSERT INTO reservations
                (location_id, customer_id, table_id, party_size, reservation_date,
                 reservation_time, duration_min, notes, confirmation_code)
             VALUES (?,?,?,?,?,?,?,?,?)",
            [
                $locationId,
                post('customer_id') ?: null,
                post('table_id')    ?: null,
                (int) post('party_size'),
                post('reservation_date'),
                post('reservation_time'),
                (int) post('duration_min') ?: 90,
                post('notes'),
                $confCode,
            ]
        );
        flash('success', "Reservation created. Confirmation: $confCode");
        redirect('/reservations.php?date=' . post('reservation_date'));
    }

    if ($action === 'update_status') {
        Database::query(
            "UPDATE reservations SET status = ? WHERE reservation_id = ?",
            [post('status'), (int)post('reservation_id')]
        );
        flash('info', 'Reservation status updated');
        redirect('/reservations.php?date=' . $date);
    }
}

ob_start();
?>
<!-- Date Nav -->
<div class="d-flex justify-content-between align-items-center mb-4">
    <div class="d-flex align-items-center gap-3">
        <a href="?date=<?= date('Y-m-d', strtotime($date . ' -1 day')) ?>"
           class="btn btn-sm btn-outline-secondary"><i class="bi bi-chevron-left"></i></a>
        <input type="date" class="form-control form-control-sm" value="<?= $date ?>"
               onchange="location.href='?date='+this.value" style="width:160px">
        <a href="?date=<?= date('Y-m-d', strtotime($date . ' +1 day')) ?>"
           class="btn btn-sm btn-outline-secondary"><i class="bi bi-chevron-right"></i></a>
        <a href="?date=<?= date('Y-m-d') ?>" class="btn btn-sm btn-outline-primary">Today</a>
    </div>
    <div class="d-flex gap-2">
        <span class="badge bg-secondary fs-6"><?= count($reservations) ?> reservations</span>
        <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addReservationModal">
            <i class="bi bi-calendar-plus me-1"></i>New Reservation
        </button>
    </div>
</div>

<!-- Summary Cards -->
<div class="row g-2 mb-4">
    <?php
    $statusCounts = ['confirmed'=>0, 'seated'=>0, 'completed'=>0, 'no-show'=>0, 'cancelled'=>0];
    foreach ($reservations as $r) $statusCounts[$r['status']] = ($statusCounts[$r['status']] ?? 0) + 1;
    $totalCovers = array_sum(array_column($reservations, 'party_size'));
    ?>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5"><?= $totalCovers ?></div><div class="text-muted small">Total Covers</div>
    </div></div></div>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5 text-primary"><?= $statusCounts['confirmed'] ?></div><div class="text-muted small">Confirmed</div>
    </div></div></div>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5 text-success"><?= $statusCounts['seated'] ?></div><div class="text-muted small">Seated</div>
    </div></div></div>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5"><?= $statusCounts['completed'] ?></div><div class="text-muted small">Completed</div>
    </div></div></div>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5 text-danger"><?= $statusCounts['no-show'] ?></div><div class="text-muted small">No-Shows</div>
    </div></div></div>
    <div class="col-6 col-md-2"><div class="card text-center shadow-sm"><div class="card-body py-2">
        <div class="fw-bold fs-5 text-warning"><?= $statusCounts['cancelled'] ?></div><div class="text-muted small">Cancelled</div>
    </div></div></div>
</div>

<!-- Reservations Timeline -->
<div class="card shadow-sm">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-hover align-middle mb-0">
                <thead class="table-light">
                    <tr><th>Time</th><th>Guest</th><th>Party</th><th>Table</th><th>Duration</th><th>Status</th><th>Notes</th><th>Actions</th></tr>
                </thead>
                <tbody>
                    <?php foreach ($reservations as $r): ?>
                    <tr>
                        <td>
                            <div class="fw-bold"><?= date('g:i A', strtotime($r['reservation_time'])) ?></div>
                        </td>
                        <td>
                            <div class="fw-600"><?= sanitize($r['customer_name'] ?? 'Walk-in') ?></div>
                            <small class="text-muted"><?= sanitize($r['customer_phone'] ?? '') ?></small>
                        </td>
                        <td><span class="badge bg-primary"><?= $r['party_size'] ?></span></td>
                        <td><?= $r['table_number'] ? 'Table ' . sanitize($r['table_number']) : 'Unassigned' ?></td>
                        <td><?= $r['duration_min'] ?>m</td>
                        <td><?= statusBadge($r['status']) ?></td>
                        <td><small class="text-muted"><?= sanitize($r['notes'] ?? '') ?></small></td>
                        <td>
                            <div class="dropdown">
                                <button class="btn btn-sm btn-outline-secondary dropdown-toggle" data-bs-toggle="dropdown">
                                    Action
                                </button>
                                <ul class="dropdown-menu">
                                    <?php foreach (['confirmed','seated','completed','no-show','cancelled'] as $st): ?>
                                    <li>
                                        <form method="POST" class="d-inline">
                                            <input type="hidden" name="action" value="update_status">
                                            <input type="hidden" name="reservation_id" value="<?= $r['reservation_id'] ?>">
                                            <input type="hidden" name="status" value="<?= $st ?>">
                                            <button type="submit" class="dropdown-item"><?= ucfirst(str_replace('-', ' ', $st)) ?></button>
                                        </form>
                                    </li>
                                    <?php endforeach; ?>
                                </ul>
                            </div>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                    <?php if (empty($reservations)): ?>
                    <tr><td colspan="8" class="text-center text-muted py-4">No reservations for this date</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Add Reservation Modal -->
<div class="modal fade" id="addReservationModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-calendar-plus me-2"></i>New Reservation</h5>
                <button class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body row g-3">
                    <input type="hidden" name="action" value="create_reservation">
                    <div class="col-6"><label class="form-label">Date *</label>
                        <input type="date" name="reservation_date" class="form-control" value="<?= $date ?>" required></div>
                    <div class="col-6"><label class="form-label">Time *</label>
                        <input type="time" name="reservation_time" class="form-control" required></div>
                    <div class="col-6"><label class="form-label">Party Size *</label>
                        <input type="number" name="party_size" class="form-control" min="1" max="50" value="2" required></div>
                    <div class="col-6"><label class="form-label">Duration (min)</label>
                        <input type="number" name="duration_min" class="form-control" value="90" min="30" max="360"></div>
                    <div class="col-12"><label class="form-label">Table</label>
                        <select name="table_id" class="form-select">
                            <option value="">Auto-assign</option>
                            <?php foreach ($tables as $t): ?>
                            <option value="<?= $t['table_id'] ?>">
                                Table <?= sanitize($t['table_number']) ?> (<?= $t['capacity'] ?> seats)
                            </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-12"><label class="form-label">Guest Phone / Name</label>
                        <input type="text" name="guest_info" class="form-control" placeholder="Search or enter name/phone"></div>
                    <div class="col-12"><label class="form-label">Notes</label>
                        <textarea name="notes" class="form-control" rows="2" placeholder="Allergies, special requests..."></textarea></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="submit" class="btn btn-success">Create Reservation</button>
                </div>
            </form>
        </div>
    </div>
</div>

<?php
$content = ob_get_clean();
require_once __DIR__ . '/templates/layout.php';
