<?php
require_once __DIR__ . '/core/bootstrap.php';

$orderId  = (int) get('order_id');
$success  = false;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $tenantId = (int) post('tenant_id');
    if ($tenantId) {
        CustomerManager::submitFeedback($_POST, $tenantId);
        $success = true;
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Rate Your Experience</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        body { background: #f8fafc; }
        .star-btn { font-size: 2rem; cursor: pointer; color: #d1d5db; transition: color .1s; }
        .star-btn.active, .star-btn:hover { color: #f59e0b; }
    </style>
</head>
<body>
<div class="container py-5">
    <div class="row justify-content-center">
        <div class="col-sm-10 col-md-7 col-lg-5">
            <?php if ($success): ?>
            <div class="card shadow text-center p-5">
                <i class="bi bi-heart-fill text-danger fs-1"></i>
                <h3 class="mt-3">Thank you!</h3>
                <p class="text-muted">Your feedback helps us improve.</p>
                <a href="javascript:history.back()" class="btn btn-outline-secondary mt-2">Back</a>
            </div>
            <?php else: ?>
            <div class="card shadow p-4">
                <h4 class="fw-bold text-center mb-1">How was your experience?</h4>
                <p class="text-center text-muted mb-4">Your feedback helps us serve you better</p>

                <form method="POST">
                    <input type="hidden" name="order_id"  value="<?= $orderId ?>">
                    <input type="hidden" name="tenant_id" value="" id="tenantField">
                    <input type="hidden" name="source"    value="qr">
                    <input type="hidden" name="rating"    id="ratingValue" value="">

                    <!-- Overall Rating -->
                    <div class="mb-4 text-center">
                        <label class="form-label fw-600 d-block">Overall Experience</label>
                        <div id="starRow">
                            <?php for ($i = 1; $i <= 5; $i++): ?>
                            <span class="star-btn" data-val="<?= $i ?>"
                                  onclick="setRating(<?= $i ?>)">&#9733;</span>
                            <?php endfor; ?>
                        </div>
                        <div class="text-muted small mt-1" id="ratingLabel"></div>
                    </div>

                    <!-- Sub-ratings -->
                    <div class="row g-3 mb-3">
                        <div class="col-4 text-center">
                            <label class="form-label small">Food</label>
                            <input type="range" name="food_rating" class="form-range" min="1" max="5" value="3">
                            <span class="small text-muted" id="foodVal">3</span>
                        </div>
                        <div class="col-4 text-center">
                            <label class="form-label small">Service</label>
                            <input type="range" name="service_rating" class="form-range" min="1" max="5" value="3">
                            <span class="small text-muted" id="serviceVal">3</span>
                        </div>
                        <div class="col-4 text-center">
                            <label class="form-label small">Ambiance</label>
                            <input type="range" name="ambiance_rating" class="form-range" min="1" max="5" value="3">
                            <span class="small text-muted" id="ambianceVal">3</span>
                        </div>
                    </div>

                    <!-- Comment -->
                    <div class="mb-4">
                        <label class="form-label">Tell us more (optional)</label>
                        <textarea name="comment" class="form-control" rows="3"
                                  placeholder="What did you love? What can we improve?"></textarea>
                    </div>

                    <div class="d-grid">
                        <button type="submit" class="btn btn-warning btn-lg fw-bold"
                                onclick="return validateFeedback()">
                            <i class="bi bi-send me-2"></i>Submit Feedback
                        </button>
                    </div>
                </form>
            </div>
            <?php endif; ?>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
function setRating(val) {
    document.getElementById('ratingValue').value = val;
    const labels = ['','Poor','Fair','Good','Very Good','Excellent'];
    document.getElementById('ratingLabel').textContent = labels[val] || '';
    document.querySelectorAll('.star-btn').forEach(s => {
        s.classList.toggle('active', parseInt(s.dataset.val) <= val);
    });
}

function validateFeedback() {
    if (!document.getElementById('ratingValue').value) {
        alert('Please select an overall rating'); return false;
    }
    return true;
}

// Sub-rating live display
['food','service','ambiance'].forEach(type => {
    const input = document.querySelector(`[name="${type}_rating"]`);
    const span  = document.getElementById(type + 'Val');
    if (input && span) {
        input.addEventListener('input', () => span.textContent = input.value);
    }
});

// Resolve tenant_id from URL context
const params = new URLSearchParams(location.search);
if (params.get('order_id')) {
    // tenant_id will be resolved server-side via order_id if provided
}
</script>
</body>
</html>
