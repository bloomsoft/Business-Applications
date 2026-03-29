<?php
require_once __DIR__ . '/core/bootstrap.php';

if (Auth::check()) {
    header('Location: /dashboard.php');
} else {
    header('Location: /login.php');
}
exit;
