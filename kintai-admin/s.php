<?php
// 短縮URL: s.php?t=TOKEN → staff_view.php?token=TOKEN
$t = trim($_GET['t'] ?? '');
if ($t === '') {
    http_response_code(400);
    echo 'URLが無効です。';
    exit;
}
header('Location: staff_view.php?token=' . urlencode($t));
exit;
