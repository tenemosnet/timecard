<?php
require_once __DIR__ . '/auth.php';
requireLogin();

header('Content-Type: application/json');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    echo json_encode(['success' => false, 'error' => '不正なリクエストです。']);
    exit;
}

$input = json_decode(file_get_contents('php://input'), true);
$token = $input['csrf_token'] ?? '';

if (!verifyCsrfToken($token)) {
    echo json_encode(['success' => false, 'error' => 'CSRFトークンが無効です。']);
    exit;
}

$staffName = trim($input['staffName'] ?? '');
if ($staffName === '') {
    echo json_encode(['success' => false, 'error' => 'スタッフ名を指定してください。']);
    exit;
}

try {
    $pdo = getDB();

    // 既存トークンを確認
    $stmt = $pdo->prepare('SELECT token FROM staff_tokens WHERE staff_name = ?');
    $stmt->execute([$staffName]);
    $existing = $stmt->fetchColumn();

    $action = $input['action'] ?? 'get_or_create';

    if ($action === 'regenerate' || !$existing) {
        // 新規発行 or 再発行
        $newToken = bin2hex(random_bytes(32));

        $stmt = $pdo->prepare(
            'INSERT INTO staff_tokens (staff_name, token) VALUES (?, ?)
             ON DUPLICATE KEY UPDATE token = VALUES(token), created_at = CURRENT_TIMESTAMP'
        );
        $stmt->execute([$staffName, $newToken]);

        echo json_encode([
            'success' => true,
            'token' => $newToken,
            'isNew' => true,
        ]);
    } else {
        echo json_encode([
            'success' => true,
            'token' => $existing,
            'isNew' => false,
        ]);
    }
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
