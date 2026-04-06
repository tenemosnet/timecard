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

$oldName = trim($input['oldName'] ?? '');
$newName = trim($input['newName'] ?? '');

if ($oldName === '' || $newName === '') {
    echo json_encode(['success' => false, 'error' => '変更前・変更後の氏名を入力してください。']);
    exit;
}

if ($oldName === $newName) {
    echo json_encode(['success' => false, 'error' => '変更前と変更後の名前が同じです。']);
    exit;
}

try {
    require_once __DIR__ . '/api.php';

    $result = callGasApi('renameStaff', [
        'oldName' => $oldName,
        'newName' => $newName,
    ]);

    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
