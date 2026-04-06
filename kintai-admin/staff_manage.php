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

$action = $input['action'] ?? '';
$staffName = trim($input['staffName'] ?? '');

if ($staffName === '') {
    echo json_encode(['success' => false, 'error' => 'スタッフ名を入力してください。']);
    exit;
}

try {
    require_once __DIR__ . '/api.php';

    if ($action === 'add') {
        $contractedHours = floatval($input['contractedHours'] ?? 8);
        $result = callGasApi('addStaff', [
            'staffName' => $staffName,
            'contractedHours' => $contractedHours,
        ]);
    } elseif ($action === 'remove') {
        $result = callGasApi('removeStaff', [
            'staffName' => $staffName,
        ]);
    } else {
        $result = ['success' => false, 'error' => '不明な操作です。'];
    }

    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
