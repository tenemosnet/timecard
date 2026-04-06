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

try {
    require_once __DIR__ . '/api.php';

    // 並び順更新
    if ($action === 'updateOrder') {
        $order = $input['order'] ?? [];
        if (!is_array($order) || count($order) === 0) {
            echo json_encode(['success' => false, 'error' => '並び順データが不正です。']);
            exit;
        }

        $result = callGasApi('updateStaffOrder', [
            'order' => $order,
        ]);
        echo json_encode($result);
        exit;
    }

    // 定時設定保存
    $staffName = trim($input['staffName'] ?? '');
    $contractedHours = floatval($input['contractedHours'] ?? 0);

    if ($staffName === '') {
        echo json_encode(['success' => false, 'error' => 'スタッフ名を指定してください。']);
        exit;
    }

    if ($contractedHours != 7.5 && $contractedHours != 8) {
        echo json_encode(['success' => false, 'error' => '定時は7.5または8を指定してください。']);
        exit;
    }

    $result = callGasApi('setStaffSetting', [
        'staffName' => $staffName,
        'contractedHours' => $contractedHours,
    ]);

    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
