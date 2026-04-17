<?php
require_once __DIR__ . '/auth.php';
requireLogin();

header('Content-Type: application/json');

require_once __DIR__ . '/api.php';

try {
    // GET: 検索
    if ($_SERVER['REQUEST_METHOD'] === 'GET') {
        $date = $_GET['date'] ?? '';
        $staffName = $_GET['staff'] ?? '';

        if ($date === '') {
            echo json_encode(['success' => false, 'error' => '日付を指定してください。']);
            exit;
        }

        $result = callGasApi('getClockLog', [
            'date' => $date,
            'staffName' => $staffName,
        ]);
        echo json_encode($result);
        exit;
    }

    // POST: 追加・編集・削除
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

    switch ($action) {
        case 'add':
            $staffName = trim($input['staffName'] ?? '');
            $type = trim($input['type'] ?? '');
            $date = trim($input['date'] ?? '');
            $time = trim($input['time'] ?? '');

            if ($staffName === '' || $type === '' || $date === '') {
                echo json_encode(['success' => false, 'error' => '必須項目が不足しています。']);
                exit;
            }

            $result = callGasApi('addClockEntry', [
                'staffName' => $staffName,
                'type' => $type,
                'date' => $date,
                'time' => $time,
            ]);
            break;

        case 'edit':
            $rowIndex = intval($input['rowIndex'] ?? 0);
            $newType = trim($input['newType'] ?? '');
            $newTime = trim($input['newTime'] ?? '');
            $expectedStaff = trim($input['expectedStaff'] ?? '');
            $expectedDate = trim($input['expectedDate'] ?? '');

            if ($rowIndex < 2) {
                echo json_encode(['success' => false, 'error' => '行番号が不正です。']);
                exit;
            }

            $result = callGasApi('editClockEntry', [
                'rowIndex' => $rowIndex,
                'newType' => $newType,
                'newTime' => $newTime,
                'expectedStaff' => $expectedStaff,
                'expectedDate' => $expectedDate,
            ]);
            break;

        case 'delete':
            $rowIndex = intval($input['rowIndex'] ?? 0);
            $expectedStaff = trim($input['expectedStaff'] ?? '');
            $expectedDate = trim($input['expectedDate'] ?? '');

            if ($rowIndex < 2) {
                echo json_encode(['success' => false, 'error' => '行番号が不正です。']);
                exit;
            }

            $result = callGasApi('deleteClockEntry', [
                'rowIndex' => $rowIndex,
                'expectedStaff' => $expectedStaff,
                'expectedDate' => $expectedDate,
            ]);
            break;

        default:
            $result = ['success' => false, 'error' => '不明な操作です: ' . $action];
    }

    echo json_encode($result);

} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
