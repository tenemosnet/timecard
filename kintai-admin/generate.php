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
$year = (int)($input['year'] ?? 0);
$month = (int)($input['month'] ?? 0);

if ($year < 2020 || $year > 2100 || $month < 1 || $month > 12) {
    echo json_encode(['success' => false, 'error' => '年月の指定が不正です。']);
    exit;
}

if ($staffName === '') {
    echo json_encode(['success' => false, 'error' => 'スタッフ名を指定してください。']);
    exit;
}

try {
    require_once __DIR__ . '/api.php';

    $result = callGasApi('generatePDF', [
        'staffName' => $staffName,
        'year' => $year,
        'month' => $month,
    ]);

    if ($result['success']) {
        // PDF生成履歴をDBに記録
        $pdo = getDB();
        $stmt = $pdo->prepare(
            'INSERT INTO pdf_records (staff_name, year, month, file_name, google_file_id, google_file_url, generated_by)
             VALUES (?, ?, ?, ?, ?, ?, ?)
             ON DUPLICATE KEY UPDATE
             file_name = VALUES(file_name),
             google_file_id = VALUES(google_file_id),
             google_file_url = VALUES(google_file_url),
             generated_by = VALUES(generated_by),
             generated_at = CURRENT_TIMESTAMP'
        );
        $stmt->execute([
            $staffName,
            $year,
            $month,
            $result['data']['fileName'],
            $result['data']['fileId'],
            $result['data']['url'] ?? '',
            $_SESSION['display_name'],
        ]);
    }

    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
