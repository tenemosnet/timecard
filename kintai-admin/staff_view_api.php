<?php
require_once __DIR__ . '/config.php';

header('Content-Type: application/json');

// トークンで認証（管理者ログイン不要）
$token = trim($_GET['token'] ?? '');
if ($token === '') {
    echo json_encode(['success' => false, 'error' => '認証エラー']);
    exit;
}

try {
    $pdo = new PDO(
        'mysql:host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=utf8mb4',
        DB_USER, DB_PASS,
        [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]
    );
    $stmt = $pdo->prepare('SELECT staff_name FROM staff_tokens WHERE token = ?');
    $stmt->execute([$token]);
    $staffName = $stmt->fetchColumn();
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => 'データベースエラー']);
    exit;
}

if (!$staffName) {
    echo json_encode(['success' => false, 'error' => '認証エラー']);
    exit;
}

$year = (int)($_GET['year'] ?? date('Y'));
$month = $_GET['month'] ?? '';

try {
    require_once __DIR__ . '/api.php';

    $params = [
        'year' => $year,
        'staffName' => $staffName,
    ];
    if ($month !== '') {
        $params['month'] = (int)$month;
    }

    $result = callGasApi('listPDFs', $params);
    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
