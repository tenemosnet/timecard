<?php
require_once __DIR__ . '/config.php';

// 管理者ログイン または スタッフトークンで認証
$staffToken = trim($_GET['token'] ?? '');
if ($staffToken !== '') {
    // トークン認証
    try {
        $pdo = new PDO('mysql:host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=utf8mb4', DB_USER, DB_PASS, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);
        $stmt = $pdo->prepare('SELECT staff_name FROM staff_tokens WHERE token = ?');
        $stmt->execute([$staffToken]);
        if (!$stmt->fetchColumn()) {
            http_response_code(403);
            echo '認証エラー';
            exit;
        }
    } catch (Exception $e) {
        http_response_code(500);
        echo 'エラー';
        exit;
    }
} else {
    // 管理者認証
    require_once __DIR__ . '/auth.php';
    requireLogin();
}

$fileId = $_GET['id'] ?? '';

if ($fileId === '') {
    http_response_code(400);
    echo 'ファイルIDが指定されていません。';
    exit;
}

try {
    require_once __DIR__ . '/api.php';

    $result = callGasApi('getPDFContent', ['fileId' => $fileId]);

    if (!$result['success']) {
        http_response_code(500);
        echo 'PDF取得エラー: ' . ($result['error'] ?? '不明なエラー');
        exit;
    }

    $pdfData = base64_decode($result['data']['base64']);
    $fileName = $result['data']['fileName'];

    header('Content-Type: application/pdf');
    header('Content-Disposition: attachment; filename="' . $fileName . '"');
    header('Content-Length: ' . strlen($pdfData));
    header('Cache-Control: no-cache, no-store, must-revalidate');
    echo $pdfData;
} catch (Exception $e) {
    http_response_code(500);
    echo 'エラー: ' . $e->getMessage();
}
