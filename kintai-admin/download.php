<?php
require_once __DIR__ . '/auth.php';
requireLogin();

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
