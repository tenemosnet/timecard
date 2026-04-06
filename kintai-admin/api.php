<?php
require_once __DIR__ . '/config.php';

function callGasApi(string $action, array $params = []): array {
    $payload = array_merge([
        'action' => $action,
        'apiKey' => GAS_API_KEY,
    ], $params);

    $ch = curl_init(GAS_API_URL);
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_POSTFIELDS => json_encode($payload),
        CURLOPT_HTTPHEADER => ['Content-Type: application/json'],
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,  // GASはリダイレクトする
        CURLOPT_TIMEOUT => 120,          // PDF生成は時間がかかる場合がある
    ]);

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error = curl_error($ch);
    curl_close($ch);

    if ($response === false) {
        throw new Exception('GAS API接続エラー: ' . $error);
    }

    if ($httpCode !== 200) {
        throw new Exception('GAS APIエラー: HTTP ' . $httpCode);
    }

    $decoded = json_decode($response, true);
    if ($decoded === null) {
        throw new Exception('GAS APIレスポンス解析エラー');
    }

    return $decoded;
}
