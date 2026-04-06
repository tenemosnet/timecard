<?php
require_once __DIR__ . '/auth.php';
requireLogin();

header('Content-Type: application/json');

$year = (int)($_GET['year'] ?? date('Y'));
$staffName = trim($_GET['staff'] ?? '');

try {
    require_once __DIR__ . '/api.php';

    $params = ['year' => $year];
    if ($staffName !== '') {
        $params['staffName'] = $staffName;
    }

    $result = callGasApi('listPDFs', $params);
    echo json_encode($result);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
