<?php
require_once __DIR__ . '/config.php';

session_start();

function requireLogin(): void {
    // セッションタイムアウトチェック
    if (isset($_SESSION['last_activity'])) {
        if (time() - $_SESSION['last_activity'] > SESSION_TIMEOUT) {
            session_unset();
            session_destroy();
            header('Location: index.php?timeout=1');
            exit;
        }
    }

    if (!isset($_SESSION['user_id'])) {
        header('Location: index.php');
        exit;
    }

    // アクティビティ更新
    $_SESSION['last_activity'] = time();
}

function getDB(): PDO {
    static $pdo = null;
    if ($pdo === null) {
        $pdo = new PDO(
            'mysql:host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=utf8mb4',
            DB_USER,
            DB_PASS,
            [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
                PDO::ATTR_EMULATE_PREPARES => false,
            ]
        );
    }
    return $pdo;
}

function generateCsrfToken(): string {
    if (empty($_SESSION['csrf_token'])) {
        $_SESSION['csrf_token'] = bin2hex(random_bytes(32));
    }
    return $_SESSION['csrf_token'];
}

function verifyCsrfToken(string $token): bool {
    return isset($_SESSION['csrf_token']) && hash_equals($_SESSION['csrf_token'], $token);
}

function isLockedOut(string $ip): bool {
    $pdo = getDB();
    $stmt = $pdo->prepare(
        'SELECT COUNT(*) FROM login_attempts
         WHERE ip_address = ? AND success = 0
         AND attempted_at > DATE_SUB(NOW(), INTERVAL ? MINUTE)'
    );
    $stmt->execute([$ip, LOCKOUT_MINUTES]);
    return $stmt->fetchColumn() >= MAX_LOGIN_ATTEMPTS;
}

function recordLoginAttempt(string $ip, string $username, bool $success): void {
    $pdo = getDB();
    $stmt = $pdo->prepare(
        'INSERT INTO login_attempts (ip_address, username, success) VALUES (?, ?, ?)'
    );
    $stmt->execute([$ip, $username, $success ? 1 : 0]);
}
