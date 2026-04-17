<?php
require_once __DIR__ . '/config.php';

$message = '';
$error = '';
$step = 'init';

// DB接続テスト
try {
    $pdo = new PDO(
        'mysql:host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=utf8mb4',
        DB_USER,
        DB_PASS,
        [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]
    );
} catch (PDOException $e) {
    $error = 'データベース接続エラー: ' . $e->getMessage();
    $step = 'error';
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && $step !== 'error') {
    $action = $_POST['action'] ?? '';

    if ($action === 'create_tables') {
        try {
            // admin_users テーブル
            $pdo->exec("
                CREATE TABLE IF NOT EXISTS admin_users (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(50) NOT NULL UNIQUE,
                    password_hash VARCHAR(255) NOT NULL,
                    display_name VARCHAR(100) NOT NULL,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            ");

            // pdf_records テーブル
            $pdo->exec("
                CREATE TABLE IF NOT EXISTS pdf_records (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    staff_name VARCHAR(100) NOT NULL,
                    year INT NOT NULL,
                    month INT NOT NULL,
                    file_name VARCHAR(255) NOT NULL,
                    google_file_id VARCHAR(255) NOT NULL,
                    google_file_url VARCHAR(500),
                    file_size_bytes INT,
                    generated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    generated_by VARCHAR(50),
                    UNIQUE KEY unique_staff_month (staff_name, year, month)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            ");

            // staff_tokens テーブル（スタッフ個人閲覧用トークン）
            $pdo->exec("
                CREATE TABLE IF NOT EXISTS staff_tokens (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    staff_name VARCHAR(100) NOT NULL UNIQUE,
                    token VARCHAR(64) NOT NULL UNIQUE,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            ");

            // login_attempts テーブル
            $pdo->exec("
                CREATE TABLE IF NOT EXISTS login_attempts (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    ip_address VARCHAR(45) NOT NULL,
                    username VARCHAR(50),
                    attempted_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    success BOOLEAN DEFAULT FALSE,
                    INDEX idx_ip_time (ip_address, attempted_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            ");

            $message = 'テーブルの作成が完了しました。';
            $step = 'create_user';
        } catch (PDOException $e) {
            $error = 'テーブル作成エラー: ' . $e->getMessage();
        }
    } elseif ($action === 'reset_password') {
        $username = trim($_POST['username'] ?? '');
        $newPassword = $_POST['new_password'] ?? '';
        $confirmPassword = $_POST['confirm_password'] ?? '';

        if ($username === '' || $newPassword === '' || $confirmPassword === '') {
            $error = 'すべての項目を入力してください。';
        } elseif (strlen($newPassword) < 8) {
            $error = 'パスワードは8文字以上にしてください。';
        } elseif ($newPassword !== $confirmPassword) {
            $error = 'パスワードが一致しません。';
        } else {
            $stmt = $pdo->prepare('SELECT id FROM admin_users WHERE username = ?');
            $stmt->execute([$username]);
            $user = $stmt->fetch();

            if (!$user) {
                $error = 'そのユーザーIDは存在しません。';
            } else {
                $hash = password_hash($newPassword, PASSWORD_BCRYPT);
                $stmt = $pdo->prepare('UPDATE admin_users SET password_hash = ? WHERE id = ?');
                $stmt->execute([$hash, $user['id']]);
                $message = 'パスワードをリセットしました。新しいパスワードでログインしてください。';
            }
        }
        $step = 'done';

    } elseif ($action === 'create_admin') {
        $username = trim($_POST['username'] ?? '');
        $password = $_POST['password'] ?? '';
        $displayName = trim($_POST['display_name'] ?? '');

        if ($username === '' || $password === '' || $displayName === '') {
            $error = 'すべての項目を入力してください。';
            $step = 'create_user';
        } elseif (strlen($password) < 8) {
            $error = 'パスワードは8文字以上にしてください。';
            $step = 'create_user';
        } else {
            try {
                $hash = password_hash($password, PASSWORD_BCRYPT);
                $stmt = $pdo->prepare(
                    'INSERT INTO admin_users (username, password_hash, display_name) VALUES (?, ?, ?)'
                );
                $stmt->execute([$username, $hash, $displayName]);
                $message = '管理者ユーザーを作成しました。このファイル(setup.php)を必ず削除してください。';
                $step = 'done';
            } catch (PDOException $e) {
                if ($e->getCode() == 23000) {
                    $error = 'そのユーザーIDは既に存在します。';
                } else {
                    $error = 'ユーザー作成エラー: ' . $e->getMessage();
                }
                $step = 'create_user';
            }
        }
    }
} else {
    // 初期表示: テーブルが既に存在するかチェック
    if ($step !== 'error') {
        $tables = $pdo->query("SHOW TABLES")->fetchAll(PDO::FETCH_COLUMN);
        if (in_array('admin_users', $tables)) {
            $count = $pdo->query("SELECT COUNT(*) FROM admin_users")->fetchColumn();
            $step = $count > 0 ? 'done' : 'create_user';
        }
    }
}
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>勤怠管理 - 初期セットアップ</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #1a1a2e; color: #e0e0e0; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
        .setup-container { background: #16213e; border-radius: 12px; padding: 2rem; max-width: 500px; width: 90%; box-shadow: 0 4px 24px rgba(0,0,0,0.3); }
        h1 { font-size: 1.5rem; margin-bottom: 0.5rem; color: #fff; }
        .subtitle { color: #888; margin-bottom: 1.5rem; }
        .alert { padding: 0.75rem 1rem; border-radius: 8px; margin-bottom: 1rem; }
        .alert-success { background: #0a3d2a; color: #4ade80; border: 1px solid #166534; }
        .alert-error { background: #3d0a0a; color: #f87171; border: 1px solid #7f1d1d; }
        .alert-warning { background: #3d2e0a; color: #fbbf24; border: 1px solid #92400e; }
        label { display: block; margin-bottom: 0.3rem; color: #aaa; font-size: 0.9rem; }
        input[type="text"], input[type="password"] { width: 100%; padding: 0.6rem 0.8rem; border: 1px solid #334; border-radius: 8px; background: #0f1629; color: #fff; font-size: 1rem; margin-bottom: 1rem; }
        input:focus { outline: none; border-color: #4f8cff; }
        .btn { display: inline-block; padding: 0.7rem 1.5rem; border: none; border-radius: 8px; cursor: pointer; font-size: 1rem; color: #fff; }
        .btn-primary { background: #4f8cff; }
        .btn-primary:hover { background: #3a7ae0; }
        .step-indicator { display: flex; gap: 0.5rem; margin-bottom: 1.5rem; }
        .step { width: 2rem; height: 4px; border-radius: 2px; background: #334; }
        .step.active { background: #4f8cff; }
        .step.done { background: #4ade80; }
    </style>
</head>
<body>
    <div class="setup-container">
        <h1>初期セットアップ</h1>
        <p class="subtitle">勤怠管理 管理者ページ</p>

        <div class="step-indicator">
            <div class="step <?= in_array($step, ['init']) ? 'active' : (in_array($step, ['create_user', 'done']) ? 'done' : '') ?>"></div>
            <div class="step <?= $step === 'create_user' ? 'active' : ($step === 'done' ? 'done' : '') ?>"></div>
            <div class="step <?= $step === 'done' ? 'active done' : '' ?>"></div>
        </div>

        <?php if ($error): ?>
            <div class="alert alert-error"><?= htmlspecialchars($error) ?></div>
        <?php endif; ?>

        <?php if ($message): ?>
            <div class="alert alert-success"><?= htmlspecialchars($message) ?></div>
        <?php endif; ?>

        <?php if ($step === 'error'): ?>
            <p>config.php のデータベース設定を確認してください。</p>

        <?php elseif ($step === 'init'): ?>
            <p style="margin-bottom:1rem;">データベースにテーブルを作成します。</p>
            <form method="POST">
                <input type="hidden" name="action" value="create_tables">
                <button type="submit" class="btn btn-primary">テーブルを作成</button>
            </form>

        <?php elseif ($step === 'create_user'): ?>
            <p style="margin-bottom:1rem;">管理者ユーザーを作成します。</p>
            <form method="POST">
                <input type="hidden" name="action" value="create_admin">
                <label for="username">ユーザーID</label>
                <input type="text" id="username" name="username" required value="<?= htmlspecialchars($_POST['username'] ?? 'admin') ?>">
                <label for="display_name">表示名</label>
                <input type="text" id="display_name" name="display_name" required value="<?= htmlspecialchars($_POST['display_name'] ?? '総務担当') ?>">
                <label for="password">パスワード（8文字以上）</label>
                <input type="password" id="password" name="password" required minlength="8">
                <button type="submit" class="btn btn-primary">管理者ユーザーを作成</button>
            </form>

        <?php elseif ($step === 'done'): ?>
            <div class="alert alert-warning">
                セットアップ済みです。このページはパスワードリセットにも使用できます。
            </div>

            <h2 style="font-size:1.1rem; margin-top:1.5rem; margin-bottom:0.75rem; color:#fff;">パスワードリセット</h2>
            <form method="POST">
                <input type="hidden" name="action" value="reset_password">
                <label for="username">ユーザーID</label>
                <input type="text" id="username" name="username" required>
                <label for="new_password">新しいパスワード（8文字以上）</label>
                <input type="password" id="new_password" name="new_password" required minlength="8">
                <label for="confirm_password">新しいパスワード（確認）</label>
                <input type="password" id="confirm_password" name="confirm_password" required minlength="8">
                <button type="submit" class="btn btn-primary" style="margin-top:0.5rem;">パスワードをリセット</button>
            </form>

            <p style="margin-top:1.5rem;"><a href="index.php" style="color:#4f8cff;">ログインページへ →</a></p>
        <?php endif; ?>
    </div>
</body>
</html>
