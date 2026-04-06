<?php
require_once __DIR__ . '/auth.php';

// 既にログイン済みならダッシュボードへ
if (isset($_SESSION['user_id'])) {
    header('Location: dashboard.php');
    exit;
}

$error = '';
$timeout = isset($_GET['timeout']);

// ログイン処理
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $token = $_POST['csrf_token'] ?? '';
    if (!verifyCsrfToken($token)) {
        $error = '不正なリクエストです。';
    } else {
        $username = trim($_POST['username'] ?? '');
        $password = $_POST['password'] ?? '';
        $ip = $_SERVER['REMOTE_ADDR'];

        if (isLockedOut($ip)) {
            $error = 'ログイン試行回数の上限に達しました。' . LOCKOUT_MINUTES . '分後に再試行してください。';
        } elseif ($username === '' || $password === '') {
            $error = 'ユーザーIDとパスワードを入力してください。';
        } else {
            $pdo = getDB();
            $stmt = $pdo->prepare('SELECT id, password_hash, display_name FROM admin_users WHERE username = ?');
            $stmt->execute([$username]);
            $user = $stmt->fetch();

            if ($user && password_verify($password, $user['password_hash'])) {
                recordLoginAttempt($ip, $username, true);
                session_regenerate_id(true);
                $_SESSION['user_id'] = $user['id'];
                $_SESSION['display_name'] = $user['display_name'];
                $_SESSION['last_activity'] = time();
                header('Location: dashboard.php');
                exit;
            } else {
                recordLoginAttempt($ip, $username, false);
                $error = 'ユーザーIDまたはパスワードが正しくありません。';
            }
        }
    }
}

$csrfToken = generateCsrfToken();
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>勤怠管理 - ログイン</title>
    <link rel="stylesheet" href="assets/style.css">
</head>
<body class="login-page">
    <div class="login-container">
        <h1>勤怠管理</h1>
        <p class="login-subtitle">管理者ページ</p>

        <?php if ($timeout): ?>
            <div class="alert alert-warning">セッションがタイムアウトしました。再度ログインしてください。</div>
        <?php endif; ?>

        <?php if ($error): ?>
            <div class="alert alert-error"><?= htmlspecialchars($error) ?></div>
        <?php endif; ?>

        <form method="POST" action="">
            <input type="hidden" name="csrf_token" value="<?= htmlspecialchars($csrfToken) ?>">
            <div class="form-group">
                <label for="username">ユーザーID</label>
                <input type="text" id="username" name="username" required autofocus
                       value="<?= htmlspecialchars($_POST['username'] ?? '') ?>">
            </div>
            <div class="form-group">
                <label for="password">パスワード</label>
                <input type="password" id="password" name="password" required>
            </div>
            <button type="submit" class="btn btn-primary btn-full">ログイン</button>
        </form>
    </div>
</body>
</html>
