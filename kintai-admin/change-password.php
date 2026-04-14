<?php
require_once __DIR__ . '/auth.php';
requireLogin();

$csrfToken = generateCsrfToken();
$displayName = htmlspecialchars($_SESSION['display_name'] ?? '管理者');
$success = '';
$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $token = $_POST['csrf_token'] ?? '';
    if (!verifyCsrfToken($token)) {
        $error = '不正なリクエストです。';
    } else {
        $currentPassword = $_POST['current_password'] ?? '';
        $newPassword = $_POST['new_password'] ?? '';
        $confirmPassword = $_POST['confirm_password'] ?? '';

        if ($currentPassword === '' || $newPassword === '' || $confirmPassword === '') {
            $error = 'すべての項目を入力してください。';
        } elseif (strlen($newPassword) < 8) {
            $error = '新しいパスワードは8文字以上にしてください。';
        } elseif ($newPassword !== $confirmPassword) {
            $error = '新しいパスワードが一致しません。';
        } else {
            $pdo = getDB();
            $stmt = $pdo->prepare('SELECT password_hash FROM admin_users WHERE id = ?');
            $stmt->execute([$_SESSION['user_id']]);
            $user = $stmt->fetch();

            if (!$user || !password_verify($currentPassword, $user['password_hash'])) {
                $error = '現在のパスワードが正しくありません。';
            } else {
                $newHash = password_hash($newPassword, PASSWORD_BCRYPT);
                $stmt = $pdo->prepare('UPDATE admin_users SET password_hash = ? WHERE id = ?');
                $stmt->execute([$newHash, $_SESSION['user_id']]);
                $success = 'パスワードを変更しました。';
            }
        }
    }
    // CSRF トークンを再生成
    $_SESSION['csrf_token'] = '';
    $csrfToken = generateCsrfToken();
}
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>勤怠管理 - パスワード変更</title>
    <link rel="stylesheet" href="assets/style.css">
</head>
<body>
    <header class="header">
        <h1>勤怠管理 管理者ページ</h1>
        <div class="header-right">
            <span class="user-name"><?= $displayName ?></span>
            <a href="dashboard.php" class="btn btn-sm btn-outline">ダッシュボードへ戻る</a>
        </div>
    </header>

    <main class="container">
        <section class="card" style="max-width: 480px; margin: 0 auto;">
            <h2>パスワード変更</h2>

            <?php if ($success): ?>
                <div class="alert alert-success"><?= htmlspecialchars($success) ?></div>
            <?php endif; ?>

            <?php if ($error): ?>
                <div class="alert alert-error"><?= htmlspecialchars($error) ?></div>
            <?php endif; ?>

            <form method="POST" action="">
                <input type="hidden" name="csrf_token" value="<?= htmlspecialchars($csrfToken) ?>">
                <div class="form-group">
                    <label for="current_password">現在のパスワード</label>
                    <input type="password" id="current_password" name="current_password" required>
                </div>
                <div class="form-group">
                    <label for="new_password">新しいパスワード（8文字以上）</label>
                    <input type="password" id="new_password" name="new_password" required minlength="8">
                </div>
                <div class="form-group">
                    <label for="confirm_password">新しいパスワード（確認）</label>
                    <input type="password" id="confirm_password" name="confirm_password" required minlength="8">
                </div>
                <button type="submit" class="btn btn-primary btn-full">パスワードを変更</button>
            </form>
        </section>
    </main>
</body>
</html>
