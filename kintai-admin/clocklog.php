<?php
require_once __DIR__ . '/auth.php';
requireLogin();

$csrfToken = generateCsrfToken();
$displayName = htmlspecialchars($_SESSION['display_name'] ?? '管理者');

// スタッフ一覧をGASから取得
$staffList = [];
$staffNames = [];
$staffError = '';
try {
    require_once __DIR__ . '/api.php';
    $result = callGasApi('getStaffList');
    if ($result['success']) {
        $staffList = $result['data'];
        $staffNames = array_column($staffList, 'name');
    }
} catch (Exception $e) {
    $staffError = 'スタッフ一覧の取得に失敗しました。';
}
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>勤怠管理 - 打刻データ修正</title>
    <link rel="stylesheet" href="assets/style.css">
</head>
<body>
    <header class="header">
        <div class="header-left">
            <h1>勤怠管理</h1>
            <nav class="nav-links">
                <a href="dashboard.php">ダッシュボード</a>
                <a href="clocklog.php" class="active">打刻データ修正</a>
                <a href="#" onclick="return openStaffSelect(event)">スタッフ閲覧</a>
            </nav>
        </div>
        <div class="header-right">
            <span class="user-name"><?= $displayName ?></span>
            <a href="change-password.php" class="btn btn-sm btn-outline">パスワード変更</a>
            <a href="logout.php" class="btn btn-sm btn-outline">ログアウト</a>
        </div>
    </header>

    <main class="container">
        <!-- 検索セクション -->
        <section class="card">
            <h2>打刻データ検索</h2>

            <?php if ($staffError): ?>
                <div class="alert alert-error"><?= htmlspecialchars($staffError) ?></div>
            <?php endif; ?>

            <div class="search-form">
                <div class="form-group">
                    <label for="searchStaff">スタッフ</label>
                    <select id="searchStaff">
                        <?php foreach ($staffNames as $name): ?>
                            <option value="<?= htmlspecialchars($name) ?>"><?= htmlspecialchars($name) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                <div class="form-group">
                    <label for="searchDate">日付</label>
                    <input type="date" id="searchDate" value="<?= date('Y-m-d') ?>">
                </div>
                <div class="form-group" style="flex: 0;">
                    <label>&nbsp;</label>
                    <button id="btnSearch" class="btn btn-primary">検索</button>
                </div>
            </div>
        </section>

        <!-- 検索結果セクション -->
        <section class="card" id="resultSection" style="display: none;">
            <h2 id="resultTitle">検索結果</h2>

            <div id="resultBody">
                <!-- JSで描画 -->
            </div>

            <button id="btnAdd" class="btn-add">＋ 新規追加</button>
        </section>

        <!-- ステータスメッセージ -->
        <div id="statusMessage" class="alert" style="display: none;"></div>
    </main>

    <!-- 追加・編集モーダル -->
    <div class="modal-overlay" id="entryModal">
        <div class="modal-content">
            <h3 id="modalTitle">打刻を追加</h3>

            <div class="modal-form-row">
                <div class="form-group" id="modalStaffGroup">
                    <label for="modalStaff">スタッフ</label>
                    <select id="modalStaff">
                        <?php foreach ($staffNames as $name): ?>
                            <option value="<?= htmlspecialchars($name) ?>"><?= htmlspecialchars($name) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>

                <div class="form-group" id="modalDateGroup">
                    <label for="modalDate">日付</label>
                    <input type="date" id="modalDate">
                </div>

                <div class="form-group">
                    <label for="modalType">種別</label>
                    <select id="modalType">
                        <option value="入室">入室</option>
                        <option value="退室">退室</option>
                        <option value="有給">有給</option>
                    </select>
                </div>

                <div class="form-group" id="modalTimeGroup">
                    <label for="modalTime">時刻</label>
                    <input type="time" id="modalTime">
                </div>
            </div>

            <div class="modal-actions">
                <button class="btn btn-secondary" id="btnModalCancel">キャンセル</button>
                <button class="btn btn-primary" id="btnModalSave">保存</button>
            </div>
        </div>
    </div>

    <!-- トースト通知 -->
    <div class="toast" id="toast"></div>

    <input type="hidden" id="csrfToken" value="<?= htmlspecialchars($csrfToken) ?>">
    <input type="hidden" id="staffNamesJson" value="<?= htmlspecialchars(json_encode($staffNames)) ?>">

    <!-- スタッフ選択モーダル -->
    <div class="modal-overlay" id="staffSelectModal">
        <div class="modal-content" style="max-width:400px;">
            <h3>スタッフ閲覧</h3>
            <div class="form-group">
                <label for="staffSelectName">スタッフを選択</label>
                <select id="staffSelectName">
                    <?php foreach ($staffNames as $name): ?>
                        <option value="<?= htmlspecialchars($name) ?>"><?= htmlspecialchars($name) ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="modal-actions">
                <button class="btn btn-secondary" id="staffSelectCancel">キャンセル</button>
                <button class="btn btn-primary" id="staffSelectOpen">開く</button>
            </div>
        </div>
    </div>

    <footer style="text-align:center; padding:1.5rem; color:#8a7f6e; font-size:0.8rem;">
        勤怠管理システム ver3.0
    </footer>

    <script>
    function openStaffSelect(event) {
        event.preventDefault();
        var modal = document.getElementById('staffSelectModal');
        modal.classList.add('active');
        return false;
    }
    document.getElementById('staffSelectCancel').addEventListener('click', function() {
        document.getElementById('staffSelectModal').classList.remove('active');
    });
    document.getElementById('staffSelectOpen').addEventListener('click', function() {
        var name = document.getElementById('staffSelectName').value;
        if (!name) return;
        document.getElementById('staffSelectModal').classList.remove('active');
        var csrfToken = document.getElementById('csrfToken').value;
        fetch('staff_token.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ csrf_token: csrfToken, staffName: name, action: 'get_or_create' }),
        })
        .then(function(res) { return res.json(); })
        .then(function(result) {
            if (result.success) window.location.href = 'staff_view.php?token=' + result.token;
            else alert('エラー: ' + (result.error || ''));
        })
        .catch(function(err) { alert('通信エラー: ' + err.message); });
    });
    document.getElementById('staffSelectModal').addEventListener('click', function(e) {
        if (e.target === this) this.classList.remove('active');
    });
    </script>
    <script src="assets/clocklog.js"></script>
</body>
</html>
