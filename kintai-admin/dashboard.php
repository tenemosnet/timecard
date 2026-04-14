<?php
require_once __DIR__ . '/auth.php';
requireLogin();

$csrfToken = generateCsrfToken();
$displayName = htmlspecialchars($_SESSION['display_name'] ?? '管理者');

// スタッフ一覧をGASから取得（定時設定付き）
$staffList = [];    // [{name, contractedHours, sortOrder}, ...]
$staffNames = [];   // [name, ...]
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

$currentYear = (int)date('Y');
$currentMonth = (int)date('n');
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>勤怠管理 - ダッシュボード</title>
    <link rel="stylesheet" href="assets/style.css">
</head>
<body>
    <header class="header">
        <h1>勤怠管理 管理者ページ</h1>
        <div class="header-right">
            <a href="<?= htmlspecialchars(GAS_API_URL) ?>" target="_blank" rel="noopener" class="btn btn-sm btn-primary">打刻画面を開く</a>
            <span class="user-name"><?= $displayName ?></span>
            <a href="change-password.php" class="btn btn-sm btn-outline">パスワード変更</a>
            <a href="logout.php" class="btn btn-sm btn-outline">ログアウト</a>
        </div>
    </header>

    <main class="container">
        <!-- PDF生成セクション -->
        <section class="card">
            <h2>PDF生成</h2>

            <?php if ($staffError): ?>
                <div class="alert alert-error"><?= htmlspecialchars($staffError) ?></div>
            <?php endif; ?>

            <div class="form-row">
                <div class="form-group">
                    <label for="gen-year">年</label>
                    <select id="gen-year">
                        <?php for ($y = $currentYear; $y >= $currentYear - 2; $y--): ?>
                            <option value="<?= $y ?>" <?= $y === $currentYear ? 'selected' : '' ?>><?= $y ?></option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group">
                    <label for="gen-month">月</label>
                    <select id="gen-month">
                        <?php for ($m = 1; $m <= 12; $m++): ?>
                            <option value="<?= $m ?>" <?= $m === $currentMonth ? 'selected' : '' ?>><?= $m ?></option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group form-group-btn">
                    <button id="btn-generate-all" class="btn btn-primary">全員分を生成</button>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label for="gen-staff">または個別</label>
                    <select id="gen-staff">
                        <option value="">-- スタッフ選択 --</option>
                        <?php foreach ($staffNames as $name): ?>
                            <option value="<?= htmlspecialchars($name) ?>"><?= htmlspecialchars($name) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                <div class="form-group form-group-btn">
                    <button id="btn-generate-one" class="btn btn-secondary" disabled>この人だけ生成</button>
                </div>
            </div>

            <div id="generate-progress" class="progress-area" style="display:none;">
                <div class="progress-bar">
                    <div class="progress-fill" id="progress-fill"></div>
                </div>
                <p id="progress-text"></p>
            </div>
        </section>

        <!-- PDF一覧セクション -->
        <section class="card">
            <h2>PDF一覧</h2>

            <div class="form-row">
                <div class="form-group">
                    <label for="list-year">年</label>
                    <select id="list-year">
                        <?php for ($y = $currentYear; $y >= $currentYear - 2; $y--): ?>
                            <option value="<?= $y ?>" <?= $y === $currentYear ? 'selected' : '' ?>><?= $y ?></option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group">
                    <label for="list-staff">スタッフ</label>
                    <select id="list-staff">
                        <option value="">全員</option>
                        <?php foreach ($staffNames as $name): ?>
                            <option value="<?= htmlspecialchars($name) ?>"><?= htmlspecialchars($name) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                <div class="form-group form-group-btn">
                    <button id="btn-search" class="btn btn-primary">検索</button>
                </div>
            </div>

            <!-- PDF選択アクションバー -->
            <div id="pdf-actions-bar" class="pdf-actions-bar" style="display:none;">
                <label><input type="checkbox" id="pdf-select-all"> 全選択</label>
                <span class="selected-count" id="pdf-selected-count">0件選択中</span>
                <button id="btn-download-selected" class="btn btn-sm btn-download">選択分をDL</button>
            </div>

            <div id="pdf-list-area">
                <table class="table" id="pdf-table">
                    <thead>
                        <tr>
                            <th style="width:30px;"></th>
                            <th>月</th>
                            <th>氏名</th>
                            <th>ファイル名</th>
                            <th>生成日時</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody id="pdf-tbody">
                        <tr><td colspan="6" class="text-center text-muted">「検索」を押してPDF一覧を取得してください</td></tr>
                    </tbody>
                </table>
            </div>
        </section>

        <!-- スタッフ設定セクション（折り畳み） -->
        <section class="card">
            <div class="collapsible-header" id="staff-settings-toggle">
                <h2>スタッフ設定</h2>
                <span class="collapse-toggle" id="staff-toggle-label">開く</span>
            </div>
            <div class="collapsible-body collapsed" id="staff-settings-body">
                <p class="text-muted" style="margin-bottom:1rem; font-size:0.85rem;">ドラッグで表示順を変更できます。</p>
                <table class="table" id="staff-settings-table">
                    <thead>
                        <tr>
                            <th style="width:30px;"></th>
                            <th>スタッフ名</th>
                            <th>定時（時間）</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody id="staff-sortable-tbody">
                        <?php foreach ($staffList as $i => $staff): ?>
                        <tr class="sortable-row" data-staff="<?= htmlspecialchars($staff['name']) ?>" data-order="<?= $i ?>">
                            <td><span class="drag-handle">☰</span></td>
                            <td><?= htmlspecialchars($staff['name']) ?></td>
                            <td>
                                <select class="staff-hours-select" data-staff="<?= htmlspecialchars($staff['name']) ?>">
                                    <option value="7.5" <?= ($staff['contractedHours'] ?? 8) == 7.5 ? 'selected' : '' ?>>7時間30分</option>
                                    <option value="8" <?= ($staff['contractedHours'] ?? 8) == 8 ? 'selected' : '' ?>>8時間</option>
                                </select>
                            </td>
                            <td class="actions">
                                <button class="btn btn-sm btn-secondary btn-save-hours" data-staff="<?= htmlspecialchars($staff['name']) ?>">保存</button>
                                <button class="btn btn-sm btn-rename-staff" data-staff="<?= htmlspecialchars($staff['name']) ?>">名前変更</button>
                                <button class="btn btn-sm btn-danger btn-remove-staff" data-staff="<?= htmlspecialchars($staff['name']) ?>">削除</button>
                            </td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

                <h3 class="section-subtitle">スタッフ追加</h3>
                <div class="form-row">
                    <div class="form-group">
                        <label for="new-staff-name">氏名</label>
                        <input type="text" id="new-staff-name" placeholder="例: 山田太郎">
                    </div>
                    <div class="form-group">
                        <label for="new-staff-hours">定時</label>
                        <select id="new-staff-hours">
                            <option value="8">8時間</option>
                            <option value="7.5">7時間30分</option>
                        </select>
                    </div>
                    <div class="form-group form-group-btn">
                        <button id="btn-add-staff" class="btn btn-primary">追加</button>
                    </div>
                </div>
                <p id="staff-manage-message" class="text-muted" style="margin-top:0.5rem;"></p>
            </div>
        </section>

        <!-- ステータスセクション -->
        <section class="card">
            <h2>ステータス</h2>
            <div id="status-area">
                <p>Googleドライブ接続: <span id="status-drive" class="status-indicator">確認中...</span></p>
            </div>
        </section>
    </main>

    <input type="hidden" id="csrf-token" value="<?= htmlspecialchars($csrfToken) ?>">
    <input type="hidden" id="staff-list" value="<?= htmlspecialchars(json_encode($staffList)) ?>">

    <script src="assets/app.js"></script>
</body>
</html>
