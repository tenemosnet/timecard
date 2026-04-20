<?php
require_once __DIR__ . '/config.php';

// トークン検証（ログイン不要）
$token = trim($_GET['token'] ?? '');
if ($token === '') {
    http_response_code(403);
    echo '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>アクセスエラー</title></head><body style="display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f5f0e8;font-family:sans-serif;"><p>URLが無効です。管理者にお問い合わせください。</p></body></html>';
    exit;
}

try {
    $pdo = new PDO(
        'mysql:host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=utf8mb4',
        DB_USER, DB_PASS,
        [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]
    );
    $stmt = $pdo->prepare('SELECT staff_name FROM staff_tokens WHERE token = ?');
    $stmt->execute([$token]);
    $staffName = $stmt->fetchColumn();
} catch (Exception $e) {
    http_response_code(500);
    echo 'エラーが発生しました。';
    exit;
}

if (!$staffName) {
    http_response_code(403);
    echo '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>アクセスエラー</title></head><body style="display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f5f0e8;font-family:sans-serif;"><p>URLが無効です。管理者にお問い合わせください。</p></body></html>';
    exit;
}

// 管理者ログイン状態をチェック（リダイレクトなし）
session_start();
$isAdmin = isset($_SESSION['user_id']);
$displayName = htmlspecialchars($_SESSION['display_name'] ?? '管理者');

// 管理者の場合、スタッフ一覧とCSRFトークンを取得
$staffNames = [];
$csrfToken = '';
if ($isAdmin) {
    $_SESSION['last_activity'] = time();
    require_once __DIR__ . '/auth.php';
    $csrfToken = generateCsrfToken();
    try {
        require_once __DIR__ . '/api.php';
        $result = callGasApi('getStaffList');
        if ($result['success']) {
            $staffNames = array_column($result['data'], 'name');
        }
    } catch (Exception $e) {}
}

$currentYear = (int)date('Y');
$escapedName = htmlspecialchars($staffName);
$escapedToken = htmlspecialchars($token);

// 短縮URL生成
$protocol = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') ? 'https' : 'http';
$host = $_SERVER['HTTP_HOST'] ?? '';
$dir = rtrim(dirname($_SERVER['SCRIPT_NAME']), '/');
$shortUrl = $protocol . '://' . $host . $dir . '/s.php?t=' . urlencode($token);
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>タイムカード — <?= $escapedName ?></title>
    <link rel="stylesheet" href="assets/style.css">
    <style>
        <?php if (!$isAdmin): ?>
        .staff-header {
            background: #fff;
            padding: 1.25rem 2rem;
            border-bottom: 1px solid #e8e0d4;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.5rem;
        }
        .staff-header h1 {
            font-size: 1.1rem;
            color: #2d2a24;
        }
        .staff-header h1 span {
            color: #b3712d;
        }
        .staff-name-label {
            font-size: 0.9rem;
            color: #8a7f6e;
        }
        <?php endif; ?>
        .filter-form {
            display: flex;
            gap: 0.75rem;
            align-items: flex-end;
            flex-wrap: wrap;
            margin-bottom: 1rem;
        }
        .filter-form select {
            padding: 0.45rem 0.65rem;
            border: 1px solid #d9d0c3;
            border-radius: 8px;
            font-size: 0.9rem;
            background: #fff;
            color: #3d3929;
        }
        .filter-form select:focus {
            outline: none;
            border-color: #b3712d;
        }
        .filter-form label {
            font-size: 0.8rem;
            font-weight: 600;
            color: #5c5545;
            margin-bottom: 0.2rem;
            display: block;
        }
        .share-url-box {
            background: #faf7f2;
            border: 1px solid #e8e0d4;
            border-radius: 8px;
            padding: 0.75rem 1rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            margin-bottom: 1rem;
        }
        .share-url-box label {
            font-size: 0.8rem;
            font-weight: 600;
            color: #5c5545;
            white-space: nowrap;
        }
        .share-url-box input {
            flex: 1;
            border: 1px solid #d9d0c3;
            border-radius: 6px;
            padding: 0.4rem 0.6rem;
            font-size: 0.8rem;
            color: #3d3929;
            background: #fff;
        }
        .share-url-box input:focus {
            outline: none;
            border-color: #b3712d;
        }
        .btn-copy {
            background: #6b8e5a;
            color: #fff;
            border: none;
            padding: 0.4rem 0.8rem;
            border-radius: 6px;
            font-size: 0.8rem;
            cursor: pointer;
            white-space: nowrap;
        }
        .btn-copy:hover { background: #5a7a4a; }
        .btn-view {
            background: #5a7f8e;
            color: #fff;
            border: none;
            padding: 0.3rem 0.6rem;
            border-radius: 6px;
            font-size: 0.8rem;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
        }
        .btn-view:hover { background: #4a6f7e; }

        /* スマホ用カードレイアウト */
        .pdf-card-list { display: none; }

        @media (max-width: 640px) {
            .pdf-table-wrap { display: none; }
            .pdf-card-list { display: block; }
            .pdf-card {
                background: #fff;
                border: 1px solid #e8e0d4;
                border-radius: 8px;
                padding: 0.75rem 1rem;
                margin-bottom: 0.5rem;
            }
            .pdf-card-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 0.4rem;
            }
            .pdf-card-month {
                font-size: 1rem;
                font-weight: 700;
                color: #2d2a24;
            }
            .pdf-card-date {
                font-size: 0.75rem;
                color: #8a7f6e;
            }
            .pdf-card-actions {
                display: flex;
                gap: 0.5rem;
            }
            .pdf-card-actions a {
                flex: 1;
                text-align: center;
                padding: 0.45rem 0;
                border-radius: 6px;
                font-size: 0.85rem;
                text-decoration: none;
            }
            .share-url-box {
                flex-direction: column;
                align-items: stretch;
            }
            .share-url-box label {
                margin-bottom: 0.3rem;
            }
            .share-url-box .url-row {
                display: flex;
                gap: 0.5rem;
            }
            .container { padding: 0.75rem; }
            .staff-header { padding: 0.75rem 1rem; }
        }
    </style>
</head>
<body>
    <?php if ($isAdmin): ?>
    <!-- 管理者用ヘッダー（他ページと統一） -->
    <header class="header">
        <div class="header-left">
            <h1>勤怠管理</h1>
            <nav class="nav-links">
                <a href="dashboard.php">ダッシュボード</a>
                <a href="clocklog.php">打刻データ修正</a>
                <a href="#" onclick="return openStaffSelect(event)" class="active">スタッフ閲覧</a>
            </nav>
        </div>
        <div class="header-right">
            <span class="staff-name-label" style="font-size:0.95rem; color:#b3712d; font-weight:600;"><?= $escapedName ?> さん</span>
            <span class="user-name"><?= $displayName ?></span>
            <a href="logout.php" class="btn btn-sm btn-outline">ログアウト</a>
        </div>
    </header>
    <?php else: ?>
    <!-- スタッフ用ヘッダー -->
    <header class="staff-header">
        <h1><span>tenemos</span> 勤怠管理</h1>
        <span class="staff-name-label"><?= $escapedName ?> さん</span>
    </header>
    <?php endif; ?>

    <main class="container">
        <section class="card">
            <h2><?= $escapedName ?> さんのタイムカード一覧</h2>

            <div class="share-url-box">
                <label>このページのURL:</label>
                <div class="url-row">
                    <input type="text" id="shareUrl" value="<?= htmlspecialchars($shortUrl) ?>" readonly onclick="this.select()">
                    <button class="btn-copy" id="btnCopyUrl">コピー</button>
                </div>
            </div>

            <div class="filter-form">
                <div class="form-group">
                    <label for="viewYear">年</label>
                    <select id="viewYear">
                        <?php for ($y = $currentYear - 1; $y <= $currentYear + 1; $y++): ?>
                            <option value="<?= $y ?>" <?= $y === $currentYear ? 'selected' : '' ?>><?= $y ?>年</option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group" style="flex:0;">
                    <label>&nbsp;</label>
                    <button id="btnLoad" class="btn btn-primary btn-sm">検索</button>
                </div>
            </div>

            <div id="pdfListBody">
                <p class="text-muted text-center">読み込み中...</p>
            </div>
        </section>
    </main>

    <?php if ($isAdmin): ?>
    <!-- スタッフ選択モーダル -->
    <div class="modal-overlay" id="staffSelectModal">
        <div class="modal-content" style="max-width:400px;">
            <h3>スタッフ閲覧</h3>
            <div class="form-group">
                <label for="staffSelectName">スタッフを選択</label>
                <select id="staffSelectName">
                    <?php foreach ($staffNames as $name): ?>
                        <option value="<?= htmlspecialchars($name) ?>" <?= $name === $staffName ? 'selected' : '' ?>><?= htmlspecialchars($name) ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="modal-actions">
                <button class="btn btn-secondary" id="staffSelectCancel">キャンセル</button>
                <button class="btn btn-primary" id="staffSelectOpen">開く</button>
            </div>
        </div>
    </div>
    <input type="hidden" id="csrf-token" value="<?= htmlspecialchars($csrfToken) ?>">
    <?php endif; ?>

    <footer style="text-align:center; padding:1.5rem; color:#8a7f6e; font-size:0.8rem;">
        勤怠管理システム ver3.0
    </footer>

    <script>
    document.addEventListener('DOMContentLoaded', () => {
        const viewYear = document.getElementById('viewYear');
        const btnLoad = document.getElementById('btnLoad');
        const pdfListBody = document.getElementById('pdfListBody');
        const staffName = <?= json_encode($staffName) ?>;
        const token = <?= json_encode($token) ?>;

        btnLoad.addEventListener('click', loadPDFs);
        loadPDFs();

        // URLコピー
        document.getElementById('btnCopyUrl').addEventListener('click', () => {
            const input = document.getElementById('shareUrl');
            input.select();
            navigator.clipboard.writeText(input.value).then(() => {
                const btn = document.getElementById('btnCopyUrl');
                btn.textContent = 'コピー済み';
                setTimeout(() => { btn.textContent = 'コピー'; }, 2000);
            });
        });

        <?php if ($isAdmin): ?>
        // 管理者用: スタッフ選択モーダル
        document.getElementById('staffSelectCancel').addEventListener('click', () => {
            document.getElementById('staffSelectModal').classList.remove('active');
        });
        document.getElementById('staffSelectOpen').addEventListener('click', () => {
            const name = document.getElementById('staffSelectName').value;
            if (!name) return;
            document.getElementById('staffSelectModal').classList.remove('active');
            const csrfToken = document.getElementById('csrf-token').value;
            fetch('staff_token.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ csrf_token: csrfToken, staffName: name, action: 'get_or_create' }),
            })
            .then(res => res.json())
            .then(result => {
                if (result.success) {
                    window.location.href = 'staff_view.php?token=' + result.token;
                } else {
                    alert('エラー: ' + (result.error || ''));
                }
            })
            .catch(err => alert('通信エラー: ' + err.message));
        });
        document.getElementById('staffSelectModal').addEventListener('click', (e) => {
            if (e.target === document.getElementById('staffSelectModal')) {
                document.getElementById('staffSelectModal').classList.remove('active');
            }
        });
        <?php endif; ?>

        async function loadPDFs() {
            btnLoad.disabled = true;
            pdfListBody.innerHTML = '<p class="text-muted text-center">読み込み中...</p>';

            try {
                const params = new URLSearchParams({
                    year: viewYear.value,
                    staff: staffName,
                    token: token,
                });

                const res = await fetch('staff_view_api.php?' + params.toString());
                const result = await res.json();

                if (!result.success) {
                    pdfListBody.innerHTML = '<p class="alert alert-error">' + escapeHtml(result.error) + '</p>';
                    return;
                }

                const items = result.data || [];
                if (items.length === 0) {
                    pdfListBody.innerHTML = '<p class="text-muted text-center" style="padding:2rem;">該当するタイムカードはありません。</p>';
                    return;
                }

                items.sort((a, b) => b.month - a.month);

                // PC用テーブル
                let table = '<div class="pdf-table-wrap"><table class="clocklog-table">';
                table += '<thead><tr><th>月</th><th>ファイル名</th><th>生成日時</th><th>操作</th></tr></thead>';
                table += '<tbody>';

                // スマホ用カード
                let cards = '<div class="pdf-card-list">';

                for (const item of items) {
                    const createdAt = item.createdAt ? new Date(item.createdAt).toLocaleString('ja-JP') : '-';
                    const dlUrl = 'download.php?id=' + encodeURIComponent(item.fileId) + '&token=' + encodeURIComponent(token);
                    const viewUrl = dlUrl + '&view=1';

                    // テーブル行
                    table += '<tr>';
                    table += '<td>' + item.month + '月</td>';
                    table += '<td>' + escapeHtml(item.fileName) + '</td>';
                    table += '<td>' + escapeHtml(createdAt) + '</td>';
                    table += '<td>';
                    table += '<a href="' + viewUrl + '" target="_blank" class="btn-view">表示</a> ';
                    table += '<a href="' + dlUrl + '" class="btn btn-sm btn-download">ダウンロード</a>';
                    table += '</td>';
                    table += '</tr>';

                    // カード
                    cards += '<div class="pdf-card">';
                    cards += '<div class="pdf-card-header">';
                    cards += '<span class="pdf-card-month">' + item.month + '月</span>';
                    cards += '<span class="pdf-card-date">' + escapeHtml(createdAt) + '</span>';
                    cards += '</div>';
                    cards += '<div class="pdf-card-actions">';
                    cards += '<a href="' + viewUrl + '" target="_blank" class="btn-view">表示</a>';
                    cards += '<a href="' + dlUrl + '" class="btn btn-sm btn-download">ダウンロード</a>';
                    cards += '</div>';
                    cards += '</div>';
                }

                table += '</tbody></table></div>';
                cards += '</div>';
                pdfListBody.innerHTML = table + cards;

            } catch (err) {
                pdfListBody.innerHTML = '<p class="alert alert-error">通信エラー: ' + escapeHtml(err.message) + '</p>';
            } finally {
                btnLoad.disabled = false;
            }
        }

        function escapeHtml(str) {
            const div = document.createElement('div');
            div.textContent = str || '';
            return div.innerHTML;
        }
    });

    <?php if ($isAdmin): ?>
    function openStaffSelect(event) {
        event.preventDefault();
        document.getElementById('staffSelectModal').classList.add('active');
        return false;
    }
    <?php endif; ?>
    </script>
</body>
</html>
