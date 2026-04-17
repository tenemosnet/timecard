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

$currentYear = (int)date('Y');
$currentMonth = (int)date('n');
$escapedName = htmlspecialchars($staffName);
$escapedToken = htmlspecialchars($token);
?>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>タイムカード — <?= $escapedName ?></title>
    <link rel="stylesheet" href="assets/style.css">
    <style>
        .staff-header {
            background: #fff;
            padding: 1.25rem 2rem;
            border-bottom: 1px solid #e8e0d4;
            display: flex;
            align-items: center;
            justify-content: space-between;
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
    </style>
</head>
<body>
    <header class="staff-header">
        <h1><span>tenemos</span> 勤怠管理</h1>
        <span class="staff-name-label"><?= $escapedName ?> さん</span>
    </header>

    <main class="container">
        <section class="card">
            <h2>タイムカード一覧</h2>

            <div class="filter-form">
                <div class="form-group">
                    <label for="viewYear">年</label>
                    <select id="viewYear">
                        <?php for ($y = $currentYear - 1; $y <= $currentYear + 1; $y++): ?>
                            <option value="<?= $y ?>" <?= $y === $currentYear ? 'selected' : '' ?>><?= $y ?>年</option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group">
                    <label for="viewMonth">月</label>
                    <select id="viewMonth">
                        <option value="">全月</option>
                        <?php for ($m = 1; $m <= 12; $m++): ?>
                            <option value="<?= $m ?>" <?= $m === $currentMonth ? 'selected' : '' ?>><?= $m ?>月</option>
                        <?php endfor; ?>
                    </select>
                </div>
                <div class="form-group" style="flex:0;">
                    <label>&nbsp;</label>
                    <button id="btnLoad" class="btn btn-primary btn-sm">表示</button>
                </div>
            </div>

            <div id="pdfListBody">
                <p class="text-muted text-center">読み込み中...</p>
            </div>
        </section>
    </main>

    <footer style="text-align:center; padding:1.5rem; color:#8a7f6e; font-size:0.8rem;">
        勤怠管理システム ver2.0
    </footer>

    <script>
    document.addEventListener('DOMContentLoaded', () => {
        const viewYear = document.getElementById('viewYear');
        const viewMonth = document.getElementById('viewMonth');
        const btnLoad = document.getElementById('btnLoad');
        const pdfListBody = document.getElementById('pdfListBody');
        const staffName = <?= json_encode($staffName) ?>;
        const token = <?= json_encode($token) ?>;

        btnLoad.addEventListener('click', loadPDFs);
        loadPDFs();

        async function loadPDFs() {
            btnLoad.disabled = true;
            pdfListBody.innerHTML = '<p class="text-muted text-center">読み込み中...</p>';

            try {
                const params = new URLSearchParams({
                    year: viewYear.value,
                    staff: staffName,
                    token: token,
                });
                if (viewMonth.value) params.set('month', viewMonth.value);

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

                let html = '<table class="clocklog-table">';
                html += '<thead><tr><th>月</th><th>ファイル名</th><th>生成日時</th><th>操作</th></tr></thead>';
                html += '<tbody>';

                for (const item of items) {
                    const createdAt = item.createdAt ? new Date(item.createdAt).toLocaleString('ja-JP') : '-';
                    html += '<tr>';
                    html += '<td>' + item.month + '月</td>';
                    html += '<td>' + escapeHtml(item.fileName) + '</td>';
                    html += '<td>' + escapeHtml(createdAt) + '</td>';
                    html += '<td><a href="download.php?id=' + encodeURIComponent(item.fileId) + '&token=' + encodeURIComponent(token) + '" class="btn btn-sm btn-download">ダウンロード</a></td>';
                    html += '</tr>';
                }

                html += '</tbody></table>';
                pdfListBody.innerHTML = html;

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
    </script>
</body>
</html>
