// ========================================
// 勤怠管理 管理者ページ - フロントエンドJS
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    const csrfToken = document.getElementById('csrf-token').value;
    const staffListRaw = JSON.parse(document.getElementById('staff-list').value || '[]');
    // staffListRaw は [{name, contractedHours}, ...] の配列
    const staffNames = staffListRaw.map(s => s.name || s);

    // --- 要素取得 ---
    const genYear = document.getElementById('gen-year');
    const genMonth = document.getElementById('gen-month');
    const genStaff = document.getElementById('gen-staff');
    const btnGenerateAll = document.getElementById('btn-generate-all');
    const btnGenerateOne = document.getElementById('btn-generate-one');
    const progressArea = document.getElementById('generate-progress');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');

    const listYear = document.getElementById('list-year');
    const listStaff = document.getElementById('list-staff');
    const btnSearch = document.getElementById('btn-search');
    const pdfTbody = document.getElementById('pdf-tbody');

    const statusDrive = document.getElementById('status-drive');

    // --- スタッフ選択で「この人だけ生成」ボタンの有効/無効 ---
    genStaff.addEventListener('change', () => {
        btnGenerateOne.disabled = genStaff.value === '';
    });

    // --- PDF生成: 全員分 ---
    btnGenerateAll.addEventListener('click', async () => {
        if (staffNames.length === 0) {
            alert('スタッフ一覧が取得できていません。ページを再読み込みしてください。');
            return;
        }
        const year = parseInt(genYear.value);
        const month = parseInt(genMonth.value);

        if (!confirm(`${year}年${month}月の全員分（${staffNames.length}名）のPDFを生成しますか？`)) {
            return;
        }

        setGenerateButtonsEnabled(false);
        showProgress(true);
        const batchFileIds = [];

        for (let i = 0; i < staffNames.length; i++) {
            const name = staffNames[i];
            updateProgress(i, staffNames.length, `${name} のPDFを生成中...`);

            const result = await generatePDF(name, year, month);
            if (!result.success) {
                updateProgress(i, staffNames.length, `エラー: ${name} - ${result.error}`);
            } else if (result.data && result.data.fileId) {
                batchFileIds.push(result.data.fileId);
            }
        }

        updateProgress(staffNames.length, staffNames.length, '全員分のPDF生成が完了しました。');

        // 生成されたPDFを自動ダウンロード
        for (const fileId of batchFileIds) {
            triggerDownload(fileId);
        }
        setGenerateButtonsEnabled(true);

        listYear.value = year;
        loadPDFList();
    });

    // --- PDF生成: 個別 ---
    btnGenerateOne.addEventListener('click', async () => {
        const name = genStaff.value;
        if (!name) return;

        const year = parseInt(genYear.value);
        const month = parseInt(genMonth.value);

        if (!confirm(`${name} の ${year}年${month}月 のPDFを生成しますか？`)) {
            return;
        }

        setGenerateButtonsEnabled(false);
        showProgress(true);
        updateProgress(0, 1, `${name} のPDFを生成中...`);

        const result = await generatePDF(name, year, month);

        if (result.success) {
            updateProgress(1, 1, `${name} のPDF生成が完了しました。`);
            // 生成されたPDFを自動ダウンロード
            if (result.data && result.data.fileId) {
                triggerDownload(result.data.fileId);
            }
        } else {
            updateProgress(0, 1, `エラー: ${result.error}`);
        }

        setGenerateButtonsEnabled(true);
        loadPDFList();
    });

    // --- PDF一覧検索 ---
    btnSearch.addEventListener('click', () => {
        loadPDFList();
    });

    // --- スタッフ設定: 定時保存 ---
    document.querySelectorAll('.btn-save-hours').forEach(btn => {
        btn.addEventListener('click', async () => {
            const staffName = btn.dataset.staff;
            const select = document.querySelector(`.staff-hours-select[data-staff="${staffName}"]`);
            const hours = parseFloat(select.value);

            btn.disabled = true;
            btn.textContent = '保存中...';

            try {
                const res = await fetch('staff_setting.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        csrf_token: csrfToken,
                        staffName: staffName,
                        contractedHours: hours,
                    }),
                });
                const result = await res.json();

                if (result.success) {
                    btn.textContent = '保存済';
                    setTimeout(() => { btn.textContent = '保存'; btn.disabled = false; }, 2000);
                } else {
                    alert('エラー: ' + result.error);
                    btn.textContent = '保存';
                    btn.disabled = false;
                }
            } catch (e) {
                alert('通信エラー: ' + e.message);
                btn.textContent = '保存';
                btn.disabled = false;
            }
        });
    });

    // --- スタッフ追加 ---
    const btnAddStaff = document.getElementById('btn-add-staff');
    const staffMessage = document.getElementById('staff-manage-message');

    btnAddStaff.addEventListener('click', async () => {
        const name = document.getElementById('new-staff-name').value.trim();
        const hours = parseFloat(document.getElementById('new-staff-hours').value);

        if (!name) {
            staffMessage.textContent = '氏名を入力してください。';
            staffMessage.style.color = '#f87171';
            return;
        }

        if (!confirm(`「${name}」をスタッフに追加しますか？`)) return;

        btnAddStaff.disabled = true;
        staffMessage.textContent = '追加中...';
        staffMessage.style.color = '#aaa';

        try {
            const res = await fetch('staff_manage.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    csrf_token: csrfToken,
                    action: 'add',
                    staffName: name,
                    contractedHours: hours,
                }),
            });
            const result = await res.json();

            if (result.success) {
                staffMessage.textContent = `「${name}」を追加しました。ページを再読み込みします...`;
                staffMessage.style.color = '#4ade80';
                setTimeout(() => location.reload(), 1500);
            } else {
                staffMessage.textContent = 'エラー: ' + result.error;
                staffMessage.style.color = '#f87171';
                btnAddStaff.disabled = false;
            }
        } catch (e) {
            staffMessage.textContent = '通信エラー: ' + e.message;
            staffMessage.style.color = '#f87171';
            btnAddStaff.disabled = false;
        }
    });

    // --- スタッフ削除 ---
    document.querySelectorAll('.btn-remove-staff').forEach(btn => {
        btn.addEventListener('click', async () => {
            const staffName = btn.dataset.staff;

            if (!confirm(`「${staffName}」を削除しますか？\n（シートは非表示になりますが、データは保持されます）`)) {
                return;
            }

            btn.disabled = true;
            btn.textContent = '削除中...';

            try {
                const res = await fetch('staff_manage.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        csrf_token: csrfToken,
                        action: 'remove',
                        staffName: staffName,
                    }),
                });
                const result = await res.json();

                if (result.success) {
                    const row = document.getElementById('staff-row-' + staffName);
                    if (row) row.remove();
                } else {
                    alert('エラー: ' + result.error);
                    btn.textContent = '削除';
                    btn.disabled = false;
                }
            } catch (e) {
                alert('通信エラー: ' + e.message);
                btn.textContent = '削除';
                btn.disabled = false;
            }
        });
    });

    // --- PDF生成API呼び出し ---
    async function generatePDF(staffName, year, month) {
        try {
            const res = await fetch('generate.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    csrf_token: csrfToken,
                    staffName: staffName,
                    year: year,
                    month: month,
                }),
            });
            return await res.json();
        } catch (e) {
            return { success: false, error: '通信エラー: ' + e.message };
        }
    }

    // --- PDF一覧取得 ---
    async function loadPDFList() {
        const year = listYear.value;
        const staff = listStaff.value;
        pdfTbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">読み込み中...</td></tr>';

        try {
            const params = new URLSearchParams({ year: year });
            if (staff) params.append('staff', staff);

            const res = await fetch('list_pdfs.php?' + params.toString());
            const result = await res.json();

            if (!result.success) {
                pdfTbody.innerHTML = `<tr><td colspan="5" class="text-center alert-error">${escapeHtml(result.error)}</td></tr>`;
                return;
            }

            const data = result.data || [];
            if (data.length === 0) {
                pdfTbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">PDFが見つかりません</td></tr>';
                return;
            }

            data.sort((a, b) => {
                if (b.month !== a.month) return b.month - a.month;
                return (a.staffName || a.fileName).localeCompare(b.staffName || b.fileName, 'ja');
            });

            pdfTbody.innerHTML = data.map(item => `
                <tr>
                    <td>${item.month}月</td>
                    <td>${escapeHtml(item.staffName || '-')}</td>
                    <td>${escapeHtml(item.fileName)}</td>
                    <td>${item.createdAt || '-'}</td>
                    <td class="actions">
                        <a href="download.php?id=${encodeURIComponent(item.fileId)}" class="btn btn-sm btn-download">DL</a>
                        <a href="${escapeHtml(item.url)}" target="_blank" rel="noopener" class="btn btn-sm btn-view">表示</a>
                    </td>
                </tr>
            `).join('');
        } catch (e) {
            pdfTbody.innerHTML = `<tr><td colspan="5" class="text-center alert-error">通信エラー: ${escapeHtml(e.message)}</td></tr>`;
        }
    }

    // --- ユーティリティ ---
    function showProgress(show) {
        progressArea.style.display = show ? 'block' : 'none';
    }

    function updateProgress(current, total, text) {
        const pct = total > 0 ? Math.round((current / total) * 100) : 0;
        progressFill.style.width = pct + '%';
        progressText.textContent = `(${current}/${total}) ${text}`;
    }

    function setGenerateButtonsEnabled(enabled) {
        btnGenerateAll.disabled = !enabled;
        btnGenerateOne.disabled = !enabled || genStaff.value === '';
    }

    function triggerDownload(fileId) {
        const a = document.createElement('a');
        a.href = 'download.php?id=' + encodeURIComponent(fileId);
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str || '';
        return div.innerHTML;
    }

    // --- 初期処理: ドライブ接続ステータス確認 ---
    async function checkDriveStatus() {
        try {
            const res = await fetch('list_pdfs.php?year=' + new Date().getFullYear());
            const result = await res.json();
            if (result.success) {
                statusDrive.textContent = '● 正常';
                statusDrive.className = 'status-indicator status-ok';
            } else {
                statusDrive.textContent = '● エラー';
                statusDrive.className = 'status-indicator status-error';
            }
        } catch {
            statusDrive.textContent = '● 接続不可';
            statusDrive.className = 'status-indicator status-error';
        }
    }

    checkDriveStatus();
});
