// ========================================
// 勤怠管理 管理者ページ - フロントエンドJS
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    const csrfToken = document.getElementById('csrf-token').value;
    const staffListRaw = JSON.parse(document.getElementById('staff-list').value || '[]');
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
    const listMonth = document.getElementById('list-month');
    const listStaff = document.getElementById('list-staff');
    const btnSearch = document.getElementById('btn-search');
    const pdfTbody = document.getElementById('pdf-tbody');

    const statusDrive = document.getElementById('status-drive');

    // ===========================
    // 折り畳み: スタッフ設定
    // ===========================
    const staffToggle = document.getElementById('staff-settings-toggle');
    const staffBody = document.getElementById('staff-settings-body');
    const staffToggleLabel = document.getElementById('staff-toggle-label');

    staffToggle.addEventListener('click', () => {
        const isCollapsed = staffBody.classList.contains('collapsed');
        if (isCollapsed) {
            staffBody.classList.remove('collapsed');
            staffBody.classList.add('expanded');
            staffToggleLabel.textContent = '閉じる';
        } else {
            staffBody.classList.remove('expanded');
            staffBody.classList.add('collapsed');
            staffToggleLabel.textContent = '開く';
        }
    });

    // ===========================
    // ドラッグ並び替え
    // ===========================
    const sortableTbody = document.getElementById('staff-sortable-tbody');
    let dragRow = null;

    sortableTbody.addEventListener('dragstart', (e) => {
        dragRow = e.target.closest('tr');
        if (dragRow) {
            dragRow.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        }
    });

    sortableTbody.addEventListener('dragover', (e) => {
        e.preventDefault();
        const target = e.target.closest('tr');
        if (target && target !== dragRow && target.parentNode === sortableTbody) {
            const rect = target.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            if (e.clientY < midY) {
                sortableTbody.insertBefore(dragRow, target);
            } else {
                sortableTbody.insertBefore(dragRow, target.nextSibling);
            }
        }
    });

    sortableTbody.addEventListener('dragend', () => {
        if (dragRow) {
            dragRow.classList.remove('dragging');
            dragRow = null;
            saveSortOrder();
        }
    });

    // ドラッグハンドルでのみドラッグ開始
    sortableTbody.querySelectorAll('.sortable-row').forEach(row => {
        row.setAttribute('draggable', 'false');
        const handle = row.querySelector('.drag-handle');
        if (handle) {
            handle.addEventListener('mousedown', () => row.setAttribute('draggable', 'true'));
            handle.addEventListener('mouseup', () => row.setAttribute('draggable', 'false'));
        }
    });

    async function saveSortOrder() {
        const rows = sortableTbody.querySelectorAll('.sortable-row');
        const order = [];
        rows.forEach((row, i) => {
            order.push({ name: row.dataset.staff, sortOrder: i });
        });

        try {
            await fetch('staff_setting.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    csrf_token: csrfToken,
                    action: 'updateOrder',
                    order: order,
                }),
            });
        } catch (e) {
            // 並び替え保存失敗は静かに無視
        }
    }

    // ===========================
    // PDF生成
    // ===========================
    genStaff.addEventListener('change', () => {
        btnGenerateOne.disabled = genStaff.value === '';
    });

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

        for (const fileId of batchFileIds) {
            triggerDownload(fileId);
        }
        setGenerateButtonsEnabled(true);

        listYear.value = year;
        loadPDFList();
    });

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
            if (result.data && result.data.fileId) {
                triggerDownload(result.data.fileId);
            }
        } else {
            updateProgress(0, 1, `エラー: ${result.error}`);
        }

        setGenerateButtonsEnabled(true);
        loadPDFList();
    });

    // ===========================
    // PDF一覧
    // ===========================
    btnSearch.addEventListener('click', () => {
        loadPDFList();
    });

    // PDF選択バー
    const pdfActionsBar = document.getElementById('pdf-actions-bar');
    const pdfSelectAll = document.getElementById('pdf-select-all');
    const pdfSelectedCount = document.getElementById('pdf-selected-count');
    const btnDownloadSelected = document.getElementById('btn-download-selected');

    pdfSelectAll.addEventListener('change', () => {
        const checkboxes = pdfTbody.querySelectorAll('.pdf-checkbox');
        checkboxes.forEach(cb => { cb.checked = pdfSelectAll.checked; });
        updateSelectedCount();
    });

    btnDownloadSelected.addEventListener('click', () => {
        const checked = pdfTbody.querySelectorAll('.pdf-checkbox:checked');
        if (checked.length === 0) {
            alert('ダウンロードするPDFを選択してください。');
            return;
        }
        checked.forEach(cb => {
            triggerDownload(cb.dataset.fileId);
        });
    });

    function updateSelectedCount() {
        const checked = pdfTbody.querySelectorAll('.pdf-checkbox:checked');
        pdfSelectedCount.textContent = checked.length + '件選択中';
    }

    // ===========================
    // スタッフ設定: 定時保存
    // ===========================
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

    // ===========================
    // スタッフ追加
    // ===========================
    const btnAddStaff = document.getElementById('btn-add-staff');
    const staffMessage = document.getElementById('staff-manage-message');

    btnAddStaff.addEventListener('click', async () => {
        const name = document.getElementById('new-staff-name').value.trim();
        const hours = parseFloat(document.getElementById('new-staff-hours').value);

        if (!name) {
            staffMessage.textContent = '氏名を入力してください。';
            staffMessage.style.color = '#c45a5a';
            return;
        }

        if (!confirm(`「${name}」をスタッフに追加しますか？`)) return;

        btnAddStaff.disabled = true;
        staffMessage.textContent = '追加中...';
        staffMessage.style.color = '#8a7f6e';

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
                staffMessage.style.color = '#5a8f5a';
                setTimeout(() => location.reload(), 1500);
            } else {
                staffMessage.textContent = 'エラー: ' + result.error;
                staffMessage.style.color = '#c45a5a';
                btnAddStaff.disabled = false;
            }
        } catch (e) {
            staffMessage.textContent = '通信エラー: ' + e.message;
            staffMessage.style.color = '#c45a5a';
            btnAddStaff.disabled = false;
        }
    });

    // ===========================
    // スタッフ削除
    // ===========================
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
                    const row = btn.closest('tr');
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

    // ===========================
    // スタッフ名前変更
    // ===========================
    document.querySelectorAll('.btn-rename-staff').forEach(btn => {
        btn.addEventListener('click', async () => {
            const oldName = btn.dataset.staff;
            const newName = prompt(`「${oldName}」の新しい名前を入力してください:`, oldName);

            if (!newName || newName.trim() === '' || newName.trim() === oldName) return;

            if (!confirm(`「${oldName}」→「${newName.trim()}」に変更しますか？\n（シート名・打刻ログ・設定がすべて更新されます）`)) {
                return;
            }

            btn.disabled = true;
            btn.textContent = '変更中...';

            try {
                const res = await fetch('staff_rename.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        csrf_token: csrfToken,
                        oldName: oldName,
                        newName: newName.trim(),
                    }),
                });
                const result = await res.json();

                if (result.success) {
                    alert(`「${oldName}」を「${newName.trim()}」に変更しました。\nページを再読み込みします。`);
                    location.reload();
                } else {
                    alert('エラー: ' + result.error);
                    btn.textContent = '名前変更';
                    btn.disabled = false;
                }
            } catch (e) {
                alert('通信エラー: ' + e.message);
                btn.textContent = '名前変更';
                btn.disabled = false;
            }
        });
    });

    // ===========================
    // API呼び出し
    // ===========================
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

    async function loadPDFList() {
        const year = listYear.value;
        const month = listMonth.value;
        const staff = listStaff.value;
        pdfTbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">読み込み中...</td></tr>';
        pdfActionsBar.style.display = 'none';

        try {
            const params = new URLSearchParams({ year: year });
            if (month) params.append('month', month);
            if (staff) params.append('staff', staff);

            const res = await fetch('list_pdfs.php?' + params.toString());
            const result = await res.json();

            if (!result.success) {
                pdfTbody.innerHTML = `<tr><td colspan="6" class="text-center alert-error">${escapeHtml(result.error)}</td></tr>`;
                return;
            }

            const data = result.data || [];
            if (data.length === 0) {
                pdfTbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">PDFが見つかりません</td></tr>';
                return;
            }

            data.sort((a, b) => {
                if (b.month !== a.month) return b.month - a.month;
                return (a.staffName || a.fileName).localeCompare(b.staffName || b.fileName, 'ja');
            });

            pdfTbody.innerHTML = data.map(item => `
                <tr>
                    <td><input type="checkbox" class="pdf-checkbox" data-file-id="${escapeHtml(item.fileId)}"></td>
                    <td>${item.month}月</td>
                    <td>${escapeHtml(item.staffName || '-')}</td>
                    <td>${escapeHtml(item.fileName)}</td>
                    <td>${item.createdAt ? new Date(item.createdAt).toLocaleString('ja-JP') : '-'}</td>
                    <td class="actions">
                        <a href="download.php?id=${encodeURIComponent(item.fileId)}" class="btn btn-sm btn-download">DL</a>
                        <a href="${escapeHtml(item.url)}" target="_blank" rel="noopener" class="btn btn-sm btn-view">表示</a>
                    </td>
                </tr>
            `).join('');

            // チェックボックスイベント
            pdfActionsBar.style.display = 'flex';
            pdfSelectAll.checked = false;
            updateSelectedCount();

            pdfTbody.querySelectorAll('.pdf-checkbox').forEach(cb => {
                cb.addEventListener('change', updateSelectedCount);
            });
        } catch (e) {
            pdfTbody.innerHTML = `<tr><td colspan="6" class="text-center alert-error">通信エラー: ${escapeHtml(e.message)}</td></tr>`;
        }
    }

    // ===========================
    // ユーティリティ
    // ===========================
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

    // ===========================
    // ドライブ接続確認
    // ===========================
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

    // ===========================
    // スタッフ個人ページ: トークン取得 → URLコピー
    // ===========================
    document.querySelectorAll('.btn-staff-page').forEach(btn => {
        btn.addEventListener('click', async () => {
            const name = btn.dataset.staff;
            btn.disabled = true;
            btn.textContent = '取得中...';

            try {
                const res = await fetch('staff_token.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        csrf_token: csrfToken,
                        staffName: name,
                        action: 'get_or_create',
                    }),
                });
                const result = await res.json();

                if (result.success) {
                    const url = location.origin + location.pathname.replace(/[^/]*$/, '') + 'staff_view.php?token=' + result.token;

                    if (navigator.clipboard) {
                        await navigator.clipboard.writeText(url);
                        alert(name + 'さんの個人ページURLをコピーしました。\n\n' + url);
                    } else {
                        prompt(name + 'さんの個人ページURL:', url);
                    }
                } else {
                    alert('エラー: ' + (result.error || '不明なエラー'));
                }
            } catch (e) {
                alert('通信エラー: ' + e.message);
            } finally {
                btn.disabled = false;
                btn.textContent = '個人ページ';
            }
        });
    });
});

// ===========================
// スタッフ閲覧ナビ: スタッフ選択ダイアログ
// ===========================
function openStaffSelect(event) {
    event.preventDefault();
    const staffListRaw = JSON.parse(document.getElementById('staff-list').value || '[]');
    const names = staffListRaw.map(s => s.name || s);

    if (names.length === 0) {
        alert('スタッフが登録されていません。');
        return false;
    }

    const choice = prompt('閲覧するスタッフ名を入力してください:\n\n' + names.join('、'));
    if (!choice || !choice.trim()) return false;

    const name = choice.trim();
    if (!names.includes(name)) {
        alert('「' + name + '」は登録されていません。');
        return false;
    }

    // トークンを取得して遷移
    const csrfToken = document.getElementById('csrf-token').value;
    fetch('staff_token.php', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            csrf_token: csrfToken,
            staffName: name,
            action: 'get_or_create',
        }),
    })
    .then(res => res.json())
    .then(result => {
        if (result.success) {
            window.open('staff_view.php?token=' + result.token, '_blank');
        } else {
            alert('エラー: ' + (result.error || ''));
        }
    })
    .catch(err => alert('通信エラー: ' + err.message));

    return false;
}
