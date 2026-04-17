// ========================================
// 勤怠管理 打刻データ修正 - フロントエンドJS
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    const csrfToken = document.getElementById('csrfToken').value;

    // --- 要素取得 ---
    const searchStaff = document.getElementById('searchStaff');
    const searchDate = document.getElementById('searchDate');
    const btnSearch = document.getElementById('btnSearch');
    const resultSection = document.getElementById('resultSection');
    const resultTitle = document.getElementById('resultTitle');
    const resultBody = document.getElementById('resultBody');
    const btnAdd = document.getElementById('btnAdd');
    const statusMessage = document.getElementById('statusMessage');

    // モーダル要素
    const entryModal = document.getElementById('entryModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalStaffGroup = document.getElementById('modalStaffGroup');
    const modalStaff = document.getElementById('modalStaff');
    const modalDateGroup = document.getElementById('modalDateGroup');
    const modalDate = document.getElementById('modalDate');
    const modalType = document.getElementById('modalType');
    const modalTimeGroup = document.getElementById('modalTimeGroup');
    const modalTime = document.getElementById('modalTime');
    const btnModalSave = document.getElementById('btnModalSave');
    const btnModalCancel = document.getElementById('btnModalCancel');

    // 現在の検索結果
    let currentResults = [];
    let editingRowIndex = null;  // 編集中の行番号（null=新規追加モード）

    // ===========================
    // 検索
    // ===========================
    btnSearch.addEventListener('click', doSearch);

    // Enterキーでも検索
    searchDate.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') doSearch();
    });

    async function doSearch() {
        const staff = searchStaff.value;
        const date = searchDate.value;

        if (!date) {
            showStatus('日付を選択してください。', 'error');
            return;
        }

        btnSearch.disabled = true;
        btnSearch.textContent = '検索中...';

        try {
            const params = new URLSearchParams({ date, staff });
            const res = await fetch('clocklog_api.php?' + params.toString());
            const result = await res.json();

            if (!result.success) {
                showStatus(result.error || '検索に失敗しました。', 'error');
                return;
            }

            currentResults = result.data || [];
            renderResults(staff, date);
            resultSection.style.display = '';
            hideStatus();

        } catch (err) {
            showStatus('通信エラー: ' + err.message, 'error');
        } finally {
            btnSearch.disabled = false;
            btnSearch.textContent = '検索';
        }
    }

    // ===========================
    // 結果表示
    // ===========================
    function renderResults(staff, date) {
        const dateObj = new Date(date + 'T00:00:00');
        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
        const dateLabel = (dateObj.getMonth() + 1) + '/' + dateObj.getDate() + '(' + weekdays[dateObj.getDay()] + ')';

        resultTitle.textContent = escapeHtml(staff) + '  ' + dateLabel + '  の打刻データ';

        if (currentResults.length === 0) {
            resultBody.innerHTML = '<div class="no-data">該当する打刻データはありません。</div>';
            return;
        }

        let html = '<table class="clocklog-table">';
        html += '<thead><tr><th>種別</th><th>記録時刻</th><th>操作</th></tr></thead>';
        html += '<tbody>';

        for (const entry of currentResults) {
            const badgeClass = getBadgeClass(entry.type);
            const timeDisplay = entry.type === '有給' ? '-' : entry.timestamp.split(' ')[1] || '';

            html += '<tr>';
            html += '<td><span class="badge ' + badgeClass + '">' + escapeHtml(entry.type) + '</span></td>';
            html += '<td>' + escapeHtml(timeDisplay) + '</td>';
            html += '<td class="actions">';
            html += '<button class="btn-edit" data-row="' + entry.rowIndex + '">編集</button>';
            html += '<button class="btn-delete" data-row="' + entry.rowIndex + '" data-staff="' + escapeHtml(entry.staffName) + '" data-date="' + escapeHtml(entry.date) + '" data-type="' + escapeHtml(entry.type) + '">削除</button>';
            html += '</td>';
            html += '</tr>';
        }

        html += '</tbody></table>';
        resultBody.innerHTML = html;

        // 編集ボタン
        resultBody.querySelectorAll('.btn-edit').forEach(btn => {
            btn.addEventListener('click', () => openEditModal(Number(btn.dataset.row)));
        });

        // 削除ボタン
        resultBody.querySelectorAll('.btn-delete').forEach(btn => {
            btn.addEventListener('click', () => doDelete(
                Number(btn.dataset.row),
                btn.dataset.staff,
                btn.dataset.date,
                btn.dataset.type
            ));
        });
    }

    function getBadgeClass(type) {
        if (type === '入室') return 'badge-enter';
        if (type === '退室') return 'badge-leave';
        if (type === '有給') return 'badge-paid';
        return '';
    }

    // ===========================
    // 新規追加
    // ===========================
    btnAdd.addEventListener('click', () => {
        editingRowIndex = null;
        modalTitle.textContent = '打刻を追加';

        // スタッフ・日付は検索条件から自動入力
        modalStaff.value = searchStaff.value;
        modalDate.value = searchDate.value;
        modalStaffGroup.style.display = '';
        modalDateGroup.style.display = '';
        modalType.value = '入室';
        modalTime.value = '';
        modalTimeGroup.style.display = '';

        entryModal.classList.add('active');
    });

    // ===========================
    // 編集モーダル
    // ===========================
    function openEditModal(rowIndex) {
        const entry = currentResults.find(e => e.rowIndex === rowIndex);
        if (!entry) return;

        editingRowIndex = rowIndex;
        modalTitle.textContent = '打刻を編集';

        // 編集時はスタッフ・日付は変更不可（表示のみ）
        modalStaffGroup.style.display = 'none';
        modalDateGroup.style.display = 'none';
        modalType.value = entry.type;

        if (entry.type === '有給') {
            modalTimeGroup.style.display = 'none';
            modalTime.value = '';
        } else {
            modalTimeGroup.style.display = '';
            // timestamp から時刻部分を取得
            const timePart = entry.timestamp.split(' ')[1] || '';
            modalTime.value = timePart;
        }

        entryModal.classList.add('active');
    }

    // ===========================
    // モーダル: 種別変更時の時刻表示切替
    // ===========================
    modalType.addEventListener('change', () => {
        if (modalType.value === '有給') {
            modalTimeGroup.style.display = 'none';
        } else {
            modalTimeGroup.style.display = '';
        }
    });

    // ===========================
    // モーダル: キャンセル
    // ===========================
    btnModalCancel.addEventListener('click', closeModal);

    entryModal.addEventListener('click', (e) => {
        if (e.target === entryModal) closeModal();
    });

    function closeModal() {
        entryModal.classList.remove('active');
        editingRowIndex = null;
    }

    // ===========================
    // モーダル: 保存
    // ===========================
    btnModalSave.addEventListener('click', async () => {
        const type = modalType.value;
        const time = modalTime.value;

        if (type !== '有給' && !time) {
            showStatus('時刻を入力してください。', 'error');
            return;
        }

        btnModalSave.disabled = true;
        btnModalSave.textContent = '保存中...';

        try {
            let body;

            if (editingRowIndex === null) {
                // 新規追加
                body = {
                    csrf_token: csrfToken,
                    action: 'add',
                    staffName: modalStaff.value,
                    type: type,
                    date: modalDate.value,
                    time: type === '有給' ? '' : time,
                };
            } else {
                // 編集
                const entry = currentResults.find(e => e.rowIndex === editingRowIndex);
                body = {
                    csrf_token: csrfToken,
                    action: 'edit',
                    rowIndex: editingRowIndex,
                    newType: type,
                    newTime: type === '有給' ? '' : time,
                    expectedStaff: entry ? entry.staffName : '',
                    expectedDate: entry ? searchDate.value : '',
                };
            }

            const res = await fetch('clocklog_api.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            const result = await res.json();

            if (result.success) {
                showStatus(result.message || '保存しました。', 'success');
                closeModal();
                // 再検索して画面を更新
                await doSearch();
            } else {
                showStatus(result.error || '保存に失敗しました。', 'error');
            }

        } catch (err) {
            showStatus('通信エラー: ' + err.message, 'error');
        } finally {
            btnModalSave.disabled = false;
            btnModalSave.textContent = '保存';
        }
    });

    // ===========================
    // 削除
    // ===========================
    async function doDelete(rowIndex, staffName, date, type) {
        if (!confirm(staffName + 'さんの「' + type + '」を削除しますか？')) {
            return;
        }

        try {
            const res = await fetch('clocklog_api.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    csrf_token: csrfToken,
                    action: 'delete',
                    rowIndex: rowIndex,
                    expectedStaff: staffName,
                    expectedDate: searchDate.value,
                }),
            });
            const result = await res.json();

            if (result.success) {
                showStatus(result.message || '削除しました。', 'success');
                await doSearch();
            } else {
                showStatus(result.error || '削除に失敗しました。', 'error');
            }

        } catch (err) {
            showStatus('通信エラー: ' + err.message, 'error');
        }
    }

    // ===========================
    // ステータスメッセージ
    // ===========================
    const toast = document.getElementById('toast');

    function showStatus(msg, type) {
        if (type === 'success') {
            // 成功時はトースト（ポップアップ）で表示
            showToast(msg, 'success');
            statusMessage.style.display = 'none';
        } else {
            // エラー時は画面内メッセージ
            statusMessage.textContent = msg;
            statusMessage.className = 'alert alert-' + (type || 'info');
            statusMessage.style.display = '';
        }
    }

    function hideStatus() {
        statusMessage.style.display = 'none';
    }

    function showToast(msg, type) {
        toast.textContent = msg;
        toast.className = 'toast toast-' + (type || 'success');
        // 表示アニメーション
        requestAnimationFrame(() => {
            toast.classList.add('show');
        });
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    // ===========================
    // ユーティリティ
    // ===========================
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str || '';
        return div.innerHTML;
    }
});
