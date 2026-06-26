# タイムカード入退勤管理システム

小規模事業所（5名以下）向けの打刻・勤務時間集計システム。
タブレット端末のブラウザから打刻し、Google スプレッドシートで自動集計。管理者はXserver上のPHP画面からPDF出力・打刻修正を行う。

**現在バージョン: ver3.3**

---

## システム構成

```
[タブレット端末]
  Index.html (打刻画面)
      │ google.script.run
      ▼
[Google Apps Script]  ──API──▶  [Google スプレッドシート]
  Code.gs (サーバーロジック)         打刻ログ / スタッフ別シート
      ▲
      │ HTTP POST (APIキー認証)
[管理者ページ]
  kintai-admin/ (PHP + MySQL)
  Xserver上でホスティング
```

### ファイル構成

| ファイル / ディレクトリ | 役割 |
|----------------------|------|
| `Code.gs.js` | GASサーバーロジック（GASエディタでは `Code.gs`） |
| `Index.html` | タブレット打刻画面（Vanilla JS + CSS、ダークモードUI） |
| `kintai-admin/` | PHP管理画面（PDF生成・打刻修正・スタッフ管理） |
| `タイムカード設計仕様書.md` | GAS・スプレッドシート詳細仕様 |
| `管理者ページ設計仕様書.md` | 管理画面詳細仕様 |

---

## 機能一覧

### タブレット打刻画面 (Index.html)
- スタッフ名タップによる入室 / 退室の打刻
- 当日の打刻状況リアルタイム表示（入室済み / 退室済み）
- 有給申請・取消（入室済みの日はブロック）
- 当月の有給申請済み一覧確認

### 管理者ページ (kintai-admin/)
- ログイン認証（セッション管理、5回/15分でロックアウト）
- PDF勤怠表の生成・一括印刷・ダウンロード
- 打刻ログの検索・追加・修正・削除
- スタッフ管理（追加・削除・改名・定時設定）
- スタッフ個人閲覧ページ（トークン認証でURL直接アクセス可）
- 祝日自動判定（Google Calendar API連携）

---

## セットアップ手順

### 1. GAS / スプレッドシート

1. Google スプレッドシートを新規作成し、URLからスプレッドシートIDをコピー
2. 「拡張機能 → Apps Script」でエディタを開く
3. `Code.gs.js` の内容を `Code.gs` に貼り付け、冒頭の `SPREADSHEET_ID` を設定
4. 「ファイル → 新規 → HTML」で `Index` という名前のHTMLファイルを作成し `Index.html` の内容を貼り付け
5. スプレッドシートをリロード → メニュー「勤怠管理 → 初期設定」を実行
6. 「デプロイ → 新しいデプロイ → ウェブアプリ」で公開
   - 実行するユーザー: 自分
   - アクセスできるユーザー: 全員
7. 発行されたURLをタブレット端末のブラウザで開く

> **コード変更後の反映**: 「デプロイ → デプロイを管理 → 新バージョン」で再デプロイが必要

### 2. 管理者ページ (Xserver)

1. `kintai-admin/config.php` を作成（`.gitignore` 対象）
   ```php
   <?php
   // DB接続情報
   define('DB_HOST', 'localhost');
   define('DB_NAME', 'tenemosnet_kintai');
   define('DB_USER', 'tenemosnet_time');
   define('DB_PASS', '...');
   // GAS APIエンドポイント・キー
   define('GAS_API_URL', 'https://script.google.com/macros/s/.../exec');
   define('GAS_API_KEY', '...');
   ```
2. Xserverへデプロイ
   ```bash
   scp -i ~/.ssh/xserver_rsa -P 10022 kintai-admin/* \
     tenemosnet@sv8542.xserver.jp:~/tenemosnet.xsrv.jp/public_html/kintai-admin/
   ```
3. ブラウザで `setup.php` にアクセスしてDBテーブルを初期化

---

## 主要設定値 (Code.gs.js)

```javascript
const SPREADSHEET_ID = '...';        // スプレッドシートID（必須）
const LOG_SHEET_NAME = '打刻ログ';
const SETTINGS_SHEET_NAME = 'スタッフ設定';
const DEFAULT_CONTRACTED_HOURS = 8;  // デフォルト定時（7.5 or 8）
const BREAK_HOURS = 1;               // 休憩時間（固定1時間）
const CUTOFF_DAY = 15;               // 給与締め日（16日〜翌15日）
```

---

## スプレッドシート構成

| シート名 | 役割 |
|---------|------|
| 打刻ログ | 生打刻データ（A:記録日時 / B:氏名 / C:種別 / D:日付） |
| 備考ログ | 月移行時の備考バックアップ |
| スタッフ設定 | スタッフ名と定時（時間）の対応表 |
| スタッフ別シート | シート名＝スタッフ名。16日〜翌15日の月次勤務表 |

スタッフ別シート列構成（A〜I）: 日 / 曜日 / 開始時間 / 終了時間 / 休憩時間 / 定時内時間 / 残業時間 / 深夜残業時間 / 備考

---

## 管理者ページ ファイル構成

| ファイル | 役割 |
|---------|------|
| `config.php` | DB接続・GAS API設定（**.gitignore対象**） |
| `auth.php` | セッション管理・ログイン試行制限 |
| `dashboard.php` | メインダッシュボード・PDF生成UI |
| `clocklog.php` | 打刻データ検索・修正・追加・削除 |
| `staff_view.php` | スタッフ個人閲覧（トークン認証） |
| `api.php` | GAS APIへのcurlラッパー |
| `assets/` | CSS / JS（app.js, clocklog.js） |

---

## バージョン履歴

| バージョン | 主な変更 |
|-----------|---------|
| ver3.3 | 別月PDF生成時の備考欄混入バグ修正・UI改善 |
| ver3.2 | 打刻時刻の秒切り捨てによる給与計算端数誤差修正 |
| ver3.1 | 有給申請済み一覧確認機能・入室済み日の申請ブロック |
| ver3.0 | 管理画面UI改善・タイムカード様式変更 |
| ver2.0 | 管理画面に打刻データ修正機能追加 |
| ver1.x | GAS基本機能・PDF生成・スタッフ個人閲覧ページ |

> バージョン表記の更新箇所: `Code.gs.js` 3行目 / `dashboard.php` footer / `clocklog.php` footer / `staff_view.php` footer

---

## GAS 開発上の注意点

- WebApp関数では `getActiveSpreadsheet()` 禁止 → `openById(SPREADSHEET_ID)` を使う
- `google.script.run` は非同期。`withSuccessHandler` / `withFailureHandler` でコールバック
- HTMLからGAS関数に返すDateオブジェクトは文字列に変換してから返す
- 曜日表示は `TEXT(date,"aaa")` ではなく `CHOOSE(WEEKDAY())` を使用（地域設定非依存）
