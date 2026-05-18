# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

タイムカード入退勤管理システム — 小規模事業所（5名以下）向けの打刻・勤務時間集計システム。Google Apps Script (GAS) + Google スプレッドシートで構成。タブレット端末のブラウザから打刻画面を操作する。

## Architecture

```
[タブレット端末]  --google.script.run-->  [Code.gs (GAS)]  --API-->  [Google スプレッドシート]
  (Index.html)                           (Webアプリとして公開)          (打刻ログ + スタッフ別シート)

[Xserver PHP管理画面]  --HTTP POST(doPost)-->  [Code.gs (GAS API)]  --API-->  [Google スプレッドシート]
  (kintai-admin/)                              (APIキー認証)
```

### 2つの通信経路

- **タブレットUI → GAS**: `google.script.run` で直接呼び出し（同一GASドメイン内）
- **管理画面PHP → GAS**: `doPost` エンドポイントへHTTP POST（APIキー認証、JSON形式）

### ファイル構成

- **Code.gs.js**: サーバーサイドロジック。GASエディタでは `Code.gs` として配置
- **Index.html**: タブレット打刻画面。Vanilla JS + CSS、ダークモードUI
- **kintai-admin/**: Xserver上のPHP管理画面（PDF生成・打刻修正・スタッフ管理）
- **タイムカード設計仕様書.md**: 全仕様の詳細ドキュメント

## Key Configuration (Code.gs.js top section)

```javascript
const SPREADSHEET_ID = '...';        // 本番スプレッドシートID（必須）
const STAFF_NAMES = [...];           // 初期設定時のみ使用。以降はシート名で管理
const LOG_SHEET_NAME = '打刻ログ';
const SETTINGS_SHEET_NAME = 'スタッフ設定';  // スタッフごとの定時設定
const DEFAULT_CONTRACTED_HOURS = 8;  // デフォルト定時（7.5 or 8）
const BREAK_HOURS = 1;               // 休憩時間（固定1時間）
const CUTOFF_DAY = 15;               // 給与締め日（16日〜翌15日）
```

## GAS-Specific Constraints

- **WebApp関数では `getActiveSpreadsheet()` 禁止** → `openById(SPREADSHEET_ID)` を使う
- **Index.html** はGASエディタ上で「ファイル→新規→HTML」で作成。ファイル名は `Index`（拡張子なし）
- 曜日表示は `TEXT(date,"aaa")` ではなく `CHOOSE(WEEKDAY())` を使用（地域設定非依存）
- `google.script.run` は非同期。`withSuccessHandler` / `withFailureHandler` でコールバック
- HTMLからGAS関数を呼ぶ際、返り値にDateオブジェクトは使えない（文字列に変換して返す）
- コード変更後、Webアプリに反映するには「デプロイ→デプロイを管理→新バージョン」で再デプロイが必要
- メニュー関数の変更は保存＋スプレッドシートリロードのみで反映

## GAS Functions（タブレットUI用 — google.script.run）

| 関数 | 用途 |
|------|------|
| `getStaffWithStatus()` | スタッフ一覧＋当日の出退勤ステータス取得 |
| `recordClock(name, type)` | 入室/退室の打刻記録 |
| `recordPaidLeave(name, dateStr, memo)` | 有給申請（入室済み日はブロック） |
| `deletePaidLeave(name, dateStr)` | 有給取消 |
| `getPaidLeaveList(staffName)` | 当月給与期間の有給申請済み一覧 |

## GAS API Functions（管理画面用 — doPost経由）

`doPost` でJSON受信 → `action` フィールドでディスパッチ。`api*_()` プレフィクスの内部関数群。

主な操作: `getClockLog`, `addClockEntry`, `editClockEntry`, `deleteClockEntry`, `getStaffList`, `setStaffSetting`, `addStaff`, `removeStaff`, `renameStaff`, `generatePDF`, `listPDFs`, `getPDFContent`

## Spreadsheet Structure

- **打刻ログ**: 生データ（A:記録日時, B:氏名, C:種別[入室/退室/有給], D:日付）
- **備考ログ**: 月移行時の備考バックアップ
- **スタッフ設定**: スタッフ名と定時（時間）の対応表
- **スタッフ別シート**: シート名＝スタッフ名。D2=年, F2=月, I2=定時。行5-35に日別データ（16日〜翌15日）、行36に月次合計
- **列構成（A〜I）**: 日, 曜日, 開始時間, 終了時間, 休憩時間, 定時内時間, 残業時間, 深夜残業時間, 備考

## Admin Page (kintai-admin/)

Xserver上のPHP管理画面。主要ファイル:

| ファイル | 役割 |
|---------|------|
| `config.php` | DB接続・GAS API設定（.gitignore対象） |
| `auth.php` | セッション管理・ログイン試行制限（5回/15分） |
| `dashboard.php` | メインダッシュボード（PDF生成UI） |
| `clocklog.php` | 打刻データ検索・修正・追加・削除 |
| `staff_view.php` | スタッフ個人閲覧ページ（トークン認証で直接アクセス可） |
| `api.php` | GAS APIへのcurlラッパー |
| `assets/` | CSS・JS（app.js, clocklog.js） |

- **DB**: tenemosnet_kintai（MySQL, ユーザー: tenemosnet_time）
- **GAS API通信**: `api.php` → doPostエンドポイント（APIキー認証）

## Version Management

バージョン表記は以下の4ファイルに存在し、すべて同時に更新する:

- `Code.gs.js` 3行目コメント
- `kintai-admin/dashboard.php` footer
- `kintai-admin/clocklog.php` footer
- `kintai-admin/staff_view.php` footer

## Deployment

### GAS（タブレットUI）
1. GASエディタで Code.gs / Index.html を更新
2. 「デプロイ → デプロイを管理 → 新バージョン」で再デプロイ

### PHP（管理画面 → Xserver）
```bash
scp -i ~/.ssh/xserver_rsa -P 10022 <files> tenemosnet@sv8542.xserver.jp:~/tenemosnet.xsrv.jp/public_html/kintai-admin/
```

### 初期セットアップ
1. スプレッドシート作成 → URLからIDをコピー
2. Apps Scriptエディタで Code.gs / Index.html を配置
3. `SPREADSHEET_ID` を設定
4. メニュー「勤怠管理 → 初期設定」実行
5. 「デプロイ → 新しいデプロイ → ウェブアプリ」で公開
