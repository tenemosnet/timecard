# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

タイムカード入退勤管理システム — 小規模事業所（5名以下）向けの打刻・勤務時間集計システム。Google Apps Script (GAS) + Google スプレッドシートで構成。タブレット端末のブラウザから打刻画面を操作する。

## Architecture

```
[タブレット端末]  --google.script.run-->  [Code.gs (GAS)]  --API-->  [Google スプレッドシート]
  (Index.html)                           (Webアプリとして公開)          (打刻ログ + スタッフ別シート)
```

- **Code.gs.js**: サーバーサイドロジック。GASエディタでは `Code.gs` として配置。Web API関数（`recordClock`, `getStaffWithStatus`）はWebApp経由で呼ばれるため `SpreadsheetApp.openById()` を使用（`getActiveSpreadsheet()` はWebApp経由では `null` を返す）。メニュー関数（`initialize`, `promptNewMonth` 等）はスプレッドシートUIから呼ばれるため `getActiveSpreadsheet()` で問題ない。
- **Index.html**: フロントエンド。GASの `HtmlService.createHtmlOutputFromFile('Index')` で配信。外部ライブラリなし、Vanilla JS + CSS Grid。ダークモードUI。
- **タイムカード設計仕様書.md**: 全仕様の詳細ドキュメント。数式ロジック、UI仕様、運用フローを含む。

## Key Configuration (Code.gs.js top section)

```javascript
const SPREADSHEET_ID = '...';        // 本番スプレッドシートID（必須）
const STAFF_NAMES = [...];           // 初期設定時のみ使用。以降はシート名で管理
const LOG_SHEET_NAME = '打刻ログ';
const OVERTIME_THRESHOLD = 8;        // 残業基準: 8時間超
const LUNCH_DEDUCT_6H = 0.75;       // 6h以上: 45分控除
const LUNCH_DEDUCT_8H = 1.0;        // 8h以上: 60分控除
```

## GAS-Specific Constraints

- **WebApp関数では `getActiveSpreadsheet()` 禁止** → `openById(SPREADSHEET_ID)` を使う
- **Index.html** はGASエディタ上で「ファイル→新規→HTML」で作成。ファイル名は `Index`（拡張子なし）
- 曜日表示は `TEXT(date,"aaa")` ではなく `CHOOSE(WEEKDAY())` を使用（地域設定非依存）
- `google.script.run` は非同期。`withSuccessHandler` / `withFailureHandler` でコールバック
- HTMLからGAS関数を呼ぶ際、返り値にDateオブジェクトは使えない（文字列に変換して返す）
- コード変更後、Webアプリに反映するには「デプロイ→デプロイを管理→新バージョン」で再デプロイが必要
- メニュー関数の変更は保存＋スプレッドシートリロードのみで反映

## Deployment

1. スプレッドシート作成 → URLからIDをコピー
2. Apps Scriptエディタで Code.gs / Index.html を配置
3. `SPREADSHEET_ID` を設定
4. メニュー「勤怠管理 → 初期設定」実行
5. 「デプロイ → 新しいデプロイ → ウェブアプリ」で公開

## Spreadsheet Structure

- **打刻ログ**: 生データ（記録日時, 氏名, 種別, 日付）。アプリが自動追記。
- **スタッフ別シート**: シート名＝スタッフ名。D1=年, F1=月で表示月を切り替え。行4-34に日別データ（FILTER数式で打刻ログから自動取得）、行36-41に月次集計。
