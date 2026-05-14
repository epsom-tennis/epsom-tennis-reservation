# EPSOM&CO テニスイベント予約管理

## WHY
テニスイベントの予約受付・当落通知をLINEで自動化する。スタッフの手作業を減らし、参加者への連絡ミスをなくす。

## HOW
- 技術: GAS (Google Apps Script) + Google Forms + スプレッドシート + LINE Messaging API
- 構成: `docs/input/` に要件・参考資料、`docs/output/` に仕様書等の成果物

### 仕組み
1. 参加者がLINE公式アカウントを友だち追加
2. 「応募」と送信 → GAS WebhookがUser IDを保存し、ランダム受付コードを返信
3. Google Formsで参加者情報を入力（受付コード欄あり）
4. フォーム送信トリガー → 受付コードでUser IDを紐付け → LINE個人宛に応募完了通知
5. 当落確定後 → スプレッドシートから一括でLINE個人宛に当落通知

### ファイル構成（予定）
- `src/webhook.gs` — LINE Webhookの受信・受付コード生成・User ID保存
- `src/form-trigger.gs` — フォーム送信時の応募完了LINE通知
- `src/notify-result.gs` — 当落結果の一括LINE送信

## ルール
- コメントは日本語
- 変更前に目的を説明する
- `rm -rf` / `git push --force` 禁止
- 仕様書・ドキュメント等の成果物は `docs/output/` に保存する
- ユーザーからの提供資料は `docs/input/` に整理する
