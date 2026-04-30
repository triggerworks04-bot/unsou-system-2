# unsou-system-2（運行システム2）

運行関連の業務システム4つを、**単独利用可能**でありながら、**共通ID**でバックグラウンド連携する運行管理の設計・実装を管理するリポジトリです。

## 目的

以下の4システムをそれぞれ単体で運用できるようにしつつ、裏側では共通IDにより一貫した運行データとしてつなぐ。

| システム | 役割 |
|----------|------|
| 配車予定表 | **最上流データ**。運行予定の起点 |
| 点呼記録簿 | 点呼業務記録 |
| 運行日報 | 日々の運行実績・報告 |
| 売上管理表 | 売上・案件に紐づく収益情報 |

## 技術スタック（基本方針）

- **保存先DB**: Google スプレッドシート
- **現場入力**: AppSheet
- **参考・現物資料**: Google Drive 上のファイル（設計の正本は本リポジトリの `docs/`）

連携の中心キーは **運行予定ID**。あわせて **運転者ID・車両ID・案件ID・取引先ID** をマスタ共通IDとして全システムで参照する。

## ドキュメント構成

| パス | 内容 |
|------|------|
| [docs/system_design.md](docs/system_design.md) | 全体構成・データフロー・方針 |
| [docs/schema_master.md](docs/schema_master.md) | **列定義の元本**（新規列は必ずここを先に更新） |
| [docs/source_materials.md](docs/source_materials.md) | Drive／既存資料の一覧と扱い |
| [docs/calendar_sync_design.md](docs/calendar_sync_design.md) | 配車予定表（カレンダー／予定系）設計 |
| [docs/tenko_record_design.md](docs/tenko_record_design.md) | 点呼記録簿設計 |
| [docs/daily_report_design.md](docs/daily_report_design.md) | 運行日報設計 |
| [docs/sales_management_design.md](docs/sales_management_design.md) | 売上管理表設計 |
| [notes/progress_memo.md](notes/progress_memo.md) | 作業進捗メモ |

## AI・ツールの役割分担（目安）

| 担当 | 役割 |
|------|------|
| **ChatGPT** | 全体設計、方針決定、仕様整理、他AIへの指示書作成 |
| **Cursor** | 本リポジトリ内の `docs` 編集、`schema_master.md` 整備、GAS・スクリプト作成 |
| **Claude Code** | 設計レビュー、抜け漏れ・整合性の確認 |

## 運用上のルール

1. 設計の正本は **GitHub の `docs/`** とする。
2. スプレッドシート／AppSheet の列追加・変更は、**先に `docs/schema_master.md` を更新**してから実装する。
3. 配車予定表をデータの上流とし、運行予定IDで下流システムと論理的に結ぶ。

---

*リポジトリ名／作業フォルダ名: `unsou-system-2`*
