# unsou-system-2（運行システム2）

運行関連の業務システム4つを、**単独利用可能**でありながら、**共通ID**でバックグラウンド連携する運行管理の設計・実装を管理するリポジトリです。あわせて **LINE通知** を補助機能として設計する。

## 目的

以下の4システムをそれぞれ単体で運用できるようにしつつ、裏側では共通IDにより一貫した運行データとしてつなぐ。

| システム | 役割 |
|----------|------|
| 配車（システム連携） | **10_配車予定**。月別**横持ち**表から変換した**縦持ち**。運行予定の起点。**ステータス**でカレンダー・LINEほか連携を制御 |
| 配車（印刷・人用） | `配車予定表_2026年マスター.xlsx` の**横持ち**。**A3**・掲示・全体調整。システム管理列は原則持たず、**変換により 10_配車予定** を生成 |
| 点呼記録簿 | 点呼業務記録 |
| 運行日報 | 日々の運行実績・報告（**ヘッダ**と**運行明細**の2層） |
| 売上管理表 | 売上・案件に紐づく収益情報 |

**補助機能**

| 名称 | 役割 |
|------|------|
| LINE通知 | 翌日運行予定の通知等。キューと GAS で重複送信を防止 |

## 技術スタック（基本方針）

- **保存先DB**: Google スプレッドシート
- **現場入力**: AppSheet
- **参考・現物資料**: Google Drive 上のファイル（設計の正本は本リポジトリの **`docs/`**）

連携の中心キーは **運行予定ID**（内部Keyは **AppSheet UNIQUEID()**）。人が読む番号は **表示用番号** 列で別管理する。あわせて **運転者ID・車両ID・案件ID・取引先ID** をマスタ共通IDとして参照する。**車両マスタ** の **車両呼称** は現場表示用であり **AppSheet の Label 候補**。帳票用の並びとは **車両番号_表示用** で区別する（詳細は `schema_master.md`）。

## 配車ステータス・予定変更の方針

- 予定の内容が変わっても **原則として同一の運行予定ID（内部Key）を維持**し、**ステータス**（変更／キャンセル等）と **更新日時** で履歴を追う。
- 表示用の **運行予定番号** は帳票・説明用。内部Keyとは混同しない。

## ドキュメント構成

| パス | 内容 |
|------|------|
| [docs/system_design.md](docs/system_design.md) | 全体構成・データフロー・方針 |
| [docs/schema_master.md](docs/schema_master.md) | **列定義の元本**（新規列は必ずここを先に更新） |
| [docs/dispatch_conversion_design.md](docs/dispatch_conversion_design.md) | **月別横持ち配車表**から **10_配車予定** へ変換するルール |
| [docs/source_materials.md](docs/source_materials.md) | Drive／既存資料の一覧と扱い |
| [docs/calendar_sync_design.md](docs/calendar_sync_design.md) | **10_配車予定** と Google カレンダー同期設計 |
| [docs/tenko_record_design.md](docs/tenko_record_design.md) | 点呼記録簿設計 |
| [docs/daily_report_design.md](docs/daily_report_design.md) | 運行日報設計 |
| [docs/sales_management_design.md](docs/sales_management_design.md) | 売上管理表設計 |
| [docs/line_notification_design.md](docs/line_notification_design.md) | LINE通知（キュー・重複防止・GAS） |
| [notes/progress_memo.md](notes/progress_memo.md) | 作業進捗メモ |

## AI・ツールの役割分担（目安）

| 担当 | 役割 |
|------|------|
| **ChatGPT** | 全体設計、方針決定、仕様整理、他AIへの指示書作成 |
| **Cursor** | 本リポジトリ内の `docs` 編集、`schema_master.md` 整備、GAS・スクリプト作成 |
| **Claude Code** | 設計レビュー、抜け漏れ・整合性の確認 |

## 運用上のルール

1. 設計の正本は **GitHub の `docs/`** とする。Google Drive は参考資料である。
2. スプレッドシート／AppSheet の列追加・変更は、**先に `docs/schema_master.md` を更新**してから実装する。
3. **10_配車予定**（変換済み縦持ち）をデータの上流（システム連携）とし、**運行予定ID**で下流システムと論理的に結ぶ。印刷用の**月別横持ち配車表**からは直接連携しない。
4. **内部Key** は **UNIQUEID()**。**表示用番号** は別列・後続整備可（詳細は `schema_master.md` の ID・採番ルール）。

---

*リポジトリ名／作業フォルダ名: `unsou-system-2`*
