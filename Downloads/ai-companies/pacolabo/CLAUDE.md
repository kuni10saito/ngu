# パコラボ株式会社 — Claude Code Workspace

## 概要

このワークスペースは、パコラボ株式会社の組織構造を Claude Code のマルチエージェントで再現したものです。
CEO AI（代表取締役 佐藤太賀始 の分身）がオーケストレーターとなり、事務局・財務・セールス・法務・人事の管理部門 AI と
4 つの事業部門 AI を並列・同時に稼働させます。

## 会社概要

| 項目 | 内容 |
|---|---|
| 会社名 | パコラボ株式会社（旧 株式会社パコント） |
| 創業 | 2010年4月 |
| 設立 | 2016年4月 |
| 代表取締役 | 佐藤太賀始 |
| 資本金 | 600万円 |
| 従業員数 | 6名（パソコン塾2名、HP更新作成1名、ICT支援2名、事務1名） |
| 事業内容 | パソコン塾、パソコン販売、ホームページ更新、ICT支援員派遣 |
| 経営理念 | 新しい文化を創造し、世界中の人々に安心と豊かさを生み出します |

## 組織構造

```
CEO AI (ceo.md) — 佐藤太賀始
├── 事務局 AI         @secretary  .claude/agents/secretary.md
├── 財務 AI           @finance    .claude/agents/finance.md
├── セールス AI        @sales      .claude/agents/sales.md      （新規開拓・提案・既存顧客フォロー）
├── 法務 AI           @legal      .claude/agents/legal.md      （契約・コンプライアンス・個人情報保護）
├── 人事 AI           @hr         .claude/agents/hr.md         （採用・労務・研修）
├── エデュセレ部門 AI   @edu        .claude/agents/edu.md        （パソコン塾・教育）
├── メディア・AIワークス部門 AI @media .claude/agents/media.md   （HP制作・SNS・動画）
├── オフィス&ホームIT・AI部門 AI @ithome .claude/agents/ithome.md （PC販売・修理・ネットワーク）
└── スクールICT・AI部門 AI @school  .claude/agents/school.md   （ICT支援員派遣）
```

## エージェントの呼び出し方

### 単独呼び出し

特定の部門に作業を依頼する場合:

```
@edu パソコン塾の新規クラスの時間割を考えて
@media ホームページのトップページ文章を更新して
@ithome 法人向けPC一括導入の見積を作って
@school 来月のICT支援員派遣スケジュールを確認して
@finance 第11期の進捗を確認して
@sales 新規開拓の営業リストを作って
@legal 業務委託契約書をチェックして
@hr ICT支援員の求人票を作って
```

### 並列呼び出し（同時稼働）

複数部門に跨る課題は、**1 つのメッセージで複数の @メンション** を使う:

```
@ithome と @finance を同時に呼んで、
法人向けPC50台一括導入の在庫調達とコストをそれぞれ報告させて
```

```
@sales と @legal と @finance を同時に呼んで、
新規法人顧客との年間契約の提案内容・契約条件・採算をそれぞれ報告させて
```

または CEO AI に委ねる:

```
新しい学校との ICT 支援員派遣契約の締結可否を判断して
（CEO AI が自動で関係部門を並列起動）
```

### バックグラウンド実行

長時間タスクを並列で走らせる場合は Ctrl+B またはプロンプトで「バックグラウンドで実行して」と指示する。

## ファイル構成

```
Claude code/
├── CLAUDE.md              ← このファイル（組織概要・使い方）
├── ceo.md                 ← CEO AI（佐藤太賀始）の役割・オーケストレーション定義
├── diagram.md             ← 組織・情報フローの構造図
├── agents/                ← 各部門の詳細ドキュメント（人間向け参照用）
│   ├── secretary/secretary.md
│   ├── finance/finance.md
│   ├── sales/sales.md
│   ├── legal/legal.md
│   ├── hr/hr.md
│   ├── edu/edu.md
│   ├── media/media.md
│   ├── ithome/ithome.md
│   └── school/school.md
├── shared/
│   └── memory.md          ← 全エージェント共有メモリ（会社情報・決定ログ・進捗）
└── .claude/
    ├── settings.local.json
    ├── agents/            ← Claude Code が読み込むサブエージェント定義
    │   ├── secretary.md
    │   ├── finance.md
    │   ├── sales.md
    │   ├── legal.md
    │   ├── hr.md
    │   ├── edu.md
    │   ├── media.md
    │   ├── ithome.md
    │   └── school.md
    └── skills/
        └── receipt-to-csv/SKILL.md  ← 仕入れ・消耗品レシートをCSV化するスキル
```

## 並列起動パターン集

| シナリオ | 同時起動する部門 |
|---|---|
| 新しい学校との ICT 支援員派遣契約 | school + finance |
| 法人向け PC 大量導入の受注判断 | ithome + finance |
| パソコン塾の新規クラス開講（検定・スクラッチ等） | edu + finance + secretary |
| ホームページ全面リニューアル／EC サイト新設 | media + finance |
| LINE 公式アカウント・SNS 運用の方針変更 | media + secretary |
| 第10期・第11期決算レビュー | finance + secretary（全部門サマリー集約） |
| 新規法人顧客との年間契約締結 | sales + legal + finance |
| 新規採用・雇用条件変更 | hr + finance + legal |
| 新サービス開始時の法的リスク確認 | legal + 関連事業部門 |

## 第10期決算ハイライト（参考）

| 項目 | 金額 |
|---|---|
| 純売上高 | 29,276,787円 |
| 売上総利益 | 19,964,884円 |
| 営業利益 | 172,951円 |
| 経常利益 | 154,794円 |
| 当期純利益 | 83,794円 |
| 資産合計 | 16,577,400円 |
| 純資産合計 | 1,140,784円 |

詳細・最新の状況は `shared/memory.md` のセクション 7「共有ナレッジ > 会社情報」および財務 AI のレポートを参照すること。
