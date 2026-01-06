
# comparison-table-v2

> 規程・契約書等の改訂前後を比較する新旧対比表を自動生成

[![Version](https://img.shields.io/badge/version-2.0-blue.svg)](https://github.com/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/)
[![Status](https://img.shields.io/badge/status-stable-success.svg)](https://github.com/)

## 📖 Overview

**comparison-table-v2**は、Word文書やPDFの改訂前後を比較し、プロフェッショナルな新旧対比表を自動生成するAgent Skillです。就業規則、契約書、社内規程などの改訂履歴管理に最適です。

### ✨ Features

* 📄 **自動階層検出** - 14種類の見出しパターンに対応（第○条、第○章、(1)等）
* 🎨 **色分け表示** - 追加・変更・削除を視覚的に識別
* 🔍 **精密比較** - 文字レベルでの差分検出
* 📝 **Word出力** - 編集可能なdocx形式で出力
* 🇯🇵 **日本語完全対応** - 条文番号・階層構造の自動認識

## 🚀 Quick Start

### Installation

```bash
# Agent Skillとして利用（Claudeに話しかけるだけ）
"新旧対比表を作成して"
```

### Basic Usage

```plaintext
User: "新旧対比表を作成して"
Claude: 旧版と新版のファイルをアップロードしてください。

[旧版ファイルをアップロード]
[新版ファイルをアップロード]

Claude: 対比表を作成しました。
        → comparison_table_就業規則_20250106.docx
```

## 📋 Table of Contents

* [Installation](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#installation)
* [Usage](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#usage)
* [Supported Formats](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#supported-formats)
* [Output Structure](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#output-structure)
* [Examples](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#examples)
* [Troubleshooting](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#troubleshooting)
* [Best Practices](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#best-practices)
* [FAQ](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#faq)
* [Contributing](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#contributing)
* [License](https://claude.ai/chat/f920e95c-d7fd-436d-acc8-ce4341c14cca#license)

## 💻 Usage

### Trigger Phrases

以下のフレーズでスキルを起動できます：

```plaintext
"新旧対比表を作成して"
"新旧対比表を作って"
"対比表を生成"
"比較表を作成"
```

### Input Requirements

| Item         | Description  | Format       |
| ------------ | ------------ | ------------ |
| 旧版ファイル | 改訂前の文書 | .docx / .pdf |
| 新版ファイル | 改訂後の文書 | .docx / .pdf |

### Output

```plaintext
comparison_table_{文書名}_{YYYYMMDD}.docx
```

## 📁 Supported Formats

### Document Types

* ✅ **Word文書** (.docx) - **推奨**
* ✅ **PDF文書** (.pdf) - テキスト抽出可能なもの
* ❌ スキャンPDF - 非対応

### Heading Patterns (14 types)

| Pattern  | Example       | Use Case   |
| -------- | ------------- | ---------- |
| 条番号   | 第1条、第2条  | 法令、規程 |
| 章番号   | 第1章、第2章  | マニュアル |
| 項番号   | 第1項、第2項  | 契約書     |
| 号番号   | 第1号、第2号  | 細則       |
| 括弧数字 | (1)、(2)、(3) | 詳細項目   |
| 丸数字   | ①、②、③    | リスト項目 |

## 📊 Output Structure

### Table Format

```
┌────────────┬──────────────┬──────────────┐
│   項目     │     新       │     旧       │
├────────────┼──────────────┼──────────────┤
│ 第1条      │ 改訂後の内容 │ 改訂前の内容 │
│ 第2条      │ 改訂後の内容 │ 改訂前の内容 │
└────────────┴──────────────┴──────────────┘
```

### Color Coding

| Change Type | Color     | Description        |
| ----------- | --------- | ------------------ |
| 追加        | 🟢 Green  | 新規追加された内容 |
| 変更        | 🟡 Yellow | 修正された内容     |
| 削除        | 🔴 Red    | 削除された内容     |
| 変更なし    | ⚪ Normal | 変更のない内容     |

## 📝 Examples

### Example 1: 就業規則の改訂

```plaintext
Input:
  - 就業規則_2024年版.docx (旧版)
  - 就業規則_2025年版.docx (新版)

Output:
  - comparison_table_就業規則_20250106.docx
  
主な変更点:
  - 第15条（休暇）: 年次有給休暇の付与日数を変更
  - 第22条（退職金）: 新規追加
  - 第30条（懲戒）: 条文の一部を削除
```

### Example 2: 契約書の変更確認

```plaintext
Input:
  - 業務委託契約書_初版.pdf (旧版)
  - 業務委託契約書_改訂版.pdf (新版)

Output:
  - comparison_table_業務委託契約書_20250106.docx

主な変更点:
  - 第5条（報酬）: 月額20万円→30万円
  - 第8条（損害賠償）: 上限額を追加（300万円）
  - 第12条（契約期間）: 自動更新条項を追加
```

### Example 3: 社内規程の更新

```plaintext
Input:
  - 情報セキュリティ規程_v1.0.docx (旧版)
  - 情報セキュリティ規程_v2.0.docx (新版)

Output:
  - comparison_table_情報セキュリティ規程_20250106.docx

主な変更点:
  - 第3章: 「生成AI利用規定」を新規追加
  - 第7条: パスワード要件を強化
  - 第15条: インシデント対応手順を詳細化
```

## 🔧 Troubleshooting

### Issue: 見出しが正しく認識されない

 **Cause** : 見出しスタイルが適用されていない、または非標準的な書式

 **Solution** :

```plaintext
Word文書の場合:
  1. 見出しスタイル（見出し1、見出し2等）を適用
  2. 条文番号を明確に記載（例: 第1条、第2条）
  3. 階層構造を統一

PDF文書の場合:
  1. テキスト抽出可能なPDFか確認
  2. 必要に応じてWordに変換してから処理
```

### Issue: 変更箇所が多すぎて見づらい

 **Solution** :

```plaintext
1. Excelに変換してフィルター機能を活用
   Word → 表をコピー → Excelに貼り付け
   
2. 重要な変更箇所のみを抜粋した要約版を作成
   「主要な変更点のみをハイライトして」と依頼
```

### Issue: 列の順序を変更したい

 **Current** : デフォルトは「新→旧」の順序

 **Workaround** :

```plaintext
1. 生成後にWord文書内で列を手動で入れ替え
2. または「旧版を左、新版を右にして」と明示的に指示
```

## 🎯 Best Practices

### 準備段階

```plaintext
✅ ファイル名に版情報を明記
   Good: 就業規則_2024年版.docx
   Bad:  就業規則.docx

✅ 見出しスタイルを統一
   - Word文書では見出しスタイルを活用
   - 条文番号を一貫した形式で記載

✅ 日付情報を含める
   - ファイル名またはヘッダーに改訂日を記載
```

### 生成後

```plaintext
✅ 目視確認を実施
   - 自動生成結果を必ず確認
   - 重要な変更点にコメント追加

✅ レビュープロセス
   - 承認者による最終確認
   - 必要に応じて修正

✅ 適切な保管
   - 原本と一緒に保管
   - バージョン管理システムと連携
```

### Version Control

```plaintext
推奨フォルダ構成:

/documents
  /就業規則
    ├── 就業規則_2024年版.docx
    ├── 就業規則_2025年版.docx
    └── /comparisons
        └── comparison_table_就業規則_20250106.docx
```

## 🏢 Use Cases

### 法務・コンプライアンス部門

* 社内規程の改訂履歴管理
* 契約書のバージョン管理
* コンプライアンス監査対応

### 人事部門

* 就業規則の改訂記録
* 給与規程の変更履歴
* 福利厚生制度の新旧比較

### 品質管理部門

* 品質マニュアルの改訂管理
* 作業手順書の更新記録
* ISO文書の版管理

## ❓ FAQ

### Q: どのような文書に適していますか？

A: 条文番号や章構成が明確な文書に最適です：

* 就業規則、社内規程
* 契約書、覚書
* マニュアル、手順書
* ISO文書、品質管理文書

### Q: PDFとWordのどちらが推奨ですか？

A: **Word文書（.docx）を強く推奨**します。理由：

* 見出しスタイルの情報を活用できる
* 階層構造の認識精度が高い
* テキスト抽出が確実

### Q: 大量の文書を一括処理できますか？

A: 現在のバージョンでは1組ずつの処理です。
複数文書を処理したい場合は、繰り返し実行してください。

### Q: 生成された対比表は編集できますか？

A: はい、Word形式（.docx）で出力されるため自由に編集可能です。

### Q: データの機密性は保たれますか？

A: はい、以下の保護措置があります：

* 処理後にファイルは自動削除
* 通信は暗号化
* ログには残りません

## 🛠️ Technical Details

### Architecture

```plaintext
Input (Word/PDF)
    ↓
Text Extraction
    ↓
Heading Detection (14 patterns)
    ↓
Hierarchical Structure Analysis
    ↓
Diff Calculation
    ↓
Color Coding & Formatting
    ↓
Output (Word .docx)
```

### Dependencies

* `python-docx`: Word文書処理
* `PyPDF2`: PDF読み込み
* `difflib`: 差分計算
* `openpyxl`: Excel連携（オプション）

## 🤝 Contributing

改善提案や不具合報告は以下まで：

```plaintext
Email: ken@miraichi.jp
Project: MIRAICHI Agent Skills
```

## 📄 License

MIT License - 自由に使用・改変・配布が可能です。

## 📞 Support

### Contact

* **Developer** : 株式会社けんけん (MIRAICHI)
* **Email** : ken@miraichi.jp
* **Documentation** : 本READMEおよび社内Wiki

### Version History

| Version | Date       | Changes                        |
| ------- | ---------- | ------------------------------ |
| 2.0     | 2025-01-06 | 14パターン対応、色分け機能強化 |
| 1.0     | 2024-12-01 | 初回リリース                   |

---

**Made with by けんけん**
