# Claude Code vs Copilot：Word / Excel / PowerPoint 活用比較

> 最終更新：2026年4月（最新情報に基づく）

---

## まず「どのCopilot」かを整理する

Microsoft Copilotには複数の製品があり、**Word/Excel/PowerPointに統合されているもの**は限られる。

| 製品名 | 対象 | Word/Excel/PPT統合 | 月額（目安） |
| ------ | ---- | :---: | ------ |
| **Microsoft 365 Copilot Business** | 中小企業（〜300名） | **◎ 深く統合** | $18〜21/人（2026年6月まで促進価格$18） |
| **Microsoft 365 Copilot Enterprise** | 中大企業（300名〜） | **◎ 深く統合** | $30/人 |
| **Copilot Pro** | 個人・家族向け | ○ 統合あり | $20/人 または $199/年（家族プラン込み） |
| **Microsoft Copilot（無料）** | 誰でも | ~~統合あり~~ **2026/4/15以降は廃止** | 無料 |
| **GitHub Copilot** | 開発者向け | × 対象外 | $10〜/人 |

### 重要：2026年4月15日に無料版Copilot ChatがOfficeから廃止

Microsoft は 2026/4/15 以降、ライセンスなしユーザーのWord・Excel・PowerPoint・OneNoteからCopilot Chatを撤去した。

- **2,000名以上の組織**：完全削除。有料ライセンスなしでは使用不可
- **2,000名未満の組織**：使用可能だが品質・速度が制限される「標準アクセス」に降格
- **Outlookのみ**：引き続き無料でCopilot Chat利用可

**以降の比較は「Microsoft 365 Copilot Business/Enterprise」を想定する。**

---

### Business と Enterprise の主な違い

| 項目 | Copilot Business | Copilot Enterprise |
| ---- | ---------------- | ------------------ |
| 対象規模 | 〜300名 | 制限なし |
| 月額 | $18〜21/人 | $30/人 |
| データ参照範囲 | 自分のメール・OneDrive・参加しているTeams | **組織全体**（Microsoft Graph経由） |
| コンプライアンス | 基本 | Microsoft Purview統合・eDiscovery・監査ログ |
| カスタムプラグイン | × | ○（Graph Connector等） |

---

## 機能比較：アプリ別（2026年4月時点）

### Word

| できること | Microsoft 365 Copilot | Claude Code |
| ---------- | :---: | :---: |
| 文章の下書き生成（リボンから直接） | ◎ | × |
| 既存文書の要約・改善提案 | ◎ | △（ファイル読み込み後にコピペ） |
| スタイル・トーンの変更 | ◎ | ○（コピペ作業が必要） |
| 社内テンプレートを参照した文書生成 | ○（SharePoint連携） | × |
| **Word Agent**（対話形式で長文レポート・提案書を一気に生成） | ◎（2026年新機能） | ○（コピペ前提） |
| 引用元の自動表示（Web・社内資料） | ◎（2026年新機能） | × |
| VBAマクロの生成 | △（基本的なもの） | ◎（複雑なマクロも対応） |

### Excel

| できること | Microsoft 365 Copilot | Claude Code |
| ---------- | :---: | :---: |
| 数式の生成・説明（セルを見ながら） | ◎ | △（スクリーンショット・コピペ必要） |
| ピボットテーブル・グラフの自動作成 | ◎ | × |
| **Excel Agent**（自然言語から多タブ・数式・データ付きブックを一括生成） | ◎（2026年新機能） | × |
| **ローカルファイルへの多段階編集**（Windows/Mac対応） | ◎（2026年新機能） | × |
| データの整形・クリーニング提案 | ◎ | ○（Pythonスクリプト生成で間接対応） |
| 複雑なVBA/マクロの生成 | △ | ◎ |
| Pythonスクリプトでのバッチ処理 | × | ◎ |
| 大量データの自動分析・洞察抽出 | ◎ | ○（スクリプト経由） |

### PowerPoint

| できること | Microsoft 365 Copilot | Claude Code |
| ---------- | :---: | :---: |
| テキスト・アジェンダからスライド生成 | ◎ | × |
| 既存スライドの要約・改善提案 | ◎ | △（pptx読み取りスクリプト経由） |
| **全スライドのフォント・デザイン一括統一** | ◎（2026年新機能） | × |
| **Researcher機能**（調査結果をPPT/PDF/インフォグラフィック/音声で出力） | ◎（2026年新機能） | × |
| スピーカーノートの自動生成 | ◎ | × |
| 社内資料を参照した資料生成 | ○（SharePoint連携） | × |
| VBAでスライド一括操作 | △ | ◎ |
| 特定フォーマット・構成の厳密な制御 | △ | ◎（スクリプトで正確に制御） |

---

## 総合評価

| 観点 | Microsoft 365 Copilot | Claude Code |
| ---- | :---: | :---: |
| **Officeアプリとの統合（リボン・セル操作）** | ◎ | × |
| **操作の手軽さ** | ◎ クリックひとつ | △ プロンプト・スクリプトが必要 |
| **コンテンツ生成（文章・スライド）** | ◎ | ○（コピペ前提） |
| **複雑なロジック・マクロ生成** | △ | ◎ |
| **大量ファイルの一括処理・自動化** | × | ◎ |
| **汎用的な思考・分析・壁打ち** | ○ | ◎ |
| **社内データ連携（SharePoint・Graph）** | ◎ | × |
| **コスト（1ライセンスあたり）** | 高 | 安（Claude Pro $20/月で複数用途） |

---

## 結論：どちらを選ぶか

### Microsoft 365 Copilot が向いているケース

- Word・Excel・PowerPointを日常業務の中心で使っている
- ITリテラシーが高くないメンバーも使う（クリックだけで動く）
- SharePoint・Teamsと連携して社内資料を参照させたい
- 組織全体で標準化したい

### Claude Code が向いているケース

- 複雑なVBA・Python処理でOfficeファイルを自動化したい
- Officeは出力先の一つで、**調査・分析・文章生成・戦略立案が主目的**
- コスト重視（1つのサブスクで幅広い用途をカバーしたい）
- プログラミングができるメンバーが中心

### 組み合わせ案

```text
日常のOffice操作・コンテンツ生成   → Microsoft 365 Copilot
複雑なマクロ・スクリプト生成       → Claude Code
調査・分析・戦略立案・壁打ち       → Claude Code
```

---

## コスト試算（150名規模）

| ライセンス | 月額/人 | 150名全員 | 備考 |
| ---------- | ------- | --------- | ---- |
| Copilot Business | $18（促進）〜$21 | 約40〜47万円/月 | 〜300名向け。2026/6末まで促進価格 |
| Copilot Enterprise | $30 | 約68万円/月 | 組織全体のGraph参照・コンプライアンス機能付き |
| Claude Pro（代表者のみ10名） | $20 | 約3万円/月 | 推進チーム・技術スタッフ向け |

**現実的な落としどころ：**
Copilot Business を頻度の高い利用者（数十名）に絞って付与 ＋ Claude Code を推進チームに付与。
全員付与は月額40〜70万円規模になるため、費用対効果の検証が先決。

---

Sources:

- [Microsoft 365 Copilot Plans and Pricing | Microsoft](https://www.microsoft.com/en-us/microsoft-365-copilot/pricing)
- [Advancing Microsoft 365: New capabilities and pricing update | Microsoft 365 Blog](https://www.microsoft.com/en-us/microsoft-365/blog/2025/12/04/advancing-microsoft-365-new-capabilities-and-pricing-update/)
- [What's New in Microsoft 365 Copilot – March 2026 | Microsoft Community Hub](https://techcommunity.microsoft.com/blog/microsoft365copilotblog/what%E2%80%99s-new-in-microsoft-365-copilot--march-2026/4506322)
- [Word, Excel and PowerPoint Agents in Microsoft 365 Copilot – February 2026 | Microsoft Community Hub](https://techcommunity.microsoft.com/blog/drivingadoptionblog/word-excel-and-powerpoint-agents-in-microsoft-365-copilot-overview--live-demo--f/4497690)
- [Microsoft Kills Free Copilot Chat in Word, Excel and PowerPoint: What Happens on April 15 | Office Watch](https://office-watch.com/2026/microsoft-removes-copilot-chat-word-excel-powerpoint-april-2026/)
- [Microsoft Copilot Pricing Explained: Plans, Cost & Licensing (2026 Guide) | Copilot Experts](https://copilot-experts.com/microsoft-copilot-pricing-guide/)
- [Copilot Business vs Enterprise: Which Plan for Executives? | AIA Copilot](https://aiacopilot.com/articles/copilot-business-vs-enterprise.html)
