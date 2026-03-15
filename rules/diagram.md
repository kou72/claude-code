# 作図ルール

SESプラットフォーム設計文書（2026年3月）の作図改善作業をもとに抽象化したプロセス。
生成した図は必ずセルフレビューを経てからドキュメントに埋め込む。

---

## ツール選定基準

**図の生成には必ず Graphviz DOT（kroki.io）を使う。Mermaid は使用禁止。**
Mermaid はラベル折り返し・絵文字文字化け・レイアウト制御の制約が多く、Graphviz の方が安定して読みやすい図が得られる。

| 用途 | ツール |
| ---- | ------ |
| フローチャート・関係図・プロセス図・その他すべての図 | **Graphviz DOT（kroki.io）** |
| 数値データ・グラフ | Chart.js（quickchart.io） |
| Mermaid | **使用禁止** |

---

## 生成コマンド

### Mermaid → PNG（kroki.io）

```bash
cat > diagram.mmd << 'EOF'
%%{init: {'theme': 'base', 'flowchart': {'htmlLabels': true}}}%%
flowchart TD
    ...
EOF

curl -s -X POST https://kroki.io/mermaid/png \
  -H "Content-Type: text/plain" \
  --data-binary @diagram.mmd \
  -o output.png

size=$(stat -c%s output.png)
echo "Size: $size bytes"
```

### Graphviz → PNG（kroki.io）

```bash
cat > diagram.dot << 'EOF'
digraph G {
    rankdir=LR
    compound=true   # クラスタ間矢印を使う場合は必須
    ...
}
EOF

curl -s -X POST https://kroki.io/graphviz/png \
  -H "Content-Type: text/plain" \
  --data-binary @diagram.dot \
  -o output.png
```

### Chart.js → PNG（quickchart.io）

```bash
cat > config.json << 'EOF'
{
  "version": 3,
  "chart": { ... },
  "width": 900,
  "height": 500,
  "backgroundColor": "white"
}
EOF

curl -s -X POST https://quickchart.io/chart \
  -H "Content-Type: application/json" \
  --data-binary "@config.json" \
  -o output.png

size=$(stat -c%s output.png)
if [ "$size" -lt 5000 ]; then cat output.png; fi   # エラーレスポンスを検出
```

---

## レイアウト方向の選び方（rankdir）

図を設計する前に、フローの形状から `rankdir` を決める。

| フローの形状 | rankdir | 理由 |
| --- | --- | --- |
| 一方向の直線フロー（A→B→C→D） | **TB（縦）** | 縦の方がスクロールで自然に読め、矢印ラベルが左右に配置されて見やすい |
| 分岐・網目・双方向が混在する図 | LR（横）または TB | 内容に応じて選ぶ。クラスタ横並びなら LR が見やすいことが多い |

**判断基準：図のノードが一本の経路に並ぶかどうかで決める。**
一本の経路 → TB。それ以外 → 実際にどちらが読みやすいか考えて選ぶ。

---

## 既知の制約と対処法

### Mermaid

| 問題 | 原因 | 対処 |
| ---- | ---- | ---- |
| ラベル内の `\n` がリテラル表示される | デフォルトは htmlLabels 無効 | `%%{init: {'flowchart': {'htmlLabels': true}}}%%` を追加し `<br/>` を使う |
| アイコン文字（絵文字）が `[]` で表示される | フォントに依存 | 絵文字を使わず、テキストのみにする |
| subgraph の枠・背景色が消える | Mermaid の制約 | 複雑なグループ図は Graphviz に切り替える |
| LR レイアウトでも長いラベルが折り返す | ボックス幅がノードサイズに依存 | Graphviz に切り替える（`shape=box` でテキストに合わせてボックスが広がる） |

### Graphviz

| 問題 | 原因 | 対処 |
| ---- | ---- | ---- |
| クラスタ間の矢印が描画されない | compound=true が未設定 | `graph [compound=true]` を追加し `ltail`/`lhead` を指定 |
| ノードが重なる | ranksep/nodesep が小さい | `ranksep=1.2`, `nodesep=0.5` 程度に調整 |
| 日本語フォントが文字化け | （kroki.io では発生しない） | kroki.io 経由では問題なし |

### Chart.js（quickchart.io）

| 問題 | 原因 | 対処 |
| ---- | ---- | ---- |
| `missing variable c or chart` エラー | JSON の `chart` キーが欠落 | `{"version": 3, "chart": {...}}` の形式で送る |
| callback 文字列が「is not a function」エラー | v2 では callback 文字列評価が不安定 | `"version": 3` を指定する |
| 対数スケールの目盛りが科学記数法（`1e+3`） | デフォルト表示 | `ticks: {display: false}` にして datalabels プラグインでバー上に数値を直接表示する |
| データラベルが上端で切れる | chart 領域の外にはみ出す | `layout.padding.top` を増やす（40px 程度）か y 軸の `max` を実データの 3 倍程度に設定する |

---

## セルフレビュープロセス（必須）

作図後は必ず Read ツールで画像を確認し、以下のチェックリストを実行する。
**1 項目でも NG があれば修正してから次へ進む。**

### チェックリスト

#### 技術的な正確性

- [ ] `\n` がリテラルテキストとして表示されていないか
- [ ] ラベルが途中で切れていないか（ボックスからはみ出し・省略）
- [ ] アイコン文字が `[]` などに文字化けしていないか
- [ ] エラーメッセージが画像に表示されていないか（Chart error: ... など）
- [ ] ファイルサイズが 5,000 bytes 以上か（小さい場合はエラーレスポンスの疑い）

#### レイアウト・視認性

- [ ] ノード/ボックスが重なっていないか
- [ ] 矢印の経路がクロスしすぎて読みにくくないか
- [ ] フォントサイズが小さすぎて読めない箇所がないか
- [ ] 色が他のノードと区別できているか（同系色が隣り合っていないか）
- [ ] タイトル・凡例・軸ラベルが表示されているか

#### 情報の正確性

- [ ] 図が表現したい概念・関係を正しく表しているか
- [ ] ラベルのテキストが内容と一致しているか
- [ ] 矢印の向きが正しいか（因果・フローの方向）

---

## NG 判定時の対応フロー

```text
NG 発見
  ↓
原因を特定（ツール制約か、設定ミスか、設計問題か）
  ↓
ツール制約 → 別ツールへ切り替えを検討（Mermaid → Graphviz など）
設定ミス   → パラメータを修正して再生成
設計問題   → ラベルを短縮 / レイアウト方向を変更 / 分割を検討
  ↓
再生成 → チェックリストを再実行（最大 3 回リトライ）
  ↓
3 回リトライ後も NG → ユーザーに状況を報告して代替手段を相談
```

---

## Markdown への埋め込み

```markdown
![図のキャプション](charts/figN_name.png)
```

- 画像ファイルは `<ドキュメントと同階層>/charts/` に保存する
- ファイル名は `figN_短い名称.png`（例: `fig1_current_structure.png`）
- キャプションはグラフタイトルと揃える
- 埋め込み前にセルフレビューを完了させること
