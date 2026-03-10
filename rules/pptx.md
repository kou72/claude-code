# PowerPoint 読み込みルール

PowerPoint ファイル（`.pptx`）を扱う際は、必ず以下のプロセスを実行する。

## 処理スクリプト

`scripts/pptx_read.ps1` を使う。

```bash
powershell.exe -ExecutionPolicy Bypass -File "scripts/pptx_read.ps1" -PptxPath "path/to/file.pptx"
```

**出力先**（デフォルト：pptxと同じディレクトリに同名フォルダ）:

```
{pptxファイル名}/
  slides/        # 各スライドのPNG画像
  content.md     # 全スライドのテキスト＋画像参照
```

出力先を変えたい場合は `-OutDir` オプションで指定する。

## 前提条件

- **PowerPoint を閉じた状態**で実行すること（起動中だと COM エラーになる）
- PowerPoint がインストールされていること

## 実行後の読み込み方

スクリプト実行後、以下の順で内容を把握する：

1. `content.md` を Read ツールで読む（テキスト全体の把握）
2. 必要なスライドの PNG を Read ツールで読む（図・レイアウトの把握）

## PowerPoint が開いている場合

COM エラー（`RPC_E_CALL_REJECTED`）が出た場合は PowerPoint を閉じてから再実行する。
