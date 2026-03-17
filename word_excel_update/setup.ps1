# ============================================================
# setup.ps1
# Excel 更新ツール（Word更新ツール.xlsm）を自動生成するスクリプト
#
# 【実行方法】
#   PowerShell を開き、このファイルがあるフォルダで実行：
#   powershell -ExecutionPolicy Bypass -File setup.ps1
# ============================================================

param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($OutputPath -eq "") {
    $OutputPath = Join-Path $scriptDir "Word更新ツール.xlsm"
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Word 更新ツール セットアップ" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "出力先: $OutputPath" -ForegroundColor Gray
Write-Host ""

$excel = $null
$vbaImported = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    # シートを1枚だけ残してリネーム（変更箇所シート）
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    $sheetMain = $workbook.Worksheets.Item(1)
    $sheetMain.Name = "変更箇所"

    # README シートを追加（変更箇所の後ろ）
    $sheetReadme = $workbook.Worksheets.Add([System.Type]::Missing, $sheetMain)
    $sheetReadme.Name = "README"

    # ============================================================
    # README シート
    # ============================================================
    $r = 1

    # タイトル
    $cell = $sheetReadme.Cells.Item($r, 1)
    $cell.Value2 = "Word テンプレート自動更新ツール - 使い方"
    $cell.Font.Bold = $true
    $cell.Font.Size = 14
    $cell.Font.Color = [int]0x2E5B9E
    $r += 2

    # セクション見出しと本文を書くヘルパー関数
    function Write-Section($row, $title, $lines) {
        $hCell = $sheetReadme.Cells.Item($row, 1)
        $hCell.Value2 = $title
        $hCell.Font.Bold = $true
        $hCell.Font.Size = 11
        $hCell.Font.Color = [int]0x4472C4
        $hCell.Interior.Color = [int]0xDCE6F1
        $row++
        foreach ($line in $lines) {
            $sheetReadme.Cells.Item($row, 1).Value2 = $line
            $row++
        }
        return $row + 1  # セクション後に空行
    }

    $r = Write-Section $r "■ 概要" @(
        "Word テンプレートをコピーして、`$変数を置換した納品ドキュメントを作成します。",
        "テンプレートファイル自体は変更されません。"
    )

    $r = Write-Section $r "■ Word テンプレートの準備" @(
        "1. Word で納品物のテンプレートを作成する",
        "2. 差し替えたい箇所を `$変数名 の形式で記述する",
        "   例:  会社名の箇所 → `$company_name",
        "        日付の箇所   → `$date",
        "        金額の箇所   → `$amount",
        "3. 変数テキストを赤字にしておく（置換後も赤字のまま残ります）",
        "4. テンプレートを任意の場所に保存する"
    )

    $r = Write-Section $r "■ 毎回の操作手順" @(
        "1. 「変更箇所」シートを開く",
        "2. B1 にテンプレートファイルの絶対パスを入力",
        "   例: C:\Users\username\Documents\template\納品書_テンプレート.docx",
        "3. B2 に出力ファイルの絶対パスを入力（新しいファイル名を指定）",
        "   例: C:\Users\username\Documents\output\納品書_株式会社ABC_20260401.docx",
        "4. テーブルに変数と変更後テキストを入力",
        "   A列: `$変数名（例: `$company_name）",
        "   C列: 変更後テキスト（例: 株式会社ABC）",
        "5. 「Word を更新」ボタンをクリック"
    )

    $r = Write-Section $r "■ 処理の流れ" @(
        "1. テンプレートファイルを出力ファイル名でコピー",
        "2. コピーしたファイルを Word で開く",
        "3. テーブルの `$変数 を一括置換（書式はそのまま維持）",
        "4. 保存して完了",
        "",
        "→ テンプレートファイルは変更されません"
    )

    $r = Write-Section $r "■ 変数の命名規則" @(
        "・$ で始める（例: `$company_name）",
        "・半角英数字とアンダースコアのみ使用可",
        "・スペース・日本語は使用不可",
        "",
        "例: `$company_name / `$date / `$amount / `$project_title / `$delivery_date"
    )

    # 列幅調整
    $sheetReadme.Columns.Item(1).ColumnWidth = 80
    $sheetReadme.Columns.Item(1).WrapText = $true

    # ============================================================
    # 変更箇所シート
    # ============================================================

    # 行1: テンプレートファイルパス
    $sheetMain.Cells.Item(1, 1).Value2 = "テンプレートファイル："
    $sheetMain.Cells.Item(1, 1).Font.Bold = $true
    $pathRange1 = $sheetMain.Range("B1:G1")
    $pathRange1.Merge()
    $pathRange1.Value2 = "C:\Users\username\Documents\template\納品書_テンプレート.docx"
    $pathRange1.Font.Color = [int]0x333333

    # 行2: 出力ファイルパス
    $sheetMain.Cells.Item(2, 1).Value2 = "出力ファイル名："
    $sheetMain.Cells.Item(2, 1).Font.Bold = $true
    $pathRange2 = $sheetMain.Range("B2:G2")
    $pathRange2.Merge()
    $pathRange2.Value2 = "C:\Users\username\Documents\output\納品書_株式会社ABC_20260401.docx"
    $pathRange2.Font.Color = [int]0x333333

    # 行3: 注釈
    $noteRange = $sheetMain.Range("A3:G3")
    $noteRange.Merge()
    $noteRange.Value2 = "※ テンプレートをコピーして出力ファイルを作成します。テンプレートは変更されません。変数は `$変数名 の形式で Word に赤字で記述してください。"
    $noteRange.Font.Color = [int]0x888888
    $noteRange.Font.Size = 9
    $noteRange.Font.Italic = $true

    # 行5〜: 変数テーブル
    $dataStart = 5
    $samples = @(
        @('$company_name', "会社名",   "株式会社サンプル商事"),
        @('$date',         "契約日付", "2026年4月1日"),
        @('$amount',       "契約金額", "1,500,000円（税込）")
    )

    $sheetMain.Cells.Item($dataStart, 1).Value2 = "変数名（`$xxx）"
    $sheetMain.Cells.Item($dataStart, 2).Value2 = "説明（任意）"
    $sheetMain.Cells.Item($dataStart, 3).Value2 = "変更後テキスト"

    for ($i = 0; $i -lt $samples.Length; $i++) {
        $row = $dataStart + 1 + $i
        $sheetMain.Cells.Item($row, 1).Value2 = $samples[$i][0]
        $sheetMain.Cells.Item($row, 2).Value2 = $samples[$i][1]
        $sheetMain.Cells.Item($row, 3).Value2 = $samples[$i][2]
    }

    $lastDataRow = $dataStart + $samples.Length
    $tableRange = $sheetMain.Range("A${dataStart}:C${lastDataRow}")
    $table = $sheetMain.ListObjects.Add(1, $tableRange, [System.Type]::Missing, 1)
    $table.Name = "変数テーブル"
    $table.TableStyle = "TableStyleMedium2"

    $table.ListColumns.Item(1).Range.ColumnWidth = 22
    $table.ListColumns.Item(2).Range.ColumnWidth = 18
    $table.ListColumns.Item(3).Range.ColumnWidth = 45

    # 「Word を更新」ボタン（行1〜2 の右側）
    $btn = $sheetMain.Shapes.AddShape(1, 460, 3, 130, 36)
    $btn.Name = "btnUpdate"
    $btn.TextFrame.Characters().Text = "Word を更新"
    $btn.TextFrame.Characters().Font.Bold = $true
    $btn.TextFrame.Characters().Font.Size = 11
    $btn.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btn.Fill.ForeColor.RGB = [int]0x4472C4
    $btn.Line.ForeColor.RGB = [int]0x2E5B9E
    $btn.Line.Weight = 1
    $btn.TextFrame.HorizontalAlignment = -4108
    $btn.TextFrame.VerticalAlignment   = -4108
    $btn.OnAction = "WordUpdater.UpdateWordDocument"

    # ウィンドウ枠の固定（4行目まで）
    $sheetMain.Activate()
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # 変更箇所シートをアクティブにして保存
    $sheetMain.Activate()

    # ============================================================
    # VBA マクロを埋め込む
    # ============================================================
    $basFilePath = Join-Path $scriptDir "WordUpdater.bas"
    if (Test-Path $basFilePath) {
        try {
            if ($null -eq $workbook.VBProject) {
                throw "VBProject が null です。"
            }
            $vbaCode = [System.IO.File]::ReadAllText($basFilePath, [System.Text.Encoding]::GetEncoding(932))
            $vbaCode = $vbaCode -replace 'Attribute VB_Name = "WordUpdater"\r?\n', ''

            $vbaModule = $workbook.VBProject.VBComponents.Add(1)
            $vbaModule.Name = "WordUpdater"
            $vbaModule.CodeModule.AddFromString($vbaCode)
            $vbaImported = $true
            Write-Host "VBA マクロを追加しました。" -ForegroundColor Green
        } catch {
            Write-Host ""
            Write-Host "[警告] VBA マクロの自動埋め込みに失敗しました。" -ForegroundColor Yellow
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  【対処方法】Excel のセキュリティ設定を変更してください:" -ForegroundColor Yellow
            Write-Host "    1. Excel を起動"
            Write-Host "    2. ファイル → オプション → セキュリティセンター"
            Write-Host "    3. セキュリティセンターの設定 → マクロの設定"
            Write-Host "    4. [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] にチェック"
            Write-Host "    5. OK で閉じて、再度 setup.ps1 を実行"
            Write-Host ""
            Write-Host "  または、Excel を開いた後に手動でインポートしてください:"
            Write-Host "    Alt+F11 → ファイル → ファイルのインポート → WordUpdater.bas"
            Write-Host ""
        }
    } else {
        Write-Host "WordUpdater.bas が見つかりません。マクロは手動でインポートしてください。" -ForegroundColor Yellow
    }

    # 保存
    $workbook.SaveAs($OutputPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    Write-Host "ファイルを作成しました: $OutputPath" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
} finally {
    if ($null -ne $excel) {
        try { $workbook.Close($false) } catch {}
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
if ($vbaImported) {
    Write-Host " セットアップ完了！" -ForegroundColor Cyan
} else {
    Write-Host " Excel ファイルを作成しました（マクロは手動で追加が必要）" -ForegroundColor Yellow
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "次のステップ:"
if (-not $vbaImported) {
    Write-Host "  [0. まず] Word更新ツール.xlsm を開き、Alt+F11 → ファイル → ファイルのインポート → WordUpdater.bas" -ForegroundColor Yellow
}
Write-Host "  1. Word でテンプレートを作成し、変数を `$変数名（赤字）で記述する"
Write-Host "  2. 「変更箇所」シートの B1 にテンプレートパス、B2 に出力パスを入力"
Write-Host "  3. テーブルに `$変数名と変更後テキストを入力"
Write-Host "  4. [Word を更新] ボタンをクリック"
Write-Host ""
