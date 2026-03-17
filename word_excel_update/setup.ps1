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

    # シートを1枚だけ残して「変更箇所」にリネーム
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    $sheetMain = $workbook.Worksheets.Item(1)
    $sheetMain.Name = "変更箇所"

    # 「実行」シートを変更箇所の前に追加
    $sheetExec = $workbook.Worksheets.Add($sheetMain)
    $sheetExec.Name = "実行"

    # 「README」シートを実行の前（先頭）に追加
    $sheetReadme = $workbook.Worksheets.Add($sheetExec)
    $sheetReadme.Name = "README"

    # シート順: README | 実行 | 変更箇所

    # ============================================================
    # README シート
    # ============================================================
    $r = 1

    # タイトル
    $cell = $sheetReadme.Cells.Item($r, 1)
    $cell.Value2 = "Word テンプレート自動更新ツール - 使い方"
    $cell.Font.Bold = $true
    $cell.Font.Size = 14
    $cell.Font.Color = [int]0x9E5B2E
    $r += 2

    function Write-Section($row, $title, $lines) {
        $hCell = $sheetReadme.Cells.Item($row, 1)
        $hCell.Value2 = $title
        $hCell.Font.Bold = $true
        $hCell.Font.Size = 11
        $hCell.Font.Color = [int]0xC47244
        $hCell.Interior.Color = [int]0xF1E6DC
        $row++
        foreach ($line in $lines) {
            $sheetReadme.Cells.Item($row, 1).Value2 = $line
            $row++
        }
        return $row + 1
    }

    $r = Write-Section $r "■ 概要" @(
        "Word テンプレートをコピーして、`$変数を置換した納品ドキュメントを作成します。",
        "「実行」シートでシートと実行フラグを管理し、複数ドキュメントを一括生成できます。",
        "テンプレートファイル自体は変更されません。"
    )

    $r = Write-Section $r "■ Word テンプレートの準備" @(
        "1. Word でテンプレートを作成する",
        "2. 差し替えたい箇所を `$変数名 の形式で記述する",
        "   例:  会社名の箇所 → `$company_name",
        "        日付の箇所   → `$date",
        "        金額の箇所   → `$amount",
        "3. 変数テキストを赤字にしておく（置換後も赤字のまま残ります）",
        "4. テンプレートを任意の場所に保存する"
    )

    $r = Write-Section $r "■ 毎回の操作手順" @(
        "【各変更箇所シートの設定】",
        "  C1: テンプレートファイルの絶対パス",
        "  C2: 出力ファイルの絶対パス（新しいファイル名）",
        "  テーブル A列: `$変数名、C列: 変更後テキスト",
        "",
        "【実行】",
        "1. 「実行」シートを開く",
        "2. 実行テーブルのシート名列に処理するシート名を入力",
        "3. 実行フラグ列を「yes」に設定（プルダウンで選択）",
        "4. 「一括実行」ボタンをクリック",
        "5. 確認ダイアログで「はい」を選択"
    )

    $r = Write-Section $r "■ 処理の流れ" @(
        "1. 実行テーブルで「yes」のシートを順番に処理",
        "2. テンプレートファイルを出力ファイル名でコピー",
        "3. コピーしたファイルを Word で開く",
        "4. テーブルの `$変数 を一括置換（書式はそのまま維持）",
        "5. 保存して次のシートへ",
        "6. 全シート完了後に結果サマリーを表示"
    )

    $sheetReadme.Columns.Item(1).ColumnWidth = 80
    $sheetReadme.Columns.Item(1).WrapText = $true

    # ============================================================
    # 実行シート
    # ============================================================

    # 行1: タイトル
    $execTitle = $sheetExec.Cells.Item(1, 1)
    $execTitle.Value2 = "■ 実行設定"
    $execTitle.Font.Bold = $true
    $execTitle.Font.Size = 13
    $execTitle.Font.Color = [int]0x9E5B2E

    # 行2: 説明
    $execNote = $sheetExec.Cells.Item(2, 1)
    $execNote.Value2 = "実行フラグを「yes」に設定したシートを一括処理します。「一括実行」ボタンをクリックしてください。"
    $execNote.Font.Color = [int]0xC47244
    $execNote.Font.Size = 9
    $execNote.Font.Italic = $true

    # 行3: 空行
    # 行4〜: 実行テーブル
    $execDataStart = 4
    $sheetExec.Cells.Item($execDataStart,     1).Value2 = "シート名"
    $sheetExec.Cells.Item($execDataStart,     2).Value2 = "実行フラグ"
    $sheetExec.Cells.Item($execDataStart + 1, 1).Value2 = "変更箇所"
    $sheetExec.Cells.Item($execDataStart + 1, 2).Value2 = "yes"

    $execLastRow = $execDataStart + 1
    $execTableRange = $sheetExec.Range("A${execDataStart}:B${execLastRow}")
    $execTable = $sheetExec.ListObjects.Add(1, $execTableRange, [System.Type]::Missing, 1)
    $execTable.Name = "実行テーブル"
    $execTable.TableStyle = "TableStyleMedium9"

    $execTable.ListColumns.Item(1).Range.ColumnWidth = 28
    $execTable.ListColumns.Item(2).Range.ColumnWidth = 14

    # 実行フラグ列にドロップダウン検証を追加
    $flagBody = $execTable.ListColumns.Item(2).DataBodyRange
    $flagBody.Validation.Delete()
    $flagBody.Validation.Add(3, 1, 1, "yes,no")   # xlValidateList=3
    $flagBody.Validation.IgnoreBlank = $true
    $flagBody.Validation.InCellDropdown = $true

    # 「一括実行」ボタン（右上）
    $btnExec = $sheetExec.Shapes.AddShape(1, 350, 3, 140, 36)
    $btnExec.Name = "btnRunAll"
    $btnExec.TextFrame.Characters().Text = "一括実行"
    $btnExec.TextFrame.Characters().Font.Bold = $true
    $btnExec.TextFrame.Characters().Font.Size = 11
    $btnExec.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btnExec.Fill.ForeColor.RGB = [int]0xC47244
    $btnExec.Line.ForeColor.RGB = [int]0x9E5B2E
    $btnExec.Line.Weight = 1
    $btnExec.TextFrame.HorizontalAlignment = -4108
    $btnExec.TextFrame.VerticalAlignment   = -4108
    $btnExec.OnAction = "WordUpdater.RunAll"

    # ウィンドウ枠の固定（行4 = テーブルヘッダーを常時表示）
    $sheetExec.Activate()
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # ============================================================
    # 変更箇所シート
    # ============================================================

    # 行1: テンプレートファイルパス
    $labelCell1 = $sheetMain.Cells.Item(1, 1)
    $labelCell1.Value2 = "テンプレートファイル："
    $labelCell1.Font.Bold = $true
    $labelCell1.Font.Color = [int]0x333333
    $labelCell1.Interior.ColorIndex = -4142
    $pathRange1 = $sheetMain.Range("B1:G1")
    $pathRange1.Merge()
    $pathRange1.Value2 = "C:\Users\username\Documents\template\納品書_テンプレート.docx"
    $pathRange1.Font.Color = [int]0x9E5B2E
    $pathRange1.Interior.ColorIndex = -4142

    # 行2: 出力ファイルパス
    $labelCell2 = $sheetMain.Cells.Item(2, 1)
    $labelCell2.Value2 = "出力ファイル名："
    $labelCell2.Font.Bold = $true
    $labelCell2.Font.Color = [int]0x333333
    $labelCell2.Interior.ColorIndex = -4142
    $pathRange2 = $sheetMain.Range("B2:G2")
    $pathRange2.Merge()
    $pathRange2.Value2 = "C:\Users\username\Documents\output\納品書_株式会社ABC_20260401.docx"
    $pathRange2.Font.Color = [int]0x9E5B2E
    $pathRange2.Interior.ColorIndex = -4142

    # 行3: 注釈
    $noteRange = $sheetMain.Range("A3:G3")
    $noteRange.Merge()
    $noteRange.Value2 = "※ テンプレートをコピーして出力ファイルを作成します。テンプレートは変更されません。変数は `$変数名 の形式で Word に赤字で記述してください。"
    $noteRange.Font.Color = [int]0xC47244
    $noteRange.Font.Size = 9
    $noteRange.Font.Italic = $true
    $noteRange.Interior.Color = [int]0xFBF3EE

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
    $table.TableStyle = "TableStyleMedium9"

    $table.ListColumns.Item(1).Range.ColumnWidth = 22
    $table.ListColumns.Item(2).Range.ColumnWidth = 18
    $table.ListColumns.Item(3).Range.ColumnWidth = 45

    # ウィンドウ枠の固定（4行目まで）
    $sheetMain.Activate()
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # 実行シートをアクティブにして保存
    $sheetExec.Activate()

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
Write-Host "  2. 「変更箇所」シートの C1 にテンプレートパス、C2 に出力パスを入力"
Write-Host "  3. テーブルに `$変数名と変更後テキストを入力"
Write-Host "  4. 「実行」シートの実行テーブルにシート名を追加し、フラグを「yes」に設定"
Write-Host "  5. 「一括実行」ボタンをクリック"
Write-Host ""
