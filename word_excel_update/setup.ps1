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

# 出力先の決定
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

# ---- Excel COM オブジェクト起動 ----
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    # 既存シートを削除（1枚だけ残す）
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "変更箇所"

    # ============================================================
    # レイアウト設定
    # ============================================================

    # ---- 行1: Wordファイルパス ----
    $cell = $sheet.Cells.Item(1, 1)
    $cell.Value2 = "Wordファイルパス："
    $cell.Font.Bold = $true
    $cell.Font.Size = 11

    $mergeRange = $sheet.Range("B1:F1")
    $mergeRange.Merge()
    $mergeRange.Value2 = "C:\Users\username\Documents\納品物.docx"
    $mergeRange.Font.Color = [int]0x333333

    # ---- 行2: 補足説明 ----
    $sheet.Range("A2:F2").Merge()
    $noteCell = $sheet.Cells.Item(2, 1)
    $noteCell.Value2 = "※ B1 に Word ファイルの絶対パスを入力してください。ブックマーク名は Word 側で事前に設定が必要です。"
    $noteCell.Font.Color = [int]0x888888
    $noteCell.Font.Size = 9
    $noteCell.Font.Italic = $true

    # ---- 行3: ヘッダー ----
    $headers = @("ブックマーク名", "説明（任意）", "変更後テキスト", "状態")
    $headerBgColor = [int]0x4472C4
    $headerFontColor = [int]0xFFFFFF

    for ($col = 1; $col -le $headers.Length; $col++) {
        $hCell = $sheet.Cells.Item(3, $col)
        $hCell.Value2 = $headers[$col - 1]
        $hCell.Font.Bold = $true
        $hCell.Font.Color = $headerFontColor
        $hCell.Interior.Color = $headerBgColor
        $hCell.HorizontalAlignment = -4108  # xlCenter
        $hCell.VerticalAlignment   = -4108  # xlCenter
        $hCell.RowHeight = 22
    }

    # ---- 行4〜6: サンプルデータ ----
    $samples = @(
        @("bm_company_name", "会社名",   "株式会社サンプル商事"),
        @("bm_date",         "契約日付", "2026年4月1日"),
        @("bm_amount",       "契約金額", "1,500,000円（税込）")
    )

    $row = 4
    foreach ($s in $samples) {
        $sheet.Cells.Item($row, 1).Value2 = $s[0]
        $sheet.Cells.Item($row, 2).Value2 = $s[1]
        $sheet.Cells.Item($row, 3).Value2 = $s[2]
        if ($row % 2 -eq 0) {
            $sheet.Range("A${row}:D${row}").Interior.Color = [int]0xF2F7FF
        }
        $row++
    }

    # ---- 列幅調整 ----
    $sheet.Columns.Item(1).ColumnWidth = 25
    $sheet.Columns.Item(2).ColumnWidth = 18
    $sheet.Columns.Item(3).ColumnWidth = 45
    $sheet.Columns.Item(4).ColumnWidth = 8

    # ---- ウィンドウ枠の固定（3行目まで固定）----
    $sheet.Activate()
    $excel.ActiveWindow.SplitRow = 3
    $excel.ActiveWindow.FreezePanes = $true

    # ============================================================
    # ボタン追加
    # ============================================================

    # 「Word を更新」ボタン
    $btnUpdate = $sheet.Shapes.AddShape(1, 280, 4, 130, 26)
    $btnUpdate.Name = "btnUpdate"
    $btnUpdate.TextFrame.Characters().Text = "Word を更新"
    $btnUpdate.TextFrame.Characters().Font.Bold = $true
    $btnUpdate.TextFrame.Characters().Font.Size = 10
    $btnUpdate.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btnUpdate.Fill.ForeColor.RGB = [int]0x4472C4
    $btnUpdate.Line.ForeColor.RGB = [int]0x2E5B9E
    $btnUpdate.Line.Weight = 1
    $btnUpdate.TextFrame.HorizontalAlignment = -4108
    $btnUpdate.TextFrame.VerticalAlignment   = -4108
    $btnUpdate.OnAction = "WordUpdater.UpdateWordDocument"

    # 「状態リセット」ボタン
    $btnReset = $sheet.Shapes.AddShape(1, 420, 4, 110, 26)
    $btnReset.Name = "btnReset"
    $btnReset.TextFrame.Characters().Text = "状態リセット"
    $btnReset.TextFrame.Characters().Font.Size = 9
    $btnReset.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btnReset.Fill.ForeColor.RGB = [int]0x808080
    $btnReset.Line.ForeColor.RGB = [int]0x606060
    $btnReset.Line.Weight = 1
    $btnReset.TextFrame.HorizontalAlignment = -4108
    $btnReset.TextFrame.VerticalAlignment   = -4108
    $btnReset.OnAction = "WordUpdater.ResetStatus"

    # 「ブックマーク一覧取得」ボタン
    $btnImport = $sheet.Shapes.AddShape(1, 540, 4, 150, 26)
    $btnImport.Name = "btnImport"
    $btnImport.TextFrame.Characters().Text = "ブックマーク一覧取得"
    $btnImport.TextFrame.Characters().Font.Size = 9
    $btnImport.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btnImport.Fill.ForeColor.RGB = [int]0x538135
    $btnImport.Line.ForeColor.RGB = [int]0x3B5E26
    $btnImport.Line.Weight = 1
    $btnImport.TextFrame.HorizontalAlignment = -4108
    $btnImport.TextFrame.VerticalAlignment   = -4108
    $btnImport.OnAction = "WordUpdater.ImportBookmarkList"

    # ============================================================
    # VBA マクロコードを追加
    # ============================================================
    $basFilePath = Join-Path $scriptDir "WordUpdater.bas"
    $vbaImported = $false
    if (Test-Path $basFilePath) {
        # VBProject へのアクセスを試みる
        # 失敗する場合は Excel のセキュリティ設定で
        # 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください
        try {
            if ($null -eq $workbook.VBProject) {
                throw "VBProject が null です。"
            }
            $vbaCode = [System.IO.File]::ReadAllText($basFilePath, [System.Text.Encoding]::UTF8)
            $vbaCode = $vbaCode -replace 'Attribute VB_Name = "WordUpdater"\r?\n', ''

            $vbaModule = $workbook.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
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
            Write-Host "  または、Excel ファイルを開いた後に手動でマクロをインポートできます:"
            Write-Host "    Alt+F11 → ファイル → ファイルのインポート → WordUpdater.bas"
            Write-Host ""
        }
    } else {
        Write-Host "WordUpdater.bas が見つかりません。マクロは手動でインポートしてください。" -ForegroundColor Yellow
    }

    # ============================================================
    # .xlsm として保存
    # ============================================================
    $workbook.SaveAs($OutputPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
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
Write-Host "  1. Word ドキュメントを開き、Alt+F11 で VBA エディタを起動"
Write-Host "  2. WordBookmarkHelper.bas をインポートしてブックマークを設定"
Write-Host "  3. Word更新ツール.xlsm を開き、B1 に Word ファイルのパスを入力"
Write-Host "  4. A列にブックマーク名、C列に変更後テキストを入力"
Write-Host "  5. [Word を更新] ボタンをクリック"
Write-Host ""
