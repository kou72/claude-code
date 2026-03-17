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

$excel = $null
$vbaImported = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    # シートを1枚だけ残す
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "変更箇所"

    # ============================================================
    # 行1: Word ファイルパス入力欄
    # ============================================================
    $labelCell = $sheet.Cells.Item(1, 1)
    $labelCell.Value2 = "Wordファイルパス："
    $labelCell.Font.Bold = $true

    $pathRange = $sheet.Range("B1:G1")
    $pathRange.Merge()
    $pathRange.Value2 = "C:\Users\username\Documents\テンプレート.docx"
    $pathRange.Font.Color = [int]0x333333

    # ============================================================
    # 行2: 注釈
    # ============================================================
    $noteRange = $sheet.Range("A2:G2")
    $noteRange.Merge()
    $noteRange.Value2 = "※ Word テンプレートの変数は `$変数名 の形式で赤字にしておいてください。置換後も赤字のまま維持されます。"
    $noteRange.Font.Color = [int]0x888888
    $noteRange.Font.Size = 9
    $noteRange.Font.Italic = $true

    # ============================================================
    # 行4〜: 変数テーブル（Excel テーブル形式）
    # ============================================================

    # サンプルデータをセルに書き込む
    $dataStart = 4
    $samples = @(
        @('$company_name', "会社名",   "株式会社サンプル商事"),
        @('$date',         "契約日付", "2026年4月1日"),
        @('$amount',       "契約金額", "1,500,000円（税込）")
    )

    # ヘッダー行（行4）
    $sheet.Cells.Item($dataStart, 1).Value2 = "変数名（`$xxx）"
    $sheet.Cells.Item($dataStart, 2).Value2 = "説明（任意）"
    $sheet.Cells.Item($dataStart, 3).Value2 = "変更後テキスト"

    # データ行（行5〜7）
    for ($r = 0; $r -lt $samples.Length; $r++) {
        $row = $dataStart + 1 + $r
        $sheet.Cells.Item($row, 1).Value2 = $samples[$r][0]
        $sheet.Cells.Item($row, 2).Value2 = $samples[$r][1]
        $sheet.Cells.Item($row, 3).Value2 = $samples[$r][2]
    }

    # Excel テーブルとして定義（ヘッダー込みの範囲）
    $lastDataRow = $dataStart + $samples.Length
    $tableRange = $sheet.Range("A${dataStart}:C${lastDataRow}")
    $table = $sheet.ListObjects.Add(1, $tableRange, [System.Type]::Missing, 1)
    $table.Name = "変数テーブル"
    $table.TableStyle = "TableStyleMedium2"

    # 列幅
    $table.ListColumns.Item(1).Range.ColumnWidth = 22
    $table.ListColumns.Item(2).Range.ColumnWidth = 18
    $table.ListColumns.Item(3).Range.ColumnWidth = 45

    # ============================================================
    # 「Word を更新」ボタン（行1 右側に配置）
    # ============================================================
    $btn = $sheet.Shapes.AddShape(1, 380, 3, 130, 22)
    $btn.Name = "btnUpdate"
    $btn.TextFrame.Characters().Text = "Word を更新"
    $btn.TextFrame.Characters().Font.Bold = $true
    $btn.TextFrame.Characters().Font.Size = 10
    $btn.TextFrame.Characters().Font.Color = [int]0xFFFFFF
    $btn.Fill.ForeColor.RGB = [int]0x4472C4
    $btn.Line.ForeColor.RGB = [int]0x2E5B9E
    $btn.Line.Weight = 1
    $btn.TextFrame.HorizontalAlignment = -4108   # xlCenter
    $btn.TextFrame.VerticalAlignment   = -4108   # xlCenter
    $btn.OnAction = "WordUpdater.UpdateWordDocument"

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

            $vbaModule = $workbook.VBProject.VBComponents.Add(1)   # vbext_ct_StdModule
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

    # ============================================================
    # 保存
    # ============================================================
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
Write-Host "  1. Word テンプレートに変数を設定"
Write-Host "     例: 会社名の箇所を `$company_name と書き、赤字にしておく"
Write-Host "  2. Word更新ツール.xlsm の B1 に Word ファイルのパスを入力"
Write-Host "  3. テーブルの A列に `$変数名、C列に変更後テキストを入力"
Write-Host "  4. [Word を更新] ボタンをクリック"
Write-Host ""
