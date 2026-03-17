Attribute VB_Name = "WordUpdater"
' ============================================================
' WordUpdater.bas
' テンプレート Word をコピーし、$変数 を一括置換して納品ドキュメントを作成するマクロ
' ============================================================

Option Explicit

Private Const SHEET_NAME    As String = "変更箇所"
Private Const TEMPLATE_CELL As String = "B1"   ' テンプレートファイルパス
Private Const OUTPUT_CELL   As String = "B2"   ' 出力ファイルパス
Private Const TABLE_NAME    As String = "変数テーブル"
Private Const COL_VAR       As Long   = 1      ' テーブル列1: $変数名
Private Const COL_NEW       As Long   = 3      ' テーブル列3: 変更後テキスト

' ============================================================
' メイン処理
'   1. テンプレートを出力パスへコピー
'   2. コピー先を開いて $変数 を一括置換
'   3. 保存（テンプレートは変更しない）
' ============================================================
Public Sub UpdateWordDocument()
    Dim ws           As Worksheet
    Dim tbl          As ListObject
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim templatePath As String
    Dim outputPath   As String
    Dim i            As Long
    Dim varName      As String
    Dim newText      As String
    Dim updatedCount As Long
    Dim notFoundList As String

    ' シート取得
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' テンプレートパス確認
    templatePath = Trim(ws.Range(TEMPLATE_CELL).Value)
    If templatePath = "" Then
        MsgBox "B1 にテンプレートファイルの絶対パスを入力してください。", vbExclamation
        Exit Sub
    End If
    If Dir(templatePath) = "" Then
        MsgBox "テンプレートファイルが見つかりません：" & vbNewLine & templatePath, vbExclamation
        Exit Sub
    End If

    ' 出力パス確認
    outputPath = Trim(ws.Range(OUTPUT_CELL).Value)
    If outputPath = "" Then
        MsgBox "B2 に出力ファイルの絶対パスを入力してください。", vbExclamation
        Exit Sub
    End If

    ' 出力先ディレクトリの存在確認
    Dim outputDir As String
    outputDir = Left(outputPath, InStrRev(outputPath, ""))
    If outputDir <> "" And Dir(outputDir, vbDirectory) = "" Then
        MsgBox "出力先のフォルダが存在しません：" & vbNewLine & outputDir, vbExclamation
        Exit Sub
    End If

    ' 上書き確認
    If Dir(outputPath) <> "" Then
        If MsgBox("出力ファイルが既に存在します。上書きしますか？" & vbNewLine & outputPath, _
                  vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    ' テーブル取得
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "テーブル「" & TABLE_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If
    If tbl.ListRows.Count = 0 Then
        MsgBox "テーブルにデータがありません。", vbExclamation
        Exit Sub
    End If

    ' テンプレートを出力パスへコピー
    On Error GoTo ErrHandler
    FileCopy templatePath, outputPath

    ' Word 起動 / 接続
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrHandler
    wdApp.Visible = True

    ' コピー先を開く
    Set wdDoc = wdApp.Documents.Open(outputPath)

    ' テーブル行をループして置換
    updatedCount = 0
    notFoundList = ""

    For i = 1 To tbl.ListRows.Count
        varName = Trim(tbl.ListRows(i).Range.Cells(1, COL_VAR).Value)
        newText = tbl.ListRows(i).Range.Cells(1, COL_NEW).Value

        If varName = "" Or Left(varName, 1) <> "$" Then GoTo NextRow

        If CountInDocument(wdDoc, varName) > 0 Then
            Call ReplaceInDocument(wdDoc, varName, newText)
            updatedCount = updatedCount + 1
        Else
            notFoundList = notFoundList & vbNewLine & "  ・" & varName
        End If
NextRow:
    Next i

    wdDoc.Save

    ' 結果メッセージ
    Dim msg As String
    msg = "【完了】納品ドキュメントを作成しました。" & vbNewLine & vbNewLine
    msg = msg & "出力先: " & outputPath & vbNewLine & vbNewLine
    msg = msg & "置換した変数: " & updatedCount & " 件"
    If notFoundList <> "" Then
        msg = msg & vbNewLine & vbNewLine & "Word に見つからなかった変数：" & notFoundList
    End If
    MsgBox msg, vbInformation, "Word 更新ツール"
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました：" & vbNewLine & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' 補助: 文書内の varName の出現回数を返す
' ============================================================
Private Function CountInDocument(wdDoc As Object, varName As String) As Long
    Dim rng   As Object
    Dim count As Long
    count = 0
    Set rng = wdDoc.Range
    With rng.Find
        .ClearFormatting
        .Text           = varName
        .Forward        = True
        .Wrap           = 0     ' wdFindStop: 末尾で停止（ループ防止）
        .MatchCase      = True
        .MatchWildcards = False
        Do While .Execute
            count = count + 1
        Loop
    End With
    CountInDocument = count
End Function

' ============================================================
' 補助: 文書内の varName を newText で置換（書式は維持）
' ============================================================
Private Sub ReplaceInDocument(wdDoc As Object, varName As String, newText As String)
    Dim rng As Object
    Set rng = wdDoc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text             = varName
        .Replacement.Text = newText
        .Forward          = True
        .Wrap             = 0       ' wdFindStop: 末尾で停止（ループ防止）
        .Format           = False   ' 書式を変更しない（赤字のまま維持）
        .MatchCase        = True
        .MatchWildcards   = False
        .Execute Replace:=2         ' wdReplaceAll
    End With
End Sub
