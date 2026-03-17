Attribute VB_Name = "WordUpdater"
' ============================================================
' WordUpdater.bas
' 実行テーブルで yes のシートを一括処理して Word 納品ドキュメントを生成する
' ============================================================
Option Explicit

Private Const EXEC_SHEET_NAME As String = "実行"
Private Const EXEC_TABLE_NAME As String = "実行テーブル"
Private Const EXEC_COL_SHEET  As Long = 1    ' 実行テーブル列1: シート名
Private Const EXEC_COL_FLAG   As Long = 2    ' 実行テーブル列2: 実行フラグ

Private Const TEMPLATE_CELL   As String = "C1"   ' 各変更箇所シート: テンプレートパス
Private Const OUTPUT_CELL     As String = "C2"   ' 各変更箇所シート: 出力パス
Private Const TABLE_NAME      As String = "変数テーブル"
Private Const COL_VAR         As Long = 1    ' 変数テーブル列1: $変数名
Private Const COL_NEW         As Long = 3    ' 変数テーブル列3: 変更後テキスト

' ============================================================
' エントリポイント: 実行テーブルで yes のシートを一括処理
' ============================================================
Public Sub RunAll()
    Dim execWs      As Worksheet
    Dim execTbl     As ListObject
    Dim wdApp       As Object
    Dim i           As Long
    Dim sName       As String
    Dim sFlag       As String
    Dim targets()   As String
    Dim targetCount As Long
    Dim ws          As Worksheet
    Dim outPath     As String
    Dim updCnt      As Long
    Dim notFound    As String
    Dim errMsg      As String
    Dim confirmMsg  As String
    Dim summaryLines As String
    Dim successCount As Long
    Dim resultMsg   As String

    ' 実行シート取得
    On Error Resume Next
    Set execWs = ThisWorkbook.Sheets(EXEC_SHEET_NAME)
    On Error GoTo 0
    If execWs Is Nothing Then
        MsgBox "シート「" & EXEC_SHEET_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 実行テーブル取得
    On Error Resume Next
    Set execTbl = execWs.ListObjects(EXEC_TABLE_NAME)
    On Error GoTo 0
    If execTbl Is Nothing Then
        MsgBox "テーブル「" & EXEC_TABLE_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If
    If execTbl.ListRows.Count = 0 Then
        MsgBox "実行テーブルにデータがありません。", vbExclamation
        Exit Sub
    End If

    ' flag="yes" の対象シートを収集
    targetCount = 0
    ReDim targets(execTbl.ListRows.Count - 1)
    For i = 1 To execTbl.ListRows.Count
        sName = Trim(execTbl.ListRows(i).Range.Cells(1, EXEC_COL_SHEET).Value)
        sFlag = LCase(Trim(execTbl.ListRows(i).Range.Cells(1, EXEC_COL_FLAG).Value))
        If sName <> "" And sFlag = "yes" Then
            targets(targetCount) = sName
            targetCount = targetCount + 1
        End If
    Next i
    If targetCount = 0 Then
        MsgBox "実行フラグが「yes」のシートがありません。", vbExclamation
        Exit Sub
    End If

    ' 実行確認
    confirmMsg = "以下の " & targetCount & " シートを処理します：" & vbNewLine & vbNewLine
    For i = 0 To targetCount - 1
        confirmMsg = confirmMsg & "  ・" & targets(i) & vbNewLine
    Next i
    confirmMsg = confirmMsg & vbNewLine & "実行しますか？"
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "一括実行確認") = vbNo Then Exit Sub

    ' Word 起動 / 接続
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrHandler
    wdApp.Visible = True

    ' 各シートを処理
    successCount = 0
    summaryLines = ""
    For i = 0 To targetCount - 1
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(targets(i))
        On Error GoTo ErrHandler
        If ws Is Nothing Then
            summaryLines = summaryLines & vbNewLine & "■ " & targets(i) & " [エラー]" & vbNewLine
            summaryLines = summaryLines & "  シートが見つかりません"
        Else
            outPath = ""
            updCnt = 0
            notFound = ""
            errMsg = ProcessSheet(ws, wdApp, outPath, updCnt, notFound)
            If errMsg = "" Then
                successCount = successCount + 1
                summaryLines = summaryLines & vbNewLine & "■ " & targets(i) & vbNewLine
                summaryLines = summaryLines & "  出力: " & outPath & "（" & updCnt & " 変数置換）"
                If notFound <> "" Then
                    summaryLines = summaryLines & vbNewLine & "  ※ 見つからなかった変数:" & notFound
                End If
            Else
                summaryLines = summaryLines & vbNewLine & "■ " & targets(i) & " [エラー]" & vbNewLine
                summaryLines = summaryLines & "  " & errMsg
            End If
        End If
    Next i

    ' 結果表示
    resultMsg = "【完了】一括処理が終わりました。" & vbNewLine & vbNewLine
    resultMsg = resultMsg & "成功: " & successCount & " 件 / 対象: " & targetCount & " 件"
    If summaryLines <> "" Then
        resultMsg = resultMsg & vbNewLine & summaryLines
    End If
    MsgBox resultMsg, vbInformation, "Word 更新ツール"
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました：" & vbNewLine & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' 補助: 1シート分の Word 更新処理
'   戻り値: "" = 成功 / それ以外 = エラーメッセージ
' ============================================================
Private Function ProcessSheet(ws As Worksheet, wdApp As Object, _
                               ByRef outPath As String, _
                               ByRef updatedCount As Long, _
                               ByRef notFoundList As String) As String
    Dim tbl          As ListObject
    Dim wdDoc        As Object
    Dim templatePath As String
    Dim outputDir    As String
    Dim i            As Long
    Dim varName      As String
    Dim newText      As String

    ProcessSheet = ""
    outPath = ""
    updatedCount = 0
    notFoundList = ""

    ' テンプレートパス確認
    templatePath = Trim(ws.Range(TEMPLATE_CELL).Value)
    If templatePath = "" Then
        ProcessSheet = "C1 にテンプレートパスが入力されていません"
        Exit Function
    End If
    If Dir(templatePath) = "" Then
        ProcessSheet = "テンプレートが見つかりません: " & templatePath
        Exit Function
    End If

    ' 出力パス確認
    outPath = Trim(ws.Range(OUTPUT_CELL).Value)
    If outPath = "" Then
        ProcessSheet = "C2 に出力パスが入力されていません"
        Exit Function
    End If

    ' 出力先ディレクトリが存在しない場合は自動作成
    outputDir = Left(outPath, InStrRev(outPath, "\"))
    If outputDir <> "" And Dir(outputDir, vbDirectory) = "" Then
        Call CreateFolderRecursive(outputDir)
    End If

    ' テーブル取得
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0
    If tbl Is Nothing Then
        ProcessSheet = "テーブル「" & TABLE_NAME & "」が見つかりません"
        Exit Function
    End If
    If tbl.ListRows.Count = 0 Then
        ProcessSheet = "テーブルにデータがありません"
        Exit Function
    End If

    ' テンプレートを出力パスへコピー
    On Error GoTo ErrExit
    FileCopy templatePath, outPath

    ' コピー先を Word で開く
    Set wdDoc = wdApp.Documents.Open(outPath)

    ' 置換ループ
    For i = 1 To tbl.ListRows.Count
        varName = Trim(tbl.ListRows(i).Range.Cells(1, COL_VAR).Value)
        newText = tbl.ListRows(i).Range.Cells(1, COL_NEW).Value
        If varName = "" Or Left(varName, 1) <> "$" Then GoTo NextRow
        If CountInDocument(wdDoc, varName) > 0 Then
            Call ReplaceInDocument(wdDoc, varName, newText)
            updatedCount = updatedCount + 1
        Else
            notFoundList = notFoundList & " " & varName
        End If
NextRow:
    Next i
    wdDoc.Save
    Exit Function

ErrExit:
    ProcessSheet = Err.Description
    If Not wdDoc Is Nothing Then
        On Error Resume Next
        wdDoc.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Function

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
        .Text = varName
        .Forward = True
        .Wrap = 0               ' wdFindStop: 末尾で停止（ループ防止）
        .MatchCase = True
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
        .Text = varName
        .Replacement.Text = newText
        .Forward = True
        .Wrap = 0                   ' wdFindStop: 末尾で停止（ループ防止）
        .Format = False             ' 書式を変更しない（赤字のまま維持）
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=2         ' wdReplaceAll
    End With
End Sub

' ============================================================
' 補助: フォルダを再帰的に作成（存在しない親フォルダも含めて作成）
' ============================================================
Private Sub CreateFolderRecursive(folderPath As String)
    If folderPath = "" Then Exit Sub
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub
    Dim parentPath As String
    parentPath = Left(folderPath, InStrRev(folderPath, "\") - 1)
    If parentPath <> "" And parentPath <> Left(folderPath, 2) Then
        Call CreateFolderRecursive(parentPath)
    End If
    MkDir folderPath
End Sub
