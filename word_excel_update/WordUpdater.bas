Attribute VB_Name = "WordUpdater"
' ============================================================
' WordUpdater.bas
' Excel の変数テーブルを使って Word の $変数 を一括置換するマクロ
'
' 【Wordテンプレートの準備】
'   置換したい箇所を $変数名（例: $company_name）にして赤字にしておく
'   → 置換後も文字色はそのまま維持されます
'
' 【Excelの使い方】
'   1. B1 に Word ファイルの絶対パスを入力
'   2. テーブルに $変数名 と 変更後テキスト を入力
'   3. 「Word を更新」ボタンをクリック
' ============================================================

Option Explicit

Private Const SHEET_NAME As String = "変更箇所"
Private Const PATH_CELL  As String = "B1"
Private Const TABLE_NAME As String = "変数テーブル"
Private Const COL_VAR    As Long   = 1   ' テーブル列1: $変数名
Private Const COL_NEW    As Long   = 3   ' テーブル列3: 変更後テキスト

' ============================================================
' メイン処理: $変数 を Excel テーブルの値で置換する
' ============================================================
Public Sub UpdateWordDocument()
    Dim ws           As Worksheet
    Dim tbl          As ListObject
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wordFilePath As String
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

    ' Word ファイルパス確認
    wordFilePath = Trim(ws.Range(PATH_CELL).Value)
    If wordFilePath = "" Then
        MsgBox "セル " & PATH_CELL & " に Word ファイルの絶対パスを入力してください。", vbExclamation
        Exit Sub
    End If
    If Dir(wordFilePath) = "" Then
        MsgBox "Word ファイルが見つかりません：" & vbNewLine & wordFilePath, vbExclamation
        Exit Sub
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

    ' Word 起動 / 接続
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrHandler
    wdApp.Visible = True

    ' ドキュメントを開く
    Set wdDoc = GetOrOpenDocument(wdApp, wordFilePath)
    If wdDoc Is Nothing Then
        MsgBox "Word ドキュメントを開けませんでした。", vbCritical
        Exit Sub
    End If

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
    msg = "【更新完了】" & vbNewLine & vbNewLine & "置換した変数: " & updatedCount & " 件"
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

' ============================================================
' 補助: ドキュメントを取得または開く
' ============================================================
Private Function GetOrOpenDocument(wdApp As Object, filePath As String) As Object
    Dim doc As Object
    For Each doc In wdApp.Documents
        If LCase(doc.FullName) = LCase(filePath) Then
            Set GetOrOpenDocument = doc
            Exit Function
        End If
    Next doc
    On Error Resume Next
    Set GetOrOpenDocument = wdApp.Documents.Open(filePath)
    On Error GoTo 0
End Function
