Attribute VB_Name = "WordUpdater"
' ============================================================
' WordUpdater.bas
' Excel から Word の指定箇所を赤字で更新するマクロ
' ============================================================
' 【使い方】
'   1. Excelシート「変更箇所」のB1にWordファイルの絶対パスを入力
'   2. 4行目以降に ブックマーク名 / 説明 / 変更後テキスト を入力
'   3. 「Word を更新」ボタンをクリック
' ============================================================

Option Explicit

' ---- 定数 ----
Private Const SHEET_NAME    As String = "変更箇所"
Private Const PATH_CELL     As String = "B1"
Private Const DATA_START    As Long   = 4   ' データ開始行
Private Const COL_BOOKMARK  As Long   = 1   ' A列: ブックマーク名
Private Const COL_DESC      As Long   = 2   ' B列: 説明（任意）
Private Const COL_NEW_TEXT  As Long   = 3   ' C列: 変更後テキスト
Private Const COL_STATUS    As Long   = 4   ' D列: 状態（済 / 空欄）
Private Const COLOR_RED     As Long   = 255
Private Const COLOR_DONE    As Long   = 32768
Private Const COLOR_BLACK   As Long   = 0

' ============================================================
' メイン処理: Wordドキュメントを更新する
' ============================================================
Public Sub UpdateWordDocument()
    Dim ws           As Worksheet
    Dim wdApp        As Object   ' Word.Application
    Dim wdDoc        As Object   ' Word.Document
    Dim wordFilePath As String
    Dim lastRow      As Long
    Dim i            As Long
    Dim bookmarkName As String
    Dim newText      As String
    Dim updatedCount As Long
    Dim skippedCount As Long
    Dim notFoundList As String
    Dim bmRange      As Object

    ' ---- シート取得 ----
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' ---- Wordファイルパス取得 ----
    wordFilePath = Trim(ws.Range(PATH_CELL).Value)
    If wordFilePath = "" Then
        MsgBox "セル " & PATH_CELL & " に Word ファイルの絶対パスを入力してください。", vbExclamation
        Exit Sub
    End If
    If Dir(wordFilePath) = "" Then
        MsgBox "Word ファイルが見つかりません：" & vbNewLine & wordFilePath, vbExclamation
        Exit Sub
    End If

    ' ---- Word アプリ取得 / 起動 ----
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler
    wdApp.Visible = True

    ' ---- ドキュメントを開く（既に開いていれば再利用）----
    Set wdDoc = GetOrOpenDocument(wdApp, wordFilePath)
    If wdDoc Is Nothing Then
        MsgBox "Word ドキュメントを開けませんでした。", vbCritical
        Exit Sub
    End If

    ' ---- データ処理 ----
    lastRow      = ws.Cells(ws.Rows.Count, COL_BOOKMARK).End(xlUp).Row
    updatedCount = 0
    skippedCount = 0
    notFoundList = ""

    Dim confirmSkip As Boolean
    confirmSkip = False

    For i = DATA_START To lastRow
        bookmarkName = Trim(ws.Cells(i, COL_BOOKMARK).Value)
        newText      = ws.Cells(i, COL_NEW_TEXT).Value

        ' ブックマーク名が空の行はスキップ
        If bookmarkName = "" Then GoTo NextRow

        ' 既に適用済みの行はスキップ
        If ws.Cells(i, COL_STATUS).Value = "済" Then
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If

        ' ---- ブックマーク存在確認 ----
        If wdDoc.Bookmarks.Exists(bookmarkName) Then
            Set bmRange = wdDoc.Bookmarks(bookmarkName).Range

            ' テキストを置換
            bmRange.Text = newText

            ' 赤字に設定
            bmRange.Font.Color = COLOR_RED

            ' ブックマークを再設定（テキスト変更後に失われるため）
            wdDoc.Bookmarks.Add bookmarkName, bmRange

            ' Excel側のステータスを更新
            updatedCount = updatedCount + 1
            With ws.Cells(i, COL_STATUS)
                .Value      = "済"
                .Font.Color = COLOR_DONE
                .Font.Bold  = True
            End With
        Else
            notFoundList = notFoundList & vbNewLine & "  ・行" & i & ": " & bookmarkName
        End If

NextRow:
    Next i

    ' ---- Word ドキュメントを保存 ----
    wdDoc.Save

    ' ---- 結果メッセージ ----
    Dim msg As String
    msg = "【更新完了】" & vbNewLine & vbNewLine
    msg = msg & "更新した箇所: " & updatedCount & " 件" & vbNewLine
    If skippedCount > 0 Then
        msg = msg & "適用済みでスキップ: " & skippedCount & " 件" & vbNewLine
    End If
    If notFoundList <> "" Then
        msg = msg & vbNewLine & "! ブックマークが見つかりませんでした：" & notFoundList & vbNewLine
        msg = msg & vbNewLine & "→ Word ドキュメントにブックマークを設定してください。"
    End If

    MsgBox msg, vbInformation, "Word 更新ツール"
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました：" & vbNewLine & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' 補助: ドキュメントを取得または開く
' ============================================================
Private Function GetOrOpenDocument(wdApp As Object, filePath As String) As Object
    Dim wdDocItem As Object
    For Each wdDocItem In wdApp.Documents
        If LCase(wdDocItem.FullName) = LCase(filePath) Then
            Set GetOrOpenDocument = wdDocItem
            Exit Function
        End If
    Next wdDocItem
    On Error Resume Next
    Set GetOrOpenDocument = wdApp.Documents.Open(filePath)
    On Error GoTo 0
End Function

' ============================================================
' 「済」ステータスをリセットする
' ============================================================
Public Sub ResetStatus()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    If MsgBox("「済」状態をリセットして、全行を再適用対象にしますか？", _
              vbQuestion + vbYesNo, "状態リセット") = vbNo Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, COL_BOOKMARK).End(xlUp).Row
    For i = DATA_START To lastRow
        If ws.Cells(i, COL_STATUS).Value = "済" Then
            With ws.Cells(i, COL_STATUS)
                .Value      = ""
                .Font.Color = COLOR_BLACK
                .Font.Bold  = False
            End With
        End If
    Next i

    MsgBox "リセットしました。次回の更新で全行が再適用されます。", vbInformation
End Sub

' ============================================================
' Wordの全ブックマーク一覧をExcelに取り込む（確認用）
' ============================================================
Public Sub ImportBookmarkList()
    Dim ws           As Worksheet
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wordFilePath As String
    Dim bm           As Object
    Dim i            As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    wordFilePath = Trim(ws.Range(PATH_CELL).Value)

    If wordFilePath = "" Or Dir(wordFilePath) = "" Then
        MsgBox "有効な Word ファイルパスを " & PATH_CELL & " に入力してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    wdApp.Visible = True

    Set wdDoc = GetOrOpenDocument(wdApp, wordFilePath)
    If wdDoc Is Nothing Then Exit Sub

    If wdDoc.Bookmarks.Count = 0 Then
        MsgBox "Word ドキュメントにブックマークが設定されていません。", vbInformation
        Exit Sub
    End If

    ' 既存データの下に追記するか確認
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_BOOKMARK).End(xlUp).Row
    Dim startRow As Long
    startRow = IIf(lastRow < DATA_START, DATA_START, lastRow + 1)

    i = startRow
    For Each bm In wdDoc.Bookmarks
        ws.Cells(i, COL_BOOKMARK).Value  = bm.Name
        ws.Cells(i, COL_DESC).Value      = "（" & Left(bm.Range.Text, 20) & "…）"
        ws.Cells(i, COL_NEW_TEXT).Value  = bm.Range.Text  ' 現在のテキストをデフォルト値として挿入
        ws.Cells(i, COL_STATUS).Value    = ""
        i = i + 1
    Next bm

    MsgBox wdDoc.Bookmarks.Count & " 件のブックマークを取り込みました。" & vbNewLine & _
           "C列（変更後テキスト）を編集してから「Word を更新」を押してください。", vbInformation
End Sub
