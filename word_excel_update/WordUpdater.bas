Attribute VB_Name = "WordUpdater"
' ============================================================
' WordUpdater.bas
' Excel から Word の $変数 を一括置換するマクロ
'
' 【使い方】
'   1. Excelシート「変更箇所」のB1にWordファイルの絶対パスを入力
'   2. 4行目以降に $変数名 / 説明 / 変更後テキスト を入力
'   3. 「Word を更新」ボタンをクリック
'
' 【Wordテンプレートの準備】
'   変数箇所を $変数名（例: $company_name）にして赤字にしておく
'   → 置換後も文字色はそのまま維持されます
' ============================================================

Option Explicit

' ---- 定数 ----
Private Const SHEET_NAME   As String = "変更箇所"
Private Const PATH_CELL    As String = "B1"
Private Const DATA_START   As Long   = 4
Private Const COL_VARIABLE As Long   = 1   ' A列: $変数名
Private Const COL_DESC     As Long   = 2   ' B列: 説明（任意）
Private Const COL_NEW_TEXT As Long   = 3   ' C列: 変更後テキスト
Private Const COL_STATUS   As Long   = 4   ' D列: 状態
Private Const COLOR_DONE   As Long   = 32768  ' RGB(0,128,0) 緑
Private Const COLOR_BLACK  As Long   = 0      ' RGB(0,0,0)

' ============================================================
' メイン処理: $変数 を Excel の値で置換する
' ============================================================
Public Sub UpdateWordDocument()
    Dim ws           As Worksheet
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wordFilePath As String
    Dim lastRow      As Long
    Dim i            As Long
    Dim varName      As String
    Dim newText      As String
    Dim updatedCount As Long
    Dim notFoundList As String
    Dim cnt          As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_NAME & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    wordFilePath = Trim(ws.Range(PATH_CELL).Value)
    If wordFilePath = "" Then
        MsgBox "セル " & PATH_CELL & " に Word ファイルの絶対パスを入力してください。", vbExclamation
        Exit Sub
    End If
    If Dir(wordFilePath) = "" Then
        MsgBox "Word ファイルが見つかりません：" & vbNewLine & wordFilePath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler
    wdApp.Visible = True

    Set wdDoc = GetOrOpenDocument(wdApp, wordFilePath)
    If wdDoc Is Nothing Then
        MsgBox "Word ドキュメントを開けませんでした。", vbCritical
        Exit Sub
    End If

    lastRow      = ws.Cells(ws.Rows.Count, COL_VARIABLE).End(xlUp).Row
    updatedCount = 0
    notFoundList = ""

    For i = DATA_START To lastRow
        varName = Trim(ws.Cells(i, COL_VARIABLE).Value)
        newText = ws.Cells(i, COL_NEW_TEXT).Value

        ' $で始まる行のみ処理
        If varName = "" Then GoTo NextRow
        If Left(varName, 1) <> "$" Then GoTo NextRow

        cnt = CountInDocument(wdDoc, varName)
        If cnt > 0 Then
            Call ReplaceInDocument(wdDoc, varName, newText)
            updatedCount = updatedCount + 1
            With ws.Cells(i, COL_STATUS)
                .Value      = "済"
                .Font.Color = COLOR_DONE
                .Font.Bold  = True
            End With
        Else
            notFoundList = notFoundList & vbNewLine & "  ・行" & i & ": " & varName
        End If

NextRow:
    Next i

    wdDoc.Save

    Dim msg As String
    msg = "【更新完了】" & vbNewLine & vbNewLine
    msg = msg & "置換した変数: " & updatedCount & " 件" & vbNewLine
    If notFoundList <> "" Then
        msg = msg & vbNewLine & "Word に見つからなかった変数：" & notFoundList
    End If
    MsgBox msg, vbInformation, "Word 更新ツール"
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました：" & vbNewLine & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' ドキュメント内の varName の出現回数を返す
' Wrap=0(wdFindStop) で末尾で止まりループしない
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
        .Wrap           = 0         ' wdFindStop: 末尾で停止（ループ防止）
        .MatchCase      = True
        .MatchWildcards = False
        Do While .Execute
            count = count + 1
        Loop
    End With
    CountInDocument = count
End Function

' ============================================================
' ドキュメント内の varName を newText で置換（文字色はそのまま維持）
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
        .Format           = False   ' 書式は変更しない（赤字のまま維持）
        .MatchCase        = True
        .MatchWildcards   = False
        .Execute Replace:=2         ' wdReplaceAll
    End With
End Sub

' ============================================================
' Word の $変数 をスキャンして Excel に一覧取り込み
' ============================================================
Public Sub ImportVariableList()
    Dim ws           As Worksheet
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wordFilePath As String
    Dim rng          As Object
    Dim varList()    As String
    Dim varCount     As Long
    Dim i            As Long
    Dim matchText    As String
    Dim isDup        As Boolean

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

    ' ワイルドカードで $xxx パターンをスキャン（重複除去）
    ' Wrap=0(wdFindStop) で末尾で止まりループしない
    ReDim varList(0)
    varCount = 0

    Set rng = wdDoc.Range
    With rng.Find
        .ClearFormatting
        .Text           = "\$[A-Za-z_][A-Za-z0-9_]*"
        .Forward        = True
        .Wrap           = 0         ' wdFindStop: 末尾で停止（ループ防止）
        .MatchCase      = False
        .MatchWildcards = True
        Do While .Execute
            matchText = rng.Text
            isDup = False
            For i = 0 To varCount - 1
                If varList(i) = matchText Then
                    isDup = True
                    Exit For
                End If
            Next i
            If Not isDup Then
                ReDim Preserve varList(varCount)
                varList(varCount) = matchText
                varCount = varCount + 1
            End If
        Loop
    End With

    If varCount = 0 Then
        MsgBox "Word ドキュメントに $変数 が見つかりませんでした。" & vbNewLine & _
               "変数は $variable_name の形式で赤字で記述してください。", vbInformation
        Exit Sub
    End If

    Dim lastRow  As Long
    Dim startRow As Long
    lastRow  = ws.Cells(ws.Rows.Count, COL_VARIABLE).End(xlUp).Row
    startRow = IIf(lastRow < DATA_START, DATA_START, lastRow + 1)

    For i = 0 To varCount - 1
        ws.Cells(startRow + i, COL_VARIABLE).Value = varList(i)
    Next i

    MsgBox varCount & " 件の変数を取り込みました。" & vbNewLine & _
           "C列（変更後テキスト）を入力してから「Word を更新」を押してください。", vbInformation
End Sub

' ============================================================
' 「済」ステータスをリセット
' ============================================================
Public Sub ResetStatus()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    If MsgBox("「済」状態をリセットしますか？", vbQuestion + vbYesNo, "状態リセット") = vbNo Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, COL_VARIABLE).End(xlUp).Row
    For i = DATA_START To lastRow
        If ws.Cells(i, COL_STATUS).Value = "済" Then
            With ws.Cells(i, COL_STATUS)
                .Value      = ""
                .Font.Color = COLOR_BLACK
                .Font.Bold  = False
            End With
        End If
    Next i

    MsgBox "リセットしました。", vbInformation
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
