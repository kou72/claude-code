Attribute VB_Name = "WordBookmarkHelper"
' ============================================================
' WordBookmarkHelper.bas
' Word 側のブックマーク管理ヘルパーマクロ
'
' 【導入方法】
'   Word の VBA エディタ（Alt+F11）を開き、
'   「ファイル」→「ファイルのインポート」でこのファイルを取り込む
' ============================================================

Option Explicit

' ============================================================
' 選択範囲にブックマークを設定する
' ============================================================
Public Sub AddBookmark()
    Dim bmName As String

    If Selection.Type = wdSelectionIP Then
        MsgBox "テキストを選択してからマクロを実行してください。", vbExclamation
        Exit Sub
    End If

    bmName = InputBox( _
        "ブックマーク名を入力してください。" & vbNewLine & vbNewLine & _
        "【命名規則の推奨】" & vbNewLine & _
        "  bm_会社名       → bm_company_name" & vbNewLine & _
        "  bm_日付         → bm_date" & vbNewLine & _
        "  bm_金額         → bm_amount" & vbNewLine & vbNewLine & _
        "※ 半角英数字とアンダースコアのみ使用可（スペース不可）", _
        "ブックマーク名の入力")

    ' キャンセル or 空欄
    If bmName = "" Then Exit Sub

    ' 使用不可文字チェック
    If Not IsValidBookmarkName(bmName) Then
        MsgBox "ブックマーク名に使用できない文字が含まれています。" & vbNewLine & _
               "半角英数字とアンダースコア（_）のみ使用してください。", vbExclamation
        Exit Sub
    End If

    ' 重複チェック
    If ActiveDocument.Bookmarks.Exists(bmName) Then
        If MsgBox("「" & bmName & "」は既に存在します。上書きしますか？", _
                  vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    ' ブックマーク追加
    ActiveDocument.Bookmarks.Add Name:=bmName, Range:=Selection.Range

    MsgBox "ブックマーク「" & bmName & "」を設定しました。" & vbNewLine & vbNewLine & _
           "Excel の変更箇所シートの A 列に「" & bmName & "」を入力してください。", _
           vbInformation

End Sub

' ============================================================
' ドキュメント内の全ブックマーク一覧を表示する
' ============================================================
Public Sub ListAllBookmarks()
    Dim doc   As Document
    Dim bm    As Bookmark
    Dim msg   As String
    Dim count As Long

    Set doc = ActiveDocument
    count = doc.Bookmarks.Count

    If count = 0 Then
        MsgBox "このドキュメントにブックマークはありません。", vbInformation
        Exit Sub
    End If

    msg = "【ブックマーク一覧】(" & count & " 件)" & vbNewLine & vbNewLine
    msg = msg & String(40, "-") & vbNewLine

    For Each bm In doc.Bookmarks
        Dim previewText As String
        previewText = bm.Range.Text
        If Len(previewText) > 30 Then previewText = Left(previewText, 30) & "..."
        msg = msg & "■ " & bm.Name & vbNewLine
        msg = msg & "  現在値: " & previewText & vbNewLine & vbNewLine
    Next bm

    MsgBox msg, vbInformation, "ブックマーク一覧"
End Sub

' ============================================================
' 選択したブックマークを削除する
' ============================================================
Public Sub DeleteBookmark()
    Dim bmName As String
    Dim doc    As Document

    Set doc = ActiveDocument

    If doc.Bookmarks.Count = 0 Then
        MsgBox "削除できるブックマークがありません。", vbInformation
        Exit Sub
    End If

    bmName = InputBox("削除するブックマーク名を入力してください：", "ブックマーク削除")
    If bmName = "" Then Exit Sub

    If Not doc.Bookmarks.Exists(bmName) Then
        MsgBox "「" & bmName & "」というブックマークは存在しません。", vbExclamation
        Exit Sub
    End If

    If MsgBox("「" & bmName & "」を削除しますか？（テキスト自体は残ります）", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    doc.Bookmarks(bmName).Delete
    MsgBox "「" & bmName & "」を削除しました。", vbInformation
End Sub

' ============================================================
' ブックマーク箇所の赤字をすべて黒字に戻す（最終納品前用）
' ============================================================
Public Sub ResetAllRedToBlack()
    Dim doc   As Document
    Dim bm    As Bookmark
    Dim count As Long

    Set doc = ActiveDocument

    If doc.Bookmarks.Count = 0 Then
        MsgBox "ブックマークがありません。", vbInformation
        Exit Sub
    End If

    If MsgBox("ブックマーク箇所の赤字をすべて黒字に戻しますか？" & vbNewLine & _
              "（納品前の最終処理として使用）", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    count = 0
    For Each bm In doc.Bookmarks
        Dim bmRange As Range
        Set bmRange = bm.Range
        If bmRange.Font.Color = RGB(255, 0, 0) Then
            bmRange.Font.Color = RGB(0, 0, 0)
            count = count + 1
        End If
    Next bm

    doc.Save
    MsgBox count & " 箇所の赤字を黒字に変換して保存しました。", vbInformation
End Sub

' ============================================================
' 内部関数: ブックマーク名の文字種チェック
' ============================================================
Private Function IsValidBookmarkName(name As String) As Boolean
    Dim i    As Long
    Dim c    As String
    Dim code As Long

    If Len(name) = 0 Then
        IsValidBookmarkName = False
        Exit Function
    End If

    For i = 1 To Len(name)
        c = Mid(name, i, 1)
        code = Asc(c)
        ' 半角英字 (A-Z, a-z)、数字 (0-9)、アンダースコア (_) のみ許可
        ' 先頭が数字の場合も Word は許容するが念のため英字推奨
        If Not ((code >= 65 And code <= 90) Or _
                (code >= 97 And code <= 122) Or _
                (code >= 48 And code <= 57) Or _
                code = 95) Then
            IsValidBookmarkName = False
            Exit Function
        End If
    Next i

    IsValidBookmarkName = True
End Function
