Attribute VB_Name = "ForMiyo"
Option Explicit

'==================================================
' JLPT N3 語彙復習プリント 自動作成マクロ
'
' 問題バンクから指定した2課の問題を抽出し、
' ランダムに並び替えて問題プリント・解答を作成する。
'==================================================

Sub ForMiyo()

'----------------------------------------------
' 使用するシート
'----------------------------------------------
Dim wsTop As Worksheet
Dim wsBank As Worksheet
Dim wsPrint As Worksheet
Dim wsAnswer As Worksheet

'----------------------------------------------
' ユーザー入力
'----------------------------------------------
Dim lesson1 As String
Dim lesson2 As String
Dim questionCount As Long

'----------------------------------------------
' 問題抽出用
'----------------------------------------------
Dim lastRow As Long
Dim searchRow As Long

Dim selectedRows() As Long
Dim selectedCount As Long

'----------------------------------------------
' 出力用
'----------------------------------------------
Dim printRow As Long
Dim answerRow As Long

Dim questionNo As Long
Dim sourceRow As Long

'==================================================
' 1. 対象課と問題数を入力
'==================================================

Set wsTop = ThisWorkbook.Worksheets("TOP")

lesson1 = wsTop.Range("C3").Value
If lesson1 = "" Then
    MsgBox "第1課を入力してください。", vbExclamation, "エラー！！"
    Exit Sub
End If

lesson2 = wsTop.Range("C4").Value
If lesson2 = "" Then
    MsgBox "第2課を入力してください。", vbExclamation, "エラー！！"
    Exit Sub
End If

questionCount = wsTop.Range("C5").Value
If questionCount <= 0 Then
    MsgBox "問題数を入力してください。", vbExclamation, "エラー！！"
    Exit Sub
End If

'==================================================
' 2. 使用するシートを取得
'==================================================

Set wsBank = ThisWorkbook.Worksheets("問題バンク")

Set wsPrint = GetOrCreateSheet("問題プリント")
Set wsAnswer = GetOrCreateSheet("解答")

' 前回の出力内容を削除
wsPrint.Cells.Clear
wsAnswer.Cells.Clear

'==================================================
' 3. 問題バンクから指定された課を抽出
'==================================================

lastRow = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row

selectedCount = 0

For searchRow = 2 To lastRow

    If wsBank.Cells(searchRow, 1).Value = lesson1 _
    Or wsBank.Cells(searchRow, 1).Value = lesson2 Then

        selectedCount = selectedCount + 1

        ReDim Preserve selectedRows(1 To selectedCount)

        selectedRows(selectedCount) = searchRow

    End If

Next searchRow

' 対象問題が1件も存在しない場合
If selectedCount = 0 Then

    MsgBox _
        "第" & lesson1 & "課・第" & lesson2 & _
        "課の問題がありません。", _
        vbExclamation, _
        "問題なし"

    Exit Sub

End If

'==================================================
' 4. 問題をランダムに並び替える
'==================================================

Randomize

ShuffleArray selectedRows

' 指定問題数より問題バンクの件数が少ない場合
If questionCount > selectedCount Then
    questionCount = selectedCount
End If

'==================================================
' 5. プリントのヘッダーを作成
'==================================================

CreatePrintHeader _
    wsPrint, _
    "JLPT N3 語彙復習プリント", _
    "第" & lesson1 & "課＋第" & lesson2 & "課"

CreatePrintHeader _
    wsAnswer, _
    "JLPT N3 語彙復習プリント　解答", _
    "第" & lesson1 & "課＋第" & lesson2 & "課"

'==================================================
' 6. 問題・解答を書き込む
'==================================================

printRow = 5
answerRow = 5

For questionNo = 1 To questionCount

    ' ランダムに選択された問題バンクの行番号
    sourceRow = selectedRows(questionNo)

    '----------------------------------------------
    ' 問題
    '----------------------------------------------

    wsPrint.Cells(printRow, 1).Value = _
        "【" & questionNo & "】" & _
        wsBank.Cells(sourceRow, 4).Value

    wsPrint.Cells(printRow, 1).WrapText = True
    wsPrint.Cells(printRow, 1).Font.Size = 11

    wsPrint.Cells(printRow + 1, 1).Value = _
        wsBank.Cells(sourceRow, 5).Value

    wsPrint.Cells(printRow + 1, 1).WrapText = True

    '----------------------------------------------
    ' 解答
    '----------------------------------------------

    wsAnswer.Cells(answerRow, 1).Value = _
        questionNo & ". " & _
        wsBank.Cells(sourceRow, 6).Value

    ' 解説
    wsAnswer.Cells(answerRow + 1, 1).Value = _
        wsBank.Cells(sourceRow, 7).Value

    ' 次の問題の出力位置へ移動
    printRow = printRow + 4
    answerRow = answerRow + 3

Next questionNo

'==================================================
' 7. 書式設定
'==================================================

FormatPrint wsPrint
FormatAnswer wsAnswer

'==================================================
' 8. 完了通知
'==================================================

MsgBox _
    questionCount & "問の語彙復習プリントを作成しました！", _
    vbInformation, _
    "作成完了"

End Sub


'==================================================
' 指定した名前のシートを取得する。
' 存在しない場合は新しく作成する。
'==================================================

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next

    Set GetOrCreateSheet = _
        ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then

        Set GetOrCreateSheet = _
            ThisWorkbook.Worksheets.Add

        GetOrCreateSheet.Name = sheetName

    End If

End Function


'==================================================
' 問題・解答プリント共通のヘッダーを作成する
'==================================================

Sub CreatePrintHeader( _
    ByVal ws As Worksheet, _
    ByVal title As String, _
    ByVal subtitle As String)

    With ws

        .Cells(1, 1).Value = title
        .Cells(1, 1).Font.Size = 18
        .Cells(1, 1).Font.Bold = True

        .Cells(2, 1).Value = subtitle
        .Cells(2, 1).Font.Size = 11

        .Cells(3, 1).Value = _
            "名前：____________________________    日付：____________"

        .Cells(3, 1).Font.Size = 10

        .Columns("A").ColumnWidth = 90

    End With

End Sub


'==================================================
' 問題プリントの書式設定
'==================================================

Sub FormatPrint(ByVal ws As Worksheet)

    With ws

        .Columns("A").ColumnWidth = 90

        .Rows("1:3").RowHeight = 25

        .Range("A1:A3").HorizontalAlignment = xlLeft

    End With

    ApplyCommonPageSetup ws

End Sub


'==================================================
' 解答プリントの書式設定
'==================================================

Sub FormatAnswer(ByVal ws As Worksheet)

    With ws

        .Columns("A").ColumnWidth = 90

        .Rows("1:3").RowHeight = 25

        .Columns("A").WrapText = True

    End With

    ApplyCommonPageSetup ws

End Sub


'==================================================
' 問題・解答プリント共通の印刷設定
'==================================================

Sub ApplyCommonPageSetup(ByVal ws As Worksheet)

    With ws.PageSetup

        .Orientation = xlPortrait
        .PaperSize = xlPaperA4

        .FitToPagesWide = 1
        .FitToPagesTall = False

        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)

    End With

End Sub


'==================================================
' Long型配列をランダムに並び替える
'
' Fisher-Yates Shuffle を使用。
'==================================================

Sub ShuffleArray(ByRef arr() As Long)

    Dim i As Long
    Dim j As Long
    Dim temp As Long

    For i = UBound(arr) To LBound(arr) + 1 Step -1

        j = Int( _
            (i - LBound(arr) + 1) * Rnd _
        ) + LBound(arr)

        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp

    Next i

End Sub

