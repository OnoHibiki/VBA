Attribute VB_Name = "ShowProjectInfo"
Sub ShowProjectInfo() '選択した都道府県の物件情報を表示する

    Dim prefecture As String  '検索対象の都道府県
    Dim result As VbMsgBoxResult  ' Yes or No
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim searchRow As Long 'Forで使う検索行
    
    
    prefecture = Range("H3").Value
    
    result = MsgBox( _
        prefecture & " の物件情報を検索しますか？", _
        vbYesNo + vbQuestion, _
        "都道府県確認")
    
    'Yes no 分岐
    If result = vbNo Then
        MsgBox "キャンセルしますね"
        Exit Sub
    End If
    
    '以下、検索処理
    
    Dim hitInfoMsg As String 'For分の中で結果をためていく
    Dim hitCount As Long: hitCount = 0
    
    For searchRow = 2 To lastRow
        
        If Cells(searchRow, 2).Value = prefecture Then
        
            hitCount = hitCount + 1
            
            hitInfoMsg = hitInfoMsg & _
                hitCount & "件目 : " & vbCrLf & _
                "物件ID:" & Cells(searchRow, 1).Value & vbCrLf & _
                "施工会社:" & Cells(searchRow, 3).Value & vbCrLf & _
                "工事開始予定日:" & Cells(searchRow, 4).Value & vbCrLf & _
                "工事終了予定日:" & Cells(searchRow, 5).Value & vbCrLf & _
                "ステータス:" & Cells(searchRow, 6).Value & vbCrLf & _
                "--------------------------------" & vbCrLf & vbCrLf

        End If
        
    Next
    
    MsgBox hitInfoMsg, , prefecture
    
End Sub
