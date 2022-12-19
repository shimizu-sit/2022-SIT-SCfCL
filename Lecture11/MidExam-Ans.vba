Option Explicit

Sub 不快指数()
    
    Dim i As Integer
    
    For i = 5 To 369
        Sheets("data").Cells(i, 5) = 0.81 * Sheets("data").Cells(i, 3) + 0.0099 * Sheets("data").Cells(i, 3) * Sheets("data").Cells(i, 4) - 0.143 * Sheets("data").Cells(i, 4) + 46.3
        Sheets("data").Cells(i, 5).NumberFormatLocal = "0.0"    '小数点第一位
    Next
    
End Sub
Sub 体感()
    
    Dim i As Integer
    
    For i = 5 To 369
        
        If Sheets("data").Cells(i, 5) >= 75 Then
            Sheets("data").Cells(i, 6) = "暑い"
        ElseIf Sheets("data").Cells(i, 5) >= 60 Then
            Sheets("data").Cells(i, 6) = "ふつう"
        Else
            Sheets("data").Cells(i, 6) = "寒い"
        End If
        
    Next
    
End Sub
Sub 分析()
    Dim Bunseki As String       '分析条件を入れる変数
    Dim Data_No As Integer      '項目数を入れる変数
    Dim Tenp_Ave As Single      '平均温度を入れる変数
    
    Bunseki = InputBox("分析する体感を入力してください（例：寒い・ふつう・暑い)")
    
    Worksheets().Add After:=Worksheets(Worksheets.Count)    'シートの追加
    ActiveSheet.Name = "分析"                               'シート名の変更
    
    Data_No = 0
    Tenp_Ave = 0
    
    Worksheets("data").Activate
    
    For i = 5 To 369
        If Sheets("data").Cells(i, 6) = Bunseki Then
            
            Sheets("data").Range(Cells(i, 1), Cells(i, 6)).Interior.Color = RGB(255, 255, 0) ' 背景色
            
            Tenp_Ave = Tenp_Ave + Sheets("data").Cells(i, 3)    '合計
            Data_No = Data_No + 1                               '項目数のカウント
        
        End If
       
    Next
    
    Sheets("分析").Range("A1") = "分析結果のデータ数"
    Sheets("分析").Range("A2") = Data_No
    
    Sheets("分析").Range("B1") = "分析結果の平均気温(℃)"
    Sheets("分析").Range("B2") = Tenp_Ave / Data_No
    Sheets("分析").Range("B2").NumberFormatLocal = "0.0"    '小数点第一位
    
    Worksheets("data").Activate     '課題Sheetをアクティブにする
End Sub

Sub 初期化()
    
    Range("A5: F369 ").Interior.Color = xlNone
    Range("E5:F369").Clear
    Worksheets("分析").Delete   'シートを削除
    
End Sub
