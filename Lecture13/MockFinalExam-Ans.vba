Option Explicit

Sub 小計()
    Dim i As Long
    
    For i = 2 To 31
        Sheets("売上表").Cells(i, 6).Value = Sheets("売上表").Cells(i, 4).Value * Sheets("売上表").Cells(i, 5).Value
    Next
End Sub

Sub ランク()
    Dim i As Long
    
    For i = 2 To 31
        If Sheets("売上表").Cells(i, 6).Value >= 50000 Then
            Sheets("売上表").Cells(i, 7).Value = "S"
            Sheets("売上表").Cells(i, 7).Font.Color = RGB(0, 0, 255) '文字の色の変更

        ElseIf Sheets("売上表").Cells(i, 6).Value >= 30000 Then
            Sheets("売上表").Cells(i, 7).Value = "A"

        ElseIf Sheets("売上表").Cells(i, 6).Value >= 10000 Then
            Sheets("売上表").Cells(i, 7).Value = "B"

        Else
            Sheets("売上表").Cells(i, 7).Value = "C"
            Sheets("売上表").Cells(i, 7).Font.Color = RGB(255, 0, 0) '文字の色の変更
        End If
    Next
End Sub

Sub 抽出()

    Dim Bunseki As String       '分析条件を入れる変数
    Dim i As Long
    Dim Total As Long
    Dim Data_No As Long         '項目数を入れる変数
    
    Bunseki = InputBox("抽出するランクを入力してください（例：S,A,B,C,D)")
    
    Worksheets().Add After:=Worksheets(Worksheets.Count)    'シートの追加
    ActiveSheet.Name = "抽出"
    
    For i = 2 To 31
        If Sheets("売上表").Cells(i, 7) = Bunseki Then
            Sheets("売上表").Range(Sheets("売上表").Cells(i, 1), Sheets("売上表").Cells(i, 8)).Interior.Color = RGB(255, 255, 0) ' 背景色
            Total = Total + Sheets("売上表").Cells(i, 6)     '合計
            Data_No = Data_No + 1                                   '項目数のカウント
        
            '行のコピー
            'Sheets("売上表").Rows(i).Copy
            Sheets("売上表").Range(Sheets("売上表").Cells(i, 1), Sheets("売上表").Cells(i, 8)).Copy
            'Sheets("抽出").Rows(Data_No + 4).PasteSpecial (xlPasteAll)
            Sheets("抽出").Range(Sheets("抽出").Cells(Data_No + 5, 1), Sheets("抽出").Cells(Data_No + 5, 8)).PasteSpecial (xlPasteAll)
            Sheets("抽出").Range(Sheets("抽出").Cells(Data_No + 5, 1), Sheets("抽出").Cells(Data_No + 5, 8)).Interior.Color = xlNone ' 背景色
        End If
    Next
    
    Sheets("売上表").Range(Sheets("売上表").Cells(1, 1), Sheets("売上表").Cells(1, 8)).Copy
    Sheets("抽出").Range(Sheets("抽出").Cells(5, 1), Sheets("抽出").Cells(5, 8)).PasteSpecial (xlPasteAll)

    Sheets("抽出").Range("A1") = "抽出結果のデータ数"
    Sheets("抽出").Range("A2") = Data_No
    
    Sheets("抽出").Range("B1") = "抽出結果の平均"
    Sheets("抽出").Range("B2") = Total / Data_No
    Sheets("抽出").Range("B2").NumberFormatLocal = "0.0"     '小数点第一位
    
    Worksheets("売上表").Activate         '課題Sheetをアクティブにする
End Sub

Sub 初期化()
    Range("A2: H31").Interior.Color = xlNone
    Range("F2:G31").Clear
    Worksheets("抽出").Delete 'シートを削除
End Sub