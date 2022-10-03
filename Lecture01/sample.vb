sub 集計()
    Dim Week As String  '曜日を入れる変数
    
    Week = InputBox("集計する曜日を入力してください（例：日，月，火）")

    Worksheets().Add After:=Worksheets(Worksheets.Count) 'シートの追加
    ActiveSheet.Name = "集計" 'シート名の変更

    Worksheets().Add After:=Worksheets(Worksheets.Count) 'シートの追加
    ActiveSheet.Name = "曜日のデータ" 'シート名の変更

    Worksheets("集計").Activate '辻堂8月をアクティブにする

    Range("A1") = "集計結果"
    Range("A2") = Week & "曜日の最高気温の平均（℃）"
    Range("A3") = Week & "曜日の最低気温の平均（℃）"
    Range("A4") = Week & "曜日の日数"

    Dim MaxT As Single '最高気温計算用
    Dim MinT As Single '最低気温計算用
    Dim Sun_No As Integer '日曜日のカウント用

    Dim i As Integer

    MaxT = 0 '変数の初期化
    MinT = 0
    Sun_No = 0

    Sheets("辻堂8月").Range("A1:G1").Copy
    Sheets("曜日のデータ").Rows(1).PasteSplecial(xlPasteAll)

    For i = 2 To 32

        If Sheets("辻堂8月").Range("B" & i) = Week Then
            MaxT = MaxT + Sheets("辻堂8月").Range("E" & i) '最高気温を足す
            MinT = MinT + Sheets("辻堂8月").Range("F" & i) '最低気温を足す
            Sun_No = Sun_No + 1 '日曜日のカウント

            Sheets("辻堂8月").Rows(i).Copy
            Sheets("曜日のデータ").Rows(Sun_No + 1).PasteSplecial(xlPasteAll)

        End If

    Next

    Sheets("集計").Range("B2") = MaxT / Sun_No
    Sheets("集計").Range("B2").NumberFormatLocal = "0.00" '小数点第二位
    
    Sheets("集計").Range("B3") = MinT / Sun_No
    Sheets("集計").Range("B3").NumberFormatLocal = "0.00" '小数点第二位

    Sheets("集計").Range("B4") = Sun_No

    Worksheets("辻堂8月").Activate '辻堂8月をアクティブにする

End Sub