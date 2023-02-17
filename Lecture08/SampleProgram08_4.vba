Option Explicit

Sub データ統合()
    ' 指定ファイルまでの絶対パスを保存する変数「パス」をString型で宣言
    Dim パス As String
    ' ファイル名を保存する変数「ファイル名」をString型で宣言
    Dim ファイル名 As String
    
    ' パスはC4，C5，C6に入力されているのでそれぞれを繋げる
    パス = Range("C4").Value & "¥" & Range("C5").Value & "¥" & Range("C6").Value & "¥"
    ' ファイル名は検索するためDirを使用する
    ファイル名 = Dir(パス & "*.xlsx")
    
    ' ファイル名が空欄になるまで繰り返す
    Do While ファイル名 <> ""
        ' 指定ファイルを開く
        Workbooks.Open パス & ファイル名

        ' 指定ファイルのデータ数をカウントする
        データ数 = Range("A1").CurrentRegion.Rows.Count - 1
        ' データ数を表示する
        MsgBox ファイル名 & " : " & データ数
        ' メッセージボックスの「OK」が押されたら開いたブックを閉じる
        ActiveWorkbook.Close
        ' 次に見つかったブックの名前を変数に代入する
        ファイル名 = Dir
    Loop
End Sub
