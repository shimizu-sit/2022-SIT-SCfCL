Option Explicit

Sub データ統合()
    ' 指定ファイルまでの絶対パスを保存する変数「パス」をString型で宣言
    Dim パス As String
    ' ファイル名を保存する変数「ファイル名」をString型で宣言
    Dim ファイル名 As String
    
    ' パスはC4，C5，C6に入力されているのでそれぞれを繋げる
    パス = Range("C4").Value & "\" & Range("C5").Value & "\" & Range("C6").Value & "\"
    ' 今回はファイル名はプログラム内で指定する
    ファイル名 = "東京_0107.xlsx"
    
    ' 指定ファイルを開く
    Workbooks.Open パス & ファイル名
    ' 指定ファイルが開いたことをメッセージボックスに表示
    MsgBox ファイル名 & "を開きました"
    ' メッセージボックスの「OK」が押されたら開いたブックを閉じる
    ActiveWorkbook.Close
End Sub
