Option Explicit

Sub 表のデータ抽出()
  Range("A1").AutoFilter field:=3, Criteria1:="生活"
  Range("A1").AutoFilter field:=4, Criteria1:=">=10000"
End Sub

Sub Sample10_02_1()
  ActiveSheet.AutoFilterMode = False
End Sub

Sub Sample10_02_2()
  Range("A1").AutoFilter field:=3
End Sub
