Option Explicit

Sub Sample10_01_1()
  Range("A1").AutoFilter field:=3, Criteria1:="生活"
  Range("A1").AutoFilter field:=4, Criteria1:=">=10000"
End Sub

Sub Sample10_01_2()
  Range("A1").AutoFilter field:=4, Criteria1:=">=10000", Operator:=xlAnd, Criteria2:="<=20000"
End Sub
