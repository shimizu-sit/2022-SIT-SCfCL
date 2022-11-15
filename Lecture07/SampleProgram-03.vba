Dim i As Long
i = 3
Do
    Cells(i, 3).Value = "済"
    i = i + 1
Loop While Cells(i, 2).Value <> ""