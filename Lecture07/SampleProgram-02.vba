Dim i As Long
i = 3
Do Until Cells(i, 2).Value = ""
    Cells(i, 3).Value = "済"
    i = i + 1
Loop