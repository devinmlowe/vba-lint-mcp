Sub Test()
    Dim x As Long
    x = 1
UsedLabel:
    MsgBox x
UnusedLabel:
    x = 2
    GoTo UsedLabel
End Sub
