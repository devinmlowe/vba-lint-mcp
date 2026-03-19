Sub Test()
    Dim strName As String
    Dim intCount As Integer
    Dim goodName As String
    Dim lngTotal As Long
    strName = "test"
    intCount = 1
    goodName = "ok"
    lngTotal = 42
    MsgBox strName & CStr(intCount) & goodName & CStr(lngTotal)
End Sub
