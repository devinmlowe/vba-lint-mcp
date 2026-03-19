Option Explicit

Private total As Long

Sub Initialize()
    total = 0
End Sub

Sub AddValue(ByVal amount As Long)
    total = total + amount
End Sub

Function GetTotal() As Long
    GetTotal = total
End Function
