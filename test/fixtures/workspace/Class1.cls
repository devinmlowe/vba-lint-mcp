Option Explicit

Private mValue As String

Public Property Get Value() As String
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As String)
    mValue = newValue
End Property
