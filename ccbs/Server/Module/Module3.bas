Attribute VB_Name = "Module3"
Public Function CtrlValidate(KeyIn As Integer, ValidateString As String) As Boolean
Dim ValidateList As String
Dim KeyOut As Integer

If KeyIn = 8 Or KeyIn = 9 Then
    CtrlValidate = True
    Exit Function
End If
If InStr(1, ValidateString, Chr(KeyIn), 1) > 0 Then
   CtrlValidate = True
Else
   CtrlValidate = False
   Beep
End If
End Function
