Attribute VB_Name = "Module1"
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0

Function ShutDown()
Dim lngresult
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function
