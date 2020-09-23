Attribute VB_Name = "g_host"
Public mIP As String
Public mHost As String
Public mUser As String
Public Type SavedData
     sIP As String
     sHost As String
     sUser As String
End Type
Public ConnectInfo(20) As SavedData

Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const conHwndTopmost = -1
Global Const conHwndNoTopmost = -2
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Public Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Public nid As NOTIFYICONDATA
