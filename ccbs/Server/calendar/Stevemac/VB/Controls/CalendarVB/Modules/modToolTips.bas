Attribute VB_Name = "modToolTips"
Option Explicit
' ------------------------------------------------------------------------
'      Copyright © 1997 Microsoft Corporation.  All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
'-------------------------------------------------------------------------
'This module is needed because it provides a WinProc used for subclassing.  Tooltips
'must be provided in this manner because VB5's tooltips for an intrinsic
'control does not work with the SetCapture API in use.  Also, VB's provided
'tooltip is container provided.  Therefore, if a control is used in a
'container that does not provide a tooltiptext property on the extender
'object, the tooltip would not be provided.  The tooltip in this control is
'provided regardless of the container hosting it.
'-------------------------------------------------------------------------
Public gHWndToolTip As Long                 'Hwnd of Tooltip created by this object
Public gbToolTipsInstanciated As Boolean    'If true the ToolTip class window has been created
Public glToolsCount As Long                 'The number of controls using tool tips
                                            
'Get Windows Long Constants
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)
                                            
'Misc Constants
Private Const H_MAX As Long = &HFFFFFFFF + 1
Private Const TOOLTIPS_CLASS As String = "tooltips_class32"
Private Const WS_EX_TOPMOST = &H8&
Private Const CW_USEDEFAULT  As Long = &H80000000
Private Const glSUNKEN_OFFSET = 1
Private Const GDI_ERROR = &HFFFFFFFF

'Messages to relay to ToolTip
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

'ToolTip style
Private Const TTF_IDISHWND = &H1

'Tool Tip messages
Private Const TTM_ACTIVATE = (WM_USER + 1)
#If UNICODE Then
    Private Const TTM_ADDTOOLW = (WM_USER + 50)
    Private Const TTM_ADDTOOL = TTM_ADDTOOLW
#Else
    Private Const TTM_ADDTOOLA = (WM_USER + 4)
    Private Const TTM_ADDTOOL = TTM_ADDTOOLA
#End If
Private Const TTM_RELAYEVENT = (WM_USER + 7)

'ToolTip Notification
Private Const TTN_FIRST = (H_MAX - 520&)
#If UNICODE Then
    Private Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
    Private Const TTN_NEEDTEXT = TTN_NEEDTEXTW
#Else
    Private Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
    Private Const TTN_NEEDTEXT = TTN_NEEDTEXTA
#End If

'Misc ToolTip
Private Const LPSTR_TEXTCALLBACK As Long = -1
                                            
#If UNICODE Then
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

#If DEBUGSUBCLASS Then
    Public goWindowProcHookCreator As Object
#End If

'-------------------------------------------------------------------------
'Purpose:   Address used for subclassing.  Calls instance of the UserControl
'           whose hWnd is stored in USERDATA of window matching passed hWnd
'-------------------------------------------------------------------------
Public Function ToolTipSubWndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    SubWndProc = ConvertUserDataToObject(hwnd).WindowProc(hwnd, MSG, wParam, lParam)
End Function 'ToolTipSubWndProc()

'-------------------------------------------------------------------------
'Purpose:   Gets the hWnd of a UserControl object, and converts it
'           to a reference to the UserControl object without increasing
'           VB's ref count of that object
'-------------------------------------------------------------------------
Private Function ConvertUserDataToObject(hwnd As Long) As CalendarVB
    Dim Obj As CalendarVB
    Dim pObj As Long
    pObj = GetWindowLong(hwnd, GWL_USERDATA)
    CopyMemory Obj, pObj, 4
    Set ConvertUserDataToObject = Obj
    CopyMemory Obj, 0&, 4
End Function 'ConvertUserDataToObject()

