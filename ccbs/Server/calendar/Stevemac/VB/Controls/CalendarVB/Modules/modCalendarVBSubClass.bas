Attribute VB_Name = "modCalendarVBSubClass"
'----------------------------------------------------------------------
' Copyright (c) 1997, CTR Business Systems, Inc.
'              All Rights Reserved
'
' Information Contained Herin is Proprietary and Confidential
'----------------------------------------------------------------------
Option Explicit

'Window API Message Constants
Public Const WM_CLOSE = &H10
Public Const WM_PAINT = &HF
Public Const WM_SIZE = &H5

'Subclassing Constants
Public Const GWL_WNDPROC = (-4&)
Public Const GWL_USERDATA = (-21&)
Public Const MIN_INSTANCES = 1
Public Const MAX_INSTANCES = 256

Type Instances
    in_use As Boolean       'This instance is alive
    ClassAddr As Long       'Pointer to self
    hwnd As Long            'hWnd being hooked
    PrevWndProc As Long     'Stored for unhooking
End Type

'Hooking Related Declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal _
        hwnd As Long, ByVal nIndex As Long)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
        ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public g_udtSubclassInstances(MIN_INSTANCES To MAX_INSTANCES) As Instances

'Replace MyUC with your usercontrol reference name, also modify for use
'with your specific messages
Public Function SwitchBoard(ByVal hwnd As Long, ByVal MSG As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim nInstance_Check As Integer
    Dim cMyUC As CalendarVB
    Dim lPrevWndProc As Long
    
    On Error GoTo Err_Handler
    Log "SwitchBoard Enter"
    
    'Do this early as we may unhook
    lPrevWndProc = Is_Hooked(hwnd)
    
    If MSG = WM_SIZE Or MSG = WM_CLOSE Then
        For nInstance_Check = MIN_INSTANCES To MAX_INSTANCES
            If g_udtSubclassInstances(nInstance_Check).hwnd = hwnd Then
                On Error Resume Next
                CopyMemory cMyUC, g_udtSubclassInstances(nInstance_Check).ClassAddr, 4
                cMyUC.UserControlResized MSG
                CopyMemory cMyUC, 0&, 4
            End If
        Next nInstance_Check
    End If
    If MSG = WM_PAINT Then
        For nInstance_Check = MIN_INSTANCES To MAX_INSTANCES
            If g_udtSubclassInstances(nInstance_Check).hwnd = hwnd Then
                On Error Resume Next
                CopyMemory cMyUC, g_udtSubclassInstances(nInstance_Check).ClassAddr, 4
                cMyUC.UserControlPaint
                CopyMemory cMyUC, 0&, 4
            End If
        Next nInstance_Check
    End If
    
    SwitchBoard = CallWindowProc(lPrevWndProc, hwnd, MSG, wParam, lParam)

Exit_Proc:
    Log "SwitchBoard Exit"
    Exit Function
    
Err_Handler:
    Log "SwitchBoard Error"
    MsgBox "Error: " & Err & vbCrLf & Error$, vbCritical, "modCalendarVBSubClass::SwitchBoard"
    GoTo Exit_Proc
    
End Function


'Hooks a window or acts as if it does if the window is
'already hooked by a previous instance of myUC.
Public Sub Hook_Window(ByVal hwnd As Long, ByVal iInstance As Integer)
    
    g_udtSubclassInstances(iInstance).PrevWndProc = Is_Hooked(hwnd)
    If g_udtSubclassInstances(iInstance).PrevWndProc = 0& Then
        g_udtSubclassInstances(iInstance).PrevWndProc = SetWindowLong(hwnd, _
            GWL_WNDPROC, AddressOf SwitchBoard)
    End If
    g_udtSubclassInstances(iInstance).hwnd = hwnd
    
End Sub

' Unhooks only if no other instances need the hWnd
Public Sub UnHookWindow(ByVal iInstance As Integer)

    If TimesHooked(g_udtSubclassInstances(iInstance).hwnd) = 1 Then
        SetWindowLong g_udtSubclassInstances(iInstance).hwnd, GWL_WNDPROC, _
            g_udtSubclassInstances(iInstance).PrevWndProc
    End If
    g_udtSubclassInstances(iInstance).hwnd = 0&

End Sub

'Determine if we have already hooked a window,
'and returns the PrevWndProc if true, 0& if false
Private Function Is_Hooked(ByVal hwnd As Long) As Long
    
    Dim iIndex As Integer
    On Error GoTo Err_Handler
    
    Log "Is_Hooked Enter"
    
    Is_Hooked = 0&
    For iIndex = MIN_INSTANCES To MAX_INSTANCES
        If g_udtSubclassInstances(iIndex).hwnd = hwnd Then
            Is_Hooked = g_udtSubclassInstances(iIndex).PrevWndProc
            Exit For
        End If
    Next iIndex
    
Exit_Proc:
    Log "Is_Hooked Exit"
    Exit Function
    
Err_Handler:
    MsgBox "Error: " & Err & vbCrLf & Error$, vbCritical, "modCalendarVBSubClass::Is_Hooked"
    Exit Function
    
End Function

'Returns a count of the number of times a given
'window has been hooked by instances of myUC.
Private Function TimesHooked(ByVal hwnd As Long) As Long
    Dim iIndex As Integer
    Dim nCount As Integer
    
    For iIndex = MIN_INSTANCES To MAX_INSTANCES
        If g_udtSubclassInstances(iIndex).hwnd = hwnd Then
            nCount = nCount + 1
        End If
    Next iIndex

    TimesHooked = nCount

End Function

Public Sub Log(ByVal sMsg As String)

    Dim nFileNum As Integer
    
    nFileNum = FreeFile
    
    Open "c:\eventlog.log" For Append As #nFileNum
    Print #nFileNum, sMsg
    Close nFileNum
    
End Sub
