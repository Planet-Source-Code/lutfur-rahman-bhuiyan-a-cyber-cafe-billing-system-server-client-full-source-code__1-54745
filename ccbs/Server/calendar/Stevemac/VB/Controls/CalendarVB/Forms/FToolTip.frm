VERSION 5.00
Begin VB.Form FToolTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "ToolTip"
   ClientHeight    =   480
   ClientLeft      =   2055
   ClientTop       =   2505
   ClientWidth     =   1860
   Icon            =   "FToolTip.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   124
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrToolTip 
      Enabled         =   0   'False
      Left            =   90
      Top             =   60
   End
End
Attribute VB_Name = "FToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================
'  FToolTip.frm
'  Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        2/12/97
'=======================================================

'$Runtime Dependencies:
'$DesignTime Dependencies:

'=======================================================
'  Methods and Properties
'
'  Method DisplayToolTip(vTipID As Variant, sToolTipText As String)
'    Initiates the process of displaying a tooltip
'    vTipID:= An id value that is used internal to
'            determine how the tooltip window should
'            react
'    sToolTipText:= The text too display in the tooltip
'                   window
'  Method HideToolTip()
'    Hides a currently displayed tooltip
'  Property HideInterval(Default:=1500) As Long
'    The time in millseconds that the tooltip will
'    wait before hiding the window
'  Property ShowInterval(Default:=1000) As Long
'    The time in millseconds that the tooltip will
'    wait before displaying the window
'  Property TipText() As String
'=======================================================

'=======================================================
'  Usage Notes
'  This form is used to display a floating tooltips
'  window above a control.
'
'=======================================================

Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const HWND_TOP& = 0
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOACTIVATE& = &H10
Private Const SWP_NOSIZE& = &H1
Private Const SWP_SHOWWINDOW& = &H40

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

Private mbEnabled As Boolean
Private mvTipID As Variant
Private msTipText As String
Private iMouseY As Long
Private iMouseX As Long
Private mnTimerMode As Integer
Private mlTimerStart As Long
Private mlShowInterval As Long
Private mlHideInterval As Long
Private Const TIP_SHOW = 1
Private Const TIP_HIDE = 2

Public Sub DisplayToolTip(vTipID As Variant, sToolTipText As String)

    mbEnabled = True
    
    'Check to see if were already displaying this tip
    If tmrToolTip.Enabled And vTipID = mvTipID Then Exit Sub
    
    'Set our values to use when the tip window is displayed
    tmrToolTip.Enabled = False
    mvTipID = vTipID
    msTipText = " " & sToolTipText & "  "
    mnTimerMode = TIP_SHOW
    If ((Timer - mlTimerStart) * 1000) <= mlHideInterval Then
        tmrToolTip.Interval = 1
        mlTimerStart = Timer
    Else
        'Need to account for the start time delay
        mlTimerStart = Timer + (mlShowInterval \ 1000)
        tmrToolTip.Interval = mlShowInterval
    End If
    tmrToolTip.Enabled = True
    
End Sub

Private Sub Form_Initialize()
    'Set some of our default values
    mlShowInterval = 1000
    mlHideInterval = 5000
End Sub

Public Property Get ShowInterval() As Long
    ShowInterval = mlShowInterval
End Property

Public Property Let ShowInterval(lInterval As Long)
    mlShowInterval = lInterval
End Property

Public Property Get HideInterval() As Long
    HideInterval = mlHideInterval
End Property

Public Property Let HideInterval(lInterval As Long)
    mlHideInterval = lInterval
End Property

Public Sub HideToolTip()
    'Reset some of our values and then
    'unload the tip window
    tmrToolTip.Enabled = False
    mlTimerStart = 0
    mvTipID = ""
    Unload Me
End Sub

Public Property Get TipText() As String
    TipText = msTipText
End Property

Public Property Let TipText(sTipText As String)
    msTipText = " " & sTipText & " "
    Call DisplayForm
End Property

Private Sub DisplayForm()
    
    Dim pt As POINTAPI  'New ssPoint
    'Dim oSystem As New ssSystem
    Dim iTextWidth As Long
    Dim i As Integer
    
    If Not mbEnabled Then Exit Sub
    
    'Get the current mouse pointer position
    Call GetCursorPos(pt)  'oSystem.GetCursorPos pt
    'If we have not moved then no need to redisplay
    If (pt.y = iMouseY) And (pt.x = iMouseX) Then Exit Sub
    iMouseY = pt.y
    iMouseX = pt.x
    
   'Finds its width and ajust the tooltip form's width accordingly
   iTextWidth = Me.TextWidth(msTipText)
   Me.Width = Me.Width * iTextWidth / Me.ScaleWidth
   
   'The tooltip form's height is fixed at 19 pixels
   Me.Height = Me.Height * 19 / Me.ScaleHeight
   
   'Print the tooltip help string on the tooltip form
   Me.Cls
   Me.CurrentY = 3
   Me.Print msTipText

   'Just to make it really look like win95, we'll draw the shadow borders on the form...
   Me.Line (0, 0)-(Me.ScaleWidth, Me.ScaleHeight), &HFFFFFF, B
   Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1), &H0&
   Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), &H0&

   '...and adjust the tooltip's position at an offset of (X+2,Y+18) from the mouse cursor.
   Me.Top = (iMouseY + 18) * Screen.TwipsPerPixelY
   Me.Left = (iMouseX + 2) * Screen.TwipsPerPixelY
   
   'Make sure that the tooltip form is over the active form
   'Me.ZOrder 0
   
   'Show the tooltip form using Windows' ShowWindow function with SW_SHOWNOACTIVATE attribute
   'so that the tooltip won't get the focus. This avoids a flashing Title Bar on the
   'main form which could confuse the user.
   'i = ShowWindow(Me.hWnd, SW_SHOWNOACTIVATE)
    i = SetWindowPos(Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
End Sub

Private Sub tmrToolTip_Timer()

    'Stop our timer
    tmrToolTip.Enabled = False
    'Check to see if we should display or hide
    'our tooltip
    Select Case mnTimerMode
        Case TIP_SHOW
            'Display the tip window
            Call DisplayForm
            'Setup for hiding our tip window
            mnTimerMode = TIP_HIDE
            tmrToolTip.Interval = mlHideInterval
            If mlHideInterval > 0 Then tmrToolTip.Enabled = True
        Case TIP_HIDE
            mlTimerStart = 0
            Unload Me
    End Select

End Sub

Public Sub ResetTimer()
    mlTimerStart = Timer
End Sub
