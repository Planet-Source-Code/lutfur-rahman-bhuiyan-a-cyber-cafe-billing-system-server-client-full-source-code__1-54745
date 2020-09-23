VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wLogin 
      Left            =   720
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   3000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF0000&
      Height          =   1935
      Left            =   1680
      ScaleHeight     =   1875
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   7
         Top             =   1440
         Width           =   3975
         Begin VB.CommandButton Command2 
            Caption         =   "Close"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Login"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6525
      TabIndex        =   0
      Top             =   3855
      Width           =   6585
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   3255
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
Dim cnt
Private Type wPos
X As Double
Y As Double
End Type
Private posW As wPos

Private Sub Command1_Click()
'If IsNumeric(Left(Text1.Text, 1)) Then
'Text1.SetFocus
'Exit Sub
'End If
'mn = Trim(Text1.Text)
'For i = 1 To Len(mn)
'If Mid(mn, i, 1) = " " Then
'Text1.SetFocus
'Exit Sub
'End If
'Next
If Len(Trim(Text1.Text)) = 0 Then
Text1.SetFocus
Exit Sub
End If
Winsock1.Close
Winsock1.RemoteHost = "202.174.157.155"
Winsock1.LocalPort = 11111
Winsock1.RemotePort = 7777
Winsock1.Bind Winsock1.LocalPort
Winsock1.SendData "[Log in]," & Winsock1.LocalHostName & "," & Text1.Text & "," & Str(Time) & "," & Text3.Text & "," & Text2.Text
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
'Unload sc
Unload Me
'ShutdownSystem EWX_LOGOFF
End Sub

Private Sub Form_Load()

Timer1.Enabled = True

wServer.Close
wServer.Bind 7778, wServer.LocalIP

wLogin.Close
wLogin.Bind 7779, wServer.LocalIP
End Sub

Private Sub Form_Resize()
Picture2.Left = (Me.Width - Picture2.Width) / 2
Picture2.Top = (Me.Height - Picture2.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload sc
End Sub

Private Sub Timer1_Timer()
r = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 1)
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData gotstring, vbString
Label1.Caption = gotstring

If Label1.Caption = "OK" Then
Timer1.Enabled = False
sttime = Time
stdt = Date
Form2.Text1.Text = sttime
Winsock1.Close
Me.Hide
Form2.Show
Else
'Label1.Caption = "Server not reachable."
End If

If gotstring = "Yes?OK." And gotstring <> "OK" Then
Exit Sub
Else
Label1.Caption = "Server not reachable."
Exit Sub
End If


End Sub


Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub


Private Sub wLogin_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
wLogin.GetData gotstring, vbString
helloa = Len(gotstring)
List1.Clear
For i = 1 To helloa
helloc = helloc + 1
hellob = Mid(gotstring, i, 1)
If hellob = "," Then
hellob = ""
hellod = hellod + hellob
If hellod = "" Then
Else
List1.AddItem hellod
End If
hellod = ""
Else
hellod = hellod + hellob
If helloc = helloa Then List1.AddItem hellod
End If
Next i

If List1.ListCount > 0 Then
List1.ListIndex = 0
If List1 = "[Request Login]" Then
List1.ListIndex = 1
Text3.Text = List1
List1.ListIndex = 2
Text1.Text = List1
If List1.ListCount = 4 Then
List1.ListIndex = 3
Text2.Text = List1
End If
End If

End If


End Sub

Private Sub wServer_DataArrival(ByVal bytesTotal As Long)
wServer.GetData gotstring, vbString

Select Case gotstring

Case "[SHUT DOWN]"
Unload Form2
ShutdownSystem EWX_SHUTDOWN
End

Case "[LOGOFF]"
Unload Form2
ShutdownSystem EWX_LOGOFF
Unload Form2
End

Case "[REBOOT]"
Unload Form2
ShutdownSystem EWX_REBOOT
End

End Select
End Sub


