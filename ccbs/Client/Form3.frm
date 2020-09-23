VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log Out"
      Height          =   275
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   2520
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   2400
   End
   Begin MSWinsockLib.Winsock wServer 
      Left            =   3000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minite"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Used:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Timer2.Enabled = False
Dim TimeIn As Date, TimeOut As Date, TimeDifferent As Date
Dim HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer

TimeIn = TimeValue(Text1)
TimeOut = TimeValue(Text2)
TimeDifferent = (TimeOut - TimeIn)
HourDiff = Hour(TimeDifferent)
MinuteDiff = Minute(TimeDifferent)
SecondDiff = Second(TimeDifferent)
'Label1.Caption = HourDiff & ":" & MinuteDiff & ":" & SecondDiff

If Val(Text3.Text) = 0 Then
Text3.Text = 1
End If

'Winsock1.Close
'Winsock1.RemoteHost = "202.174.157.155"
'Winsock1.LocalPort = 11111
''Winsock1.RemotePort = 7777
'Winsock1.Bind Winsock1.LocalPort
'Winsock1.SendData "[Log out]," & Winsock1.LocalHostName & "," & Time & "," & Text3.Text
Unload Me
Form1.Show

End Sub

Private Sub Form_Resize()
Me.Left = Screen.Width - Me.Width
Me.Top = (Screen.Height - Me.Height) - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = False
Dim TimeIn As Date, TimeOut As Date, TimeDifferent As Date
Dim HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer

TimeIn = TimeValue(Text1)
TimeOut = TimeValue(Text2)
TimeDifferent = (TimeOut - TimeIn)
HourDiff = Hour(TimeDifferent)
MinuteDiff = Minute(TimeDifferent)
SecondDiff = Second(TimeDifferent)
'Label1.Caption = HourDiff & ":" & MinuteDiff & ":" & SecondDiff

If Val(Text3.Text) = 0 Then
Text3.Text = 1
End If

Winsock1.Close
Winsock1.RemoteHost = "202.174.157.155"
Winsock1.LocalPort = 11111
Winsock1.RemotePort = 7777
Winsock1.Bind Winsock1.LocalPort
Winsock1.SendData "[Log out]," & Winsock1.LocalHostName & "," & Time & "," & Text3.Text

End Sub

Private Sub Timer1_Timer()
Text2.Text = Time
End Sub

Private Sub Timer2_Timer()
Dim TimeIn As Date, TimeOut As Date, TimeDifferent As Date
Dim HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer

TimeIn = TimeValue(Text1)
TimeOut = TimeValue(Text2)
TimeDifferent = (TimeOut - TimeIn)
HourDiff = Hour(TimeDifferent)
MinuteDiff = Minute(TimeDifferent)
SecondDiff = Second(TimeDifferent)
Label1.Caption = HourDiff & " Hour " & MinuteDiff & " Minute " & SecondDiff & " Second "
Text3.Text = (HourDiff * 60) + MinuteDiff

If Text3.Text > Text4.Text Then
Winsock1.SendData "[Time]," & Winsock1.LocalHostName & "," & Text3.Text
Text4.Text = Text3.Text
End If

Winsock1.Close
Winsock1.RemoteHost = "202.174.157.155"
Winsock1.LocalPort = 11111
Winsock1.RemotePort = 7777

End Sub

