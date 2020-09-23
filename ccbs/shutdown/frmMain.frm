VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown/Reboot/Logoff Win (9x/NT/2000) and Remote Shutdown (NT)"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":000D
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A&ction!!!"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Action!!!"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Remote System Shutdown:"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Options:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "My System S/R/L (9x/NT/2000)"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   600
      Y2              =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Shutdown/Reboot/Logoff Win (9x/NT/2000) and Remote Shutdown (NT)"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''
'Written by Marty Forde                    '
''''''''''''''''''''''''''''''''''''''''''''
'About Me:                                 '
'Occupation: Sophmore HighSchool Student   '
'    and VB consultant                     '
'Age: 16                                   '
'Expertise Areas: VB Api programming and   '
'    multimedia programming                '
'E-mail: Marty1149@aol.com                 '
''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Dim nUser As String
Dim nMessage As String
Dim nTimeAllow As Long

Private Sub Command1_Click()
    
    If UCase(Combo1.Text) = "SHUTDOWN" Then ShutdownSystem EWX_SHUTDOWN
    If UCase(Combo1.Text) = "REBOOT" Then ShutdownSystem EWX_REBOOT
    If UCase(Combo1.Text) = "LOGOFF" Then ShutdownSystem EWX_LOGOFF
    
End Sub

Private Sub Command2_Click()
    
    'uses the initiatesystemshutdown to remote shutdown a system
    'i used inputboxs' for the info becuase i thought it looked nicer
    
    On Error Resume Next
    
    'get user
    nUser = InputBox("Please enter the username to shutdown.", "Remote Shutdown")
    If nUser = "" Then
        Exit Sub
    End If
    
    'get message
    nMessage = InputBox("Please enter the message you want to send.", "Remote Shutdown")
    
    'get time limit
1:  nTimeAllow = InputBox("Please enter in number of seconds for then computer to shutdown (Please enter in only number of 1 and above)", "Remote Shutdown")
    
    'take action based upon input
    If IsNumeric(nTimeAllow) = False Or nTimeAllow <= Val(0) Then
        MsgBox "Please enter in only numbers as the shutdown time", , "Remote Shutdown"
        GoTo 1
    ElseIf nTimeAllow = "" Then
        Exit Sub
    ElseIf IsNumeric(nTimeAllow) = True And nTimeAllow > Val(0) Then
        InitiateSystemShutdown nUser, nMessage, nTimeAllow, False, False
    End If

End Sub
