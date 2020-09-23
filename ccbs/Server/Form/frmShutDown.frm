VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmShutDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shut Down"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   2385
      Width           =   5715
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   300
         Left            =   4110
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Logoff"
         Height          =   300
         Left            =   2910
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restart"
         Height          =   300
         Left            =   1710
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Shutdown"
         Height          =   300
         Left            =   510
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView l 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr. #"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Machine Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP Address"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Login Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin MSWinsockLib.Winsock wControl 
      Left            =   5160
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmShutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
wControl.Close
wControl.RemoteHost = Trim(l.SelectedItem.SubItems(2))
wControl.RemotePort = 7778
wControl.SendData "[SHUT DOWN]"
'MsgBox wControl.RemotePort
End Sub

Private Sub Command2_Click()
wControl.Close
wControl.RemoteHost = Trim(l.SelectedItem.SubItems(2))
wControl.RemotePort = 7778
wControl.SendData "[REBOOT]"
End Sub

Private Sub Command3_Click()
wControl.Close
wControl.RemoteHost = Trim(l.SelectedItem.SubItems(2))
wControl.RemotePort = 7778
wControl.SendData "[LOGOFF]"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
l.ListItems.Clear
If frm_status.l.ListItems.Count > 1 Then
For i = 1 To frm_status.l.ListItems.Count
frm_status.l.ListItems(i).Selected = True
Set x = l.ListItems.Add(, , frm_status.l.SelectedItem.Text)
x.SubItems(1) = frm_status.l.SelectedItem.SubItems(1)
x.SubItems(2) = frm_status.l.SelectedItem.SubItems(2)
x.SubItems(3) = frm_status.l.SelectedItem.SubItems(3)
x.SubItems(4) = frm_status.l.SelectedItem.SubItems(4)
Next
End If
End Sub

