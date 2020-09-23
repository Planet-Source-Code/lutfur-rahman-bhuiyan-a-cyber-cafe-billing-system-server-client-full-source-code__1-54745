VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_remote_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Login"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wLogin 
      Left            =   4800
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   300
      Left            =   2160
      TabIndex        =   11
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5535
      TabIndex        =   8
      Top             =   4155
      Width           =   5535
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member ID"
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member Name"
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3435
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1080
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3780
      Width           =   2055
   End
   Begin MSComctlLib.ListView l 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4260
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
      NumItems        =   4
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Login Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client ID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   645
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client Name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3435
      Width           =   900
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3780
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remote Machine:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frm_remote_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS1 As Recordset
Attribute adoPrimaryRS1.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Public acmode As Boolean

Private Sub Command1_Click()
If Len(Trim(txtFields(0))) = 0 Then
txtFields(0).SetFocus
Exit Sub
End If
wLogin.Close
wLogin.RemoteHost = Trim(l.SelectedItem.SubItems(2))
wLogin.RemotePort = 7779
wLogin.SendData "[Request Login]," & txtFields(0) & "," & txtFields(1) & "," & txtFields(2)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
frm_Client_List.Show
End Sub

Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from Member", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear
If frm_status.l.ListItems.Count > 1 Then
For i = 1 To frm_status.l.ListItems.Count
frm_status.l.ListItems(i).Selected = True
Set x = l.ListItems.Add(, , frm_status.l.SelectedItem.Text)
x.SubItems(1) = frm_status.l.SelectedItem.SubItems(1)
x.SubItems(2) = frm_status.l.SelectedItem.SubItems(2)
x.SubItems(3) = frm_status.l.SelectedItem.SubItems(4)
Next
End If
End Sub
