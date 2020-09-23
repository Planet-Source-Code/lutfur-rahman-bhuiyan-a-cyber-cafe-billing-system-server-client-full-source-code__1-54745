VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Details Bill"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11115
      TabIndex        =   5
      Top             =   5805
      Width           =   11115
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   300
         Left            =   8760
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView l 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8705
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Machine Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "St. Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "End Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tot. Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Tot.l Bill"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tot. Dis."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Net Bill"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "dfdsf"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "dfdsfsdfsdf"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Client Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Client ID:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmdetbill"
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
Unload Me
End Sub

Private Sub Form_Load()
Label3.Caption = clnID
Label4.Caption = clnName
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from bill where [Client ID]='" & clnID & "'", db, adOpenStatic, adLockOptimistic
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields("Machine"))
x.SubItems(1) = adoPrimaryRS.Fields("st time")
x.SubItems(2) = adoPrimaryRS.Fields("end time")
x.SubItems(3) = adoPrimaryRS.Fields("tot time")
x.SubItems(4) = adoPrimaryRS.Fields("rate")
x.SubItems(5) = adoPrimaryRS.Fields("discount")
x.SubItems(6) = adoPrimaryRS.Fields("tot bill")
x.SubItems(7) = adoPrimaryRS.Fields("tot discount")
x.SubItems(8) = adoPrimaryRS.Fields("net")
adoPrimaryRS.MoveNext
Loop


End Sub

