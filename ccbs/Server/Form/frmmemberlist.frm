VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmemberlist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client List"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pick"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView l 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Name"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmmemberlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Command1_Click()
If l.ListItems.Count > 0 Then
frmSalesOrder.txtFields(4) = l.SelectedItem.Text
frmSalesOrder.txtFields(5) = l.SelectedItem.SubItems(1)
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from member", db, adOpenStatic, adLockOptimistic
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields("member id"))
x.SubItems(1) = adoPrimaryRS.Fields("member name")
adoPrimaryRS.MoveNext
Loop
End Sub

