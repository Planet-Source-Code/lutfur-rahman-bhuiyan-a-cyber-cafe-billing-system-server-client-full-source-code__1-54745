VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Client_List1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client List"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5700
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
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
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Client Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Client Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Password"
         Object.Width           =   2
      EndProperty
   End
End
Attribute VB_Name = "frm_Client_List1"
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
frmReceipt.txtFields(0).Text = l.SelectedItem.Text
frmReceipt.txtFields(1).Text = l.SelectedItem.SubItems(1)
frmReceipt.txtFields(2).Text = l.SelectedItem.SubItems(3)
Unload Me
End Sub

Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from Member", db, adOpenStatic, adLockOptimistic
Call addlist
End Sub

Private Sub addlist()
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If
l.ListItems.Clear
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields("Member ID"))
If Not IsNull(adoPrimaryRS.Fields("Member Name")) Then
x.SubItems(1) = adoPrimaryRS.Fields("Member Name")
End If
If adoPrimaryRS.Fields("type") = 0 Then
x.SubItems(2) = "Regular"
Else
x.SubItems(2) = "Temporary"
End If
If Len(Trim(adoPrimaryRS.Fields("password"))) > 0 Then
x.SubItems(3) = adoPrimaryRS.Fields("password")
End If
adoPrimaryRS.MoveNext
Loop
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveLast
End If
End Sub

