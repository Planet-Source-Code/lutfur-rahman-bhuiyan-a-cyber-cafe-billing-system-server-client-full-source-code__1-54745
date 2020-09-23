VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form DueList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Due List"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pick"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView l1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tot. Time"
         Object.Width           =   2646
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
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Net Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Received"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Receivable"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "DueList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS1 As Recordset
Attribute adoPrimaryRS1.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS4 As Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS5 As Recordset
Attribute adoPrimaryRS5.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim gotstring As String
Public DX, dy
Public dwn As Boolean
Public CLTYPE, rate, discount


Private Sub Command1_Click()
If l1.ListItems.Count > 0 Then
frmReceipt1.txtFields(0).Text = l1.SelectedItem.Text
frmReceipt1.txtFields(1).Text = l1.SelectedItem.SubItems(1)
frmReceipt1.txtFields(3).Text = l1.SelectedItem.SubItems(6)
frmReceipt1.txtFields(5).Text = l1.SelectedItem.SubItems(7)
'frmReceipt.txtFields(6).Text = l1.SelectedItem.SubItems(8)
frmReceipt1.txtFields(8).Text = l1.SelectedItem.SubItems(9)
frmReceipt1.txtFields(7).Text = l1.SelectedItem.SubItems(10)
frmReceipt1.txtFields(4).Text = Format(Date, "dd/mm/yyyy")
frmReceipt1.txtFields(3).Text = frmReceipt1.txtFields(3).Text - frmReceipt1.txtFields(8).Text
frmReceipt1.txtFields(6).Text = frmReceipt1.txtFields(3).Text - frmReceipt1.txtFields(5).Text
frmReceipt1.txtFields(7).Text = frmReceipt1.txtFields(6).Text
End If
frmReceipt1.Show
Unload Me
End Sub

Private Sub Form_Load()
Call BillProcc
End Sub


Private Sub BillProcc()

On Error Resume Next

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

db.Execute "delete * from tmpbill"

Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from bill", db, adOpenStatic, adLockOptimistic

Set adoPrimaryRS2 = New Recordset
adoPrimaryRS2.Open "select * from member", db, adOpenStatic, adLockOptimistic

Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic

Set adoPrimaryRS4 = New Recordset
adoPrimaryRS4.Open "select * from [user group]", db, adOpenStatic, adLockOptimistic


If adoPrimaryRS2.RecordCount > 0 Then
adoPrimaryRS2.MoveFirst
End If

Do Until adoPrimaryRS2.EOF
With adoPrimaryRS3
.AddNew
.Fields("Client ID") = adoPrimaryRS2.Fields("Member ID")
.Fields("user") = adoPrimaryRS2.Fields("Member Name")
.Fields("category") = adoPrimaryRS2.Fields("Group Name")
.Fields("rate") = adoPrimaryRS2.Fields("rate")
.Fields("discount") = adoPrimaryRS2.Fields("discount")
.Update
End With
adoPrimaryRS2.MoveNext
Loop


If adoPrimaryRS3.RecordCount > 0 Then
adoPrimaryRS3.MoveFirst
End If

Do Until adoPrimaryRS3.EOF

Set adoPrimaryRS2 = New Recordset
clid = adoPrimaryRS3.Fields("Client ID")
adoPrimaryRS2.Open "select * from bill where [Client ID]='" & clid & "'", db, adOpenStatic, adLockOptimistic

tm = 0
dis = 0

If adoPrimaryRS2.RecordCount > 0 Then adoPrimaryRS2.MoveFirst
Do Until adoPrimaryRS2.EOF
tm = adoPrimaryRS2.Fields("tot time") + tm
adoPrimaryRS2.MoveNext
Loop

With adoPrimaryRS3
.Fields("tot time") = tm
.Update
End With

adoPrimaryRS3.MoveNext
Loop


If adoPrimaryRS3.RecordCount > 0 Then
adoPrimaryRS3.MoveFirst
End If

If adoPrimaryRS4.RecordCount > 0 Then adoPrimaryRS4.MoveFirst
Do Until adoPrimaryRS4.EOF
Set adoPrimaryRS3 = New Recordset
clid = adoPrimaryRS4.Fields("Group Name")
adoPrimaryRS3.Open "select * from tmpbill where [category]='" & clid & "'", db, adOpenStatic, adLockOptimistic
Do Until adoPrimaryRS3.EOF
With adoPrimaryRS3
mt = Val(adoPrimaryRS4.Fields("Munimum Time"))
.Fields("min Time") = mt
.Update
End With
adoPrimaryRS3.MoveNext
Loop
adoPrimaryRS4.MoveNext
Loop

Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic

Do Until adoPrimaryRS3.EOF
tb = 0
td = 0
nb = 0
clid = adoPrimaryRS3.Fields("Client ID")
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from bill where [Client ID]='" & clid & "'", db, adOpenStatic, adLockOptimistic
Do Until adoPrimaryRS1.EOF
tb = tb + adoPrimaryRS1.Fields("tot bill")
td = td + adoPrimaryRS1.Fields("tot discount")
adoPrimaryRS1.MoveNext
Loop

With adoPrimaryRS3
.Fields("total bill") = tb
.Fields("dicount") = td
.Fields("net") = tb - td
.Fields("receipt") = 0
End With

adoPrimaryRS3.MoveNext
Loop

Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic

Do Until adoPrimaryRS3.EOF
tb = 0
td = 0
nb = 0
clid = adoPrimaryRS3.Fields("Client ID")
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from member where [Member ID]='" & clid & "'", db, adOpenStatic, adLockOptimistic
With adoPrimaryRS3
.Fields("receipt") = .Fields("receipt") + adoPrimaryRS1.Fields("amount")
.Fields("receivable") = adoPrimaryRS3.Fields("net") - .Fields("receipt")
.Update
End With
adoPrimaryRS3.MoveNext
Loop

Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic

Do Until adoPrimaryRS3.EOF
rc = 0
clid = adoPrimaryRS3.Fields("Client ID")
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from receipt where [Client ID]='" & clid & "'", db, adOpenStatic, adLockOptimistic
Do Until adoPrimaryRS1.EOF
If IsNumeric(adoPrimaryRS1.Fields("amount")) Then
rc = rc + adoPrimaryRS1.Fields("amount")
End If
adoPrimaryRS1.MoveNext
Loop
With adoPrimaryRS3
.Fields("receipt") = .Fields("receipt") + rc
.Fields("receivable") = adoPrimaryRS3.Fields("net") - .Fields("receipt")
.Update
End With
adoPrimaryRS3.MoveNext
Loop

Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic
Do Until adoPrimaryRS3.EOF
If adoPrimaryRS3.Fields("receivable") > 0 Then
Set x = l1.ListItems.Add(, , adoPrimaryRS3.Fields("Client ID"))
x.SubItems(1) = adoPrimaryRS3.Fields("user")
x.SubItems(2) = adoPrimaryRS3.Fields("category")
x.SubItems(3) = adoPrimaryRS3.Fields("tot time")
x.SubItems(4) = adoPrimaryRS3.Fields("rate")
x.SubItems(5) = adoPrimaryRS3.Fields("discount")
If Not IsNull(adoPrimaryRS3.Fields("total bill")) Then
x.SubItems(6) = adoPrimaryRS3.Fields("total bill")
End If
If Not IsNull(adoPrimaryRS3.Fields("dicount")) Then
x.SubItems(7) = adoPrimaryRS3.Fields("dicount")
End If
If Not IsNull(adoPrimaryRS3.Fields("net")) Then
x.SubItems(8) = Format(adoPrimaryRS3.Fields("net"), "#,##0.00")
End If
If Not IsNull(adoPrimaryRS3.Fields("receipt")) Then
x.SubItems(9) = Format(adoPrimaryRS3.Fields("receipt"), "#,##0.00")
End If
If Not IsNull(adoPrimaryRS3.Fields("receivable")) Then
x.SubItems(10) = Format(adoPrimaryRS3.Fields("receivable"), "#,##0.00")
End If


End If
adoPrimaryRS3.MoveNext
Loop

End Sub

