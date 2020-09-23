VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOpenOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Order"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pick"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ListView l 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Client ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Client Name"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmOpenOrder"
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
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Public aFlag As Boolean, eFlag As Boolean

Private Sub Command1_Click()
If l.ListItems.Count > 0 Then
frmSales.txtDoc.Text = l.SelectedItem.Text
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]where doc='" & frmSales.txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic
frmSales.l.ListItems.Clear
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = frmSales.l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
frmSales.txtFields(4).Text = adoPrimaryRS.Fields(4)
frmSales.txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
frmSales.txtFields(2).Text = adoPrimaryRS.Fields(8)
frmSales.txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If frmSales.l.ListItems.Count > 0 Then
For i = 1 To frmSales.l.ListItems.Count
frmSales.l.ListItems(i).Selected = True
tot = tot + Val(frmSales.l.SelectedItem.SubItems(5)) * Val(frmSales.l.SelectedItem.SubItems(4))
Next
End If

frmSales.Text1.Text = tot
frmSales.C1.Clear
frmSales.C2.Clear
frmSales.C3.Clear
frmSales.C4.Clear

If frmSales.l.ListItems.Count > 0 Then
frmSales.l.ListItems(1).Selected = True
C1 = frmSales.l.SelectedItem.Text
C2 = frmSales.l.SelectedItem.SubItems(1)
C3 = frmSales.l.SelectedItem.SubItems(2)
C4 = frmSales.l.SelectedItem.SubItems(3)
frmSales.txtFields(8).Text = frmSales.l.SelectedItem.SubItems(4)
frmSales.txtFields(9).Text = frmSales.l.SelectedItem.SubItems(5)
End If

End If
Unload Me
End Sub

Private Sub Form_Load()
clck = False
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from [Order No]", db, adOpenStatic, adLockOptimistic
If adoPrimaryRS3.RecordCount > 0 Then
adoPrimaryRS3.MoveFirst
Do Until adoPrimaryRS3.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS3.Fields("Voucher No"))
x.SubItems(1) = adoPrimaryRS3.Fields("Date")
x.SubItems(2) = adoPrimaryRS3.Fields("Vendor ID")
x.SubItems(3) = adoPrimaryRS3.Fields("Vendor Name")
adoPrimaryRS3.MoveNext
Loop
End If
End Sub

