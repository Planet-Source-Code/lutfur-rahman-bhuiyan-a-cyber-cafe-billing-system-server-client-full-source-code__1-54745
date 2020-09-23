VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_status 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3240
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
      Begin Project1.xpButton Command1 
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         TX              =   "Refreash"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frm_status.frx":0000
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   4695
         TabIndex        =   10
         Top             =   0
         Width           =   4695
         Begin VB.CommandButton Command3 
            Caption         =   "X"
            Height          =   255
            Left            =   4320
            TabIndex        =   11
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pending Bill"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   15
            Width           =   1020
         End
      End
      Begin MSComctlLib.ListView l1 
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12632064
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
      Begin Project1.xpButton command4 
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         TX              =   "Details"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frm_status.frx":001C
      End
      Begin Project1.xpButton command5 
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         TX              =   "Receipt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frm_status.frx":0038
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Bill"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5160
      ScaleHeight     =   2175
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   2760
      Width           =   4695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   4695
         TabIndex        =   4
         Top             =   0
         Width           =   4695
         Begin VB.CommandButton Command2 
            Caption         =   "X"
            Height          =   255
            Left            =   4320
            TabIndex        =   5
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Console:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   15
            Width           =   750
         End
      End
      Begin VB.TextBox console 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2640
      Top             =   5520
   End
   Begin MSWinsockLib.Winsock wot1 
      Left            =   600
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1920
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComctlLib.ListView l 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632064
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   16
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
         Text            =   "St. Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "End Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tot. Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Client Category"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Client ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Min. Time"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Tot. Bill"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Tot. Dis."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Net Bill"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSWinsockLib.Winsock port80 
      Left            =   1560
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture5 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture6 
      Height          =   2175
      Left            =   5160
      ScaleHeight     =   2115
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frm_status"
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
l1.ListItems.Clear
Call BillProcc
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
End Sub

Private Sub Command3_Click()
Picture3.Visible = False
End Sub

Private Sub Command4_Click()
If l1.ListItems.Count > 0 Then
clnID = l1.SelectedItem.Text
clnName = l1.SelectedItem.SubItems(1)
frmdetbill.Show
End If
End Sub

Private Sub command5_Click()
If l1.ListItems.Count > 0 Then
frmReceipt.txtFields(0).Text = l1.SelectedItem.Text
frmReceipt.txtFields(1).Text = l1.SelectedItem.SubItems(1)
frmReceipt.txtFields(3).Text = l1.SelectedItem.SubItems(6)
frmReceipt.txtFields(5).Text = l1.SelectedItem.SubItems(7)
'frmReceipt.txtFields(6).Text = l1.SelectedItem.SubItems(8)
frmReceipt.txtFields(8).Text = l1.SelectedItem.SubItems(9)
frmReceipt.txtFields(7).Text = l1.SelectedItem.SubItems(10)
frmReceipt.txtFields(4).Text = Format(Date, "dd/mm/yyyy")
frmReceipt.txtFields(3).Text = frmReceipt.txtFields(3).Text - frmReceipt.txtFields(8).Text
frmReceipt.txtFields(6).Text = frmReceipt.txtFields(3).Text - frmReceipt.txtFields(5).Text
frmReceipt.txtFields(7).Text = frmReceipt.txtFields(6).Text
frmReceipt.Show
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
runserver = True
'MDIForm1.Image1.Picture = MDIForm1.Image2.Picture
'MDIForm1.Picture2.BackColor = vbGreen
wot1.LocalPort = 80
wot1.Listen
'wot1.Bind 8080, wot1.LocalIP

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select sl,machine,ip from ip", db, adOpenStatic, adLockOptimistic


Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from bill", db, adOpenStatic, adLockOptimistic

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If

Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
Load Winsock1(adoPrimaryRS.Fields(0))
'Winsock1(adoPrimaryRS.Fields(0)).RemoteHost = adoPrimaryRS.Fields(1)
'Winsock1(adoPrimaryRS.Fields(0)).RemotePort = 11111
'Winsock1(adoPrimaryRS.Fields(0)).LocalPort = adoPrimaryRS.Fields(0)
'Winsock1(adoPrimaryRS.Fields(0)).Bind Winsock1(adoPrimaryRS.Fields(0)).LocalPort
adoPrimaryRS.MoveNext
Loop

'Winsock1(0).RemoteHost = adoPrimaryRS.Fields(1)
'Winsock1(0).RemotePort = 11111
Winsock1(0).LocalPort = 7777
Winsock1(0).Bind Winsock1(0).LocalPort

Call BillProcc


'If adoPrimaryRS3.RecordCount > 0 Then
'adoPrimaryRS3.MoveFirst
'End If

'Set adoPrimaryRS3 = New Recordset
'adoPrimaryRS3.Open "select * from tmpbill", db, adOpenStatic, adLockOptimistic"

'Do Until adoPrimaryRS3.EOF
'Set x = l1.ListItems.Add(, , adoPrimaryRS3.Fields("Client ID"))
'x.SubItems(1) = adoPrimaryRS3.Fields("user")
'x.SubItems(2) = adoPrimaryRS3.Fields("category")
'x.SubItems(3) = adoPrimaryRS3.Fields("tot time")
'x.SubItems(4) = adoPrimaryRS3.Fields("rate")
'x.SubItems(5) = adoPrimaryRS3.Fields("discount")
'If Not IsNull(adoPrimaryRS3.Fields("total bill")) Then
'x.SubItems(6) = adoPrimaryRS3.Fields("total bill")
'End If
'x.SubItems(6) = adoPrimaryRS3.Fields(5)
'If Not IsNull(adoPrimaryRS3.Fields(7)) Then
'x.SubItems(7) = adoPrimaryRS3.Fields(7)
'End If
'adoPrimaryRS3.MoveNext
'Loop



End Sub

Private Sub Form_Resize()
l.left = 100
l.Width = Me.Width - 400
End Sub

Private Sub mnuSD_Click()
End Sub

Private Sub l1_DblClick()
If l1.ListItems.Count > 0 Then
clnID = l1.SelectedItem.Text
clnName = l1.SelectedItem.SubItems(1)
frmdetbill.Show
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Picture1.hWnd, _
            WM_NCLBUTTONDOWN, _
            HTCAPTION, 0&
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Picture3.hWnd, _
            WM_NCLBUTTONDOWN, _
            HTCAPTION, 0&
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
wot1.Close
wot1.Listen
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
List1.Clear
Dim rmhost As String
Winsock1(0).GetData gotstring, vbString
'tmhost = Winsock1(Index).RemoteHost
'MsgBox rmhost
If Len(console.Text) > 1024 * 10 Then console.Text = ""
console.Text = console.Text + Chr$(13) & Chr$(10) + gotstring
console.Text = console.Text + Chr$(13) & Chr$(10) + "Request received from: " & Winsock1(0).RemoteHostIP

helloa = Len(gotstring)
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
End If

If List1 = "[Log out]" Then
List1.ListIndex = 1
clm = List1
List1.ListIndex = 2
cln = List1
List1.ListIndex = 3
tm = List1

sl = 0

Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from bill", db, adOpenStatic, adLockOptimistic

If adoPrimaryRS1.RecordCount > 0 Then
adoPrimaryRS1.MoveFirst
Do Until adoPrimaryRS1.EOF
If IsNumeric(adoPrimaryRS1.Fields(0)) And adoPrimaryRS1.Fields(0) > sl Then
sl = adoPrimaryRS1.Fields(0)
End If
adoPrimaryRS1.MoveNext
Loop
End If


If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True

If l.SelectedItem.SubItems(1) = clm Then
l.SelectedItem.SubItems(4) = "Not Reachable"
l.SelectedItem.SubItems(6) = cln
l.SelectedItem.SubItems(7) = tm
If tm < Val(l.SelectedItem.SubItems(12)) Then
l.SelectedItem.SubItems(13) = Val(l.SelectedItem.SubItems(12)) * Val(l.SelectedItem.SubItems(10))
Else
l.SelectedItem.SubItems(13) = tm * Val(l.SelectedItem.SubItems(10))
End If
l.SelectedItem.SubItems(14) = (l.SelectedItem.SubItems(13) * Val(l.SelectedItem.SubItems(11)) / 100)
l.SelectedItem.SubItems(15) = Val(l.SelectedItem.SubItems(13)) - Val(l.SelectedItem.SubItems(14))
With adoPrimaryRS1
.AddNew
.Fields(0) = sl + 1
.Fields(1) = l.SelectedItem.SubItems(1)
.Fields(2) = l.SelectedItem.SubItems(2)
.Fields(6) = l.SelectedItem.SubItems(3)
.Fields(3) = l.SelectedItem.SubItems(5)
.Fields(4) = l.SelectedItem.SubItems(6)
.Fields(5) = l.SelectedItem.SubItems(7)
.Fields(7) = l.SelectedItem.SubItems(9)
.Fields(8) = l.SelectedItem.SubItems(8)
.Fields(9) = l.SelectedItem.SubItems(10)
.Fields(10) = l.SelectedItem.SubItems(11)
.Fields(14) = l.SelectedItem.SubItems(12)
.Fields(11) = l.SelectedItem.SubItems(13)
.Fields(12) = l.SelectedItem.SubItems(14)
.Fields(13) = l.SelectedItem.SubItems(15)
.Update
End With
End If
Next
End If
Command1_Click
Exit Sub
End If

If List1 = "[Start?]" Then
Winsock1(0).RemoteHost = Winsock1(0).RemoteHost
Winsock1(0).RemotePort = Winsock1(0).RemotePort
Winsock1(0).SendData "Yes?OK."
End If


If List1 = "[Log in]" Then
If List1.ListCount < 3 Then
Winsock1(0).RemoteHost = Winsock1(0).RemoteHost
Winsock1(0).RemotePort = Winsock1(0).RemotePort
Winsock1(0).SendData "Invalid logged in"
Exit Sub
Else

List1.ListIndex = 1
clm = List1
List1.ListIndex = 2
cln = List1
List1.ListIndex = 3
tm = List1
List1.ListIndex = 4
clid = List1

Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS2 = New Recordset
adoPrimaryRS2.Open "select * from member where [member id]='" & clid & "'", db, adOpenStatic, adLockOptimistic
CLTYPE = adoPrimaryRS2.Fields("Group Name")
rate = adoPrimaryRS2.Fields("rate")
discount = adoPrimaryRS2.Fields("discount")
mintime = adoPrimaryRS2.Fields("Munimum Time")
If adoPrimaryRS2.Fields("member name") <> cln Then
Exit Sub
End If

Winsock1(0).RemoteHost = Winsock1(0).RemoteHost
Winsock1(0).RemotePort = Winsock1(0).RemotePort
Winsock1(0).SendData "OK"
List1.ListIndex = 0
End If
End If

If List1 = "[Log in]" Then
List1.ListIndex = 1
clm = List1
List1.ListIndex = 2
cln = List1
List1.ListIndex = 3
tm = List1
List1.ListIndex = 4
clid = List1

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
If l.SelectedItem.SubItems(1) = clm Then
l.SelectedItem.SubItems(4) = "In Use"
l.SelectedItem.SubItems(3) = cln
l.SelectedItem.SubItems(5) = tm
l.SelectedItem.SubItems(6) = ""
l.SelectedItem.SubItems(7) = ""
l.SelectedItem.SubItems(8) = CLTYPE
l.SelectedItem.SubItems(9) = clid
l.SelectedItem.SubItems(10) = rate
l.SelectedItem.SubItems(11) = discount
l.SelectedItem.SubItems(12) = mintime
End If
Next
End If
Exit Sub
End If


If List1 = "[Time]" Then
List1.ListIndex = 1
clm = List1
List1.ListIndex = 2
tm = List1
If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
If l.SelectedItem.SubItems(1) = clm Then
l.SelectedItem.SubItems(4) = "In Use"
'l.SelectedItem.SubItems(3) = cln
l.SelectedItem.SubItems(7) = tm
End If
Next
End If
Exit Sub
End If


End Sub

Private Sub wot1_ConnectionRequest(ByVal requestID As Long)
wot1.Close
wot1.Accept requestID
console.Text = console.Text & vbCrLf & wot1.RemoteHostIP + ":" + wot1.Tag + " Connection Attempted! Using port 80. Remote port: " & wot1.RemotePort & "." + vbCrLf
End Sub

Private Sub wot1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim s As String
If Len(console.Text) > 1024 * 10 Then console.Text = ""
wot1.GetData s, vbString
console.Text = console.Text & vbCrLf & s
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
