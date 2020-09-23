VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_search_machine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Search"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   3480
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock w 
      Left            =   4800
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5430
      TabIndex        =   5
      Top             =   2850
      Width           =   5430
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   300
         Left            =   3480
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   300
         Left            =   2640
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   300
         Left            =   4440
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ListView l 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Host Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "to:"
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start from:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frm_search_machine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public srec As Integer
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean


Private Sub Command1_Click()

Call extract(Text1, List1)

If List1.ListCount < 4 Then
MsgBox "Invalid IP address (1)", vbOKOnly + vbExclamation, "Invalid"
Text1.SetFocus
Exit Sub
End If

For i = 0 To 3
List1.ListIndex = i
If List1 < 1 And List1 > 255 Then
Text1.SetFocus
MsgBox "Invalid IP address (2)", vbOKOnly + vbExclamation, "Invalid"
Exit Sub
End If
Next

Call extract(Text2, List2)

If List2.ListCount < 4 Then
MsgBox "Invalid IP address", vbOKOnly + vbExclamation, "Invalid"
Text2.SetFocus
Exit Sub
End If

For i = 0 To 3
List2.ListIndex = i
If List2 < 1 Or List2 > 255 Then
Text2.SetFocus
MsgBox "Invalid IP address", vbOKOnly + vbExclamation, "Invalid"
Exit Sub
End If
Next

Dim minip As String
Dim maxip As String
Dim minlip As String
Dim maxlip As String

List1.ListIndex = 0
minip = List1
List1.ListIndex = 1
minip = minip + "." + List1
List1.ListIndex = 2
minip = minip + "." + List1
List1.ListIndex = 3
minlip = List1

List2.ListIndex = 3
maxlip = List2

If minlip > maxlip Then
MsgBox "Invalid IP address range.", vbOKOnly + vbExclamation, "Invalid"
Text1.SetFocus
Exit Sub
End If
srec = 0

Label4.Caption = ""

Dim sendmass As Integer
l.ListItems.Clear
If Command1.Caption = "Start" Then
Command2.Visible = True
Command1.Visible = False
DoEvents
For i = minlip To maxlip
If srec = 1 Then
srec = 0
Exit For
Exit Sub
End If

ipfind = minip + "." + Trim(Str(i))
Label3.Caption = "Finding IP: " & ipfind
Host = HostByAddress(ipfind)
'Host = getHost.GetHostNameFromIP(ipfind)
If Len(Trim(Host)) > 0 Then
Set x = l.ListItems.Add(, , ipfind)
x.SubItems(1) = Host
Label4.Caption = "Found Host: " & l.ListItems.Count
End If
DoEvents
Next
End If

srec = 1
Command1.Visible = True
Command2.Visible = False

End Sub

Private Sub Command2_Click()
srec = 1
Command1.Visible = True
Command2.Visible = False
End Sub

Private Sub Command3_Click()
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If

Do Until adoPrimaryRS.EOF
If Trim(adoPrimaryRS.Fields("machine")) = l.SelectedItem.SubItems(1) Then
MsgBox "Machine Name allready exist .", vbExclamation + vbOKOnly, "Invalid"
Exit Sub
End If

If Trim(adoPrimaryRS.Fields("ip")) = l.SelectedItem.Text Then
MsgBox "IP address allready exist .", vbExclamation + vbOKOnly, "Invalid"
Exit Sub
End If
adoPrimaryRS.MoveNext
Loop

sl = 0
If adoPrimaryRS.RecordCount = 0 Then
sl = 1
Else
sl = adoPrimaryRS.RecordCount + 1
End If

With adoPrimaryRS
.AddNew
.Fields(0) = sl
.Fields(1) = l.SelectedItem.SubItems(1)
.Fields(2) = l.SelectedItem.Text
.Update
End With

Set x = Mname.l.ListItems.Add(, , sl)
x.SubItems(1) = l.SelectedItem.SubItems(1)
x.SubItems(2) = l.SelectedItem.Text
End Sub

Private Sub Command4_Click()
srec = 1
Unload Me
End Sub

Private Sub Form_Load()

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select sl,machine,ip from ip", db, adOpenStatic, adLockOptimistic

Command2.Visible = False
Text1.Text = w.LocalIP
Call extract(Text1, List1)

If List1.ListCount < 4 Then
MsgBox "Invalid IP address", vbOKOnly + vbExclamation, "Invalid"
Text1.SetFocus
Exit Sub
End If

For i = 0 To 3
List1.ListIndex = i
If List1 < 1 And List1 > 255 Then
Text1.SetFocus
MsgBox "Invalid IP address", vbOKOnly + vbExclamation, "Invalid"
Exit Sub
End If
Next

Dim minip As String
Dim maxip As String

List1.ListIndex = 0
minip = List1
List1.ListIndex = 1
minip = minip + "." + List1
List1.ListIndex = 2
minip = minip + "." + List1
List1.ListIndex = 3
maxip = minip + ".255"
minip = minip + "." + "1"
Text1.Text = minip
Text2.Text = maxip
End Sub

Private Sub extract(txtbox As TextBox, lb As ListBox)
lb.Clear
helloa = Len(txtbox)
For i = 1 To helloa
helloc = helloc + 1
hellob = Mid(txtbox, i, 1)
If hellob = "." Then
hellob = ""
hellod = hellod + hellob
If hellod = "" Then
Else
lb.AddItem hellod
End If
hellod = ""
Else
hellod = hellod + hellob
If helloc = helloa Then lb.AddItem hellod
End If
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
srec = 1
End Sub

