VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Mname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Master"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Detect"
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5220
      TabIndex        =   6
      Top             =   4425
      Width           =   5220
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         Height          =   300
         Left            =   3360
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   300
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   300
         Left            =   1200
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      MaxLength       =   15
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      MaxLength       =   255
      TabIndex        =   4
      Top             =   3960
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   4935
      TabIndex        =   1
      Top             =   3480
      Width           =   4935
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   1800
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   1275
      End
   End
   Begin MSComctlLib.ListView l 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
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
      NumItems        =   3
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
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1920
      Y1              =   3840
      Y2              =   4320
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   3480
      Width           =   4935
   End
End
Attribute VB_Name = "Mname"
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
If IsNumeric(left(Text1.Text, 1)) Then
MsgBox "Invalid Machine Name or IP address.", vbExclamation + vbOKOnly, "Invalid"
Text1.SetFocus
Exit Sub
End If
mn = Trim(Text1.Text)
For i = 1 To Len(mn)
If Mid(mn, i, 1) = " " Then
MsgBox "Invalid Machine Name or IP address.", vbExclamation + vbOKOnly, "Invalid"
Text1.SetFocus
Exit Sub
End If
Next
If Len(Trim(Text2.Text)) = 0 Then
MsgBox "Invalid Machine Name or IP address.", vbExclamation + vbOKOnly, "Invalid"
Text2.SetFocus
Exit Sub
End If
If Len(Trim(Text1.Text)) = 0 Then
MsgBox "Invalid Machine Name or IP address.", vbExclamation + vbOKOnly, "Invalid"
Text1.SetFocus
Exit Sub
End If

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If

Do Until adoPrimaryRS.EOF
If Trim(adoPrimaryRS.Fields("machine")) = Trim(Text1.Text) Then
MsgBox "Machine Name allready exist .", vbExclamation + vbOKOnly, "Invalid"
Text1.SetFocus
Exit Sub
End If
If Trim(adoPrimaryRS.Fields("ip")) = Trim(Text2.Text) Then
MsgBox "IP address allready exist .", vbExclamation + vbOKOnly, "Invalid"
Text2.SetFocus
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
.Fields(1) = Text1.Text
.Fields(2) = Text2.Text
.Update
End With

Set x = l.ListItems.Add(, , sl)
x.SubItems(1) = Text1.Text
x.SubItems(2) = Text2.Text

End Sub

Private Sub Command2_Click()
If l.ListItems.Count > 0 Then
l.ListItems.Remove (l.SelectedItem.Index)
End If
updatedata
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
frm_search_machine.Show
End Sub

Private Sub command5_Click()
Text1.Text = ""
If Len(Text2.Text) > 0 Then
WinsockInit
Text1.Text = HostByAddress(Text2.Text)
WSACleanUp
End If
End Sub

Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select sl,machine,ip from ip", db, adOpenStatic, adLockOptimistic

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If

Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
adoPrimaryRS.MoveNext
Loop

End Sub

Private Sub updatedata()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select sl,machine,ip from ip", db, adOpenStatic, adLockOptimistic
db.Execute "delete from ip"
If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
With adoPrimaryRS
.AddNew
.Fields(0) = i
.Fields(1) = l.SelectedItem.SubItems(1)
.Fields(2) = l.SelectedItem.SubItems(2)
.Update
End With
Next
End If
adoPrimaryRS.Requery
l.ListItems.Clear
If adoPrimaryRS.RecordCount > 0 Then adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
adoPrimaryRS.MoveNext
Loop

End Sub
