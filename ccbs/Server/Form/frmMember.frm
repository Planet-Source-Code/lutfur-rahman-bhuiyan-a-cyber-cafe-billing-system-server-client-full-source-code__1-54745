VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\..\..\..\..\Documents and Settings\Administrator.OMEGA\Desktop\ccbs\Server\calendar\Stevemac\VB\Controls\CalendarVB\CalendarVB.vbp"
Begin VB.Form frmMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Master"
   ClientHeight    =   5955
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin ctrCalendarVB.CalendarVB MonthView1 
      Height          =   2535
      Left            =   3360
      TabIndex        =   39
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      LineStyle       =   2
      DayHeaderBackColor=   16761024
      PrePeriodBackColor=   16761024
      PostPeriodBackColor=   16761024
      ActiveDayFontBold=   0   'False
      ActiveDayFontItalic=   0   'False
      ActiveDayFontSize=   8.25
      ActiveDayFontName=   "MS Sans Serif"
      DayHeaderFontBold=   0   'False
      DayHeaderFontItalic=   0   'False
      DayHeaderFontSize=   8.25
      DayHeaderFontName=   "MS Sans Serif"
      DaysFontBold    =   0   'False
      DaysFontItalic  =   0   'False
      DaysFontSize    =   8.25
      DaysFontName    =   "MS Sans Serif"
   End
   Begin VB.TextBox txtFields 
      DataField       =   "prn yn"
      Height          =   285
      Index           =   7
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Mr"
      Height          =   285
      Index           =   6
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   36
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3120
      TabIndex        =   34
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Join Date"
      Height          =   285
      Index           =   5
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      Height          =   645
      Index           =   4
      Left            =   2040
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Amount"
      Height          =   285
      Index           =   3
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin MSComctlLib.ListView l 
      Height          =   1815
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3201
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
   End
   Begin VB.ComboBox c2 
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1935
      Width           =   1455
   End
   Begin VB.ComboBox c1 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1935
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox c 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7125
      TabIndex        =   18
      Top             =   5355
      Width           =   7125
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   4680
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5880
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7125
      TabIndex        =   12
      Top             =   5655
      Width           =   7125
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4560
         Picture         =   "frmMember.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmMember.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmMember.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmMember.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   17
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member Name"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   2
      Top             =   495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member ID"
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MR No."
      Height          =   195
      Left            =   3600
      TabIndex        =   35
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Join Date:"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   33
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   32
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   31
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Existing Client:"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   28
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "User Group:"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   27
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   26
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Password Require?"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1545
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client Name:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   495
      Width           =   900
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client ID:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   645
   End
End
Attribute VB_Name = "frmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS1 As Recordset
Attribute adoPrimaryRS1.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS6 As Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Public acmode As Boolean

Private Sub c_Click()
If c.ListIndex = 0 Then
txtFields(2).Locked = False
Else
txtFields(2).Locked = True
End If

End Sub

Private Sub c1_Click()
If acmode = True And C1.ListCount > 0 Then
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS1 = New Recordset
sqlstr = "Select * from [User Group] where [Group Name]='" & C1.Text & "'"
adoPrimaryRS1.Open sqlstr, db, adOpenStatic, adLockOptimistic
C2.ListIndex = adoPrimaryRS1.Fields("type")
End If
End Sub

Private Sub Calendario1_GotFocus()

End Sub

Private Sub cmdPrint_Click()
adoPrimaryRS.Fields("prn yn") = "Y"
adoPrimaryRS.Update
If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
      For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS6 = New Recordset
adoPrimaryRS6.Open "select Name,Address,Email,Phone,Web from Company", db, adOpenStatic, adLockOptimistic

Open App.Path + "\report\voucher.htm" For Output As #1
Print #1, "<html>"
Print #1, "<head>"
Print #1, "<meta http-equiv=""Content-Language"" content=""en-us"">"
Print #1, "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">"
Print #1, "<meta name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">"
Print #1, "<meta name=""ProgId"" content=""FrontPage.Editor.Document"">"
Print #1, "<title>OMEGA COMPUTERS</title>"
Print #1, "</head>"
Print #1, "<body>"
Print #1, "<p align=""center""><font face=""Arial Black"" size=""5"">Mony Receipt</font></p>"
Print #1, "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
Print #1, "<tr>"
Print #1, "<td width=""100%""><b><font face=""Arial Black"" size=""5"">" & adoPrimaryRS6.Fields(0) & "</font></b></td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""100%""><b><font size=""3"" face=""MS Serif"">" & adoPrimaryRS6.Fields(1) & "</font></b></td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""100%""><b><font size=""3"" face=""MS Serif"">Phine:" & adoPrimaryRS6.Fields(2) & "</font></b></td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""100%""><b><font size=""3"" face=""MS Serif"">Email:" & adoPrimaryRS6.Fields(3) & "</font></b></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<p>Date: " & adoPrimaryRS.Fields("join date") & "</p>"
Print #1, "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
Print #1, "<tr>"
Print #1, "<p>Recipt No.: " & adoPrimaryRS.Fields("Mr") & "</p>"
Print #1, "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Client ID:</b></td>"
Print #1, "<td width=""80%"">" & adoPrimaryRS.Fields("member id") & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Cient Name:</b></td>"
Print #1, "<td width=""80%"">" & adoPrimaryRS.Fields("member name") & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Address:</b></td>"
Print #1, "<td width=""80%"">" & adoPrimaryRS.Fields("address") & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td Width=""20%""><b>Amount:</b></td>"
Print #1, "<td width=""80%"">" & adoPrimaryRS.Fields("amount") & "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<p>&nbsp;</p>"
Print #1, "<p>Singnature:</p>"
Print #1, "</body>"
Print #1, "</html>"
Close #1

frmBrowser.brwWebBrowser.Navigate App.Path + "\report\voucher.htm"
frmBrowser.Show

End Sub

Private Sub Command1_Click()
MonthView1.Visible = True
End Sub

Private Sub Form_Activate()
acmode = True
End Sub

Private Sub Form_Load()
 acmode = False
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from Member", db, adOpenStatic, adLockOptimistic
  
  Set adoPrimaryRS1 = New Recordset
  adoPrimaryRS1.Open "select [Group ID],[Group Name],[Munimum Time],Rate,Discount,type from [User Group]", db, adOpenStatic, adLockOptimistic
  
  If adoPrimaryRS1.RecordCount > 0 Then adoPrimaryRS1.MoveFirst
  Do Until adoPrimaryRS1.EOF
  C1.AddItem adoPrimaryRS1.Fields("Group Name")
  adoPrimaryRS1.MoveNext
  Loop
    
  If C1.ListCount > 0 Then
  C1.ListIndex = 0
  End If
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
c.AddItem "Yes"
c.AddItem "No"
c.ListIndex = 0

C2.AddItem "Regular"
C2.AddItem "Temporary"
C2.ListIndex = 0
addlist
mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.left = lblStatus.Width + 700
  cmdLast.left = cmdNext.left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
 On Error Resume Next
 
  If adoPrimaryRS.RecordCount > 0 Then
  adoPrimaryRS.MoveFirst
  End If
  gid = 0
  
   For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
   
  mr = 0
    
  Do Until adoPrimaryRS.EOF
  If Val(adoPrimaryRS.Fields("Member ID")) > gid Then gid = Val(adoPrimaryRS.Fields("Member ID"))
  adoPrimaryRS.MoveNext
  Loop
  
  
  If adoPrimaryRS.RecordCount > 0 Then
  adoPrimaryRS.MoveFirst
  End If
  
  Do Until adoPrimaryRS.EOF
  If Val(adoPrimaryRS.Fields("Mr")) > mr And IsNumeric(adoPrimaryRS.Fields("Mr")) Then mr = Val(adoPrimaryRS.Fields("mr"))
  adoPrimaryRS.MoveNext
  Loop
  
  mr = mr + 1
  gid = gid + 1
  If adoPrimaryRS.RecordCount > 0 Then
  adoPrimaryRS.MoveFirst
  End If
    
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = gid
    txtFields(6) = mr
    txtFields(7) = "N"
    txtFields(1).SetFocus
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error Resume Next
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error Resume Next

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  Call addlist
End Sub

Private Sub cmdUpdate_Click()

On Error Resume Next

If Len(Trim(txtFields(3))) = 0 Then
txtFields(3).SetFocus
Exit Sub
End If

If Not IsDate(txtFields(5)) Then
txtFields(5).SetFocus
Exit Sub
End If

If Len(Trim(txtFields(1))) = 0 Then
txtFields(1).SetFocus
Exit Sub
End If



If c.ListIndex = 0 And Len(Trim(txtFields(2))) = 0 Then
txtFields(2).SetFocus
Exit Sub
End If

If C1.ListCount = 0 Then Exit Sub

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS1 = New Recordset
sqlstr = "Select * from [User Group] where [Group Name]='" & C1.Text & "'"
adoPrimaryRS1.Open sqlstr, db, adOpenStatic, adLockOptimistic
C2.ListIndex = adoPrimaryRS1.Fields("type")

adoPrimaryRS.Fields("type") = adoPrimaryRS1.Fields("type")
adoPrimaryRS.Fields("Group ID") = adoPrimaryRS1.Fields("Group ID")
adoPrimaryRS.Fields("Group Name") = adoPrimaryRS1.Fields("Group Name")
adoPrimaryRS.Fields("Munimum Time") = adoPrimaryRS1.Fields("Munimum Time")
adoPrimaryRS.Fields("Rate") = adoPrimaryRS1.Fields("Rate")
adoPrimaryRS.Fields("Discount") = adoPrimaryRS1.Fields("Discount")
adoPrimaryRS.Fields("Pass Req") = c.ListIndex
'adoPrimaryRS.Fields("min time") = c.ListIndex
adoPrimaryRS.UpdateBatch adAffectAll

If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
addlist
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error Resume Next

  adoPrimaryRS.MoveFirst
  c.ListIndex = adoPrimaryRS.Fields("Pass Req")
  C2.ListIndex = adoPrimaryRS.Fields("type")
  C1.Text = adoPrimaryRS.Fields(1)
  If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next

  adoPrimaryRS.MoveLast
  c.ListIndex = adoPrimaryRS.Fields("Pass Req")
  C2.ListIndex = adoPrimaryRS.Fields("type")
  C1.Text = adoPrimaryRS.Fields(1)
  If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next

  If Not adoPrimaryRS.EOF Then
  adoPrimaryRS.MoveNext
  c.ListIndex = adoPrimaryRS.Fields("Pass Req")
  C2.ListIndex = adoPrimaryRS.Fields("type")
  C1.Text = adoPrimaryRS.Fields(1)
  
  If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  End If
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
    c.ListIndex = adoPrimaryRS.Fields("Pass Req")
    C2.ListIndex = adoPrimaryRS.Fields("type")
    C1.Text = adoPrimaryRS.Fields(1)
    
    If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
    End If
    
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error Resume Next

  If Not adoPrimaryRS.BOF Then
  adoPrimaryRS.MovePrevious
  c.ListIndex = adoPrimaryRS.Fields("Pass Req")
  C2.ListIndex = adoPrimaryRS.Fields("type")
  C1.Text = adoPrimaryRS.Fields(1)
  If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    
   For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
   
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
  For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  End If
  
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
    c.ListIndex = adoPrimaryRS.Fields("Pass Req")
    C2.ListIndex = adoPrimaryRS.Fields("type")
    C1.Text = adoPrimaryRS.Fields(1)
    If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub XPCalendar1_Change()
txtFields(5).Text = XPCalendar1.Value
End Sub

Private Sub XPCalendar1_Click()
txtFields(5).Text = XPCalendar1.Value
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
adoPrimaryRS.MoveNext
Loop
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveLast
If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
      For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
      For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
End If
End Sub

Private Sub l_Click()
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If
Do Until adoPrimaryRS.EOF
'MsgBox l.SelectedItem.Text
If Trim(adoPrimaryRS.Fields("Member ID")) = Trim(l.SelectedItem.Text) Then
If txtFields(7) = "Y" Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    
   For i = 0 To txtFields.UBound
   txtFields(i).Enabled = False
   Next
   
    Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
  For i = 0 To txtFields.UBound
   txtFields(i).Enabled = True
   Next
  End If
Exit Sub
End If
adoPrimaryRS.MoveNext
Loop
End Sub

Private Sub MonthView1_DblClick(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
MonthView1.Visible = False
txtFields(5).Text = MonthView1.DateValue
End Sub
