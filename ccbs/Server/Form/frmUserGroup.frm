VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group Master"
   ClientHeight    =   4605
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView l 
      Height          =   1335
      Left            =   120
      TabIndex        =   30
      Top             =   2520
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2355
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Group ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Group Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tyoe"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox c 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1800
      Width           =   2535
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5940
      TabIndex        =   17
      Top             =   4005
      Width           =   5940
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   18
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
      ScaleWidth      =   5940
      TabIndex        =   11
      Top             =   4305
      Width           =   5940
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmUserGroup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmUserGroup.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   360
         Picture         =   "frmUserGroup.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmUserGroup.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   16
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Discount"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1455
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Rate"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1140
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Munimum Time"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   825
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Group Name"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   495
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Group ID"
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Existing Group:"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Index           =   8
      Left            =   2880
      TabIndex        =   27
      Top             =   1440
      Width           =   120
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Per Minute"
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   26
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Minute"
      Height          =   195
      Index           =   6
      Left            =   2880
      TabIndex        =   25
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1785
      Width           =   405
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Discount:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1455
      Width           =   675
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Rate:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Minimum Time:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   825
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Group Name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   945
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Group ID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   690
   End
End
Attribute VB_Name = "frmUserGroup"
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

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\cyber.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select [Group ID],[Group Name],[Munimum Time],Rate,Discount,type from [User Group]", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  c.AddItem "Regular"
  c.AddItem "Temporary"
  c.ListIndex = 0
  Call addlist
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
  On Error GoTo AddErr
  If adoPrimaryRS.RecordCount > 0 Then
  adoPrimaryRS.MoveFirst
  End If
  gid = 0
  Do Until adoPrimaryRS.EOF
  If adoPrimaryRS.Fields(0) > gid Then gid = adoPrimaryRS.Fields(0)
  adoPrimaryRS.MoveNext
  Loop
  gid = gid + 1
  If adoPrimaryRS.RecordCount > 0 Then
  adoPrimaryRS.MoveFirst
  End If
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
  txtFields(0).Text = gid
  txtFields(1).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  With adoPrimaryRS
    .Delete
    Call addlist
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

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

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  If l.ListItems.Count > 0 Then
  For i = 1 To l.ListItems.Count
  l.ListItems(i).Selected = True
  If UCase(l.SelectedItem.SubItems(1)) = UCase(txtFields(1)) Then Exit Sub
  Next
  End If
  
  adoPrimaryRS.Fields("type") = c.ListIndex
  adoPrimaryRS.UpdateBatch adAffectAll
  Call addlist
  
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
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
  c.ListIndex = adoPrimaryRS.Fields("type")
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next

  adoPrimaryRS.MoveLast
  c.ListIndex = adoPrimaryRS.Fields("type")
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next

  If Not adoPrimaryRS.EOF Then
  adoPrimaryRS.MoveNext
  c.ListIndex = adoPrimaryRS.Fields("type")
  End If
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
    c.ListIndex = adoPrimaryRS.Fields("type")
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
  c.ListIndex = adoPrimaryRS.Fields("type")
  End If
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
    c.ListIndex = adoPrimaryRS.Fields("type")
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


Private Sub addlist()
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
End If
l.ListItems.Clear
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
If adoPrimaryRS.Fields("type") = 0 Then
x.SubItems(2) = "Regular"
Else
x.SubItems(2) = "Temporary"
End If
adoPrimaryRS.MoveNext
Loop
If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveLast
c.ListIndex = adoPrimaryRS.Fields("type")
End If

End Sub
