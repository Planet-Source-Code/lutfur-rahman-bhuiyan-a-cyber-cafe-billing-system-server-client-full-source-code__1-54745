VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item"
   ClientHeight    =   4665
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView l 
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2778
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Opening Balance"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Opening Price"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7815
      TabIndex        =   15
      Top             =   4065
      Width           =   7815
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   16
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
      ScaleWidth      =   7815
      TabIndex        =   9
      Top             =   4365
      Width           =   7815
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmItem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmItem.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmItem.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmItem.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   14
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "opening price"
      Height          =   285
      Index           =   3
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1620
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "opening balance"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1305
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "item name"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   2
      Top             =   975
      Width           =   5415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "item id"
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Available Item:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category Name:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category ID"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Opening Price:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Opening Balance:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Iitem Id:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   1815
   End
End
Attribute VB_Name = "frmItem"
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
Public act As Boolean
Dim db As Connection
Private Sub Combo2_Click()
Combo1.ListIndex = Combo2.ListIndex
'If act = True Then
'Dim db As Connection
'Set db = New Connection
'db.CursorLocation = adUseClient
'db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\cyber.mdb;"
'Dim STR
'STR = "Select * from item where [category ID]='" & Combo1 & " '"
'Set adoPrimaryRS = New Recordset
'adoPrimaryRS.Open "Select * from item where [category ID]='" & Combo1 & " '", db, adOpenStatic, adLockOptimistic
'End If
End Sub

Private Sub Combo2_LostFocus()
Call listadd
End Sub

Private Sub Form_Activate()
act = True
End Sub

Private Sub Form_Load()
act = False
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\cyber.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from Item", db, adOpenStatic, adLockOptimistic
  
  Set adoPrimaryRS1 = New Recordset
  adoPrimaryRS1.Open "select [category ID],[Category Name] from Category", db, adOpenStatic, adLockOptimistic


  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  
  If adoPrimaryRS1.RecordCount > 0 Then adoPrimaryRS1.MoveFirst
  Do Until adoPrimaryRS1.EOF
  Combo1.AddItem adoPrimaryRS1.Fields(0)
  Combo2.AddItem adoPrimaryRS1.Fields(1)
  adoPrimaryRS1.MoveNext
  Loop
  
  If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
  If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
  
  Call listadd

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
'  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
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
Do Until adoPrimaryRS.EOF
If adoPrimaryRS.Fields("item id") > catid And IsNumeric(adoPrimaryRS.Fields("item id")) Then
catid = adoPrimaryRS.Fields(0)
End If
adoPrimaryRS.MoveNext
Loop

  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = catid + 1
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
    Call listadd
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

End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next

If Len(Trim(txtFields(1))) = 0 Then
txtFields(1).SetFocus
Exit Sub
End If
adoPrimaryRS.Fields(0) = Combo1
adoPrimaryRS.Fields(1) = Combo2

adoPrimaryRS.UpdateBatch adAffectAll
Call listadd
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
Call listadd
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
  Combo2 = adoPrimaryRS.Fields(1)
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next

  adoPrimaryRS.MoveLast
  Combo2 = adoPrimaryRS.Fields(1)
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next

  If Not adoPrimaryRS.EOF Then
  adoPrimaryRS.MoveNext
  Combo2 = adoPrimaryRS.Fields(1)
  End If
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
    Combo2 = adoPrimaryRS.Fields(1)
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
  Combo2 = adoPrimaryRS.Fields(1)
  End If
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
    Combo2 = adoPrimaryRS.Fields(1)
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

Private Sub l_Click()
On Error Resume Next
Combo1 = l.SelectedItem.Text
Combo2 = l.SelectedItem.SubItems(1)
txtFields(0).Text = l.SelectedItem.SubItems(2)
txtFields(1).Text = l.SelectedItem.SubItems(3)
txtFields(2).Text = l.SelectedItem.SubItems(4)
txtFields(3).Text = l.SelectedItem.SubItems(5)
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 2, 3
            Dim keyResponse As Boolean
            keyResponse = CtrlValidate(KeyAscii, "0123456789.")
            If keyResponse = True Then
            Else
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub listadd()

On Error Resume Next
l.ListItems.Clear
If adoPrimaryRS.RecordCount > 0 Then adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
'MsgBox adoPrimaryRS.Fields("category ID") & Combo1
'If adoPrimaryRS.Fields("category ID") = Combo1 Then
If Not IsNull(adoPrimaryRS.Fields(0)) Then
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
If Not IsNull(adoPrimaryRS.Fields(1)) Then
x.SubItems(1) = adoPrimaryRS.Fields(1)
End If
If Not IsNull(adoPrimaryRS.Fields(2)) Then
x.SubItems(2) = adoPrimaryRS.Fields(2)
End If
If Not IsNull(adoPrimaryRS.Fields(3)) Then
x.SubItems(3) = adoPrimaryRS.Fields(3)
End If
If Not IsNull(adoPrimaryRS.Fields(4)) Then
x.SubItems(4) = adoPrimaryRS.Fields(4)
End If
If Not IsNull(adoPrimaryRS.Fields(5)) Then
x.SubItems(5) = adoPrimaryRS.Fields(5)
'End If
End If
End If
adoPrimaryRS.MoveNext
Loop

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveLast
If Not IsNull(adoPrimaryRS.Fields(0)) Then
Combo1 = adoPrimaryRS.Fields(0)
End If
If Not IsNull(adoPrimaryRS.Fields(1)) Then
Combo2 = adoPrimaryRS.Fields(1)
End If
If Not IsNull(adoPrimaryRS.Fields(2)) Then
txtFields(0).Text = adoPrimaryRS.Fields(2)
End If
If Not IsNull(adoPrimaryRS.Fields(3)) Then
txtFields(1).Text = adoPrimaryRS.Fields(3)
End If
If Not IsNull(adoPrimaryRS.Fields(4)) Then
txtFields(2).Text = adoPrimaryRS.Fields(4)
End If
If Not IsNull(adoPrimaryRS.Fields(5)) Then
txtFields(3).Text = adoPrimaryRS.Fields(5)
End If

End If

End Sub
