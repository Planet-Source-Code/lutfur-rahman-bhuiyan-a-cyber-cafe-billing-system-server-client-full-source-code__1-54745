VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\A..\..\..\..\..\..\..\DOCUME~1\ADMINI~1.OME\Desktop\Desk\ccbs\Server\calendar\Stevemac\VB\Controls\CalendarVB\CalendarVB.vbp"
Begin VB.Form frmSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales"
   ClientHeight    =   6765
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   255
      Left            =   2520
      TabIndex        =   38
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtDoc 
      Height          =   285
      Left            =   2880
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   5520
      Width           =   1455
   End
   Begin ctrCalendarVB.CalendarVB MonthView1 
      Height          =   2535
      Left            =   2760
      TabIndex        =   27
      Top             =   600
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
   Begin MSComctlLib.ListView l 
      Height          =   2295
      Left            =   0
      TabIndex        =   34
      Top             =   3120
      Width           =   10335
      _ExtentX        =   18230
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CatID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Name"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10335
      TabIndex        =   29
      Top             =   1800
      Width           =   10335
      Begin VB.Line Line3 
         X1              =   8520
         X2              =   8520
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   10335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price:"
         Height          =   195
         Index           =   9
         Left            =   8760
         TabIndex        =   33
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   8
         Left            =   7080
         TabIndex        =   32
         Top             =   120
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Nme:"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   31
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name:"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   120
         Width           =   1140
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   3120
      TabIndex        =   28
      Top             =   855
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox C4 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   2460
      Width           =   3135
   End
   Begin VB.ComboBox C3 
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   6795
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox C2 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2460
      Width           =   3015
   End
   Begin VB.ComboBox C1 
      Height          =   315
      Left            =   1800
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10590
      TabIndex        =   20
      Top             =   6165
      Width           =   10590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Posting"
         Height          =   300
         Left            =   59
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   1200
         TabIndex        =   21
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
      ScaleWidth      =   10590
      TabIndex        =   14
      Top             =   6465
      Width           =   10590
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmSales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmSales.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmSales.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmSales.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   19
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   9
      Left            =   8640
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2460
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   8
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   6
      Top             =   2460
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1185
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   2
      Top             =   900
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   0
      Top             =   225
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Price:"
      Height          =   195
      Left            =   7560
      TabIndex        =   35
      Top             =   5520
      Width           =   810
   End
   Begin VB.Line Line6 
      X1              =   8520
      X2              =   8520
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Line Line5 
      X1              =   6720
      X2              =   6720
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3360
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   0
      Top             =   2280
      Width           =   10335
   End
   Begin VB.Label lblLabels 
      Caption         =   "item id:"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Top             =   6780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client Name:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   1185
      Width           =   900
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client ID:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   900
      Width           =   645
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Order No.:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   225
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "category ID:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   7260
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmSales"
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
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Public aFlag As Boolean, eFlag As Boolean

Private Sub C2_Click()

C1.ListIndex = C2.ListIndex
C4.Clear

If adoPrimaryRS1.RecordCount > 0 Then
adoPrimaryRS1.MoveFirst
End If

Do Until adoPrimaryRS1.EOF
If adoPrimaryRS1.Fields("Category ID") = C1 Then
C4.AddItem adoPrimaryRS1.Fields(3)
End If
adoPrimaryRS1.MoveNext
Loop

If C4.ListCount > 0 Then
C4.ListIndex = 0
End If

End Sub

Private Sub C4_Change()
'C1.ListIndex = C4.ListIndex
'C3.ListIndex = C4.ListIndex
'C2.ListIndex = C4.ListIndex
C3.ListIndex = C4.ListIndex
End Sub

Private Sub Command1_Click()
MonthView1.Visible = True
End Sub

Private Sub Command2_Click()
frmmemberlist.Show
End Sub

Private Sub Command3_Click()

If Len(Trim(txtFields(2))) = 0 Then
txtFields(2).SetFocus
Exit Sub
End If

If adoPrimaryRS2.RecordCount > 0 Then
adoPrimaryRS2.MoveFirst
End If

Dim found As Boolean
found = False
Do Until adoPrimaryRS2.EOF
If adoPrimaryRS2.Fields("member id") = txtFields(4) Then
found = True
txtFields(5) = adoPrimaryRS2.Fields("member name")
End If
adoPrimaryRS2.MoveNext
Loop

If found = False Then
frmmemberlist.Show
Exit Sub
End If


If Not IsNumeric(txtFields(8).Text) Then
txtFields(8).SetFocus
Exit Sub
End If

If Not IsNumeric(txtFields(9).Text) Then
txtFields(9).SetFocus
Exit Sub
End If

If aFlag = True Or eFlag = True Then
Set x = l.ListItems.Add(, , C1)
x.SubItems(1) = C2
x.SubItems(2) = C3
x.SubItems(3) = C4
x.SubItems(4) = txtFields(8)
x.SubItems(5) = txtFields(9)
tot = 0
If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot
C2.SetFocus
End If

End Sub

Private Sub Command4_Click()
If aFlag = True Or eFlag = True Then
If l.ListItems.Count > 0 Then
l.ListItems.Remove (l.SelectedItem.Index)
End If
tot = 0
If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot
End If
End Sub

Private Sub command5_Click()
frmOpenOrder.Show
End Sub

Private Sub Form_Load()

aFlag = False
eFlag = False

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from Item", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS2 = New Recordset
adoPrimaryRS2.Open "select * from [member]", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from [Order No]", db, adOpenStatic, adLockOptimistic

Set adoPrimaryRS4 = New Recordset
adoPrimaryRS4.Open "select * from Sales", db, adOpenStatic, adLockOptimistic

If adoPrimaryRS3.RecordCount > 0 Then
adoPrimaryRS3.MoveFirst
If Not IsNull(adoPrimaryRS3.Fields("doc")) Then
txtDoc = adoPrimaryRS3.Fields("doc")
End If
End If

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order] where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
If l.ListItems.Count > 0 Then
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)
End If

mbDataChanged = False
End Sub

Private Sub Form_Resize()
  'on error Resume Next
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

Private Sub adoPrimaryRS3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS3.AbsolutePosition)
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
  'on error GoTo AddErr
  'vn = 0
  'If adoPrimaryRS3.RecordCount > 0 Then
  'adoPrimaryRS3.MoveFirst
  'End If
  'Do Until adoPrimaryRS3.EOF
  'If adoPrimaryRS.Fields(0) > vn Then
  'vn = adoPrimaryRS3.Fields(3)
  'End If
  'adoPrimaryRS3.EOF
  'Loop
txtFields(2).Text = ""
txtFields(4).Text = ""
txtFields(5).Text = ""
txtFields(8).Text = ""
txtFields(9).Text = ""

C1.Clear
C2.Clear
C3.Clear
C4.Clear

If adoPrimaryRS1.RecordCount > 0 Then
adoPrimaryRS1.MoveFirst
End If

Do Until adoPrimaryRS1.EOF
C1.AddItem adoPrimaryRS1.Fields(0)
C2.AddItem adoPrimaryRS1.Fields(1)
C3.AddItem adoPrimaryRS1.Fields(2)
C4.AddItem adoPrimaryRS1.Fields(3)
adoPrimaryRS1.MoveNext
Loop

If C1.ListCount > 0 Then
C1.ListIndex = 0
C2.ListIndex = 0
C3.ListIndex = 0
C4.ListIndex = 0
End If

  doc = 0
  If adoPrimaryRS3.RecordCount > 0 Then
  adoPrimaryRS3.MoveFirst
  End If
  
  Do Until adoPrimaryRS3.EOF
  If IsNumeric(adoPrimaryRS3.Fields("doc")) Then
  If Val(adoPrimaryRS3.Fields("doc")) > doc Then doc = adoPrimaryRS3.Fields("doc")
  End If
  adoPrimaryRS3.MoveNext
  Loop
   
  txtDoc.Text = doc + 1
  txtFields(2).Text = doc + 1

aFlag = True
eFlag = False

l.ListItems.Clear
txtFields(3).Text = Format(Date, "dd/mm/yyyy")
    txtFields(2).SetFocus
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
'  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()

On Error Resume Next

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

db.Execute "Delete from [Sales Order] where doc='" & txtDoc.Text & "'"

With adoPrimaryRS3
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
End With

txtDoc = adoPrimaryRS3.Fields("doc")

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order] where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)

  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'on error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  'on error GoTo EditErr
  aFlag = False
  eFlag = True
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
  
If Len(Trim(txtFields(2))) = 0 Then
txtFields(2).SetFocus
Exit Sub
End If

If adoPrimaryRS2.RecordCount > 0 Then
adoPrimaryRS2.MoveFirst
End If

Dim found As Boolean
found = False
Do Until adoPrimaryRS2.EOF
If adoPrimaryRS2.Fields("member id") = txtFields(4) Then
found = True
txtFields("member name") = adoPrimaryRS2.Fields(1)
End If
adoPrimaryRS2.MoveNext
Loop

If found = False Then
frmmemberlist.Show
Exit Sub
End If


If Not IsNumeric(txtFields(8).Text) Then
txtFields(8).SetFocus
Exit Sub
End If

If Not IsNumeric(txtFields(9).Text) Then
txtFields(8).SetFocus
Exit Sub
End If

If l.ListItems.Count = 0 Then
Command3_Click
End If

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

db.Execute "Delete from [Sales Order]where doc='" & txtDoc.Text & "'"

For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
adoPrimaryRS4.AddNew
adoPrimaryRS4.Fields(0) = l.SelectedItem.Text
adoPrimaryRS4.Fields(1) = l.SelectedItem.SubItems(1)
adoPrimaryRS4.Fields(2) = l.SelectedItem.SubItems(2)
adoPrimaryRS4.Fields(3) = l.SelectedItem.SubItems(3)
adoPrimaryRS4.Fields(4) = txtFields(4).Text
adoPrimaryRS4.Fields(5) = txtFields(5).Text
adoPrimaryRS4.Fields(6) = l.SelectedItem.SubItems(4)
adoPrimaryRS4.Fields(7) = l.SelectedItem.SubItems(5)
adoPrimaryRS4.Fields(8) = txtFields(2).Text
adoPrimaryRS4.Fields(9) = txtFields(3).Text
adoPrimaryRS4.Fields("doc") = txtDoc
adoPrimaryRS4.Update
Next

db.Execute "Delete from [Order No]where doc='" & txtDoc.Text & "'"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS1 = New Recordset
adoPrimaryRS1.Open "select * from Item", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS2 = New Recordset
adoPrimaryRS2.Open "select * from [member]", db, adOpenStatic, adLockOptimistic
  
Set adoPrimaryRS3 = New Recordset
adoPrimaryRS3.Open "select * from [Order No]", db, adOpenStatic, adLockOptimistic

Set adoPrimaryRS4 = New Recordset
adoPrimaryRS4.Open "select * from Sales", db, adOpenStatic, adLockOptimistic

If Not adoPrimaryRS3.EOF Then adoPrimaryRS3.MoveNext
  If adoPrimaryRS3.EOF And adoPrimaryRS3.RecordCount > 0 Then
    Beep
    adoPrimaryRS3.MoveLast
End If
  
txtDoc = adoPrimaryRS3.Fields("doc")

l.ListItems.Clear

cmdNext_Click

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()

On Error Resume Next

adoPrimaryRS3.MoveFirst
  
txtDoc = adoPrimaryRS3.Fields("doc")

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)
  
  
mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()

On Error Resume Next

adoPrimaryRS3.MoveLast
  
txtDoc = adoPrimaryRS3.Fields("doc")

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)

  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()

On Error Resume Next

If Not adoPrimaryRS3.EOF Then adoPrimaryRS3.MoveNext
  If adoPrimaryRS3.EOF And adoPrimaryRS3.RecordCount > 0 Then
    Beep
    adoPrimaryRS3.MoveLast
End If
  
txtDoc = adoPrimaryRS3.Fields("doc")

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)

  
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()

On Error Resume Next


If Not adoPrimaryRS3.BOF Then adoPrimaryRS3.MovePrevious
If adoPrimaryRS3.BOF And adoPrimaryRS3.RecordCount > 0 Then
    Beep
adoPrimaryRS3.MoveFirst
End If

txtDoc = adoPrimaryRS3.Fields("doc")

Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path + "\cyber.mdb;"

Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from [Sales Order]where doc='" & txtDoc.Text & "'", db, adOpenStatic, adLockOptimistic

l.ListItems.Clear

If adoPrimaryRS.RecordCount > 0 Then
adoPrimaryRS.MoveFirst
Do Until adoPrimaryRS.EOF
Set x = l.ListItems.Add(, , adoPrimaryRS.Fields(0))
x.SubItems(1) = adoPrimaryRS.Fields(1)
x.SubItems(2) = adoPrimaryRS.Fields(2)
x.SubItems(3) = adoPrimaryRS.Fields(3)
txtFields(4).Text = adoPrimaryRS.Fields(4)
txtFields(5).Text = adoPrimaryRS.Fields(5)
x.SubItems(4) = adoPrimaryRS.Fields(6)
x.SubItems(5) = adoPrimaryRS.Fields(7)
txtFields(2).Text = adoPrimaryRS.Fields(8)
txtFields(3).Text = adoPrimaryRS.Fields(9)
adoPrimaryRS.MoveNext
Loop
End If

If l.ListItems.Count > 0 Then
For i = 1 To l.ListItems.Count
l.ListItems(i).Selected = True
tot = tot + Val(l.SelectedItem.SubItems(5)) * Val(l.SelectedItem.SubItems(4))
Next
End If
Text1.Text = tot

C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)

mbDataChanged = False

Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
'  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
'  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub l_Click()
On Error Resume Next
C1.Clear
C2.Clear
C3.Clear
C4.Clear
C1 = l.SelectedItem.Text
C2 = l.SelectedItem.SubItems(1)
C3 = l.SelectedItem.SubItems(2)
C4 = l.SelectedItem.SubItems(3)
txtFields(8).Text = l.SelectedItem.SubItems(4)
txtFields(9).Text = l.SelectedItem.SubItems(5)
End Sub

Private Sub MonthView1_DblClick(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
MonthView1.Visible = False
txtFields(3).Text = MonthView1.DateValue
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command3_Click
End If
End Sub
