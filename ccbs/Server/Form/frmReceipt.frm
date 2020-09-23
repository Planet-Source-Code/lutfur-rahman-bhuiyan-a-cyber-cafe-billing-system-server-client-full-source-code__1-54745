VERSION 5.00
Object = "*\A..\calendar\Stevemac\VB\Controls\CalendarVB\CalendarVB.vbp"
Begin VB.Form frmReceipt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mony Receipt"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   8
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   22
      Top             =   2910
      Visible         =   0   'False
      Width           =   975
   End
   Begin ctrCalendarVB.CalendarVB MonthView1 
      Height          =   2535
      Left            =   2520
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5325
      TabIndex        =   17
      Top             =   2940
      Width           =   5325
      Begin Project1.xpButton xpButton1 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         TX              =   "&Save"
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
         MICON           =   "frmReceipt.frx":0000
      End
      Begin Project1.xpButton xpButton2 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         TX              =   "&Print"
         ENAB            =   0   'False
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
         MICON           =   "frmReceipt.frx":001C
      End
      Begin Project1.xpButton xpButton3 
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         TX              =   "E&xit"
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
         MICON           =   "frmReceipt.frx":0038
      End
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2610
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   6
      Top             =   2250
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1930
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   2280
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1620
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member Name"
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1290
      Width           =   3855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Member ID"
      Height          =   285
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Received"
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Receipt:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   2565
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Net Due:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   2250
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Discount:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1930
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Due:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Receipt No."
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client Name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Client ID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   645
   End
End
Attribute VB_Name = "frmReceipt"
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

Private Sub Command1_Click()
MonthView1.Visible = True
End Sub

Private Sub Command3_Click()
frm_Client_List1.Show
End Sub

Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cyber.mdb;"
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select * from receipt", db, adOpenStatic, adLockOptimistic
If adoPrimaryRS.RecordCount > 0 Then adoPrimaryRS.MoveFirst
mr = 0
Do Until adoPrimaryRS.EOF
If adoPrimaryRS.Fields("mr") > mr Then mr = adoPrimaryRS.Fields("mr")
adoPrimaryRS.MoveNext
Loop
txtFields(2).Text = mr + 1
'txtFields(7).SetFocus
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
MonthView1.Visible = False
txtFields(4).Text = MonthView1.Value
End Sub

Private Sub MonthView1_DblClick(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
txtFields(4).Text = MonthView1.DateValue
MonthView1.Visible = False
End Sub

Private Sub xpButton1_Click()
If Len(Trim(txtFields(7))) = 0 Then
txtFields(7).SetFocus
Exit Sub
End If
With adoPrimaryRS
.AddNew
.Fields("Client ID") = txtFields(0)
.Fields("Client Name") = txtFields(1)
.Fields("date") = txtFields(4)
.Fields("MR") = txtFields(2)
.Fields("Amount") = txtFields(7)
.Update
End With
xpButton1.Enabled = False
xpButton2.Enabled = True
End Sub

Private Sub xpButton2_Click()
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
Print #1, "<p>Date: " & txtFields(4).Text & "</p>"
Print #1, "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
Print #1, "<tr>"
Print #1, "<p>Recipt No.: " & txtFields(2).Text & "</p>"
Print #1, "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Client ID:</b></td>"
Print #1, "<td width=""80%"">" & txtFields(0).Text & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Cient Name:</b></td>"
Print #1, "<td width=""80%"">" & txtFields(1).Text & "</td>"
Print #1, "</tr>"

Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Total Due:</b></td>"
Print #1, "<td width=""80%"">" & Format(txtFields(3).Text, "#,##0.00") & "</td>"
Print #1, "</tr>"

Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Total Discount:</b></td>"
Print #1, "<td width=""80%"">" & Format(txtFields(5).Text, "#,##0.00") & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td width=""20%""><b>Net Due:</b></td>"
Print #1, "<td width=""80%"">" & Format(txtFields(6).Text, "#,##0.00") & "</td>"
Print #1, "</tr>"
Print #1, "<tr>"
Print #1, "<td Width=""20%""><b>Receipt:</b></td>"
Print #1, "<td width=""80%"">" & Format(txtFields(7).Text, "#,##0.00") & "</td>"
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

Private Sub xpButton3_Click()
Unload Me
End Sub
