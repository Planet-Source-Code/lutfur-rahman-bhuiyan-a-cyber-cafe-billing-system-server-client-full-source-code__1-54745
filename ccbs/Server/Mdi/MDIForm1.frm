VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "CyberServer"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock port80 
      Left            =   720
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10425
      TabIndex        =   0
      Top             =   0
      Width           =   10425
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   1
         Top             =   240
         Width           =   735
         Begin VB.Shape Shape1 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   -120
            Shape           =   3  'Circle
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin MSWinsockLib.Winsock w 
      Left            =   2160
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wControl 
      Left            =   1680
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&File"
      Begin VB.Menu mnumm 
         Caption         =   "&Machine Master"
      End
      Begin VB.Menu mnuugm 
         Caption         =   "&User Group Master"
      End
      Begin VB.Menu mnuCM 
         Caption         =   "&Client Master"
      End
      Begin VB.Menu mnuCS 
         Caption         =   "&Company Setup"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRC 
      Caption         =   "&Remote Control"
      Begin VB.Menu mnuRLogin 
         Caption         =   "&Remote Login"
      End
      Begin VB.Menu mnuSD 
         Caption         =   "&Shut Down"
      End
   End
   Begin VB.Menu mnuCl 
      Caption         =   "Client"
      Begin VB.Menu mnuMR 
         Caption         =   "Mony Receipt"
      End
   End
   Begin VB.Menu mnuSer 
      Caption         =   "Service"
      Begin VB.Menu mnuCat 
         Caption         =   "Category"
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Item"
      End
   End
   Begin VB.Menu mnupur 
      Caption         =   "Purchase"
      Begin VB.Menu mnuClient 
         Caption         =   "Vendor Create"
      End
      Begin VB.Menu mnupurs 
         Caption         =   "Purchase"
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "Sales"
      Begin VB.Menu mnuSorder 
         Caption         =   "Sales Order"
      End
      Begin VB.Menu mnuOS 
         Caption         =   "Open Order"
      End
   End
   Begin VB.Menu mmnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuPB 
         Caption         =   "Pending Bill"
      End
      Begin VB.Menu mnuConsole 
         Caption         =   "Console"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()

If runserver = False Then
MDIForm1.Picture2.BackColor = vbGreen
frm_status.Winsock1(0).Close
frm_status.Winsock1(0).LocalPort = 7777
frm_status.Winsock1(0).Bind frm_status.Winsock1(0).LocalPort
runserver = True
Image1.Picture = Image2.Picture
Exit Sub
End If

If runserver = True Then
frm_status.Winsock1(0).Close
MDIForm1.Picture2.BackColor = vbRed
runserver = False
Image1.Picture = Image3.Picture
Exit Sub
End If

End Sub

Private Sub MDIForm_Resize()
'Call HangUp
Shape1.FillColor = vbGreen
frm_status.left = 0
frm_status.Top = 0
frm_status.Width = Me.Width
frm_status.Show
'Image1.left = 0
'Image1.Top = 0
'Image1.Height = Picture2.Height
'Image1.Width = Picture2.Height
End Sub

Private Sub mnuCat_Click()
frmCategory.Show
End Sub

Private Sub mnuClient_Click()
frmvendor.Show
End Sub

Private Sub mnuCM_Click()
frmMember.Show
End Sub

Private Sub mnuConsole_Click()
frm_status.Picture1.Visible = True
End Sub

Private Sub mnuCS_Click()
frmCompany.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuItem_Click()
frmItem.Show
End Sub

Private Sub mnumm_Click()
Mname.Show
End Sub

Private Sub mnuMR_Click()
frmReceipt1.Show
End Sub

Private Sub mnuOS_Click()
frmSales.Show
End Sub

Private Sub mnuPB_Click()
frm_status.Picture3.Visible = True
End Sub

Private Sub mnupurs_Click()
frmPurchase.Show
End Sub

Private Sub mnuRLogin_Click()
frm_remote_login.Show
End Sub

Private Sub mnuSD_Click()
frmShutDown.Show
End Sub


Private Sub mnuSorder_Click()
frmSalesOrder.Show
End Sub

Private Sub mnuugm_Click()
frmUserGroup.Show
End Sub

Private Sub Picture2_Click()
If Shape1.FillColor = vbGreen Then
frm_status.Winsock1(0).Close
Shape1.FillColor = vbRed
Exit Sub
End If
If Shape1.FillColor = vbRed Then
frm_status.Winsock1(0).Close
frm_status.Winsock1(0).LocalPort = 7777
frm_status.Winsock1(0).Bind frm_status.Winsock1(0).LocalPort
Shape1.FillColor = vbGreen
Exit Sub
End If
End Sub
