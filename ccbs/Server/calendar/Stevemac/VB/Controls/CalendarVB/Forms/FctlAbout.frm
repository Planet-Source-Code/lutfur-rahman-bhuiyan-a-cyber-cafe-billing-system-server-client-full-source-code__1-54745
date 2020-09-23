VERSION 5.00
Begin VB.Form FctlAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   2820
   ClientLeft      =   2430
   ClientTop       =   4455
   ClientWidth     =   4230
   Icon            =   "FctlAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register..."
      Height          =   315
      Left            =   1770
      TabIndex        =   7
      Top             =   2430
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Tag             =   "STS:[ Close the About Box dialog]"
      Top             =   2430
      Width           =   1125
   End
   Begin VB.Label lblUnregistered 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unregistered!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2490
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgDefault 
      Height          =   480
      Left            =   150
      Picture         =   "FctlAbout.frx":000C
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblAboutBox 
      Alignment       =   2  'Center
      Caption         =   "Trademark"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2130
      Width           =   3975
   End
   Begin VB.Label lblAboutBox 
      Alignment       =   2  'Center
      Caption         =   "Copyright"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label lblAboutBox 
      Alignment       =   2  'Center
      Caption         =   "CompanyName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   3735
   End
   Begin VB.Label lblAboutBox 
      Alignment       =   2  'Center
      Caption         =   "FileDescription"
      Height          =   795
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   690
      UseMnemonic     =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblAboutBox 
      Caption         =   "ProductName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   90
      Width           =   3375
   End
   Begin VB.Image imgAppIcon 
      Height          =   555
      Left            =   60
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "FctlAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================
'  Name [FctlAbout.frm]
'
'  Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        8/11/97
'=======================================================

'$Runtime Dependencies:
'$DesignTime Dependencies:

'=======================================================
'  Usage Notes:
'       For displaying an aboutbox for custom objects.
'
'=======================================================
'  Form Methods:
'   ShowModal(sProduct, sVersion, sDescription, [sCompany], [sCopyright], [sTradeMark], [vIcon])
'=======================================================
'  Form Properties:
'    AppIcon        (Write-Only)
'=======================================================
Option Explicit

Private Sub cmdOK_Click()
 
    Hide
 
End Sub


Private Sub cmdRegister_Click()
    frmRegister.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 
    'If ESC pressed the unload the About dialog
    If KeyAscii = 27 Then Unload Me
 
End Sub

Private Sub Form_Load()
 
    'Set the icon that is displayed in the about box too our
    'default image.
    imgAppIcon.Picture = imgDefault.Picture
 
End Sub

Public Property Let AppIcon(picIcon As Picture)
    imgAppIcon.Picture = picIcon
End Property

Public Sub ShowModal(ByVal sProduct As String, ByVal sVersion As String, ByVal sDescription As String, Optional ByVal sCompany As String = "", Optional ByVal sCopyright As String = "", Optional ByVal sTradeMark As String = "", Optional vIcon, Optional ByVal bIsRegistered As Boolean = True)
 
    'Check to see if we need to display the registration stuff
'    If bIsRegistered = False Then
'        cmdRegister.Visible = True
'        lblUnregistered.Visible = True
'    End If
    'If Icon passed then use it, else display our default Icon
    If Not IsMissing(vIcon) Then
        imgAppIcon.Picture = vIcon
    Else
        imgAppIcon.Picture = imgDefault.Picture
    End If
 
    'If none was passed then use these as our default
    If Len(sCompany) = 0 Then sCompany = "CTR Business Systems, Inc."
    If Len(sCopyright) = 0 Then sCopyright = "Copyright © " & DatePart("yyyy", Now) & " CTR Business Systems, Inc."
 
    'Set the labels on the about box
    Me.Caption = "About " & sProduct
    lblAboutBox(0) = sProduct & " v" & sVersion
    'lblAboutBox(1) = sVersion
    lblAboutBox(2) = sDescription
    lblAboutBox(3) = sCompany
    lblAboutBox(4) = sCopyright
    lblAboutBox(5) = sTradeMark
 
    'All done so lets show off...
    Me.Show vbModal
 
End Sub

