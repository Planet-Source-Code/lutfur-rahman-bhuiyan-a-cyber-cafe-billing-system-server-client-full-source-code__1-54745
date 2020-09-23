VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register"
   ClientHeight    =   1140
   ClientLeft      =   2220
   ClientTop       =   4035
   ClientWidth     =   4680
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3660
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtRegistrationCode 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Registration Code:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim oLicense As New CLicense
    
    If oLicense.IsRegistrationCode(txtRegistrationCode.Text) = True Then
        Call oLicense.Register
    End If
    Set oLicense = Nothing
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRegister = Nothing
End Sub
