VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather Report"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Scroll"
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   8775
      TabIndex        =   8
      Top             =   7785
      Width           =   795
      Begin PRJMÉTÉO.xpButton cmdMore 
         Height          =   240
         Left            =   450
         TabIndex        =   9
         Top             =   225
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   423
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmMain.frx":030A
      End
      Begin PRJMÉTÉO.xpButton cmdLess 
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   423
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmMain.frx":0326
      End
   End
   Begin VB.PictureBox picPrévision 
      AutoSize        =   -1  'True
      Height          =   5190
      Left            =   8520
      Picture         =   "frmMain.frx":0342
      ScaleHeight     =   5130
      ScaleWidth      =   3705
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   3765
   End
   Begin PRJMÉTÉO.xpButton cmdStart 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   7935
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Start Animation"
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
      FOCUSR          =   0   'False
      MPTR            =   0
      MICON           =   "frmMain.frx":28B0
   End
   Begin PRJMÉTÉO.xpButton cmdPrévision 
      Height          =   375
      Left            =   3495
      TabIndex        =   4
      Top             =   7935
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Forecast"
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
      FOCUSR          =   0   'False
      MPTR            =   0
      MICON           =   "frmMain.frx":28CC
   End
   Begin PRJMÉTÉO.xpButton cmdAnimée 
      Height          =   375
      Left            =   5205
      TabIndex        =   5
      Top             =   7935
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Radar"
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
      FOCUSR          =   0   'False
      MPTR            =   0
      MICON           =   "frmMain.frx":28E8
   End
   Begin PRJMÉTÉO.xpButton cmdQuit 
      Height          =   375
      Left            =   6900
      TabIndex        =   6
      Top             =   7935
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Quit"
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
      FOCUSR          =   0   'False
      MPTR            =   0
      MICON           =   "frmMain.frx":2904
   End
   Begin PRJMÉTÉO.xpButton cmdStop 
      Height          =   375
      Left            =   105
      TabIndex        =   2
      Top             =   7935
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Stop Animation"
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
      FOCUSR          =   0   'False
      MPTR            =   0
      MICON           =   "frmMain.frx":2920
   End
   Begin PRJMÉTÉO.AnimGif ctrlAnimGif 
      Height          =   7665
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15285
      _extentx        =   26961
      _extenty        =   13520
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Infos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   9735
      TabIndex        =   1
      Top             =   7935
      Width           =   5430
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
        ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Dim CheminDuGif As String
Dim Réussite As Boolean
Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function
Private Sub cmdStop_Click()
    ctrlAnimGif.StopGif
    cmdMore.Enabled = False
    cmdLess.Enabled = False
    Me.Caption = "Weather Report      Radar Image Pause"
    lblInfo.Caption = "Pause Image Radar"
End Sub
Private Sub cmdAnimée_Click()
    cmdStart.Enabled = True
    cmdStop.Enabled = True
    cmdMore.Enabled = True
    cmdLess.Enabled = True
    picPrévision.Visible = False
    lblInfo.Caption = "Loading"
    CheminDuGif = "http://www.meteo.ec.gc.ca/radar/data/looped_image/COMPOSITE_QUE_LOOP.GIF"
    AfficheAnimationGIF
    Me.Caption = "Weather Report     Radar Image"
    lblInfo.Caption = "Radar Image"
End Sub
Private Sub cmdPrévision_Click()
    cmdStart.Enabled = False
    cmdStop.Enabled = False
    cmdMore.Enabled = False
    cmdLess.Enabled = False
    picPrévision.Visible = True
    ctrlAnimGif.StopGif
    CheminDuGif = "http://gfx.weatheroffice.ec.gc.ca/jet_stream/data/tempmapwx_e.gif"
    AfficheAnimationGIF
    Me.Caption = "Weather Report     Radar Image"
    lblInfo.Caption = "Forecast"
End Sub
Private Sub cmdStart_Click()
    ctrlAnimGif.StartGif
    cmdMore.Enabled = True
    cmdLess.Enabled = True
    Me.Caption = "Weather Report     Radar Image"
    lblInfo.Caption = "Radar Image"
End Sub
Private Sub cmdQuit_Click()
    Unload Me
    End
End Sub
Private Sub Form_Load()
    Me.Show
    lblInfo.Caption = "Loading"
    CheminDuGif = "http://www.meteo.ec.gc.ca/radar/data/looped_image/COMPOSITE_QUE_LOOP.GIF"
    AfficheAnimationGIF
    Me.Caption = "Weather Report     Radar Image"
    lblInfo.Caption = "Radar Image"
End Sub
Private Sub AfficheAnimationGIF()
    On Error Resume Next
    Réussite = DownloadFile(CheminDuGif, App.Path & "\picture.GIF")
    If Réussite Then
        ctrlAnimGif.FichierGif = App.Path & "\picture.GIF"
        ctrlAnimGif.StartGif
    Else
        MsgBox "Unable to load image, verify your connection...", vbExclamation, "Erreur de chargement"
        Exit Sub
    End If
End Sub
Private Sub cmdLess_Click()
    If ctrlAnimGif.AnimSpeed < 2000 Then
        ctrlAnimGif.AnimSpeed = ctrlAnimGif.AnimSpeed + 100
    Else
        ctrlAnimGif.AnimSpeed = 2000
    End If
End Sub
Private Sub cmdMore_Click()
    If ctrlAnimGif.AnimSpeed > 100 Then
        ctrlAnimGif.AnimSpeed = ctrlAnimGif.AnimSpeed - 100
    Else
        ctrlAnimGif.AnimSpeed = 100
    End If
End Sub
