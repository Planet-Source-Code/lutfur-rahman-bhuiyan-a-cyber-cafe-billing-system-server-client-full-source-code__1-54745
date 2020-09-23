VERSION 5.00
Begin VB.UserControl AnimGif 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   ScaleHeight     =   1155
   ScaleWidth      =   1155
   ToolboxBitmap   =   "AnimGif.ctx":0000
   Begin VB.Timer Timer 
      Left            =   180
      Top             =   660
   End
   Begin VB.Image imgSource 
      Height          =   1095
      Index           =   0
      Left            =   15
      Top             =   30
      Width           =   1095
   End
End
Attribute VB_Name = "AnimGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mTotalFrames As Long
Dim mRepeatTimes As Long
Dim FrameCount As Long
Const m_def_AnimSpeed = 500
Const m_def_FichierGif = ""
Dim m_AnimSpeed As Integer
Dim m_FichierGif As String
Private Sub Timer_Timer()
    Dim i As Long
    If FrameCount < TotalFrames Then
        imgSource(FrameCount).Visible = False
        FrameCount = FrameCount + 1
        imgSource(FrameCount).Visible = True
        Timer.Interval = m_AnimSpeed
    Else
        FrameCount = 0
        For i = 1 To imgSource.Count - 1
            imgSource(i).Visible = False
        Next i
        imgSource(FrameCount).Visible = True
        Timer.Interval = m_AnimSpeed
    End If
End Sub
Private Sub UserControl_Initialize()
    imgSource(0).Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Private Sub UserControl_Resize()
    imgSource(0).Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Public Property Get TotalFrames() As Long
    TotalFrames = mTotalFrames
End Property
Public Property Let TotalFrames(ByVal vNewValue As Long)
    mTotalFrames = vNewValue
End Property
Public Property Get RepeatTimes() As Long
    RepeatTimes = mRepeatTimes
End Property
Public Property Let RepeatTimes(ByVal vNewValue As Long)
    mRepeatTimes = vNewValue
End Property
Private Function LoadGif(sFile As String, aImg As Variant) As Boolean
    LoadGif = False
    If sFile = "" Then
        Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    For i = 1 To aImg.Count - 1
        Unload aImg(i)
    Next i
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
    buf = String(LOF(fNum), Chr(0))
    Get #fNum, , buf
    Close fNum
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = left(buf, j)
    If left$(fileHeader, 3) <> "GIF" Then
        Err.Raise vbObjectError + 2, , "not supported file format"
        Exit Function
    End If
    LoadGif = True
    i = j + 2
    If Len(fileHeader) >= 127 Then
        mRepeatTimes = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
    Else
        mRepeatTimes = 0
    End If
    Do
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + j - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, j - i)
            Put #fNum, 1, picbuf
            imgHeader = left(Mid(buf, i - 1, j - i), 16)
            Close fNum
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).left = aImg(0).left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            i = j
        End If
        DoEvents
    Loop Until j = 3
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
        picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
        picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
        Put #fNum, 1, picbuf
        imgHeader = left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).left = aImg(0).left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    LoadGif = False
    On Error GoTo 0
End Function
Public Sub StartGif()
    Timer.Enabled = False
    If LoadGif(FichierGif, imgSource) Then
        FrameCount = 0
        Timer.Interval = m_AnimSpeed
        Timer.Enabled = True
    End If
End Sub
Public Sub StopGif()
    Timer.Enabled = False
End Sub
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HC0C0C0)
    m_FichierGif = PropBag.ReadProperty("FichierGif", m_def_FichierGif)
    If m_FichierGif <> "" Then
        Set imgSource(0) = LoadPicture(m_FichierGif)
        If Ambient.UserMode = False Then
            Exit Sub
        Else
            StartGif
        End If
    End If
    m_AnimSpeed = PropBag.ReadProperty("AnimSpeed", m_def_AnimSpeed)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("FichierGif", m_FichierGif, m_def_FichierGif)
    Call PropBag.WriteProperty("AnimSpeed", m_AnimSpeed, m_def_AnimSpeed)
End Sub
Public Property Get FichierGif() As String
    FichierGif = m_FichierGif
End Property
Public Property Let FichierGif(ByVal New_FichierGif As String)
    If New_FichierGif = "" Then m_FichierGif = New_FichierGif: Exit Property
    If VerifieFichier(New_FichierGif) Then
        m_FichierGif = New_FichierGif
        PropertyChanged "FichierGif"
        Set imgSource(0) = LoadPicture(m_FichierGif)
    Else
        MsgBox "Bad file format", vbExclamation, "File error"
        New_FichierGif = ""
        m_FichierGif = New_FichierGif
        Set imgSource(0) = LoadPicture()
    End If
End Property
Private Sub UserControl_InitProperties()
    m_FichierGif = m_def_FichierGif
    m_AnimSpeed = m_def_AnimSpeed
    Timer.Interval = m_AnimSpeed
End Sub
Private Function VerifieFichier(sFile As String) As Boolean
    If sFile = "" Then
        Exit Function
    End If
    Dim fNum As Integer
    Dim fileHeader As String
    Dim buf$
    Dim imgCount As Integer
    Dim i&, j&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
    buf = String(LOF(fNum), Chr(0))
    Get #fNum, , buf
    Close fNum
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = left(buf, j)
    If left$(fileHeader, 3) <> "GIF" Then
        VerifieFichier = False
        Exit Function
    Else
        VerifieFichier = True
    End If
End Function
Public Property Get AnimSpeed() As Integer
    AnimSpeed = m_AnimSpeed
End Property
Public Property Let AnimSpeed(ByVal New_AnimSpeed As Integer)
    m_AnimSpeed = New_AnimSpeed
    PropertyChanged "AnimSpeed"
End Property
