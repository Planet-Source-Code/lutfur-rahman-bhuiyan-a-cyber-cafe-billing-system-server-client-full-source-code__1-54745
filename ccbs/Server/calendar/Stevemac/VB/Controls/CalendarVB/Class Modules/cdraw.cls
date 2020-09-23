VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Copyright © 1997-1999 by CTR Business Systems, Inc.
'----------------------------------------------------------------------
Option Explicit

'API Declares
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type Size
        cx As Long
        cy As Long
End Type

'API Text alignment
Private Const TA_BASELINE& = 24
Private Const TA_BOTTOM& = 8
Private Const TA_CENTER& = 6
Private Const TA_LEFT& = 0
Private Const TA_NOUPDATECP& = 0
Private Const TA_UPDATECP& = 1
Private Const TA_RIGHT& = 2
Private Const TA_TOP& = 0

'API Pen Constants
Private Const PS_SOLID& = 0

'API Color Constants
Private Const COLOR_SYSTEM = &H80000000
Private Const COLOR_SYSTEM_INDEX = &HFF
Private Const COLOR_BTNFACE& = 15
Private Const COLOR_BTNHIGHLIGHT& = 20
Private Const COLOR_BTNHILIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_BTNSHADOW& = 16
Private Const COLOR_BTNTEXT& = 18
Private Const COLOR_3DDKSHADOW& = 21
Private Const COLOR_3DFACE& = COLOR_BTNFACE
Private Const COLOR_3DHIGHLIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHILIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DLIGHT& = 22
Private Const COLOR_3DSHADOW& = COLOR_BTNSHADOW

'API DrawEdge Constants
Private Const BORDER_WIDTH As Long = 0
Private Const BDR_INNER& = &HC
Private Const BDR_OUTER& = &H3
Private Const BDR_RAISED& = &H5
Private Const BDR_RAISEDINNER& = &H4
Private Const BDR_RAISEDOUTER& = &H1
Private Const BDR_SUNKEN& = &HA
Private Const BDR_SUNKENINNER& = &H8
Private Const BDR_SUNKENOUTER& = &H2
Private Const EDGE_BUMP& = &H9&
Private Const EDGE_ETCHED& = &H6&
Private Const EDGE_RAISED& = &H5&
Private Const EDGE_SUNKEN& = &HA&
Private Const BF_ADJUST& = &H2000
Private Const BF_BOTTOM& = &H8
Private Const BF_BOTTOMLEFT& = &H9
Private Const BF_BOTTOMRIGHT& = &HC
Private Const BF_DIAGONAL& = &H10
Private Const BF_FLAT& = &H4000
Private Const BF_LEFT& = &H1
Private Const BF_MIDDLE& = &H800
Private Const BF_MONO& = &H8000&
Private Const BF_RECT& = &HF
Private Const BF_RIGHT& = &H4
Private Const BF_SOFT& = &H1000
Private Const BF_TOP& = &H2
Private Const BF_TOPLEFT& = &H3
Private Const BF_TOPRIGHT& = &H6

'API DrawFrameControl Constants
Private Const DFC_BUTTON& = 4
Private Const DFC_CAPTION& = 1
Private Const DFC_MENU& = 2
Private Const DFC_SCROLL& = 3
Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_BUTTON3STATE& = &H8
Private Const DFCS_BUTTONCHECK& = &H0
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_BUTTONRADIO& = &H4
Private Const DFCS_BUTTONRADIOIMAGE& = &H1
Private Const DFCS_BUTTONRADIOMASK& = &H2
Private Const DFCS_CAPTIONCLOSE& = &H0
Private Const DFCS_CAPTIONHELP& = &H4
Private Const DFCS_CAPTIONMAX& = &H2
Private Const DFCS_CAPTIONMIN& = &H1
Private Const DFCS_CAPTIONRESTORE& = &H3
Private Const DFCS_CHECKED& = &H400
Private Const DFCS_FLAT& = &H4000
Private Const DFCS_INACTIVE& = &H100
Private Const DFCS_MENUARROW& = &H0
Private Const DFCS_MENUARROWRIGHT& = &H4
Private Const DFCS_MENUBULLET& = &H2
Private Const DFCS_MENUCHECK& = &H1
Private Const DFCS_MONO& = &H8000&
Private Const DFCS_PUSHED& = &H200
Private Const DFCS_SCROLLCOMBOBOX& = &H5
Private Const DFCS_SCROLLDOWN& = &H1
Private Const DFCS_SCROLLLEFT& = &H2
Private Const DFCS_SCROLLRIGHT& = &H3
Private Const DFCS_SCROLLSIZEGRIP& = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT& = &H10
Private Const DFCS_SCROLLUP& = &H0

'API DrawText Constants
Private Const DT_BOTTOM& = &H8
Private Const DT_CALCRECT& = &H400
Private Const DT_CENTER& = &H1
Private Const DT_EXPANDTABS& = &H40
Private Const DT_EXTERNALLEADING& = &H200
Private Const DT_LEFT& = &H0
Private Const DT_NOCLIP& = &H100
Private Const DT_NOPREFIX& = &H800
Private Const DT_RIGHT& = &H2
Private Const DT_SINGLELINE& = &H20
Private Const DT_TOP& = &H0
Private Const DT_VCENTER& = &H4
Private Const DT_WORDBREAK& = &H10

Private Const OPAQUE& = 2
Private Const SRCAND& = &H8800C6
Private Const SRCCOPY& = &HCC0020
Private Const SRCERASE& = &H440328
Private Const SRCINVERT& = &H660046
Private Const DSTINVERT& = &H550009
Private Const SRCPAINT& = &HEE0086
Private Const COLOR_HIGHLIGHT& = 13
Private Const WHITENESS& = &HFF0062
Private Const BLACKNESS& = &H42

Private Const ETO_OPAQUE = 2
Private Const ETO_CLIPPED = 4

'Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
'Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
'Private Const COLOR_BTNFACE = 15
'Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
'Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
'Private Const COLOR_BTNHIGHLIGHT = 20

'API Declares
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvertRectAPI Lib "user32" Alias "InvertRect" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

#If UNICODE Then
    Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
#Else
    Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
#End If

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Private Declare Function CreatePatternBrush& Lib "gdi32" (ByVal hBitmap As Long)
Private Declare Function SetBkMode& Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long)
Private Declare Function GetDesktopWindow& Lib "user32" ()
Private Declare Function UpdateWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, lpInitData As Long)
Private Declare Function GetLastError& Lib "kernel32" ()
Private Declare Function GetCurrentPositionEx& Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI)
Private Declare Function TextOutAPI& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long)
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByVal lpDx As Long) As Long
Private Declare Function SetTextAlign& Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long)
Private Declare Function LineToAPI& Lib "gdi32" Alias "LineTo" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long)
Private Declare Function RectangleAPI& Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function RoundRect& Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long)
Private Declare Function FillRectAPI& Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long)
Private Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Private Declare Function BeginPath& Lib "gdi32" (ByVal hdc As Long)
Private Declare Function EndPath& Lib "gdi32" (ByVal hdc As Long)
Private Declare Function DrawFocusRectAPI& Lib "user32" Alias "DrawFocusRect" (ByVal hdc As Long, lpRect As RECT)
Private Declare Function DrawFrameControl& Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long)
Private Declare Function DrawTextAPI& Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long)

'Class Constants
Private Const m_lSUNKEN_OFFSET = 1
'Button States
Private Const giFLATTENED As Integer = 0
Private Const giRAISED As Integer = 1
Private Const giSUNKEN As Integer = 3
Private Const giDISABLED As Integer = 4

'VB Errors
Private Const giOBJECT_VARIABLE_NOT_SET As Integer = 91
Private Const giINVALID_PICTURE As Integer = 481
Private Const giDLL_FUNCTION_NOT_FOUND As Integer = 453

'Public Enums
Public Enum drwTextAlignment
    drwBaseLine = TA_BASELINE
    drwBottom = TA_BOTTOM
    drwCenter = TA_CENTER
    drwLeft = TA_LEFT
    drwRight = TA_RIGHT
    drwTop = TA_TOP
    drwNoCursorPositionUpdate = TA_NOUPDATECP
    drwUpdateCursorPosition = TA_UPDATECP
End Enum
Public Enum drwBorderStyle
    'drwDropShadow = 0
    drwBump = 1
    drwEtched = 2
    drwRaised = 3
    drwSunken = 4
End Enum
Public Enum drwComponentType
    drwButton = DFC_BUTTON
    drwCaption = DFC_CAPTION
    drwMenu = DFC_MENU
    drwScroll = DFC_SCROLL
End Enum
Public Enum drwComponentStyle
    drwButton3State = DFCS_BUTTON3STATE
    drwButtonCheck = DFCS_BUTTONCHECK
    drwButtonPush = DFCS_BUTTONPUSH
    drwButtonRadio = DFCS_BUTTONRADIO
    drwCaptionClose = DFCS_CAPTIONCLOSE
    drwCaptionHelp = DFCS_CAPTIONHELP
    drwCaptionMax = DFCS_CAPTIONMAX
    drwCaptionMin = DFCS_CAPTIONMIN
    drwCaptionRestore = DFCS_CAPTIONRESTORE
    drwMenuArrow = DFCS_MENUARROW
    drwMenuBullet = DFCS_MENUBULLET
    drwMenuCheck = DFCS_MENUCHECK
    drwScrollComboBox = DFCS_SCROLLCOMBOBOX
    drwScrollDown = DFCS_SCROLLDOWN
    drwScrollleft = DFCS_SCROLLLEFT
    drwScrollright = DFCS_SCROLLRIGHT
    drwScrollSizeGrip = DFCS_SCROLLSIZEGRIP
    drwScrollUp = DFCS_SCROLLUP
End Enum
Public Enum drwComponentState
    drwStateNone = 0
    drwStateChecked = DFCS_CHECKED
    drwStateFlat = DFCS_FLAT
    drwStateInactive = DFCS_INACTIVE
    drwStateMono = DFCS_MONO
    drwStatePushed = DFCS_PUSHED
End Enum
Public Enum drwTextFormatting
    drwTextAlignBottom = DT_BOTTOM
    drwTextCalRectangle = DT_CALCRECT
    drwTextMultiLineCenter = DT_CENTER
    drwTextExpandTabs = DT_EXPANDTABS
    drwTextExternalLeading = DT_EXTERNALLEADING
    drwTextAlignLeft = DT_LEFT
    drwTextNoClipping = DT_NOCLIP
    drwTextNoPrefix = DT_NOPREFIX
    drwTextAlignRight = DT_RIGHT
    drwTextSingleLine = DT_SINGLELINE
    drwTextAlignTop = DT_TOP
    drwTextSingleLineCenter = DT_VCENTER
    drwTextWordBreak = DT_WORDBREAK
End Enum
Public Enum bmpBorderStates
    'Button States
    bmpFLATTENED = 0
    bmpRAISED = 1
    bmpSUNKEN = 3
    bmpDISABLED = 4
End Enum

Public Enum CaptionAlignments
    caCenterCenter
    caCenterLeft
    caCenterRight
    caTopCenter
    caTopLeft
    caTopRight
    caBottomCenter
    caBottomLeft
    caBottomright
End Enum

Public Enum Appearances
    NoLines = 0
    Raised
    Flat
    Sunken
    Selected
End Enum

'Class Variables
Private mlhHalftonePal As Long
Private mrgb3DFace As Long      'color to use for the 3d face
Private mrgb3DHighlight As Long 'color to use for the 3d highlight
Private mrgb3DShadow As Long    'color to use for the 3d shadow

'Class Properties
Private m_MemoryDC As New CMemoryDC
Private m_hdc As Long
Private m_lErrNumber As Long
Private m_lCurrentX As Long
Private m_lCurrentY As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_nTextAlignment As Integer
Private m_lFlatLineColor As Long

Public Property Get CurrentX() As Long
    
    CurrentX = CurrentXPos
    
End Property

Public Property Let CurrentX(ByVal x As Long)

    Dim lpPoint As POINTAPI
    
    On Error Resume Next
    'Get the current location
    m_lErrNumber = GetCurrentPositionEx(m_hdc, lpPoint)
    If m_lErrNumber <> 0 Then
        'Move to the new location
        m_lErrNumber = MoveToEx(m_hdc, x, lpPoint.y, lpPoint)
    End If
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Property

Public Property Get CurrentY() As Long
    
    CurrentY = CurrentYPos
    
End Property

Public Property Let CurrentY(ByVal y As Long)
    
    Dim lpPoint As POINTAPI
    
    On Error Resume Next
    'Get the current location
    m_lErrNumber = GetCurrentPositionEx(m_hdc, lpPoint)
    'Move to the new location
    m_lErrNumber = MoveToEx(m_hdc, lpPoint.x, y, lpPoint)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Property

Public Property Get ErrNumber() As Long
    ErrNumber = m_lErrNumber
End Property

Private Function CurrentXPos() As Long

    Dim lpPoint As POINTAPI
    
    m_lErrNumber = GetCurrentPositionEx(m_hdc, lpPoint)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    CurrentXPos = lpPoint.x

End Function

Private Function CurrentYPos() As Long

    Dim lpPoint As POINTAPI
    
    On Error Resume Next
    m_lErrNumber = GetCurrentPositionEx(m_hdc, lpPoint)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    CurrentYPos = lpPoint.y

End Function

Public Property Get hdc() As Long
    hdc = m_hdc
End Property

Public Property Let hdc(ByVal lhdc As Long)
    m_hdc = lhdc
    CurrentX = 0
    CurrentY = 0
    mlhHalftonePal = CreateHalftonePalette(m_hdc)
End Property

Public Sub LineFromTo(ByVal x As Long, ByVal y As Long, ByVal XEnd As Long, ByVal YEnd As Long)

    CurrentX = x
    CurrentY = y
    m_lErrNumber = LineToAPI(m_hdc, XEnd, YEnd)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If

End Sub

Public Sub LineTo(ByVal XEnd As Long, ByVal YEnd As Long, Optional ByVal clrLineColor As Long = -1)

    Dim lhPenOld As Long
    Dim lhPenNew As Long

    If clrLineColor <> -1 Then
        If clrLineColor And COLOR_SYSTEM Then clrLineColor = GetSysColor(clrLineColor And COLOR_SYSTEM_INDEX)
        lhPenNew = CreatePen(PS_SOLID, 0, clrLineColor)
        lhPenOld = SelectObject(m_hdc, lhPenNew)
    End If
    
    m_lErrNumber = LineToAPI(m_hdc, XEnd, YEnd)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
    'Do our house cleaning
    If clrLineColor <> -1 Then
        lhPenNew = SelectObject(m_hdc, lhPenOld)
        DeleteObject lhPenNew
    End If
    
End Sub

Public Sub Rectangle(ByVal rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long, Optional ByVal clrBorderColor As OLE_COLOR = -1, Optional ByVal clrFillColor As OLE_COLOR = -1, Optional ByVal EllipseWidth As Long = 0)

    Dim lpRect As RECT
    Dim lhPenOld As Long
    Dim lhPenNew As Long
    Dim lhbrFillColorOld As Long
    Dim lhbrFillColorNew As Long
    Dim bSystemColor As Long
    
    If clrBorderColor <> -1 Then
        If clrBorderColor And COLOR_SYSTEM Then clrBorderColor = GetSysColor(clrBorderColor And COLOR_SYSTEM_INDEX)
        lhPenNew = CreatePen(PS_SOLID, 0, clrBorderColor)
        lhPenOld = SelectObject(m_hdc, lhPenNew)
    End If
    If clrFillColor <> -1 Then
        If clrFillColor And COLOR_SYSTEM Then
            clrFillColor = GetSysColor(clrFillColor And COLOR_SYSTEM_INDEX)
        End If
        lhbrFillColorNew = CreateSolidBrush(clrFillColor)
        lhbrFillColorOld = SelectObject(m_hdc, lhbrFillColorNew)
    End If
    
    If EllipseWidth > 0 Then
        m_lErrNumber = RoundRect(m_hdc, rctLeft, rctTop, rctLeft + rctWidth, rctTop + rctHeight, EllipseWidth, EllipseWidth)
    Else
        m_lErrNumber = RectangleAPI(m_hdc, rctLeft, rctTop, rctLeft + rctWidth, rctTop + rctHeight)
    End If
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
    'Do our house cleaning
    If clrBorderColor <> -1 Then
        lhPenNew = SelectObject(m_hdc, lhPenOld)
        DeleteObject lhPenNew
    End If
    If clrFillColor <> -1 Then
        lhbrFillColorNew = SelectObject(m_hdc, lhbrFillColorOld)
        DeleteObject lhbrFillColorNew
    End If
    
End Sub

Public Sub TextOut(sText As String, ByVal x As Long, ByVal y As Long)

    m_lErrNumber = TextOutAPI(m_hdc, x, y, sText, Len(sText))
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Sub

Public Property Get TextAlignment() As drwTextAlignment
    TextAlignment = m_nTextAlignment
End Property

Public Property Let TextAlignment(ByVal nTextAlignment As drwTextAlignment)
    
    m_nTextAlignment = nTextAlignment
    m_lErrNumber = SetTextAlign(m_hdc, nTextAlignment)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Property

Public Sub DrawBorder(ByVal brdLeft As Long, ByVal brdTop As Long, ByVal brdWidth As Long, ByVal brdHeight As Long, ByVal brdStyle As drwBorderStyle, Optional ByVal bFlat As Boolean = False, Optional ByVal bMono As Boolean = False, Optional ByVal bSoft As Boolean = False)

    Dim lpRect As RECT
    Dim lFlag As Long
    Dim lEdge As Long
    
    lpRect.Left = brdLeft
    lpRect.Top = brdTop
    lpRect.Right = brdLeft + brdWidth
    lpRect.Bottom = brdTop + brdHeight
    
    Select Case brdStyle
    Case drwBump
        lEdge = EDGE_BUMP
    Case drwEtched
        lEdge = EDGE_ETCHED
    Case drwSunken
        lEdge = EDGE_SUNKEN
    Case drwRaised
        lEdge = EDGE_RAISED
    Case Else
        lEdge = EDGE_SUNKEN
    End Select
    lFlag = BF_RECT
    If bSoft Then lFlag = lFlag + BF_SOFT
    If bMono Then lFlag = lFlag + BF_MONO
    If bFlat Then lFlag = lFlag + BF_FLAT
    
    m_lErrNumber = DrawEdge(m_hdc, lpRect, lEdge, lFlag)

Exit_Proc:
    
End Sub

Public Sub DrawFocusRect(ByVal rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long)

    Dim lpRect As RECT

    lpRect.Left = rctLeft
    lpRect.Top = rctTop
    lpRect.Right = rctLeft + rctWidth
    lpRect.Bottom = rctTop + rctHeight
    
    m_lErrNumber = DrawFocusRectAPI(m_hdc, lpRect)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Sub

Public Sub DrawComponent(ByVal rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long, ByVal nType As drwComponentType, ByVal Style As drwComponentStyle, ByVal State As drwComponentState)

    Dim lFlag As Long
    Dim lpRect As RECT
    
    lpRect.Left = rctLeft
    lpRect.Top = rctTop
    lpRect.Right = rctLeft + rctWidth
    lpRect.Bottom = rctTop + rctHeight
    
    lFlag = Style + State
    
    m_lErrNumber = DrawFrameControl(m_hdc, lpRect, nType, lFlag)
    If m_lErrNumber <> 0 Then
        m_lErrNumber = 0
    Else
        m_lErrNumber = GetLastError()
    End If
    
End Sub

Public Function DrawText(ByVal sText As String, ByVal rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long, Optional ByVal Format As drwTextFormatting, Optional ByVal clrForeColor As SystemColorConstants = -1)

    Dim lFormat As Long
    Dim clrCurrentTextColor As Long
    Dim lpRect As RECT
    
    lpRect.Left = rctLeft
    lpRect.Top = rctTop
    lpRect.Right = rctLeft + rctWidth
    lpRect.Bottom = rctTop + rctHeight
    
    lFormat = Format
    
    If clrForeColor <> -1 Then
        clrCurrentTextColor = GetTextColor(m_hdc)
        If clrForeColor And COLOR_SYSTEM Then clrForeColor = GetSysColor(clrForeColor And COLOR_SYSTEM_INDEX)
        m_lErrNumber = SetTextColor(m_hdc, clrForeColor)
    End If
    
    DrawText = DrawTextAPI(m_hdc, sText, Len(sText), lpRect, lFormat)
    
    If Not IsMissing(clrCurrentTextColor) Then
        m_lErrNumber = SetTextColor(m_hdc, clrCurrentTextColor)
    End If
    
End Function

Public Sub ShadeRect(ByVal rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long)
    Dim crHiColor&, crOldBkColor&, crOldTextColor&
    Dim hBrush&, hOldBrush&
    Dim hBitmap&, hBitmap2&
    Dim hBrushBitmap&, hOldMemBitmap&
    Dim oldBkMode&, nWidth&, nHeight&
    Dim hMemDC&
    Dim rcRect As RECT
    Dim Bits(7) As Integer
    Dim i%, dl&
    Dim lpRect As RECT
    
    lpRect.Left = rctLeft
    lpRect.Top = rctTop
    lpRect.Right = rctLeft + rctWidth
    lpRect.Bottom = rctTop + rctHeight
    
    For i% = 0 To 6 Step 2
        Bits(i%) = &H55
        Bits(i% + 1) = &HAA
    Next i%
    nHeight = lpRect.Bottom - lpRect.Top
    nWidth = lpRect.Right - lpRect.Left
    
    hBrushBitmap = CreateBitmap(8, 8, 1, 1, Bits(0))
    hBrush = CreatePatternBrush(hBrushBitmap)
    
    hBitmap = CreateCompatibleBitmap(m_hdc, nWidth, nHeight)
    hMemDC = CreateCompatibleDC(m_hdc)
    
    hOldMemBitmap = SelectObject(hMemDC, hBitmap)
    
    rcRect.Right = nWidth
    rcRect.Bottom = nHeight

    dl& = FillRectAPI(hMemDC, rcRect, hBrush)
    dl& = BitBlt(hMemDC, 0, 0, nWidth, nHeight, m_hdc, lpRect.Left, lpRect.Top, SRCAND)
    
    crOldTextColor = SetTextColor(m_hdc, 0)
    crOldBkColor = SetBkColor(m_hdc, RGB(0, 0, 0))
    dl& = SetBkMode(m_hdc, OPAQUE)
    
    hOldBrush = SelectObject(m_hdc, hBrush)
    dl& = FillRectAPI(m_hdc, lpRect, hBrush)
    dl& = BitBlt(m_hdc, lpRect.Left, lpRect.Top, nWidth, nHeight, hMemDC, 0, 0, SRCPAINT)
    

    dl& = SetBkMode(m_hdc, oldBkMode)
    dl& = SetBkColor(m_hdc, crOldBkColor)
    dl& = SetTextColor(m_hdc, crOldTextColor)

    dl& = SelectObject(hMemDC, hOldMemBitmap)
    dl& = DeleteObject(hBitmap)
    dl& = DeleteDC(hMemDC)
    
    dl& = DeleteObject(hBrushBitmap)
    dl& = SelectObject(m_hdc, hOldBrush)
    dl& = DeleteObject(hBrush)
    
End Sub

Public Sub DrawImage(rctLeft As Long, ByVal rctTop As Long, ByVal rctWidth As Long, ByVal rctHeight As Long, picImage As Picture, oPalette As Object, Optional ByVal clrMaskColor As OLE_COLOR = -1, Optional ByVal clrBackColor As OLE_COLOR = -1, Optional ByVal bEnabled As Integer = True)
    
    Dim lhbmMemory As Long
    Dim lhbmMemoryOld As Long
    Dim lhdcMem As Long 'HDC
    Dim lBackColor As Long
    Dim udtPictureFrame As RECT
    Dim udtPictureRect As RECT
    Dim bUseMask As Boolean
    Dim lhPal As Long
    Dim lhPalOld As Long
    Dim lhbrBack As Long
    Dim bHavePalette As Boolean
    Dim oDrawTool As New clsDrawPictures
    
    On Error GoTo DrawImage_Error
    
    udtPictureRect.Left = 0
    udtPictureRect.Top = 0
    udtPictureRect.Right = rctWidth
    udtPictureRect.Bottom = rctHeight
    udtPictureFrame.Left = rctLeft
    udtPictureFrame.Top = rctTop
    udtPictureFrame.Right = rctLeft + rctWidth
    udtPictureFrame.Bottom = rctTop + rctHeight
    
    On Error Resume Next
    'Error will occur if the Ambient.Palette is not supported
    bHavePalette = CBool(oPalette)   '(Not oPalette Is Nothing)  '(Not UserControl.Ambient.Palette Is Nothing)
    If Err.Number <> 0 Then bHavePalette = False
    Err.Clear
    If bHavePalette Then
        'If the Palette or hPal property fails
        'resume next and use the halftone palette
        lhPal = oPalette.hPal   'UserControl.Ambient.Palette.hPal
        If lhPal = 0 Then lhPal = mlhHalftonePal
        Err.Clear
    Else
        lhPal = mlhHalftonePal    'If there is no specified palette
                                  'use the halftone palette.
    End If
    On Error GoTo DrawImage_Error
    
    'Create memory DC and bitmap to do all of the painting work
    lhdcMem = CreateCompatibleDC(m_hdc)   'UserControl.hdc)
    lhbmMemory = CreateCompatibleBitmap(m_hdc, udtPictureRect.Right, udtPictureRect.Bottom)  '(UserControl.hdc, mudtButtonRect.Right, mudtButtonRect.Bottom)
    lhbmMemoryOld = SelectObject(lhdcMem, lhbmMemory)
    lhPalOld = SelectPalette(lhdcMem, lhPal, True)
    RealizePalette lhdcMem
    
    'fill the memory DC with the background color of the screen dc
    If clrBackColor = -1 Then
        clrBackColor = GetBkColor(m_hdc)
    End If
    OleTranslateColor clrBackColor, 0, lBackColor 'UserControl.BackColor, 0, lBackColor
    SetBkColor lhdcMem, lBackColor
    lhbrBack = CreateSolidBrush(lBackColor)
    FillRectAPI lhdcMem, udtPictureRect, lhbrBack  'mudtButtonRect, lhbrBack
    If Not picImage Is Nothing Then   'Not m_picPictured Is Nothing Then
        If picImage.Type = vbPicTypeBitmap Then   'If m_picPictured.Type = vbPicTypeBitmap Then
            If clrMaskColor <> -1 Then bUseMask = True
        End If
        If Not bEnabled Then
            'If button is disabled draw disabled picture on memory dc
            oDrawTool.DrawDisabledPicture lhdcMem, picImage, udtPictureRect.Left, udtPictureRect.Top, udtPictureRect.Left + udtPictureRect.Right, udtPictureRect.Top + udtPictureRect.Bottom, lBackColor, bUseMask, clrMaskColor, lhPal 'lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lBackColor, bUseMask, m_clrMaskColor, lhPal
        ElseIf bUseMask Then
            'if using mask color draw transparent bitmap on memory dc
            oDrawTool.DrawTransparentBitmap lhdcMem, picImage, udtPictureRect.Left, udtPictureRect.Top, udtPictureRect.Left + udtPictureRect.Right, udtPictureRect.Top + udtPictureRect.Bottom, clrMaskColor, lhPal 'lhdcMem, picImage, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, m_clrMaskColor, lhPal
        Else
            'otherwise draw picture with no effects on button
            If picImage.Type = vbPicTypeBitmap Then   'If m_picPictured.Type = vbPicTypeBitmap Then
                oDrawTool.DrawBitmapToHDC lhdcMem, picImage, udtPictureRect.Left, udtPictureRect.Top, udtPictureRect.Left + udtPictureRect.Right, udtPictureRect.Top + udtPictureRect.Bottom, lhPal 'lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lhPal
            ElseIf picImage.Type = vbPicTypeIcon Then   'ElseIf m_picPictured.Type = vbPicTypeIcon Then
                DrawIcon lhdcMem, udtPictureRect.Left, udtPictureRect.Top, picImage.Handle   'm_picPictured.Handle
            End If
        End If
    End If
    
DrawImage:
    BitBlt m_hdc, udtPictureFrame.Left, udtPictureFrame.Top, udtPictureFrame.Right, udtPictureFrame.Bottom, lhdcMem, 0, 0, vbSrcCopy   'UserControl.hdc, 0, 0, udtPictureFrame.Right, udtPictureFrame.Bottom, lhdcMem, 0, 0, vbSrcCopy

DrawImageCleanUp:
    DeleteObject lhbrBack
    SelectPalette lhdcMem, lhPalOld, True
    RealizePalette (lhdcMem)
    DeleteObject SelectObject(lhdcMem, lhbmMemoryOld)
    DeleteDC lhdcMem
Exit Sub

DrawImage_Error:
    Select Case Err.Number
        Case giOBJECT_VARIABLE_NOT_SET
            Resume DrawImage
        Case giINVALID_PICTURE
            Resume DrawImage
        Case Else
            Resume DrawImageCleanUp
    End Select
End Sub

Private Sub Class_Initialize()

    m_lFlatLineColor = GetSysColor(COLOR_BTNSHADOW)
    
End Sub

Private Sub Class_Terminate()
    Dim hBmp As Long
    
    If mlhHalftonePal <> 0 Then _
      DeleteObject mlhHalftonePal

End Sub

Public Sub TileBitmap(picSource As Picture, _
                                    lLeft As Long, _
                                    lTop As Long, _
                                    lWidth As Long, _
                                    lHeight As Long, _
                                    lDestLeft As Long, _
                                    lDestTop As Long, _
                                    lDestWidth As Long, _
                                    lDestHeight As Long)
    
    Dim oDrawTool As New clsDrawPictures
    
    oDrawTool.TileBitmapToHDC m_hdc, picSource, lLeft, lTop, lWidth, lHeight, lDestLeft, lDestTop, lDestWidth, lDestHeight, mlhHalftonePal
    
End Sub

Public Property Get BackColor() As OLE_COLOR
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
    'return the current background color
    BackColor = m_MemoryDC.BackColor
End Property 'BackColor Get

Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
    Dim lColor As Long
    
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
   OleTranslateColor lBackColor, 0, lColor
   lBackColor = lColor
    
   'set the new background color
   m_MemoryDC.BackColor = lBackColor
   
End Property 'BackColor Let

Public Property Get FlatLineColor() As OLE_COLOR
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
    'return the current background color
    FlatLineColor = m_lFlatLineColor
End Property 'FlatLineColor Get

Public Property Let FlatLineColor(ByVal lFlatLineColor As OLE_COLOR)
    Dim lColor As Long
    
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
    OleTranslateColor lFlatLineColor, 0, lColor
    lFlatLineColor = lColor
    
    'set the new background color
    m_lFlatLineColor = lFlatLineColor
End Property 'FlatLineColor Let

Public Property Get ForeColor() As OLE_COLOR
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
    'return the current Text color
    ForeColor = m_MemoryDC.TextColor
End Property 'ForeColor Get

Public Property Let ForeColor(ByVal lForeColor As OLE_COLOR)
    'assert that we are initialized
    Debug.Assert m_hdc <> 0
    
    Dim lColor As Long
    
    OleTranslateColor lForeColor, 0, lColor
    lForeColor = lColor
    'set the new text color
    m_MemoryDC.TextColor = lForeColor
End Property 'ForeColor Let

Public Sub DrawStart(lhdc As Long, lWidth As Long, lHeight As Long, Optional ByVal bCopySourceDC As Boolean = False)

    m_MemoryDC.Attach lhdc, lWidth, lHeight, bCopySourceDC
    m_hdc = m_MemoryDC.hdc
    
End Sub

Public Sub DrawStop(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)

    m_MemoryDC.CopyToHdc lLeft, lTop, lWidth, lHeight
    
End Sub

'Set rgbColor to -1 if you wish the rectangle not to be filled in
Public Sub FillRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, rgbColor As Long, Optional sCaption As String = "", Optional CaptionAlign As Integer = caCenterCenter)

    Dim nX As Long              'X for drawing caption text
    Dim nY As Long              'Y for drawing caption text
    Dim rc As RECT              'rect struct to pass to the GDI
    Dim szTextExtent As Size    'pixel size of caption
    Dim lColor As Long
    Dim lOrgColor As Long
    
    'assert that we've been initialized already
    'and check the inputs
    Debug.Assert m_MemoryDC.hdc <> 0
    
    If rgbColor < 0 Then
        OleTranslateColor rgbColor, 0, lColor
        rgbColor = lColor
    End If
    
    'set the back color of the DC to the color desired
    lOrgColor = m_MemoryDC.BackColor
    m_MemoryDC.BackColor = rgbColor
    
    'calculate the caption X and Y (centered) if the caption
    'is not an empty string
    If Len(sCaption) > 0 Then
        
        'get the pixel width of the Caption
        GetTextExtentPoint32 m_MemoryDC.hdc, sCaption, Len(sCaption), szTextExtent
        
        'determine the X value based on the alignment chosen
        Select Case CaptionAlign
            Case caCenterCenter, caTopCenter, caBottomCenter
                nX = ((nWidth - szTextExtent.cx) \ 2) + nLeft
            
            Case caCenterRight, caTopRight, caBottomright
                nX = nWidth - BORDER_WIDTH - szTextExtent.cx + nLeft
            
            Case caCenterLeft, caTopLeft, caBottomLeft
                nX = nLeft + BORDER_WIDTH
        End Select
        
        'determine the Y value base on the alignment chosen
        Select Case CaptionAlign
            Case caCenterCenter, caCenterRight, caCenterLeft
                nY = ((nHeight - szTextExtent.cy) \ 2) + nTop
                
            Case caTopCenter, caTopLeft, caTopRight
                nY = nTop + BORDER_WIDTH
            
            Case caBottomCenter, caBottomLeft, caBottomright
                nY = nHeight - BORDER_WIDTH - szTextExtent.cy + nTop
                
        End Select
    End If 'caption is not ""
    
    'assign the input values to the rect struct
    rc.Left = nLeft
    rc.Top = nTop
    rc.Right = nWidth + nLeft
    rc.Bottom = nHeight + nTop
    
    'ExtTextOut is one of the fastest ways to fill a rectangular
    'area on a DC and is used here to fill our rect
    ExtTextOut m_MemoryDC.hdc, nX, nY, ETO_OPAQUE + ETO_CLIPPED, rc, sCaption, Len(sCaption), 0
    
    m_MemoryDC.BackColor = lOrgColor
    
End Sub 'FillRect()

'Public Sub Draw3DRect(ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional sCaption As String = "", Optional CaptionAlign As CaptionAlignments = caCenterCenter, Optional Appearance As Appearances = Raised)
'
'    Call m_MemoryDC.Draw3DRect(nLeft, nTop, nWidth, nHeight, sCaption, CaptionAlign, Appearance)
'
'End Sub 'Draw3DRect()

'----------------------------------------------------------------------
' InvertRect()
'----------------------------------------------------------------------
' Purpose:  To invert a particular rect on the bitmap
' Inputs:   The area to invert
' Outputs:  none
'----------------------------------------------------------------------
Public Sub InvertRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long)
    Dim rc As RECT
    
    rc.Left = nLeft
    rc.Top = nTop
    rc.Right = nLeft + nWidth
    rc.Bottom = nTop + nHeight
    
    InvertRectAPI m_MemoryDC.hdc, rc
End Sub 'InvertRect()

Public Property Get Font() As StdFont
    Set Font = m_MemoryDC.Font
End Property 'Font Get

Public Property Set Font(oFont As StdFont)
    Set m_MemoryDC.Font = oFont
End Property 'Font Set

'----------------------------------------------------------------------
' Draw3DRect()
'----------------------------------------------------------------------
' Purpose:  To draw a 3D looking rectangle on the off-screen DC
' Inputs:   The rectangle to make 3d and optionally a caption to
'           display centered in the rect
' Outputs:  none
'----------------------------------------------------------------------
Public Sub Draw3DRect(ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional sCaption As String = "", Optional CaptionAlign As Integer = caCenterCenter, Optional Appearance As Appearances = Raised)
    Dim rgbLowerRight As Long   'color to use for the lower right
    Dim rgbUpperLeft As Long    'color to use for the upper left
    
    'assert that we've been initialized already
    'and check the inputs
    Debug.Assert m_MemoryDC.hdc <> 0
    
    'if we haven't gotten the system colors for 3d effects
    'get them first
    If mrgb3DFace = 0 Then
        mrgb3DFace = GetSysColor(COLOR_BTNFACE)
        mrgb3DHighlight = GetSysColor(COLOR_BTNHIGHLIGHT)
        mrgb3DShadow = GetSysColor(COLOR_BTNSHADOW)
    End If
            
    'set the lower-right and upper-left colors based on the
    'desired appearance
    Select Case Appearance
        Case Flat
            rgbLowerRight = m_lFlatLineColor
            rgbUpperLeft = m_lFlatLineColor
        
        Case Raised
            rgbLowerRight = mrgb3DShadow
            rgbUpperLeft = mrgb3DHighlight
            
        Case Sunken
            rgbLowerRight = mrgb3DHighlight
            rgbUpperLeft = mrgb3DShadow
        
        Case Selected
            rgbLowerRight = mrgb3DHighlight
            rgbUpperLeft = vbBlack
    End Select
    
    'fill the rect with the shadow color (or hightlight if sunken)
    If Appearance <> NoLines Then FillRect nLeft, nTop, nWidth, nHeight, rgbLowerRight
    
    'now pull the right and bottom edges in by 1 pixel
    nWidth = nWidth - 1
    nHeight = nHeight - 1
    
    'fill the rect with the 3d highlight color (or shadow if sunken)
    If Appearance <> NoLines Then FillRect nLeft, nTop, nWidth, nHeight, rgbUpperLeft
    
    'finally pull in the left and top edges by 1 pixel
    nLeft = nLeft + 1
    nTop = nTop + 1
    nWidth = nWidth - 1
    nHeight = nHeight - 1
    
    'change the color to the 3d face color
    'and fill the rect passing the desired caption
    FillRect nLeft, nTop, nWidth, nHeight, m_MemoryDC.BackColor, sCaption, CaptionAlign
    
    'if the appearance setting was Selected, invert the rect
    If Appearance = Selected Then
        InvertRect nLeft, nTop, nWidth, nHeight
    End If 'appearance = selected
    
End Sub 'Draw3dRect

'----------------------------------------------------------------------
' 3D Colors Properties
'----------------------------------------------------------------------
' Purpose:  To return the RGB values for 3d colors
'----------------------------------------------------------------------
Public Property Get ThreeDFaceColor() As Long
    ThreeDFaceColor = mrgb3DFace
End Property

Public Property Get ThreeDHighlightColor() As Long
    ThreeDHighlightColor = mrgb3DHighlight
End Property

Public Property Get ThreeDShadowColor() As Long
    ThreeDShadowColor = mrgb3DShadow
End Property

Public Sub CopyDC(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)

    m_MemoryDC.CopyToHdc lLeft, lTop, lWidth, lHeight
    
End Sub

