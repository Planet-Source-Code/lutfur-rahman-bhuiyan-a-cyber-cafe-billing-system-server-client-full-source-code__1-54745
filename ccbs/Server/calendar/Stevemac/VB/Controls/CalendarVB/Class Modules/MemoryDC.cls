VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' CMemoryDC.cls
'----------------------------------------------------------------------
' Implementation file for MemoryDC class
' This class represents an off screen DC that is useful
' for enabling flick-free and smooth repainting of things
' like controls.
'
' There are also a couple helper methods that do interesting
' GDI operations like drawing 3d rectangles and fast rectangles
'----------------------------------------------------------------------
' Copyright (c) 1996, Microsoft Corporation
'----------------------------------------------------------------------

Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByVal lpDx As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)

Private Const ETO_OPAQUE = 2
Private Const ETO_CLIPPED = 4

Private Const SRCCOPY = &HCC0020

'======================================================================
' Private Constants
'======================================================================
Private Const BORDER_WIDTH As Long = 0

'======================================================================
' Private Data Members
'======================================================================
Private mhdcWork As Long        'off-screen HDC
Private mhdcCtl As Long         'actual HDC of the control
Private mhbmpOld As Long        'hBmp of the old bitmap in the off-sceen DC
Private mfntCurrent As IFont    'font to use when drawing text
Private mhfntOld As Long        'hFont of the old font in the off-screen dc

Private mcxCtlWidth As Long     'width of the control's surface
Private mcyCtlHeight As Long    'height of the control's surface

'======================================================================
' Initialize and Terminate Events
'======================================================================

'----------------------------------------------------------------------
' Class_Terminate()
'----------------------------------------------------------------------
' Purpose:  Called when the object is destroyed--do clean-up work
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub Class_Terminate()
    Me.Detach
End Sub 'Class_Terminate()

'======================================================================
' Public Methods and Properties
'======================================================================

'----------------------------------------------------------------------
' BackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  To get and let the current background color of the DC
'----------------------------------------------------------------------
Public Property Get BackColor() As Long
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'return the current background color
    BackColor = GetBkColor(mhdcWork)
End Property 'BackColor Get

Public Property Let BackColor(rgbNew As Long)
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'set the new background color
    SetBkColor mhdcWork, rgbNew
End Property 'BackColor Let

'----------------------------------------------------------------------
' TextColor Get/Let
'----------------------------------------------------------------------
' Purpose:  To get and let the current text color of the DC
'----------------------------------------------------------------------
Public Property Get TextColor() As Long
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'return the current Text color
    TextColor = GetTextColor(mhdcWork)
End Property 'TextColor Get

Public Property Let TextColor(rgbNew As Long)
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'set the new text color
    SetTextColor mhdcWork, rgbNew
End Property 'TextColor Let

'----------------------------------------------------------------------
' Font Get/Set
'----------------------------------------------------------------------
' Purpose:  To get and set the current font to use on the DC
'----------------------------------------------------------------------
Public Property Get Font() As StdFont
    'just return the reference we currently are holding
    Set Font = mfntCurrent
End Property 'Font Get

Public Property Set Font(NewFont As StdFont)
    'make sure we're initialized first
    'must call Initialize before setting the font!
    Debug.Assert (mhdcWork <> 0)
    
    'below we will set a local member variable equal to the
    'object passed in.  Even though the type passed in is a
    'StdFont, our member variable is of type IFont.  A StdFont
    'can be casted (QI) to an IFont, and the IFont interface gives
    'us access to the hFont property, which we need in order to
    'set the current font of the off-screen device context.
        
    'if this is the first time the font is being set,
    'grab the existing hFont handle so we can restore it
    'before deleting the DC
    If mfntCurrent Is Nothing Then
        Set mfntCurrent = NewFont
        mhfntOld = SelectObject(mhdcWork, mfntCurrent.hFont)
    Else
        Set mfntCurrent = NewFont
        
        'if this is being set to Nothing, restore the old font
        If mfntCurrent Is Nothing Then
            SelectObject mhdcWork, mhfntOld
        Else
            SelectObject mhdcWork, mfntCurrent.hFont
        End If 'new font is nothing
        
    End If 'first time setting font

End Property 'Font Set

'----------------------------------------------------------------------
' Initialize()
'----------------------------------------------------------------------
' Purpose:  To initialize the object with the screen DC from which we
'           will create the off-screen DC
' Inputs:   The user control
' Outputs:  none
'----------------------------------------------------------------------
Public Sub Attach(CtlHdc As Long, CtlWidth As Long, CtlHeight As Long, Optional ByVal bCopySourceDC As Boolean = False)
    Dim hBmp As Long
    
    'assert that the inputs are valid
    'and that we haven't already called Initialize
    Debug.Assert (CtlHdc <> 0)
    Debug.Assert (mhdcWork = 0)
    
    'store the HDC of the control in our private variable
    mhdcCtl = CtlHdc
    
    'capture the width and height of the control
    mcxCtlWidth = CtlWidth
    mcyCtlHeight = CtlHeight
    
    'create the off-sceen DC
    mhdcWork = CreateCompatibleDC(mhdcCtl)
    
    'create a compatible bitmap from the control DC
    'that is the same size as the control itself
    hBmp = CreateCompatibleBitmap(mhdcCtl, mcxCtlWidth, mcyCtlHeight)
    
    'select that new bitmap into the off-screen DC
    'and save the old bitmap handle so we can reselect
    'it back in before we destroy the off-screen DC
    mhbmpOld = SelectObject(mhdcWork, hBmp)
    Me.BackColor = GetBkColor(mhdcCtl)
    Me.TextColor = GetTextColor(mhdcCtl)
    
    If bCopySourceDC Then
        'use bitblt to blast the contents of the control to the memory DC
        BitBlt mhdcWork, 0, 0, mcxCtlWidth, mcyCtlHeight, mhdcCtl, 0, 0, SRCCOPY
    Else
        'Clear the Memory DC
        Call FillRect(0, 0, mcxCtlWidth, mcyCtlHeight, Me.BackColor)
    End If
    
End Sub 'Initialize()

'----------------------------------------------------------------------
' CopyToHdc()
'----------------------------------------------------------------------
' Purpose:  Blasts the contents of the memory DC to the control's
'           attached hDC
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Public Sub CopyToHdc(Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1)
    If Width = -1 Then Width = mcxCtlWidth
    If Height = -1 Then Height = mcyCtlHeight
    
    'use bitblt to blast the contents of the off-screen dc to the control
    BitBlt mhdcCtl, Left, Top, Width, Height, mhdcWork, _
            Left, Top, SRCCOPY
End Sub 'CopyToHdc()

'----------------------------------------------------------------------
' CopyFromHdc()
'----------------------------------------------------------------------
' Purpose: Copies a specified rectangle from a source hDC to our
'          memory hDC
' Inputs:   hdcSource, [Left], [Top], [Width], [Height]
' Outputs:  none
'----------------------------------------------------------------------
Public Sub CopyFromHdc(ByVal hdcSource As Long, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1)
    If Width = -1 Then Width = mcxCtlWidth
    If Height = -1 Then Height = mcyCtlHeight
    
    'use bitblt to blast the contents of the Source  dc to the memory dc
    BitBlt mhdcWork, Left, Top, Width, Height, hdcSource, _
            Left, Top, SRCCOPY
End Sub 'CopyFromHdc()

Public Property Get hdc() As Long
Attribute hdc.VB_UserMemId = 0
    hdc = mhdcWork
End Property 'hDC Get

'----------------------------------------------------------------------
' Detach Method
'----------------------------------------------------------------------
' Purpose:  Free all resources that are being used
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Public Sub Detach()
    
    Dim hBmp As Long
    
        'select the old font back into the off-screen dc
        If mhdcWork <> 0 And mhfntOld <> 0 Then _
          SelectObject mhdcWork, mhfntOld
        
        'select the old bitmap back into the off-screen DC
        If mhdcWork <> 0 And mhbmpOld <> 0 Then _
          hBmp = SelectObject(mhdcWork, mhbmpOld)
        
        'delete the bitmap we were using
        If hBmp <> 0 Then DeleteObject hBmp
        
        'and now delete the off-screen DC to totally clean up
        If mhdcWork <> 0 Then DeleteDC mhdcWork
        mhdcWork = 0
        
End Sub
'----------------------------------------------------------------------
' FillRect()
'----------------------------------------------------------------------
' Purpose:  To fill a rectangle on the off-screen DC with a specified
'           color in a fast way
' Inputs:   The rectangle to fill and color to use
' Outputs:  none
'----------------------------------------------------------------------
Private Sub FillRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, rgbColor As Long, Optional sCaption As String = "", Optional CaptionAlign As CaptionAlignments = caCenterCenter)
    Dim nX As Long              'X for drawing caption text
    Dim nY As Long              'Y for drawing caption text
    Dim rc As RECT              'rect struct to pass to the GDI
    Dim szTextExtent As Size    'pixel size of caption
    Dim lColor As Long
    Dim lOrgColor As Long
    
    'assert that we've been initialized already
    'and check the inputs
    Debug.Assert mhdcWork <> 0
    
    If rgbColor < 0 Then
        OleTranslateColor rgbColor, 0, lColor
        rgbColor = lColor
    End If
    
    'set the back color of the DC to the color desired
    lOrgColor = Me.BackColor
    Me.BackColor = rgbColor
    
    'calculate the caption X and Y (centered) if the caption
    'is not an empty string
    If Len(sCaption) > 0 Then
        
        'get the pixel width of the Caption
        GetTextExtentPoint32 mhdcWork, sCaption, Len(sCaption), szTextExtent
        
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
    ExtTextOut mhdcWork, nX, nY, ETO_OPAQUE + ETO_CLIPPED, rc, sCaption, Len(sCaption), 0
    
    Me.BackColor = lOrgColor
    
End Sub 'FillRect

