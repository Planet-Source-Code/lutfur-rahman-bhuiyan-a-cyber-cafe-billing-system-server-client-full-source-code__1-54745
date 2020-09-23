VERSION 5.00
Begin VB.UserControl CalendarVB 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   KeyPreview      =   -1  'True
   PropertyPages   =   "CalendarVB.ctx":0000
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ToolboxBitmap   =   "CalendarVB.ctx":0044
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   30
   End
   Begin VB.TextBox ctlFocus 
      Height          =   285
      Left            =   -1500
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   645
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   645
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   2340
      Picture         =   "CalendarVB.ctx":013E
      Top             =   1710
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuNextPeriod 
         Caption         =   "Next Period"
      End
      Begin VB.Menu mnuPrevPeriod 
         Caption         =   "Previous Period"
      End
      Begin VB.Menu mnuCalendarType 
         Caption         =   "Calendar Types"
         Begin VB.Menu mnuCalTypeMonth 
            Caption         =   "Month"
         End
         Begin VB.Menu mnuCalTypePeriod 
            Caption         =   "Period"
         End
         Begin VB.Menu mnuCalTypeWeek 
            Caption         =   "Week"
         End
      End
   End
End
Attribute VB_Name = "CalendarVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================================
'  Name [CalendarVB.ctl]
'
'  Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        11/12/1997
'===========================================================================================

'$Runtime Dependencies: VB Runtime support files

'$DesignTime Dependencies:
'   CDraw.cls                   clsDrawPictures.cls         MemoryDC.cls
'   FctlAbout.frm               appear.pag                  FToolTip.frm
'   CCalendarVBPeriod.cls       CCalendarVBPeriods.cls      CCalendarVBMethods.cls
'   CCalendarVBVars.cls         CLicense.cls                frmRegister.frm

'===========================================================================================
'  Usage Notes:
'
'===========================================================================================
'Properties
'    *   ActiveDayFont              *   ActiveDayFontBold           *   ActiveDayFontItalic
'    *   ActiveDayFontName          *   ActiveDayFontSize           *   AutoPaint
'    *   BackColor                  *   LineStyle                   *   YearStartPlacement
'    *   CalendarType               *   CurrentPeriodBackColor      *   CurrentPeriodForeColor
'    *   DateTipFormat              *   DateValue                   *   ActiveDayForeColor
'    *   DayHeaderBackColor         *   DayHeaderForeColor          *   DayHeaderFormat
'    *   DayHeaderFont              *   DayHeaderFontBold           *   DayHeaderFontItalic
'    *   DayHeaderFontName          *   DayHeaderFontSize           *   DaysFont
'    *   DaysFontBold               *   DaysFontItalic              *   DaysFontName
'    *   DaysFontSize               *   ShowDayHeader               *   Enabled
'    *   ExtraWeekPlacement         *   FlatLineColor               *   PeriodRows
'    *   Periods                    *   PeriodValue                 *   PeriodYear
'    *   Picture                    *   PopupMenuDisabled           *   PostPeriodBackColor
'    *   PostPeriodForeColor        *   PrePeriodBackColor          *   PrePeriodForeColor
'    *   ShowDateTip                *   ShowPeriodList              *   ShowYearList
'    *   FirstDayOfWeek             *   URLPicture                  *   YearBegin
'    *   YearEnd
'
'Hidden Properties
'   *   URLLocation
'
'Events
'    *   Click                      *   DateChange                  *   DblClick
'    *   ErrorEvent                 *   WillChangeDate
'
'Methods
'    *   AboutBox                   *   PeriodDefault               *   Refresh
'===========================================================================================

Option Explicit

'===========================================================================================
'API USER DEFINED TYPES
'===========================================================================================
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'===========================================================================================
'API DECLARES
'===========================================================================================
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function GetUpdateRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_CLOSE = &H10
Private Const WM_ERASEBKGND = &H14
Private Const WM_PAINT = &HF

'===========================================================================================
'PRIVATE USER DEFINED TYPES
'===========================================================================================
Private Type VBRect
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

'===========================================================================================
'PUBLIC ENUMS
'===========================================================================================
Public Enum CalendarTypes
    calMonth = 0
    calPeriod
    calWeek
End Enum 'CalendarTypes

Public Enum CalendarLineTypes
    calNoLine = 0
    cal3D
    calFlat
    calSunken
End Enum 'CalendarLineTypes

Public Enum DaysOfTheWeek
    calSunday = 1
    calMonday
    calTuesday
    calWednesday
    calThursday
    calFriday
    calSaturday
End Enum 'DaysOfTheWeek

Public Enum CalYearStartPlacement
    calPreviousYear = 0
    calCurrentYear
End Enum 'StartWeek

Public Enum DayHeaderFormats
    calOneLetterName = 0
    calTwoLetterName
    calThreeLetterName
    calFullName
End Enum 'DayHeaderFormats

Public Enum DayNumberAlignments
    calCenterCenter
    calCenterLeft
    calCenterRight
    calTopCenter
    calTopLeft
    calTopRight
    calBottomCenter
    calBottomLeft
    calBottomright
End Enum 'DayNumberAlignments

Public Enum ExtraWeekPlacements
    calFirstPeriod = 0
    calLastPeriod
End Enum 'ExtraWeekPlacements

Public Enum CalErrorActions
    calAbort = 1
    calRetry
    calIgnore
End Enum 'CalErrorActions

Public Enum CalErrorNumbers
    calUnhandledError = vbObjectError + 513
    calInvalidDateRange = vbObjectError + 514
    calInvalidPropertyValue = vbObjectError + 515
End Enum 'CalErrorNumbers


'===========================================================================================
'INTERNAL CONSTANTS
'===========================================================================================
Private Const DEF_LEFT_MARGIN = 2           'The left margin starting point for drawing the calendar
Private Const DEF_TOP_MARGIN = 28           'The top margin starting point for drawing the calendar. Also accounts for the comboboxes.
Private Const DEF_CALENDAR_ROWS = 6
Private Const DEF_CELL_BACK_COLOR = 0       'The cell constants are used to access
Private Const DEF_CELL_LEFT = 1             'These are array positions for variant array stored in m_vaCellLocations array
Private Const DEF_CELL_TOP = 2
Private Const DEF_CELL_WIDTH = 3
Private Const DEF_CELL_HEIGHT = 4
Private Const DEF_CELL_FORE_COLOR = 5
Private Const DEF_CELL_UBOUND = 5
Private Const DEF_CALENDAR_COLS = 7         'The number of calendar columns to display
Private Const LIC_CLSID = "12427960-4ED2-11D1-9037-00A0C91EF7D6"    'License ID
Private Const LIC_KEY = "411-0948347"                               'License Key
Private Const VERSION_NUMBER = "1.0.06"

'===========================================================================================
'INTERNAL VARIABLES
'===========================================================================================
Private m_bActive               As Boolean                  'Used to indicate whether the combobox clicks should be acted upon
Private m_bClearURLOnly         As Boolean                  'Used for knowing which picture property is being used
Private m_bClearPictureOnly     As Boolean                  'Same as above
Private m_bDesign               As Boolean                  'A little faster to check this than the Ambient UserMode property
Private m_bToolTipVisible       As Boolean                  'Indicates whether the tooltip is currently visible
Private m_nCellHeight           As Integer                  'The height in pixels for a calendar cell
Private m_nCellWidth            As Integer                  'The width in pixels for a calendar cell
Private m_nCol                  As Integer                  'The current calendar cell column
Private m_nDisabledLeft         As Integer                  'These four variables hold
Private m_nLastCol              As Integer                  'The calendar cells previous column location
Private m_nLastRow              As Integer                  'The calendar cells previous row location
Private m_nRow                  As Integer                  'The current calendar cell row
Private m_sRegistered           As String                   'Used with the registration code
Private m_sEnvironment          As String                   'Used with the registration code
Private m_vaCellLocations()     As Variant                  'Deminsioned as (m_nPeriodRows, DEF_CALENDAR_COLS)
Private m_RefreshDC             As New CMemoryDC            'Used to hold an image of the control for refreshing the screen without having to redraw the calendar each time
Private m_Methods               As CCalendarVBMethods       'Some of our calendar control procs and methods
Private m_udtDisabledRect       As VBRect                   'Holds the dimensions for the disabled rectangle shading
Private m_udtFocusArea          As RECT                     'current focus area
Private m_ToolTip               As Form                     'Our tooltip object
Private m_Vars                  As New CCalendarVBVars      'Holds our DateValue, PeriodValue, PeriodYear, Periods, and some useful date values

'===========================================================================================
'PROPERTY NAME CONSTANTS
'===========================================================================================
Private Const pnActiveDayFont = "ActiveDayFont"
Private Const pnActiveDayFontBold = "ActiveDayFontBold"
Private Const pnActiveDayFontItalic = "ActiveDayFontItalic"
Private Const pnActiveDayFontName = "ActiveDayFontName"
Private Const pnActiveDayFontSize = "ActiveDayFontSize"
Private Const pnActiveDayForeColor = "ActiveDayForeColor"
Private Const pnAutoPaint = "AutoPaint"
Private Const pnBackColor = "BackColor"
Private Const pnBorderStyle = "BorderStyle"
Private Const pnCalendarType = "CalendarType"
Private Const pnCurrentPeriodBackColor = "CurrentPeriodbackColor"
Private Const pnCurrentPeriodForeColor = "CurrentPeriodForeColor"
Private Const pnDateTipFormat = "DateTipFormat"
Private Const pnDayNumberAlignment = "DayNumberAlignment"
Private Const pnDayHeaderBackColor = "DayHeaderBackColor"
Private Const pnDayHeaderForeColor = "DayHeaderForeColor"
Private Const pnDayHeaderFormat = "DayHeaderFormat"
Private Const pnDayHeaderFont = "DayHeaderFont"
Private Const pnDayHeaderFontBold = "DayHeaderFontBold"
Private Const pnDayHeaderFontItalic = "DayHeaderFontItalic"
Private Const pnDayHeaderFontName = "DayHeaderFontName"
Private Const pnDayHeaderFontSize = "DayHeaderFontSize"
Private Const pnDaysFont = "DaysFont"
Private Const pnDaysFontBold = "DaysFontBold"
Private Const pnDaysFontItalic = "DaysFontItalic"
Private Const pnDaysFontName = "DaysFontName"
Private Const pnDaysFontSize = "DaysFontSize"
Private Const pnEnabled = "Enabled"
Private Const pnExtraWeekPlacement = "ExtraWeekPlacement"
Private Const pnFirstCurrentPeriod = "FirstDayOfWeek"
Private Const pnFlatLineColor = "FlatLineColor"
Private Const pnLineStyle = "LineStyle"
Private Const pnPeriodRows = "PeriodRows"
Private Const pnPeriods = "Periods"
Private Const pnPicture = "Picture"
Private Const pnPrePeriodBackColor = "PrePeriodBackColor"
Private Const pnPrePeriodForeColor = "PrePeriodforeColor"
Private Const pnPopupMenuDisabled = "PopupMenuDisabled"
Private Const pnPostPeriodBackColor = "PostPeriodBackColor"
Private Const pnPostPeriodForeColor = "PostPeriodforeColor"
Private Const pnShowDateTip = "ShowDateTip"
Private Const pnDisplayHeader = "ShowDayHeader"
Private Const pnShowPeriodList = "ShowPeriodList"
Private Const pnShowYearList = "ShowYearList"
Private Const pnURLLocation = "URLLocation"
Private Const pnURLPicture = "URLPicture"
Private Const pnYearBegin = "YearBegin"
Private Const pnYearEnd = "YearEnd"
Private Const pnYearStartPlacement = "YearStartPlacement"

'===========================================================================================
'CONTROL EVENTS
'===========================================================================================
Public Event Click(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
Attribute Click.VB_Description = "Event"
Public Event ErrorEvent(ByVal Number As Long, ByVal Source As String, ByVal Description As String, Action As Integer, Value As Variant)
Public Event DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
Attribute DateChange.VB_Description = "Event"
Public Event DblClick(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
Attribute DblClick.VB_Description = "Event"
Attribute DblClick.VB_MemberFlags = "200"
Public Event WillChangeDate(ByVal NewDate As Date, Cancel As Boolean)
Attribute WillChangeDate.VB_Description = "An event that gets fired before the date is changed allowing the programming to cancel the date change before it takes place."

'===========================================================================================
'CONTROL PROPERTIES
'===========================================================================================
Private m_bAutoPaint            As Boolean
Private m_nBorderStyle          As CalendarLineTypes
Private m_bPopupMenuDisabled    As Boolean
Private m_bShowDateTip          As Boolean
Private m_bShowDayHeader        As Boolean
Private m_bShowPeriodList       As Boolean
Private m_bShowYearList         As Boolean
Private m_nCalendarType         As Integer
Private m_nDayNumberAlignment   As Integer
Private m_nDayHeaderFormat      As DayHeaderFormats
Private m_nExtraWeekPlacement   As ExtraWeekPlacements
Private m_nFirstCurrentPeriod   As Integer
Private m_nLineStyle            As Appearances
Private m_nPeriodRows           As Integer
Private m_nYearBegin            As Integer
Private m_nYearEnd              As Integer
Private m_nYearStartPlacement   As Integer
Private m_sDateTipFormat        As String
Private m_sURLLocation          As String
Private m_sURLPicture           As String
Private m_ActiveDayFont         As New StdFont
Private m_DayHeaderFont         As New StdFont
Private m_DaysFont              As New StdFont
Private m_picImage              As StdPicture

'===========================================================================================
'CONTROL COLOR PROPERTIES
'===========================================================================================
Private m_oActiveDayForeColor       As OLE_COLOR
Private m_oBackColor                As OLE_COLOR
Private m_oCurrentPeriodBackColor   As OLE_COLOR
Private m_oCurrentPeriodForeColor   As OLE_COLOR
Private m_oDayHeaderBackColor       As OLE_COLOR
Private m_oDayHeaderForeColor       As OLE_COLOR
Private m_oFlatLineColor            As OLE_COLOR
Private m_oPrePeriodBackColor       As OLE_COLOR
Private m_oPrePeriodForeColor       As OLE_COLOR
Private m_oPostPeriodBackColor      As OLE_COLOR
Private m_oPostPeriodForeColor      As OLE_COLOR

'===========================================================================================
'DEFAULT CONTROL PROPERTY VALUES
'===========================================================================================
Private Const DEF_ACTIVE_DAY_FORECOLOR = &H200FF
Private Const DEF_AUTOPAINT = True
Private Const DEF_BORDERSTYLE = calNoLine
Private Const DEF_BACKCOLOR = SystemColorConstants.vbButtonFace
Private Const DEF_CALENDAR_TYPE = calMonth
Private Const DEF_CURRENT_PERIOD_BACKCOLOR = SystemColorConstants.vbButtonFace
Private Const DEF_CURRENT_PERIOD_FORECOLOR = &H800000
Private Const DEF_DATE_TIP_FORMAT = "Dddd Mmm dd, yyyy"
Private Const DEF_DAY_NUMBER_ALIGNMENT = DayNumberAlignments.calCenterCenter
Private Const DEF_DAY_HEADER_BACKCOLOR = &H80FFFF
Private Const DEF_DAY_HEADER_FORECOLOR = &H800000
Private Const DEF_DAY_HEADER_FORMAT = DayHeaderFormats.calThreeLetterName
Private Const DEF_EXTRA_WEEK_PLACEMENT = ExtraWeekPlacements.calLastPeriod
Private Const DEF_FIRST_DAY_OF_WEEK = calSunday
Private Const DEF_FLAT_LINE_COLOR = SystemColorConstants.vb3DDKShadow
Private Const DEF_LINE_STYLE = CalendarLineTypes.cal3D
Private Const DEF_PERIOD_ROWS = 6
Private Const DEF_PRE_PERIOD_BACKCOLOR = &H80FFFF
Private Const DEF_PRE_PERIOD_FORECOLOR = SystemColorConstants.vbButtonText
Private Const DEF_POPUP_MENU_DISABLED = False
Private Const DEF_POST_PERIOD_BACKCOLOR = &H80FFFF
Private Const DEF_POST_PERIOD_FORECOLOR = &H800000
Private Const DEF_SHOW_DATE_TIP = True
Private Const DEF_SHOW_DAY_HEADER = True
Private Const DEF_SHOW_PERIOD_LIST = True
Private Const DEF_SHOW_YEAR_LIST = True
Private DEF_YEAR_BEGIN As Long
Private DEF_YEAR_END   As Long
Private Const DEF_YEAR_START_PLACEMENT = CalYearStartPlacement.calPreviousYear

' Window style bit functions:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
    ) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long _
    ) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_BORDER = &H800000
Private Const WS_THICKFRAME = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const HWND_NOTOPMOST = -2

Private Sub cboPeriod_Click()

    'If were working with the combobox through code then we can exit
    If m_bActive = False Then Exit Sub
    PeriodValue = cboPeriod.ListIndex + 1
    
End Sub 'Event cboPeriod_Click()

Private Sub cboYear_Click()

    Dim bAutoPaintValue     As Boolean
    Dim nPeriodYear               As Integer
    Dim dtDate              As Date

    'If were working with the combobox through code then we can exit
    If m_bActive = False Then Exit Sub
    'Save the new year to the PeriodYear property and calculate
    'the new date based on the current period and the new year
    nPeriodYear = CInt(cboYear.Text)
    If nPeriodYear = 0 Then Exit Sub
    PeriodYear = nPeriodYear

End Sub 'Event cboYear_Click()

'----------------------------------------------------------------------
' ctlFocus_GotFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the main calendar area is to get focus.
'           We use a dummy control to capture focus since we are
'           just painting the calendar days and cannot set focus
'           to the entire user control.
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_GotFocus()
    'Removes the focus rectangle around the current date
    Call FocusRect
End Sub 'Event ctlFocus_GotFocus()

'----------------------------------------------------------------------
' ctlFocus_LostFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the main calendar area has lost focus.
'           We use a dummy control to capture focus since we are
'           just painting the calendar days and cannot set focus
'           to the entire user control.
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_LostFocus()
    'Draws the focus rectangle around the current date
    Call FocusRect
End Sub 'Event ctlFocus_LostFocus()

'----------------------------------------------------------------------
' ctlFocus_KeyDown Event
'----------------------------------------------------------------------
' Purpose:  Called when the user presses a key while the dummy control
'           has focus
' Inputs:   Which key and shift state
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_KeyDown(keycode As Integer, Shift As Integer)
    Dim bAutoPaint          As Integer
    Dim bShowDateTipOrg     As Boolean
    Dim nPeriodValue        As Integer
    Dim nWeeks              As Integer
    Dim dtTemp              As Date      'temp date for date arithmetic
    
    bShowDateTipOrg = m_bShowDateTip
    'Be sure that our tooltip is not visible
    If m_bShowDateTip Then
        Call DateTipDisplay(False)
        m_bToolTipVisible = False
    End If
    
    Select Case keycode
        Case vbKeyLeft
            dtTemp = DateValue
            If (Shift And vbShiftMask) > 0 Then
            'if shift is down, move by month
                dtTemp = DateAdd("m", -1, dtTemp)
            ElseIf (Shift And vbCtrlMask) > 0 Then
                'else if control is down, move by year
                dtTemp = DateAdd("yyyy", -1, dtTemp)
            Else
                'go back on day
                dtTemp = DateAdd("d", -1, dtTemp)
            End If
            'Update the display with the new datevalue
            If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
            Call ChangeValue(dtTemp)
        Case vbKeyRight
            dtTemp = DateValue
            'if shift is down, move by month
            If (Shift And vbShiftMask) > 0 Then
                dtTemp = DateAdd("m", 1, dtTemp)
            
            ElseIf (Shift And vbCtrlMask) > 0 Then
                'else if control is down, move by year
                dtTemp = DateAdd("yyyy", 1, dtTemp)
            Else
                'go forward one day
                dtTemp = DateAdd("d", 1, dtTemp)
            End If
            'Update the display with the new datevalue
            If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
            Call ChangeValue(dtTemp)
        Case vbKeyUp
            dtTemp = DateAdd("ww", -1, DateValue)
            'go one week back
            'Update the display with the new datevalue
            If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
            Call ChangeValue(dtTemp)
        Case vbKeyDown
            'go one week forward
            dtTemp = DateAdd("ww", 1, DateValue)
            'Update the display with the new datevalue
            If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
            Call ChangeValue(dtTemp)
        Case vbKeyHome
            'if control is down, go to first day of the year
            If (Shift And vbCtrlMask) > 0 Then
                PeriodValue = 1
                cboPeriod.ListIndex = 0
                Call FocusRect
            Else
                'go to the first day of the current month
                Call ChangeValue(m_Vars.PeriodStart)
            End If
        Case vbKeyEnd
            'if control is down, go to last day of the year
            If (Shift And vbCtrlMask) > 0 Then
                'bAutoPaint = Me.AutoPaint
                'Me.AutoPaint = False
                cboPeriod.ListIndex = cboPeriod.ListCount - 1
                'Update the display with the new datevalue
                Call ChangeValue(m_Vars.PeriodEnd)
                'Me.AutoPaint = m_bAutoPaint
            Else
                'go to the last day of the current month
                'Update the display with the new datevalue
                Call ChangeValue(m_Vars.PeriodEnd)
            End If
        Case vbKeyPageUp
            Select Case CalendarType
            Case calMonth
                'go one month forward
                dtTemp = DateAdd("m", 1, DateValue)
                'Update the display with the new datevalue
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            Case calWeek
                'go one week forward
                dtTemp = DateAdd("ww", 1, DateValue)
                'Update the display with the new datevalue
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            Case calPeriod
                'go one period forward
                'Update the display with the new datevalue
                If m_Vars.Periods.Count < PeriodValue + 1 Then
                    nPeriodValue = 1
                Else
                    nPeriodValue = PeriodValue + 1
                End If
                nWeeks = m_Vars.Periods(nPeriodValue).NumberOfWeeks
                'Check to see if this is a leap week year
                If m_Methods.IsExtraWeek(m_Vars.FirstOfYear, PeriodYear, FirstDayOfWeek, YearStartPlacement) Then
                    Select Case m_nExtraWeekPlacement
                    Case ExtraWeekPlacements.calFirstPeriod
                        'See if this is the first period, if so
                        'then add one week to the total number of
                        'weeks for this period
                        If PeriodValue = 1 Then nWeeks = nWeeks + 1
                    Case ExtraWeekPlacements.calLastPeriod
                        'See if this is the last period, if so
                        'then add one week to the total number of
                        'weeks for this period
                        If PeriodValue = m_Vars.Periods.Count Then nWeeks = nWeeks + 1
                    End Select
                End If
                dtTemp = DateAdd("ww", nWeeks, DateValue)
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            End Select
        Case vbKeyPageDown
            Select Case CalendarType
            Case calMonth
                'go one month back
                dtTemp = DateAdd("m", -1, DateValue)
                'Update the display with the new datevalue
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            Case calWeek
                'go one week back
                dtTemp = DateAdd("ww", -1, DateValue)
                'Update the display with the new datevalue
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            Case calPeriod
                'go one period forward
                'Update the display with the new datevalue
                If PeriodValue - 1 > 0 Then
                    nPeriodValue = PeriodValue - 1
                Else
                    nPeriodValue = m_Vars.Periods.Count
                End If
                nWeeks = m_Vars.Periods(nPeriodValue).NumberOfWeeks
                'Check to see if this is a leap week year
                If m_Methods.IsExtraWeek(m_Vars.FirstOfYear, PeriodYear, FirstDayOfWeek, YearStartPlacement) Then
                    Select Case m_nExtraWeekPlacement
                    Case ExtraWeekPlacements.calFirstPeriod
                        'See if this is the first period, if so
                        'then add one week to the total number of
                        'weeks for this period
                        If nPeriodValue = 1 Then nWeeks = nWeeks + 1
                    Case ExtraWeekPlacements.calLastPeriod
                        'See if this is the last period, if so
                        'then add one week to the total number of
                        'weeks for this period
                        If nPeriodValue = m_Vars.Periods.Count Then nWeeks = nWeeks + 1
                    End Select
                End If
                dtTemp = DateAdd("ww", -nWeeks, DateValue)
                If IsInRange(dtTemp, CalendarType) = False Then Exit Sub
                Call ChangeValue(dtTemp)
            End Select
    End Select
    
    'Now match the dropdown list boxes with our new date value
    Call RefreshListIndexes
    'Set our DateTip property value back to its original setting
    m_bShowDateTip = bShowDateTipOrg

End Sub 'Event ctlFocus_KeyDown()

Private Sub mnuCalTypeMonth_Click()
    
    'Change the calendar style to month format
    CalendarType = calMonth
    'Call SetCalendarStyle(calMonth)
    
End Sub 'Menu mnuCalTypeMonth

Private Sub mnuCalTypePeriod_Click()
    
    'Change the calendar style to Period format
    CalendarType = calPeriod
    'Call SetCalendarStyle(calPeriod)
    
End Sub 'Menu mnuCalTypePeriod

Private Sub mnuCalTypeWeek_Click()
    
    'Change the calendar style to Week format
    CalendarType = calWeek
    'Call SetCalendarStyle(calWeek)
    
End Sub 'Menu CalTypeWeek

Private Sub mnuNextPeriod_Click()
    
    'Increment to the next period, if the last period
    'has been reached then loop back to the first period
    If PeriodValue + 1 < cboPeriod.ListCount + 1 Then
        PeriodValue = PeriodValue + 1
    Else
        PeriodValue = 1
    End If
    cboPeriod.ListIndex = PeriodValue - 1
    
End Sub 'Menu mnuNextPeriod

Private Sub mnuPrevPeriod_Click()

    'Decrement to the previous period, if the first period
    'has been reached then loop to the last period
    If PeriodValue - 1 > 0 Then
        PeriodValue = PeriodValue - 1
    Else
        PeriodValue = cboPeriod.ListCount
    End If
    cboPeriod.ListIndex = PeriodValue - 1

End Sub 'Menu mnuPrevPeriod

'----------------------------------------------------------------------
' tmrResize_Timer Method
'----------------------------------------------------------------------
' Purpose:  Used for the resize event
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub tmrResize_Timer()
    tmrResize.Enabled = False
    If m_bAutoPaint Then Refresh
End Sub 'Event trmResize_Timer

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

    On Error GoTo ErrorHandler
    
    'Check to see if this is our picture for the background
    If (AsyncProp.PropertyName = pnPicture) Then ' Picture download is complete
        m_bClearPictureOnly = True
        Set Picture = AsyncProp.Value           ' Store picture data to property...
    End If
    '=========================================================
    ' LICENSE STUFF
    '=========================================================
'    If (AsyncProp.PropertyName = pnURLLocation) Then
'        If Len(AsyncProp.Value) Then
'            m_sRegistered = "TRUE"
'            Call CheckLicense
'        End If
'    End If
    '=========================================================
    
    Exit Sub
    
ErrorHandler:
    If (AsyncProp.PropertyName = pnURLLocation) Then m_sRegistered = "FALSE"
    m_bClearPictureOnly = False
    
End Sub 'Event AsyncReadComplete

Private Sub UserControl_DblClick()

    RaiseEvent DblClick(DateValue, PeriodValue, m_nRow, m_nCol)
    
End Sub 'Event UserControl_DblClick

Private Sub UserControl_Initialize()

    'Set some of our UserControl properties the way we need them
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
    UserControl.BackColor = DEF_BACKCOLOR
    
    'Init our internal objects
    Set m_Methods = New CCalendarVBMethods
    
    'Init our period to 13 periods 4 weeks each
    Call PeriodDefault
    
    'Set the control property variables with the default values
    m_bAutoPaint = DEF_AUTOPAINT
    m_bShowDayHeader = DEF_SHOW_DAY_HEADER
    m_bPopupMenuDisabled = DEF_POPUP_MENU_DISABLED
    m_bShowDateTip = DEF_SHOW_DATE_TIP
    m_bShowPeriodList = DEF_SHOW_PERIOD_LIST
    m_bShowYearList = DEF_SHOW_YEAR_LIST
    
    m_nLineStyle = DEF_LINE_STYLE
    m_nYearStartPlacement = DEF_YEAR_START_PLACEMENT
    m_nDayNumberAlignment = DEF_DAY_NUMBER_ALIGNMENT
    m_nDayHeaderFormat = DEF_DAY_HEADER_FORMAT
    m_nExtraWeekPlacement = DEF_EXTRA_WEEK_PLACEMENT
    m_nFirstCurrentPeriod = DEF_FIRST_DAY_OF_WEEK
    m_nPeriodRows = DEF_PERIOD_ROWS
    m_Vars.PeriodValue = m_Methods.DateToPeriod(DateValue, PeriodYear, Periods _
        , DEF_CALENDAR_TYPE, m_nFirstCurrentPeriod, m_nYearStartPlacement _
        , m_nExtraWeekPlacement, DateValue)
    
    m_oBackColor = DEF_BACKCOLOR
    m_oCurrentPeriodBackColor = DEF_CURRENT_PERIOD_BACKCOLOR
    m_oCurrentPeriodForeColor = DEF_CURRENT_PERIOD_FORECOLOR
    m_oDayHeaderBackColor = DEF_DAY_HEADER_BACKCOLOR
    m_oDayHeaderForeColor = DEF_DAY_HEADER_FORECOLOR
    m_oActiveDayForeColor = DEF_ACTIVE_DAY_FORECOLOR
    m_oFlatLineColor = DEF_FLAT_LINE_COLOR
    m_oPrePeriodBackColor = DEF_PRE_PERIOD_BACKCOLOR
    m_oPrePeriodForeColor = DEF_PRE_PERIOD_FORECOLOR
    m_oPostPeriodBackColor = DEF_POST_PERIOD_BACKCOLOR
    m_oPostPeriodForeColor = DEF_POST_PERIOD_FORECOLOR
    DEF_YEAR_BEGIN = Format$(DateAdd("yyyy", -10, Date$), "yyyy")
    DEF_YEAR_END = Format$(DateAdd("yyyy", 10, Date$), "yyyy")
    m_nYearBegin = DEF_YEAR_BEGIN
    m_nYearEnd = DEF_YEAR_END
    
    m_sDateTipFormat = DEF_DATE_TIP_FORMAT
    
    m_Vars.DateValue = Format$(Date, "mm/dd/yyyy")
    m_Vars.PeriodYear = Year(m_Vars.DateValue)
    
    m_Methods.CopyFont UserControl.Font, m_DayHeaderFont
    m_Methods.CopyFont UserControl.Font, m_DaysFont
    m_Methods.CopyFont UserControl.Font, m_ActiveDayFont
    
    Set m_picImage = Nothing
    
End Sub 'Event UserControl_Initialize()

Private Sub UserControl_InitProperties()

    'Set our module level variable to indicate whether where in design or run mode
    SetDesignMode
End Sub 'Event UserControl_InitProperties

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim nCol            As Integer
    Dim nRow            As Integer
    Dim dtDateValue     As Date
    
    Static nLastCol             As Integer
    Static nLastRow             As Integer
    
    If m_bShowDateTip = False Then
        DateTipDisplay False
        Exit Sub
    End If
    
    'Check to see if were still on the calendar grid portion of the control, if not
    'then turn off our ToolTip and were out of here...
    If (x <= m_udtDisabledRect.Left) Or (y <= (((m_udtDisabledRect.Height - (m_nCellHeight * nCalendarRows()))) + DEF_TOP_MARGIN)) _
      Or (x > (UserControl.ScaleWidth - DEF_LEFT_MARGIN)) Or (y > UserControl.ScaleHeight) Then
        DateTipDisplay False
        Exit Sub
    End If
    
    'Display the date for the current row and column as a tooltip
    Call GetCellLocation(x, y, nRow, nCol, dtDateValue)
    
    'If the user clicked the mouse off our calendar grid
    'then set the tooltip to an empty string
    If nRow > nCalendarRows() Or nRow < 1 Or nCol > DEF_CALENDAR_COLS Or nCol < 1 Then
        DateTipDisplay False
    Else
        'Only need to update if the user has moved to a new cell
        'location
        If nRow <> nLastRow Or nCol <> nLastCol Then
            nLastCol = nCol
            nLastRow = nRow
            DateTipDisplay False
            DateTipDisplay True, Format$(dtDateValue, m_sDateTipFormat)
        End If
    End If

End Sub 'Event UserControl_MouseMove

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim bShowDateTipOrg     As Boolean
    Dim nCol                As Integer
    Dim nRow                As Integer
    Dim dtDateValue         As Date
    
    'If the left mouse button has been clicked then we need to indicate the new
    'selected calendar cell. If it is a period then we need to redraw the calendar
    'using the new period.
    
    'Be sure that our tooltip is not visible and disable for now
    If m_bShowDateTip Then
        bShowDateTipOrg = m_bShowDateTip
        Call DateTipDisplay(False)
        m_bToolTipVisible = False
    End If
    
    Select Case Button
    Case vbLeftButton
        'Get the current cell location
        If GetCellLocation(x, y, nRow, nCol, dtDateValue) Then Exit Sub
        
        'Check to be sure that we have not gone past our YearBegin and YearEnd
        'values, if so then we just ignore
        If IsInRange(dtDateValue, CalendarType) = False Then Exit Sub
        
        'Update the calendar with the new datevalue that has been selected
        Call ChangeValue(dtDateValue)
        Call RefreshListIndexes
        
        'Fire off the click event
        RaiseEvent Click(DateValue, PeriodValue, m_nRow, m_nCol)
    Case vbRightButton
        'The right mouse button was pressed so show our Popup Menu
        If Not m_bPopupMenuDisabled Then PopupMenu mnuPopup
    End Select
    
    'Set our ShowDateTip property back to its original value
    m_bShowDateTip = bShowDateTipOrg
    
End Sub 'Event UserControl_MouseUp

Private Sub UserControl_Paint()

    'Refresh the control display with our MemoryDC image. Much faster
    'than having to redraw the calendar each time.
    If m_bAutoPaint Then
        With UserControl
            m_RefreshDC.CopyToHdc 0, 0, .ScaleWidth, .ScaleHeight
        End With
    End If

End Sub 'Event UserControl_Paint

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Dim sUrl            As String
    Dim picMine         As Picture
    Dim TempByte()      As Byte
'    Dim oLicense        As New CLicense
    
    On Error Resume Next
    
    '=========================================================
    ' LICENSE STUFF
    '=========================================================
'    TempByte = PropBag.ReadProperty("ABOUT", "FALSE")
'    m_sRegistered = TempByte
'    TempByte = PropBag.ReadProperty("ENVIRONMENT", "")
'    m_sEnvironment = TempByte
'    Set oLicense = Nothing
'    'License file URL location used for web page runtime licensing
'    With PropBag
'        sUrl = .ReadProperty(pnURLLocation, "")
'        If Len(sUrl) <> 0 Then
'            URLLocation = sUrl
'        End If
'    End With
    '=========================================================
    
    'Turn painting off, this will be the last property that
    'gets set
    m_bAutoPaint = False
    
    ' Read in the properties that have been saved into the PropertyBag...
    m_bShowDayHeader = PropBag.ReadProperty(pnDisplayHeader, DEF_SHOW_DAY_HEADER)
    m_bPopupMenuDisabled = PropBag.ReadProperty(pnPopupMenuDisabled, DEF_POPUP_MENU_DISABLED)
    m_bShowDateTip = PropBag.ReadProperty(pnShowDateTip, DEF_SHOW_DATE_TIP)
    m_bShowPeriodList = PropBag.ReadProperty(pnShowPeriodList, DEF_SHOW_PERIOD_LIST)
        cboPeriod.Visible = m_bShowPeriodList
    m_bShowYearList = PropBag.ReadProperty(pnShowYearList, DEF_SHOW_YEAR_LIST)
        cboYear.Visible = m_bShowYearList
    
    With UserControl
        .Enabled = PropBag.ReadProperty(pnEnabled, True)
        cboPeriod.Enabled = .Enabled
        cboYear.Enabled = .Enabled
    End With
    m_nLineStyle = PropBag.ReadProperty(pnLineStyle, DEF_LINE_STYLE)
    m_nYearStartPlacement = PropBag.ReadProperty(pnYearStartPlacement, DEF_YEAR_START_PLACEMENT)
    m_nDayNumberAlignment = PropBag.ReadProperty(pnDayNumberAlignment, DEF_DAY_NUMBER_ALIGNMENT)
    m_nDayHeaderFormat = PropBag.ReadProperty(pnDayHeaderFormat, DEF_DAY_HEADER_FORMAT)
    m_nExtraWeekPlacement = PropBag.ReadProperty(pnExtraWeekPlacement, DEF_EXTRA_WEEK_PLACEMENT)
    m_nFirstCurrentPeriod = PropBag.ReadProperty(pnFirstCurrentPeriod, DEF_FIRST_DAY_OF_WEEK)
    m_nPeriodRows = PropBag.ReadProperty(pnPeriodRows, DEF_PERIOD_ROWS)
    
    m_oBackColor = PropBag.ReadProperty(pnBackColor, DEF_BACKCOLOR)
      UserControl.BackColor = m_oBackColor
    m_oCurrentPeriodBackColor = PropBag.ReadProperty(pnCurrentPeriodBackColor, DEF_CURRENT_PERIOD_BACKCOLOR)
    m_oCurrentPeriodForeColor = PropBag.ReadProperty(pnCurrentPeriodForeColor, DEF_CURRENT_PERIOD_FORECOLOR)
    m_oActiveDayForeColor = PropBag.ReadProperty(pnActiveDayForeColor, DEF_ACTIVE_DAY_FORECOLOR)
    m_oDayHeaderBackColor = PropBag.ReadProperty(pnDayHeaderBackColor, DEF_DAY_HEADER_BACKCOLOR)
    m_oDayHeaderForeColor = PropBag.ReadProperty(pnDayHeaderForeColor, DEF_DAY_HEADER_FORECOLOR)
    m_oFlatLineColor = PropBag.ReadProperty(pnFlatLineColor, DEF_FLAT_LINE_COLOR)
    m_oPrePeriodBackColor = PropBag.ReadProperty(pnPrePeriodBackColor, DEF_PRE_PERIOD_BACKCOLOR)
    m_oPrePeriodForeColor = PropBag.ReadProperty(pnPrePeriodForeColor, DEF_PRE_PERIOD_FORECOLOR)
    m_oPostPeriodBackColor = PropBag.ReadProperty(pnPostPeriodBackColor, DEF_POST_PERIOD_BACKCOLOR)
    m_oPostPeriodForeColor = PropBag.ReadProperty(pnPostPeriodForeColor, DEF_POST_PERIOD_FORECOLOR)
    m_nYearBegin = PropBag.ReadProperty(pnYearBegin, DEF_YEAR_BEGIN)
    m_nYearEnd = PropBag.ReadProperty(pnYearEnd, DEF_YEAR_END)
    
    m_sDateTipFormat = PropBag.ReadProperty(pnDateTipFormat, DEF_DATE_TIP_FORMAT)
    
    'Period definition object
    Set Periods = PropBag.ReadProperty(pnPeriods, m_Vars.Periods)
    
    'Fonts
    m_ActiveDayFont.Bold = PropBag.ReadProperty(pnActiveDayFontBold, False)
    m_ActiveDayFont.Italic = PropBag.ReadProperty(pnActiveDayFontItalic, False)
    m_ActiveDayFont.Name = PropBag.ReadProperty(pnActiveDayFontName, "Arial")
    m_ActiveDayFont.Size = PropBag.ReadProperty(pnActiveDayFontSize, 8)
        
    m_DayHeaderFont.Bold = PropBag.ReadProperty(pnDayHeaderFontBold, False)
    m_DayHeaderFont.Italic = PropBag.ReadProperty(pnDayHeaderFontItalic, False)
    m_DayHeaderFont.Name = PropBag.ReadProperty(pnDayHeaderFontName, "Arial")
    m_DayHeaderFont.Size = PropBag.ReadProperty(pnDayHeaderFontSize, 8)
    Set cboYear.Font = m_DayHeaderFont
    Set cboPeriod.Font = m_DayHeaderFont
    
    m_DaysFont.Bold = PropBag.ReadProperty(pnDaysFontBold, False)
    m_DaysFont.Italic = PropBag.ReadProperty(pnDaysFontItalic, False)
    m_DaysFont.Name = PropBag.ReadProperty(pnDaysFontName, "Arial")
    m_DaysFont.Size = PropBag.ReadProperty(pnDaysFontSize, 8)
    m_nCalendarType = PropBag.ReadProperty(pnCalendarType, DEF_CALENDAR_TYPE)
    
    'Background picture
    With PropBag
        sUrl = .ReadProperty(pnURLPicture, "")      ' Read URLPicture property value
        If Len(sUrl) <> 0 Then                      ' If a URL has been entered...
            URLPicture = sUrl                       ' Attempt to download it now, URL may be unavailable at this time
        Else
            Set picMine = PropBag.ReadProperty(pnPicture, UserControl.Picture) ' Read Picture property value
            If Not (picMine Is Nothing) Then        ' URL is not available
                Set Picture = picMine               ' Use existing picture (This is used only if URL is empty)
            End If
        End If
    End With
    
    BorderStyle = PropBag.ReadProperty(pnBorderStyle, DEF_BORDERSTYLE)
    m_bAutoPaint = PropBag.ReadProperty(pnAutoPaint, DEF_AUTOPAINT)
    Call SetCalendarStyle(CalendarType)
    
End Sub 'Event UserControl_ReadProperties

Private Sub UserControl_Resize()

    'Update our calendar display using the new UserControl size
    'Have to use a timer because the resize event fires before the
    'UserControl has updated the size properties with the actual
    'resized values.
    tmrResize.Enabled = True
    
End Sub 'Event UserControl_Resize

Private Sub UserControl_Show()
    
    'Check to see if this is a Registered copy of the control.
    'If not then it will display the about box with registration
    'information.
'    If m_sRegistered <> "AsyncRead" Then Call CheckLicense

    'Be sure that our calendar has been drawn
    Refresh
    m_bActive = True
        
End Sub 'Event UserControl_Show

Private Sub CleanUp()

    On Error Resume Next
    
    Set m_ActiveDayFont = Nothing
    Set m_DayHeaderFont = Nothing
    Set m_DaysFont = Nothing
    Set m_picImage = Nothing
    Set m_Vars.Periods = Nothing
    Unload m_ToolTip
    Set m_ToolTip = Nothing
    Set m_RefreshDC = Nothing
    
End Sub 'CleanUp()

Private Sub UserControl_Terminate()
    Call CleanUp
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    '=========================================================
    ' LICENSE STUFF
    '=========================================================
'    Dim TempByte()      As Byte
'    Dim oLicense        As New CLicense
'
'    On Error Resume Next
'
'    With oLicense
'        If .IsRegistered = False Then
'             TempByte = "FALSE"
'        Else
'             TempByte = "TRUE"
'        End If
'        Call PropBag.WriteProperty("ABOUT", TempByte, "")
'
'        TempByte = .GetProgramName(GetParent(UserControl.hwnd))
'        Call PropBag.WriteProperty("ENVIRONMENT", TempByte, "")
'    End With
'    Set oLicense = Nothing
    '=========================================================
    
    Call PropBag.WriteProperty(pnAutoPaint, m_bAutoPaint, DEF_AUTOPAINT)
    Call PropBag.WriteProperty(pnBorderStyle, m_nBorderStyle, DEF_BORDERSTYLE)
    Call PropBag.WriteProperty(pnDisplayHeader, m_bShowDayHeader, DEF_SHOW_DAY_HEADER)
    Call PropBag.WriteProperty(pnEnabled, UserControl.Enabled, True)
    Call PropBag.WriteProperty(pnPopupMenuDisabled, m_bPopupMenuDisabled, DEF_POPUP_MENU_DISABLED)
    Call PropBag.WriteProperty(pnShowDateTip, m_bShowDateTip, DEF_SHOW_DATE_TIP)
    Call PropBag.WriteProperty(pnShowPeriodList, m_bShowPeriodList, DEF_SHOW_PERIOD_LIST)
    Call PropBag.WriteProperty(pnShowYearList, m_bShowYearList, DEF_SHOW_YEAR_LIST)
    
    Call PropBag.WriteProperty(pnLineStyle, m_nLineStyle, DEF_LINE_STYLE)
    Call PropBag.WriteProperty(pnPeriodRows, m_nPeriodRows, DEF_PERIOD_ROWS)
    Call PropBag.WriteProperty(pnYearStartPlacement, m_nYearStartPlacement, DEF_YEAR_START_PLACEMENT)
    Call PropBag.WriteProperty(pnCalendarType, m_nCalendarType, DEF_CALENDAR_TYPE)
    Call PropBag.WriteProperty(pnDayNumberAlignment, m_nDayNumberAlignment, DEF_DAY_NUMBER_ALIGNMENT)
    Call PropBag.WriteProperty(pnDayHeaderFormat, m_nDayHeaderFormat, DEF_DAY_HEADER_FORMAT)
    Call PropBag.WriteProperty(pnExtraWeekPlacement, m_nExtraWeekPlacement, DEF_EXTRA_WEEK_PLACEMENT)
    Call PropBag.WriteProperty(pnFirstCurrentPeriod, m_nFirstCurrentPeriod, DEF_FIRST_DAY_OF_WEEK)
    
    Call PropBag.WriteProperty(pnBackColor, m_oBackColor, DEF_BACKCOLOR)
      UserControl.BackColor = m_oBackColor
    Call PropBag.WriteProperty(pnCurrentPeriodBackColor, m_oCurrentPeriodBackColor, DEF_CURRENT_PERIOD_BACKCOLOR)
    Call PropBag.WriteProperty(pnCurrentPeriodForeColor, m_oCurrentPeriodForeColor, DEF_CURRENT_PERIOD_FORECOLOR)
    Call PropBag.WriteProperty(pnDayHeaderBackColor, m_oDayHeaderBackColor, DEF_DAY_HEADER_BACKCOLOR)
    Call PropBag.WriteProperty(pnDayHeaderForeColor, m_oDayHeaderForeColor, DEF_DAY_HEADER_FORECOLOR)
    Call PropBag.WriteProperty(pnActiveDayForeColor, m_oActiveDayForeColor, DEF_ACTIVE_DAY_FORECOLOR)
    Call PropBag.WriteProperty(pnFlatLineColor, m_oFlatLineColor, DEF_FLAT_LINE_COLOR)
    Call PropBag.WriteProperty(pnPrePeriodBackColor, m_oPrePeriodBackColor, DEF_PRE_PERIOD_BACKCOLOR)
    Call PropBag.WriteProperty(pnPrePeriodForeColor, m_oPrePeriodForeColor, DEF_PRE_PERIOD_FORECOLOR)
    Call PropBag.WriteProperty(pnPostPeriodBackColor, m_oPostPeriodBackColor, DEF_POST_PERIOD_BACKCOLOR)
    Call PropBag.WriteProperty(pnPostPeriodForeColor, m_oPostPeriodForeColor, DEF_POST_PERIOD_FORECOLOR)
    Call PropBag.WriteProperty(pnYearBegin, m_nYearBegin, DEF_YEAR_BEGIN)
    Call PropBag.WriteProperty(pnYearEnd, m_nYearEnd, DEF_YEAR_END)
    
    Call PropBag.WriteProperty(pnDateTipFormat, m_sDateTipFormat, DEF_DATE_TIP_FORMAT)
    
    PropBag.WriteProperty pnActiveDayFontBold, m_ActiveDayFont.Bold
    PropBag.WriteProperty pnActiveDayFontItalic, m_ActiveDayFont.Italic
    PropBag.WriteProperty pnActiveDayFontSize, m_ActiveDayFont.Size
    PropBag.WriteProperty pnActiveDayFontName, m_ActiveDayFont.Name
    
    PropBag.WriteProperty pnDayHeaderFontBold, m_DayHeaderFont.Bold
    PropBag.WriteProperty pnDayHeaderFontItalic, m_DayHeaderFont.Italic
    PropBag.WriteProperty pnDayHeaderFontSize, m_DayHeaderFont.Size
    PropBag.WriteProperty pnDayHeaderFontName, m_DayHeaderFont.Name
    
    PropBag.WriteProperty pnDaysFontBold, m_DaysFont.Bold
    PropBag.WriteProperty pnDaysFontItalic, m_DaysFont.Italic
    PropBag.WriteProperty pnDaysFontSize, m_DaysFont.Size
    PropBag.WriteProperty pnDaysFontName, m_DaysFont.Name
    
    If Len(m_sURLPicture) <> 0 Then
        PropBag.WriteProperty pnURLPicture, m_sURLPicture
    Else
        PropBag.WriteProperty pnPicture, m_picImage
    End If
    If Len(m_sURLLocation) <> 0 Then PropBag.WriteProperty pnURLLocation, m_sURLLocation
        
End Sub 'Event UserControl_WriteProperties

'**********************************************************************
' PROPERTY IMPLEMENTAION
'**********************************************************************

'----------------------------------------------------------------------
' AutoPaint Get/Let
'----------------------------------------------------------------------
' Purpose: Determines whether the control is painted when changes
'          are made.
'----------------------------------------------------------------------
Public Property Get AutoPaint() As Boolean
Attribute AutoPaint.VB_Description = "If false the control will not paint itself until this property is set to true."
    AutoPaint = m_bAutoPaint
End Property 'AutoPaint Get

Public Property Let AutoPaint(ByVal bAutoPaint As Boolean)
    m_bAutoPaint = bAutoPaint
    PropertyChanged pnAutoPaint
    If m_bAutoPaint Then Refresh
End Property 'AutoPaint Let

'----------------------------------------------------------------------
' BorderStyle Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the border Style of the calendar
'----------------------------------------------------------------------
Public Property Get BorderStyle() As CalendarLineTypes
   BorderStyle = m_nBorderStyle
End Property
Public Property Let BorderStyle(ByVal nStyle As CalendarLineTypes)
   m_nBorderStyle = nStyle
   Select Case nStyle
   Case calSunken
      UserControl.BorderStyle() = 1
   Case calNoLine
      UserControl.BorderStyle() = 0
   Case cal3D
      UserControl.BorderStyle() = 0
   Case calNoLine
      UserControl.BorderStyle() = 0
   End Select
   pSetBorderStyle
   PropertyChanged pnBorderStyle
   If m_bAutoPaint Then Refresh
End Property
Private Sub pSetBorderStyle()
Dim lStyle As Long
Dim lExStyle As Long
Dim lhWnd As Long
   lhWnd = UserControl.hwnd
   lExStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
   lStyle = GetWindowLong(lhWnd, GWL_STYLE)
   lStyle = lStyle And Not (WS_BORDER Or WS_THICKFRAME)
   lExStyle = lExStyle And Not (WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
   If m_nBorderStyle = calSunken Then
      lExStyle = lExStyle Or WS_EX_CLIENTEDGE
   ElseIf m_nBorderStyle = calFlat Then
      lExStyle = lExStyle Or WS_EX_STATICEDGE
   ElseIf m_nBorderStyle = cal3D Then
      lExStyle = lExStyle Or WS_EX_WINDOWEDGE
      lStyle = lStyle Or WS_BORDER Or WS_THICKFRAME
   End If
   SetWindowLong lhWnd, GWL_STYLE, lStyle
   SetWindowLong lhWnd, GWL_EXSTYLE, lExStyle
   SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
End Sub

'----------------------------------------------------------------------
' CalendarType Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the display format of calendar
'----------------------------------------------------------------------
Public Property Get CalendarType() As CalendarTypes
Attribute CalendarType.VB_Description = "This indicates which period style will be used when displaying the calendar. Month, Period, or Week"
Attribute CalendarType.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CalendarType = m_nCalendarType
End Property 'CalendarType Get

Public Property Let CalendarType(ByVal nCalendarType As CalendarTypes)
    
    'Error related declares
    Dim nAction             As Integer
    Dim lErrNumber          As Long
    Dim sErrDescription     As String
    Const sErrSource = "CalendarType(Let)"
     
CalendarType_Begin:
    'Check to be sure that the current date is in our year ranges
    'for this new period
    If IsInRange(DateValue, nCalendarType) = False Then
        nAction = CalErrorActions.calAbort
        lErrNumber = CalErrorNumbers.calInvalidDateRange
        sErrDescription = "Invalid date. The date: " & DateValue & " does not fall within"
        sErrDescription = sErrDescription & " the BeginYear: " & YearEnd
        sErrDescription = sErrDescription & " and the EndYear: " & YearEnd & " ranges."
        RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, nCalendarType)
        Select Case nAction
        Case CalErrorActions.calRetry
            GoTo CalendarType_Begin
        Case Else
            GoTo Exit_CalendarType
        End Select
        Exit Property
    End If
    m_nCalendarType = nCalendarType
    PropertyChanged pnCalendarType
    If m_bAutoPaint Then Call SetCalendarStyle(nCalendarType)
    
Exit_CalendarType:

End Property 'CalendarType Let

'----------------------------------------------------------------------
' PopupMenuDisabled Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines if the popup menu is to be displayed when the
'           user presses the right mouse button
'----------------------------------------------------------------------
Public Property Get PopupMenuDisabled() As Boolean
Attribute PopupMenuDisabled.VB_Description = "If true then the popup menu will not be displayed when the right mouse button is pressed"
Attribute PopupMenuDisabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    PopupMenuDisabled = m_bPopupMenuDisabled
End Property 'Get PopupMenuDisabled

Public Property Let PopupMenuDisabled(ByVal bPopupMenuDisabled As Boolean)
    m_bPopupMenuDisabled = bPopupMenuDisabled
    PropertyChanged pnPopupMenuDisabled
End Property 'Let PopupMenuDisabled

'----------------------------------------------------------------------
' ShowDayHeader Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines if the days of the week will be displayed above
'           the calendar days
'----------------------------------------------------------------------
Public Property Get ShowDayHeader() As Boolean
Attribute ShowDayHeader.VB_Description = "If true then the days of the week header will be displayed"
Attribute ShowDayHeader.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowDayHeader = m_bShowDayHeader
End Property 'ShowDayHeader Get

Public Property Let ShowDayHeader(ByVal bShowDayHeader As Boolean)
    m_bShowDayHeader = bShowDayHeader
    PropertyChanged pnDisplayHeader
    If m_bAutoPaint Then Refresh
End Property 'ShowDayHeader Let

'----------------------------------------------------------------------
' BackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  The background color of the calendar control
'----------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Controls background color"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_oBackColor
End Property 'BackColor Get

Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
    m_oBackColor = lBackColor
    UserControl.BackColor = m_oBackColor
    PropertyChanged pnBackColor
    If m_bAutoPaint Then Refresh

End Property 'BackColor Let

'----------------------------------------------------------------------
' DayNumberAlignment Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the number alignment for the calendar cels
'----------------------------------------------------------------------
Public Property Get DayNumberAlignment() As DayNumberAlignments
    DayNumberAlignment = m_nDayNumberAlignment
End Property 'DayNumberAlignment Get

Public Property Let DayNumberAlignment(ByVal nDayNumberAlignment As DayNumberAlignments)
    
    m_nDayNumberAlignment = nDayNumberAlignment
    PropertyChanged pnDayNumberAlignment
    If m_bAutoPaint Then Refresh
    
End Property 'DayNumberAlignment Let

'----------------------------------------------------------------------
' DayHeaderBackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the color to use for the background of the day header
'           cells
'----------------------------------------------------------------------
Public Property Get DayHeaderBackColor() As OLE_COLOR
Attribute DayHeaderBackColor.VB_Description = "The background color used to display the days of the week header"
Attribute DayHeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayHeaderBackColor = m_oDayHeaderBackColor
End Property 'DayHeaderBackColor Get

Public Property Let DayHeaderBackColor(ByVal lDayHeaderBackColor As OLE_COLOR)
    Dim lColor As Long
    
    m_oDayHeaderBackColor = lDayHeaderBackColor
    PropertyChanged pnDayHeaderBackColor
    If m_bAutoPaint Then Refresh
    
End Property 'DayHeaderBackColor Let

'----------------------------------------------------------------------
' DayHeaderForeColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the color for the day of the week lettering
'----------------------------------------------------------------------
Public Property Get DayHeaderForeColor() As OLE_COLOR
Attribute DayHeaderForeColor.VB_Description = "The foreground color used to display the days of the week header"
Attribute DayHeaderForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayHeaderForeColor = m_oDayHeaderForeColor
End Property 'DayHeaderForeColor Get

Public Property Let DayHeaderForeColor(ByVal lDayHeaderForeColor As OLE_COLOR)
    m_oDayHeaderForeColor = lDayHeaderForeColor
    PropertyChanged pnDayHeaderForeColor
    If m_bAutoPaint Then Refresh
    
End Property 'DayHeaderForeColor Let

'----------------------------------------------------------------------
' ActiveDayForeColor Get/Let
'----------------------------------------------------------------------
' Purpose: The color that will be used to indicate the currently
'          selected date
'----------------------------------------------------------------------
Public Property Get ActiveDayForeColor() As OLE_COLOR
Attribute ActiveDayForeColor.VB_Description = "The foregound color to display the currenlty selected date"
Attribute ActiveDayForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ActiveDayForeColor = m_oActiveDayForeColor
End Property 'ActiveDayForeColor Get

Public Property Let ActiveDayForeColor(ByVal lActiveDayForeColor As OLE_COLOR)
    m_oActiveDayForeColor = lActiveDayForeColor
    PropertyChanged pnActiveDayForeColor
    If m_bAutoPaint Then Refresh
    
End Property 'ActiveDayForeColor Let

'----------------------------------------------------------------------
' FlatLineColor Get/Let
'----------------------------------------------------------------------
' Purpose: The color that will be used to draw the grid lines when
'          the flat calendar style has been selected
'----------------------------------------------------------------------
Public Property Get FlatLineColor() As OLE_COLOR
Attribute FlatLineColor.VB_Description = "The color used for drawing the flat grid lines"
Attribute FlatLineColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatLineColor = m_oFlatLineColor
End Property 'FlatLineColor Get

Public Property Let FlatLineColor(ByVal lFlatLineColor As OLE_COLOR)
    m_oFlatLineColor = lFlatLineColor
    PropertyChanged pnFlatLineColor
    If m_bAutoPaint Then Refresh
    
End Property 'FlatLineColor Let

'----------------------------------------------------------------------
' PrePeriodBackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the backcolor to use for indicating the pre-period
'           cells
'----------------------------------------------------------------------
Public Property Get PrePeriodBackColor() As OLE_COLOR
Attribute PrePeriodBackColor.VB_Description = "The background color used to display the period calendar cells just prior to the current period"
Attribute PrePeriodBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PrePeriodBackColor = m_oPrePeriodBackColor
End Property 'PrePeriodBackColor Get

Public Property Let PrePeriodBackColor(ByVal lPrePeriodBackColor As OLE_COLOR)
    m_oPrePeriodBackColor = lPrePeriodBackColor
    PropertyChanged pnPrePeriodBackColor
    If m_bAutoPaint Then Refresh
    
End Property 'PrePeriodBackColor Let

'----------------------------------------------------------------------
' PrePeriodForeColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the forecolor to use for indicating the pre-period
'           cells
'----------------------------------------------------------------------
Public Property Get PrePeriodForeColor() As OLE_COLOR
Attribute PrePeriodForeColor.VB_Description = "The foreground color used to display the period calendar cells just prior to the current period"
Attribute PrePeriodForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PrePeriodForeColor = m_oPrePeriodForeColor
End Property 'PrePeriodBackColor Get

Public Property Let PrePeriodForeColor(ByVal lPrePeriodForeColor As OLE_COLOR)
    
    m_oPrePeriodForeColor = lPrePeriodForeColor
    PropertyChanged pnPrePeriodForeColor
    If m_bAutoPaint Then Refresh
    
End Property 'PrePeriodForeColor Let

'----------------------------------------------------------------------
' CurrentPeriodBackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the backcolor to use for indicating the current period
'           cells
'----------------------------------------------------------------------
Public Property Get CurrentPeriodBackColor() As OLE_COLOR
Attribute CurrentPeriodBackColor.VB_Description = "The background color used to display the current period calendar cells"
Attribute CurrentPeriodBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CurrentPeriodBackColor = m_oCurrentPeriodBackColor
End Property 'CurrentPeriodBackColor Get

Public Property Let CurrentPeriodBackColor(ByVal lCurrentPeriodBackColor As OLE_COLOR)
    m_oCurrentPeriodBackColor = lCurrentPeriodBackColor
    PropertyChanged pnCurrentPeriodBackColor
    If m_bAutoPaint Then Refresh
    
End Property 'CurrentPeriodBackColor Let

'----------------------------------------------------------------------
' CurrentPeriodForeColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the forecolor to use for indicating the current period
'           cells
'----------------------------------------------------------------------
Public Property Get CurrentPeriodForeColor() As OLE_COLOR
Attribute CurrentPeriodForeColor.VB_Description = "The foreground color used to display the current period calendar cells"
Attribute CurrentPeriodForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CurrentPeriodForeColor = m_oCurrentPeriodForeColor
End Property 'CurrentPeriodForeColor Get

Public Property Let CurrentPeriodForeColor(ByVal lCurrentPeriodForeColor As OLE_COLOR)
    m_oCurrentPeriodForeColor = lCurrentPeriodForeColor
    PropertyChanged pnCurrentPeriodForeColor
    If m_bAutoPaint Then Refresh
    
End Property 'CurrentPeriodForeColor Let

'----------------------------------------------------------------------
' PostPeriodBackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the backcolor to use for indicating the post period
'           cells
'----------------------------------------------------------------------
Public Property Get PostPeriodBackColor() As OLE_COLOR
Attribute PostPeriodBackColor.VB_Description = "The background color used to display the period calendar cells just after the current period"
Attribute PostPeriodBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PostPeriodBackColor = m_oPostPeriodBackColor
End Property 'PostPeriodBackColor Get

Public Property Let PostPeriodBackColor(ByVal lPostPeriodBackColor As OLE_COLOR)
    m_oPostPeriodBackColor = lPostPeriodBackColor
    PropertyChanged pnPostPeriodBackColor
    If m_bAutoPaint Then Refresh
    
End Property 'PostPeriodBackColor Let

'----------------------------------------------------------------------
' PostPeriodForeColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Determines the forecolor to use for indicating the post period
'           cells
'----------------------------------------------------------------------
Public Property Get PostPeriodForeColor() As OLE_COLOR
Attribute PostPeriodForeColor.VB_Description = "The foreground color used to display the period calendar cells just after the current period"
Attribute PostPeriodForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PostPeriodForeColor = m_oPostPeriodForeColor
End Property 'PostPeriodBackColor Get

Public Property Let PostPeriodForeColor(ByVal lPostPeriodForeColor As OLE_COLOR)
    m_oPostPeriodForeColor = lPostPeriodForeColor
    PropertyChanged pnPostPeriodForeColor
    If m_bAutoPaint Then Refresh
    
End Property 'PostPeriodForeColor Let

'----------------------------------------------------------------------
' ShowDateTip Get/Let
'----------------------------------------------------------------------
' Purpose: Indicates whether a tooltip showing the date for the cell
'          that the mouse is currently over, should be displayed
'----------------------------------------------------------------------
Public Property Get ShowDateTip() As Boolean
Attribute ShowDateTip.VB_Description = "If true then the datevalue for the calendar cell that the mouse cursor is over will be displayed"
Attribute ShowDateTip.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowDateTip = m_bShowDateTip
End Property 'Get ShowDateTip

Public Property Let ShowDateTip(ByVal bShowDateTip As Boolean)
    m_bShowDateTip = bShowDateTip
End Property 'Let ShowDateTip

'----------------------------------------------------------------------
' ShowPeriodList Get/Let
'----------------------------------------------------------------------
' Purpose: Shows or hides the dropdown list
'----------------------------------------------------------------------
Public Property Get ShowPeriodList() As Boolean
Attribute ShowPeriodList.VB_Description = "Shows or hides the Period DropDown list"
    ShowPeriodList = m_bShowPeriodList
End Property 'Get ShowPeriodList

Public Property Let ShowPeriodList(ByVal bShowPeriodList As Boolean)
    m_bShowPeriodList = bShowPeriodList
    cboPeriod.Visible = m_bShowPeriodList
    If m_bAutoPaint Then Refresh
End Property 'Let ShowPeriodList

'----------------------------------------------------------------------
' ShowYearList Get/Let
'----------------------------------------------------------------------
' Purpose: Shows or hides the dropdown list
'----------------------------------------------------------------------
Public Property Get ShowYearList() As Boolean
Attribute ShowYearList.VB_Description = "Shows or Hides the Year DropDown list"
    ShowYearList = m_bShowYearList
End Property 'Get ShowYearList

Public Property Let ShowYearList(ByVal bShowYearList As Boolean)
    m_bShowYearList = bShowYearList
    cboYear.Visible = m_bShowYearList
    If m_bAutoPaint Then Refresh
End Property 'Let ShowYearList

'----------------------------------------------------------------------
' FirstDayOfWeek Get/Let
'----------------------------------------------------------------------
' Purpose: Determines which day will be the first day of the week
'----------------------------------------------------------------------
Public Property Get FirstDayOfWeek() As DaysOfTheWeek
Attribute FirstDayOfWeek.VB_Description = "Indicates which day of the week will be used as the first day of the week"
Attribute FirstDayOfWeek.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FirstDayOfWeek = m_nFirstCurrentPeriod
End Property 'Get FirstDayOfWeek()

Public Property Let FirstDayOfWeek(ByVal nFirstDayOfWeek As DaysOfTheWeek)
    'validate our inputs
    If nFirstDayOfWeek >= calSunday And nFirstDayOfWeek <= calSaturday Then
        m_nFirstCurrentPeriod = nFirstDayOfWeek
        PropertyChanged pnFirstCurrentPeriod
        If m_bAutoPaint Then Refresh
    End If 'valid inputs
End Property 'Let FirstDayOfWeek()

'----------------------------------------------------------------------
' LineStyle Get/Let
'----------------------------------------------------------------------
' Purpose: Indicates whether the calendar will be displayed with flat
'          grid lines or 3D grid lines
'----------------------------------------------------------------------
Public Property Get LineStyle() As CalendarLineTypes
Attribute LineStyle.VB_Description = "The style of grid lines that will be used when drawing the control"
Attribute LineStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    LineStyle = m_nLineStyle
End Property 'Get LineStyle

Public Property Let LineStyle(ByVal nLineStyle As CalendarLineTypes)
    m_nLineStyle = nLineStyle
    PropertyChanged pnLineStyle
    If m_bAutoPaint Then Refresh
End Property 'Let LineStyle

'----------------------------------------------------------------------
' YearStartPlacement Get/Let
'----------------------------------------------------------------------
' Purpose: Determines if the period starting year will use the previous
'          year or current year for its starting point
'----------------------------------------------------------------------
Public Property Get YearStartPlacement() As CalYearStartPlacement
Attribute YearStartPlacement.VB_Description = "Indicates whether the previous year or current year is used for the first week of a period year"
Attribute YearStartPlacement.VB_ProcData.VB_Invoke_Property = ";Behavior"
    YearStartPlacement = m_nYearStartPlacement
End Property 'Get YearStartPlacement()

Public Property Let YearStartPlacement(ByVal nYearStartPlacement As CalYearStartPlacement)
    m_nYearStartPlacement = nYearStartPlacement
    PropertyChanged pnYearStartPlacement
    If m_bAutoPaint Then Refresh
End Property 'Let YearStartPlacement()

'----------------------------------------------------------------------
' URLLocation Get/Let
'----------------------------------------------------------------------
' Purpose: Points to a URL location for the runtime license file
'----------------------------------------------------------------------
Public Property Get URLLocation() As String
Attribute URLLocation.VB_MemberFlags = "40"
'    URLLocation = m_sURLLocation    ' Return URL string value
End Property 'Get URLLocation

Public Property Let URLLocation(ByVal Url As String)
'    If (m_sURLLocation <> Url) Then            ' Do only if value has changed...
'        m_sRegistered = "FALSE"
'        m_sURLLocation = Url                   ' Save URL string value to global variable
'        PropertyChanged pnURLLocation          ' Notify property bag of property change
'
'        On Error GoTo Err_Handler               ' Handle Error if URL is unavailable or Invalid...
'        If (Url <> "") Then
'            UserControl.AsyncRead Url & "\ctrCalendarVB.www", vbAsyncTypeByteArray, pnURLLocation  ' Begin async download of license file...
'        Else
'            m_sURLLocation = ""
'        End If
'        m_sRegistered = "AsyncRead"
'    End If
'    Exit Property
'
'Err_Handler:
'    m_sURLLocation = ""
'    Call CheckLicense
End Property 'Let URLLocation


'----------------------------------------------------------------------
' URLPicture Get/Let
'----------------------------------------------------------------------
' Purpose: Points to a URL location for the background picture source
'----------------------------------------------------------------------
Public Property Get URLPicture() As String
Attribute URLPicture.VB_Description = "A URL path to a picture image"
Attribute URLPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    URLPicture = m_sURLPicture                         ' Return URL string value
End Property 'Get URLPicture

Public Property Let URLPicture(ByVal Url As String)
    If (m_sURLPicture <> Url) Then                   ' Do only if value has changed...
        m_bClearPictureOnly = Not m_bClearURLOnly      ' If Picture property is not being set by the URLPicture
                                                     ' property then clear the URLPicture value...
        m_sURLPicture = Url                          ' Save URL string value to global variable
        PropertyChanged pnURLPicture                 ' Notify property bag of property change

        If Not m_bClearURLOnly Then
            On Error GoTo ErrorHandler               ' Handle Error if URL is unavailable or Invalid...
            If (Url <> "") Then
                UserControl.AsyncRead Url, vbAsyncTypePicture, pnPicture ' Begin async download of picture file...
            Else
                Set Picture = Nothing
            End If
        End If
    End If
ErrorHandler:
    m_bClearPictureOnly = False
End Property 'Let URLPicture

'----------------------------------------------------------------------
' Version Get
'----------------------------------------------------------------------
' Purpose:  Returns the controls version number
'----------------------------------------------------------------------
Public Property Get Version() As String
    Version = VERSION_NUMBER
End Property 'Version Get

'----------------------------------------------------------------------
' YearBegin Get/Let
'----------------------------------------------------------------------
' Purpose:  Sets the beginning year that can be selected from the
'           combo list box
'----------------------------------------------------------------------
Public Property Get YearBegin() As Integer
Attribute YearBegin.VB_Description = "The begin date to use for filling the dropdown year combolist box"
Attribute YearBegin.VB_ProcData.VB_Invoke_Property = ";Data"
    YearBegin = m_nYearBegin
End Property 'YearBegin Get

Public Property Let YearBegin(ByVal nYearBegin As Integer)
    If nYearBegin = 0 Then
        GoTo Err_Handler
    Else
        If (IsDate(DateSerial(nYearBegin, 1, 1)) = False) _
          Or (nYearBegin < 100) _
          Or (nYearBegin > 9989) _
          Or (nYearBegin > Year(m_Vars.DateValue)) Then
            GoTo Err_Handler
        End If
    End If
    m_nYearBegin = nYearBegin
    PropertyChanged pnYearBegin
    If m_bAutoPaint Then Call UpdateCombos
    
Exit_Proc:
    Exit Property
    
Err_Handler:
    If m_bDesign Then
        MsgBox "Please Enter a valid year. Must be equal to or less than the DateValue.", vbInformation, "Property Error"
    End If
    GoTo Exit_Proc
    
End Property 'YearBegin Let

'----------------------------------------------------------------------
' YearEnd Get/Let
'----------------------------------------------------------------------
' Purpose:  Sets the Ending year that can be selected from the
'           combo list box
'----------------------------------------------------------------------
Public Property Get YearEnd() As Integer
Attribute YearEnd.VB_Description = "The end date to use for filling the dropdown year combolist box"
    YearEnd = m_nYearEnd
End Property 'YearEnd Get

Public Property Let YearEnd(ByVal nYearEnd As Integer)
    If Len(nYearEnd) = 0 Or Len(nYearEnd) > 4 Or IsNumeric(nYearEnd) = False Then
        GoTo Err_Handler
    Else
        If (IsDate(DateSerial(nYearEnd, 1, 1)) = False) _
          Or (nYearEnd < 100) _
          Or (nYearEnd > 9989) Then
            GoTo Err_Handler
        End If
    End If
    m_nYearEnd = nYearEnd
    PropertyChanged pnYearEnd
    If m_bAutoPaint Then Call UpdateCombos
    
Exit_Proc:
    Exit Property
    
Err_Handler:
    If m_bDesign Then
        MsgBox "Please Enter a valid year. Must be equal to or less than the DateValue.", vbInformation, "Property Error"
    End If
    GoTo Exit_Proc
    
End Property 'YearEnd Let

'----------------------------------------------------------------------
' DateValue Get/Let
'----------------------------------------------------------------------
' Purpose:  Current Date Value selected
'----------------------------------------------------------------------
Public Property Get DateValue() As Date
Attribute DateValue.VB_Description = "The date value for the currently selected calendar cell"
Attribute DateValue.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute DateValue.VB_UserMemId = 0
    DateValue = m_Vars.DateValue
End Property 'DateValue Get

Public Property Let DateValue(ByVal dtDateValue As Date)
    Dim bAutoPaint      As Boolean
    Dim nPeriodValue    As Integer
    Dim nPeriodYear     As Integer
    Dim nYear           As Integer
    Dim dtDateValueOld  As Date
    'Error related declares
    Dim nAction             As Integer
    Dim lErrNumber          As Long
    Dim sErrDescription     As String
    Const sErrSource = "DateValue(Let)"
    On Error GoTo Err_Handler
    
DateValue_Begin:
    'Make sure that we have a valid date
    If Not IsDate(dtDateValue) Then
        dtDateValue = Date$
    End If
    
    'Save current values
    dtDateValueOld = DateValue
    nPeriodValue = PeriodValue
    nPeriodYear = PeriodYear
    
    'Only set the date if it falls within our YearBegin
    'and YearEnd range
    If IsInRange(dtDateValue, CalendarType) = False Then
        nAction = CalErrorActions.calAbort
        lErrNumber = CalErrorNumbers.calInvalidDateRange
        sErrDescription = "Invalid date. The date: " & dtDateValue & " does not fall within"
        sErrDescription = sErrDescription & " the BeginYear: " & YearBegin
        sErrDescription = sErrDescription & " and the EndYear: " & YearEnd & " ranges."
        RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, dtDateValue)
        Select Case nAction
        Case CalErrorActions.calRetry
            GoTo DateValue_Begin
        Case Else
            GoTo Exit_DateValue
        End Select
    End If
    
    m_Vars.DateValue = dtDateValue
    
    'Determine our period year
    nYear = m_Methods.DateToPeriodYear(dtDateValue, CalendarType, FirstDayOfWeek, YearStartPlacement)
    ' Update our vars
    nPeriodValue = m_Methods.DateToPeriod(dtDateValueOld, nYear, Periods, CalendarType _
        , FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement, dtDateValue)
    m_Vars.PeriodValue = nPeriodValue
    m_Vars.PeriodYear = nYear
    m_Vars.SetDateRanges nPeriodValue, PeriodYear, Periods, CalendarType _
        , FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement
        
    RaiseEvent DateChange(dtDateValueOld, dtDateValue)
    If m_bAutoPaint Then Refresh
    
Exit_DateValue:
    Call RefreshListIndexes
    Exit Property
    
Err_Handler:
    nAction = CalErrorActions.calAbort
    lErrNumber = CalErrorNumbers.calUnhandledError
    sErrDescription = "Unhandled Error. " & Error$
    RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, Null)
    Select Case nAction
    Case CalErrorActions.calAbort
        GoTo Exit_DateValue
    Case CalErrorActions.calIgnore
        Resume Next
    Case CalErrorActions.calRetry
        Resume
    Case Else
        GoTo Exit_DateValue
    End Select

End Property 'DateValue Let

'----------------------------------------------------------------------
' PeriodYear Get/Let
'----------------------------------------------------------------------
' Purpose:  Current Period Year
'----------------------------------------------------------------------
Public Property Get PeriodYear() As Integer
Attribute PeriodYear.VB_Description = "The year value for the currently active period"
Attribute PeriodYear.VB_ProcData.VB_Invoke_Property = ";Data"
    PeriodYear = m_Vars.PeriodYear
End Property 'PeriodYear Get

Public Property Let PeriodYear(ByVal nPeriodYear As Integer)

    Dim bCancel         As Boolean
    Dim dtDateValue     As Date
    Dim dtFOYDate       As Date
    Dim dtEOYDate       As Date
    'Error related declares
    Dim nAction             As Integer
    Dim lErrNumber          As Long
    Dim sErrDescription     As String
    Const sErrSource = "PeriodYear(Let)"
    On Error GoTo Err_Handler
    
PeriodYear_Begin:
    If nPeriodYear < YearBegin Or nPeriodYear > YearEnd Then
        nAction = CalErrorActions.calAbort
        lErrNumber = CalErrorNumbers.calInvalidPropertyValue
        sErrDescription = "Invalid Property Value. "
        sErrDescription = sErrDescription & "The period year must be in the range of " & YearBegin & " - " & YearEnd & "."
        RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, nPeriodYear)
        Select Case nAction
        Case CalErrorActions.calRetry
            GoTo PeriodYear_Begin
        Case Else
            GoTo Exit_PeriodYear
        End Select
    End If
    
    dtDateValue = DateSerial(nPeriodYear, Month(DateValue), Day(DateValue))
    
    RaiseEvent WillChangeDate(dtDateValue, bCancel)
    If bCancel Then GoTo Exit_PeriodYear
    
    m_Vars.PeriodYear = nPeriodYear
    Select Case CalendarType
    Case calMonth
        DateValue = dtDateValue
    Case calPeriod, calWeek
        dtFOYDate = m_Methods.FirstOfYearDate(nPeriodYear, FirstDayOfWeek, YearStartPlacement)
        dtEOYDate = m_Methods.EndOfYearDate(dtFOYDate, nPeriodYear, FirstDayOfWeek, YearStartPlacement)
        If dtDateValue < dtFOYDate Or dtDateValue > dtEOYDate Then
            PeriodValue = PeriodValue
        Else
            DateValue = dtDateValue
        End If
    End Select
    
Exit_PeriodYear:
    Call RefreshListIndexes
    Exit Property
Err_Handler:
    nAction = CalErrorActions.calAbort
    lErrNumber = CalErrorNumbers.calUnhandledError + Err
    sErrDescription = "Unhandled Error. " & Error$
    RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, Null)
    Select Case nAction
    Case CalErrorActions.calAbort
        GoTo Exit_PeriodYear
    Case CalErrorActions.calIgnore
        Resume Next
    Case CalErrorActions.calRetry
        Resume
    Case Else
        GoTo Exit_PeriodYear
    End Select

End Property 'PeriodYear Let

'----------------------------------------------------------------------
' Picture Get/Set
'----------------------------------------------------------------------
' Purpose: A picture that is tiled on the background of the calendar
'          control
'----------------------------------------------------------------------
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "A picture image that is tiled on the background of the calendar control"
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_picImage
End Property 'Picture Get

Public Property Set Picture(ByVal picPicture As StdPicture)
    'Validate what kind of picture is passed
    'Only allow bitmaps and icons
    'If not in runtime display message that UseMaskColor can't be
    'used with icons, if picture is icon.
    'If picture is icon, make sure UseMaskColor is false
    'Paint Control
    If Not picPicture Is Nothing Then
        With picPicture
            If (.Type <> vbPicTypeBitmap) And (.Type <> vbPicTypeNone) And (.Type <> vbPicTypeIcon) Then
                If Not UserControl.Ambient.UserMode Then
                    MsgBox "Invalid Picture Type", vbOKOnly, UserControl.Name
                End If
                Exit Property
            End If
        End With
    End If
    If Not m_bClearPictureOnly Then
        m_bClearURLOnly = True       ' If Picture property is not being set by the URLPicture
        URLPicture = ""             ' property then clear the URLPicture value...
        m_bClearURLOnly = False
    End If
    
    If (Not picPicture Is Nothing) Then
        If (picPicture.Handle = 0) Then Set picPicture = Nothing
    End If
    Set m_picImage = picPicture
    PropertyChanged pnPicture
    If m_bAutoPaint Then Refresh
End Property 'Picture Set

'----------------------------------------------------------------------
' DateTipFormat Get/Let
'----------------------------------------------------------------------
' Purpose:  Specifies the format of the date that will be displayed
'           as the tooltip
'----------------------------------------------------------------------
Public Property Get DateTipFormat() As String
Attribute DateTipFormat.VB_Description = "The format to use when displaying the date in the tooltip. Ex. dd-Mmm-yyyy"
Attribute DateTipFormat.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DateTipFormat = m_sDateTipFormat
End Property 'DateTipFormat Get

Public Property Let DateTipFormat(ByVal sDateTipFormat As String)
    m_sDateTipFormat = sDateTipFormat
    PropertyChanged pnDateTipFormat
End Property 'DateTipFormat Let

'----------------------------------------------------------------------
' PeriodValue Get/Let
'----------------------------------------------------------------------
' Purpose:  Sets or returns the current period for the calendar
'----------------------------------------------------------------------
Public Property Get PeriodValue() As Integer
Attribute PeriodValue.VB_Description = "The period value for the currently selected calendar cell"
Attribute PeriodValue.VB_ProcData.VB_Invoke_Property = ";Data"
    
    PeriodValue = m_Vars.PeriodValue
    
End Property 'PeriodValue Get

Public Property Let PeriodValue(ByVal nPeriodValue As Integer)

    Dim bCancel         As Boolean
    Dim dtDateValueNew  As Date
    Dim dtDateValueOld  As Date
    'Error related declares
    Dim nAction             As Integer
    Dim lErrNumber          As Long
    Dim sErrDescription     As String
    Const sErrSource = "PeriodValue(Let)"
    On Error GoTo Err_Handler
    
PeriodValue_Begin:
    If nPeriodValue < 1 Or nPeriodValue > Periods.Count Then
        nAction = CalErrorActions.calAbort
        lErrNumber = CalErrorNumbers.calInvalidPropertyValue
        sErrDescription = "Invalid Property Value. "
        sErrDescription = sErrDescription & "The period must be in the range of 1 - " & Periods.Count & "."
        RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, nPeriodValue)
        Select Case nAction
        Case CalErrorActions.calRetry
            GoTo PeriodValue_Begin
        Case Else
            GoTo Exit_PeriodValue
        End Select
    End If
    
    dtDateValueOld = DateValue
    dtDateValueNew = m_Methods.PeriodToDate(nPeriodValue, PeriodYear _
            , Periods, CalendarType, FirstDayOfWeek _
            , YearStartPlacement, ExtraWeekPlacement)
    'See if the programmer wishes to change to this date value
    'abort operation if canceled
    RaiseEvent WillChangeDate(dtDateValueNew, bCancel)
    If bCancel Then GoTo Exit_PeriodValue
    
    With m_Vars
        .SetDateRanges nPeriodValue, PeriodYear, Periods, CalendarType _
            , FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement
        .DateValue = dtDateValueNew
    End With
    
    RaiseEvent DateChange(dtDateValueOld, dtDateValueNew)
    If m_bAutoPaint Then Refresh
    
Exit_PeriodValue:
    Call RefreshListIndexes
    Exit Property
Err_Handler:
    nAction = CalErrorActions.calAbort
    lErrNumber = CalErrorNumbers.calUnhandledError + Err
    sErrDescription = "Unhandled Error. " & Error$
    RaiseEvent ErrorEvent(lErrNumber, sErrSource, sErrDescription, nAction, Null)
    Select Case nAction
    Case CalErrorActions.calAbort
        GoTo Exit_PeriodValue
    Case CalErrorActions.calIgnore
        Resume Next
    Case CalErrorActions.calRetry
        Resume
    Case Else
        GoTo Exit_PeriodValue
    End Select
End Property 'PeriodValue Let

'----------------------------------------------------------------------
' DayHeaderFormat Get/Set
'----------------------------------------------------------------------
' Purpose: Allows the user to select one of the following
'   Day Of Name Lengths;
'   Short: 1, Medium: Jan, Long: January
'----------------------------------------------------------------------
Public Property Get DayHeaderFormat() As DayHeaderFormats
Attribute DayHeaderFormat.VB_Description = "A value indicating whether to display the day of the week names in short, medium, or long weekname formats"
    DayHeaderFormat = m_nDayHeaderFormat
End Property 'Get DayHeaderFormat

Public Property Let DayHeaderFormat(ByVal nDayHeaderFormat As DayHeaderFormats)
    m_nDayHeaderFormat = nDayHeaderFormat
    PropertyChanged pnDayHeaderFormat
    Refresh
End Property 'Let DayHeaderFormat

'----------------------------------------------------------------------
' DayHeaderFont Get/Set
'----------------------------------------------------------------------
' Purpose:  Day Of Week Font
'----------------------------------------------------------------------
Public Property Get DayHeaderFont() As Font
Attribute DayHeaderFont.VB_Description = "The font attributes that are used to display the day of week header"
Attribute DayHeaderFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set DayHeaderFont = m_DayHeaderFont
End Property 'Get DayHeaderFont

Public Property Set DayHeaderFont(ByVal oDayHeaderFont As Font)
    Set m_DayHeaderFont = oDayHeaderFont
    Set cboPeriod.Font = m_DayHeaderFont
    Set cboYear.Font = m_DayHeaderFont
    PropertyChanged pnDayHeaderFont
    If m_bAutoPaint Then Refresh
End Property 'Set DayHeaderFont

'----------------------------------------------------------------------
' DayHeaderFontBold Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font bold style of the Days Font
'----------------------------------------------------------------------
Public Property Get DayHeaderFontBold() As Boolean
Attribute DayHeaderFontBold.VB_Description = "For the Day Of Week header"
Attribute DayHeaderFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DayHeaderFontBold.VB_MemberFlags = "400"
    DayHeaderFontBold = m_DayHeaderFont.Bold
End Property 'Get DayHeaderFontBold

Public Property Let DayHeaderFontBold(ByVal bDayHeaderFontBold As Boolean)
    m_DayHeaderFont.Bold = bDayHeaderFontBold
    PropertyChanged pnDayHeaderFontBold
    If m_bAutoPaint Then Refresh
End Property 'Set DayHeaderFontBold

'----------------------------------------------------------------------
' DayHeaderFontItalic Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font italic style of the Days Font
'----------------------------------------------------------------------
Public Property Get DayHeaderFontItalic() As Boolean
Attribute DayHeaderFontItalic.VB_Description = "For the Day Of Week header"
Attribute DayHeaderFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DayHeaderFontItalic.VB_MemberFlags = "400"
    DayHeaderFontItalic = m_DayHeaderFont.Italic
End Property 'Get DayHeaderFontItalic

Public Property Let DayHeaderFontItalic(ByVal bDayHeaderFontItalic As Boolean)
    m_DayHeaderFont.Italic = bDayHeaderFontItalic
    PropertyChanged pnDayHeaderFontItalic
    If m_bAutoPaint Then Refresh
End Property 'Let DayHeaderFontItalic

'----------------------------------------------------------------------
' DayHeaderFontName Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font style of the Days Font
'----------------------------------------------------------------------
Public Property Get DayHeaderFontName() As String
Attribute DayHeaderFontName.VB_Description = "For the Day Of Week header"
Attribute DayHeaderFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DayHeaderFontName.VB_MemberFlags = "400"
    DayHeaderFont = m_DayHeaderFont.Name
End Property 'Get DaysFontSize

Public Property Let DayHeaderFontName(ByVal sDayHeaderFontName As String)
    m_DayHeaderFont.Name = sDayHeaderFontName
    PropertyChanged pnDayHeaderFontName
    If m_bAutoPaint Then Refresh
End Property 'Set DayHeaderFontSize

'----------------------------------------------------------------------
' DayHeaderFontSize Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the Size of the Days Font
'----------------------------------------------------------------------
Public Property Get DayHeaderFontSize() As Long
Attribute DayHeaderFontSize.VB_Description = "For the Day Of Week header"
Attribute DayHeaderFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DayHeaderFontSize.VB_MemberFlags = "400"
    DayHeaderFontSize = m_DayHeaderFont.Size
End Property 'Get DayHeaderFontSize

Public Property Let DayHeaderFontSize(ByVal lDayHeaderFontSize As Long)
    m_DayHeaderFont.Size = lDayHeaderFontSize
    PropertyChanged pnDayHeaderFontSize
    If m_bAutoPaint Then Refresh
End Property 'Set DayHeaderFontSize

'----------------------------------------------------------------------
' DaysFont Get/Set
'----------------------------------------------------------------------
' Purpose:  Days Font
'----------------------------------------------------------------------
Public Property Get DaysFont() As Font
Attribute DaysFont.VB_Description = "The font attributes used for displaying the days in the calendar grid"
Attribute DaysFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set DaysFont = m_DaysFont
End Property 'Get DaysFont

Public Property Set DaysFont(ByVal oDaysFont As Font)
    Set m_DaysFont = oDaysFont
    PropertyChanged pnDaysFont
    If m_bAutoPaint Then Refresh
End Property 'Set DaysFont

'----------------------------------------------------------------------
' DaysFontBold Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font bold style of the Days Font
'----------------------------------------------------------------------
Public Property Get DaysFontBold() As Boolean
Attribute DaysFontBold.VB_Description = "For the Calendar Days"
Attribute DaysFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DaysFontBold.VB_MemberFlags = "400"
    DaysFontBold = m_DaysFont.Bold
End Property 'Get DaysFontBold

Public Property Let DaysFontBold(ByVal bDaysFontBold As Boolean)
    m_DaysFont.Bold = bDaysFontBold
    PropertyChanged pnDaysFontBold
    If m_bAutoPaint Then Refresh
End Property 'Set DaysFontBold

'----------------------------------------------------------------------
' DaysFontItalic Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font italic style of the Days Font
'----------------------------------------------------------------------
Public Property Get DaysFontItalic() As Boolean
Attribute DaysFontItalic.VB_Description = "For the Calendar Days"
Attribute DaysFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DaysFontItalic.VB_MemberFlags = "400"
    DaysFontItalic = m_DaysFont.Italic
End Property 'Get DaysFontItalic

Public Property Let DaysFontItalic(ByVal bDaysFontItalic As Boolean)
    m_DaysFont.Italic = bDaysFontItalic
    PropertyChanged pnDaysFontItalic
    If m_bAutoPaint Then Refresh
End Property 'Set DaysFontItalic

'----------------------------------------------------------------------
' DaysFontName Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font style of the Days Font
'----------------------------------------------------------------------
Public Property Get DaysFontName() As String
Attribute DaysFontName.VB_Description = "For the Calendar Days"
Attribute DaysFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DaysFontName.VB_MemberFlags = "400"
    DaysFontName = m_DaysFont.Name
End Property 'Get DaysFontSize

Public Property Let DaysFontName(ByVal sDaysFontName As String)
    m_DaysFont.Name = sDaysFontName
    PropertyChanged pnDaysFontName
    If m_bAutoPaint Then Refresh
End Property 'Set DaysFontSize

'----------------------------------------------------------------------
' DaysFontSize Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the Size of the Days Font
'----------------------------------------------------------------------
Public Property Get DaysFontSize() As Long
Attribute DaysFontSize.VB_Description = "For the Calendar Days"
Attribute DaysFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DaysFontSize.VB_MemberFlags = "400"
    DaysFontSize = m_DaysFont.Size
End Property 'Get DaysFontSize

Public Property Let DaysFontSize(ByVal lDaysFontSize As Long)
    m_DaysFont.Size = lDaysFontSize
    PropertyChanged pnDaysFontSize
    If m_bAutoPaint Then Refresh
End Property 'Set DaysFontSize

'----------------------------------------------------------------------
' Enabled Get/Set
'----------------------------------------------------------------------
' Purpose: Enables or Disables the calendar control
'----------------------------------------------------------------------
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Enables or disables the calendar control"
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property 'Get Enabled

Public Property Let Enabled(ByVal bEnabled As Boolean)
    Dim oDC As New CDraw
    'Enable/Disable our comboboxes
    cboPeriod.Enabled = bEnabled
    cboYear.Enabled = bEnabled
    'Modify the calendar to show an enabled or disabled state
    'If disabled the the control is shaded
    With UserControl
        If bEnabled And .Enabled = False Then
            'Only need to redraw the calendar if the
            'control was previously disabled
            .Enabled = bEnabled
            Refresh
        ElseIf bEnabled = False And .Enabled Then
            'Only need to draw the shading if the control
            'is currently enabled
            oDC.DrawStart m_RefreshDC.hdc, .ScaleWidth, .ScaleHeight, True
            oDC.ShadeRect m_udtDisabledRect.Left, m_udtDisabledRect.Top, m_udtDisabledRect.Width, m_udtDisabledRect.Height
            oDC.DrawStop m_udtDisabledRect.Left, m_udtDisabledRect.Top, m_udtDisabledRect.Width, m_udtDisabledRect.Height
            m_RefreshDC.CopyToHdc
            .Enabled = bEnabled
        End If
    End With
    PropertyChanged pnEnabled
End Property

'----------------------------------------------------------------------
' ExtraWeekPlacement Get/Set
'----------------------------------------------------------------------
' Purpose:  Where to place a leap years extra week, Beginning or
'           Ending of period
'----------------------------------------------------------------------
Public Property Get ExtraWeekPlacement() As ExtraWeekPlacements
Attribute ExtraWeekPlacement.VB_Description = "Tells the calendar control where to place the extra week for some leap years. The first period or the last period."
Attribute ExtraWeekPlacement.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ExtraWeekPlacement = m_nExtraWeekPlacement
End Property 'Get ExtraWeekPlacement

Public Property Let ExtraWeekPlacement(ByVal nExtraWeekPlacement As ExtraWeekPlacements)
    m_nExtraWeekPlacement = nExtraWeekPlacement
    PropertyChanged pnExtraWeekPlacement
    If m_bAutoPaint Then Refresh
End Property 'Let ExtraWeekPlacement

'----------------------------------------------------------------------
' ActiveDayFont Get/Set
'----------------------------------------------------------------------
' Purpose:  Active Day Font
'----------------------------------------------------------------------
Public Property Get ActiveDayFont() As Font
Attribute ActiveDayFont.VB_Description = "The currently selected dates font attributes"
Attribute ActiveDayFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set ActiveDayFont = m_ActiveDayFont
End Property 'Get ActiveDayFont

Public Property Set ActiveDayFont(ByVal oActiveDayFont As Font)
    Set m_ActiveDayFont = oActiveDayFont
    PropertyChanged pnActiveDayFont
    If m_bAutoPaint Then Refresh
End Property 'Set ActiveDayFont

'----------------------------------------------------------------------
' ActiveDayFontBold Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font bold style of the Days Font
'----------------------------------------------------------------------
Public Property Get ActiveDayFontBold() As Boolean
Attribute ActiveDayFontBold.VB_Description = "For the currently selected date"
Attribute ActiveDayFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ActiveDayFontBold.VB_MemberFlags = "400"
    ActiveDayFontBold = m_ActiveDayFont.Bold
End Property 'Get ActiveDayFontBold

Public Property Let ActiveDayFontBold(ByVal bActiveDayFontBold As Boolean)
    m_ActiveDayFont.Bold = bActiveDayFontBold
    PropertyChanged pnActiveDayFontBold
    If m_bAutoPaint Then Refresh
End Property 'Set ActiveDayFontBold

'----------------------------------------------------------------------
' ActiveDayFontItalic Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font italic style of the Days Font
'----------------------------------------------------------------------
Public Property Get ActiveDayFontItalic() As Boolean
Attribute ActiveDayFontItalic.VB_Description = "For the currently selected date"
Attribute ActiveDayFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ActiveDayFontItalic.VB_MemberFlags = "400"
    ActiveDayFontItalic = m_ActiveDayFont.Italic
End Property 'Get ActiveDayFontItalic

Public Property Let ActiveDayFontItalic(ByVal bActiveDayFontItalic As Boolean)
    m_ActiveDayFont.Italic = bActiveDayFontItalic
    PropertyChanged pnActiveDayFontItalic
    If m_bAutoPaint Then Refresh
End Property 'Set ActiveDayFontItalic

'----------------------------------------------------------------------
' ActiveDayFontName Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the font style of the Days Font
'----------------------------------------------------------------------
Public Property Get ActiveDayFontName() As String
Attribute ActiveDayFontName.VB_Description = "For the currently selected date"
Attribute ActiveDayFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ActiveDayFontName.VB_MemberFlags = "400"
    ActiveDayFontName = m_ActiveDayFont.Name
End Property 'Get ActiveDayFontSize

Public Property Let ActiveDayFontName(ByVal sActiveDayFontName As String)
    m_ActiveDayFont.Name = sActiveDayFontName
    PropertyChanged pnActiveDayFontName
    If m_bAutoPaint Then Refresh
End Property 'Set ActiveDayFontSize

'----------------------------------------------------------------------
' ActiveDayFontSize Get/Let
'----------------------------------------------------------------------
' Purpose:  Set the Size of the Days Font
'----------------------------------------------------------------------
Public Property Get ActiveDayFontSize() As Long
Attribute ActiveDayFontSize.VB_Description = "For the currently selected date"
Attribute ActiveDayFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ActiveDayFontSize.VB_MemberFlags = "400"
    ActiveDayFontSize = m_ActiveDayFont.Size
End Property 'Get ActiveDayFontSize

Public Property Let ActiveDayFontSize(ByVal lActiveDayFontSize As Long)
    m_ActiveDayFont.Size = lActiveDayFontSize
    PropertyChanged pnActiveDayFontSize
    If m_bAutoPaint Then Refresh
End Property 'Set ActiveDayFontSize

'----------------------------------------------------------------------
' Periods Get/Set
'----------------------------------------------------------------------
' Purpose: Allows a program access to the periods defenition object
'----------------------------------------------------------------------
Public Property Get Periods() As CCalendarVBPeriods
Attribute Periods.VB_Description = "This is the periods object that defines the period structure."
Attribute Periods.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Set Periods = m_Vars.Periods
End Property 'Get Periods

Public Property Set Periods(ByVal oPeriods As CCalendarVBPeriods)
    Set m_Vars.Periods = oPeriods
    PropertyChanged pnPeriods
    If m_bAutoPaint Then
        Call UpdateCombos
        Refresh
    End If
End Property 'Set periods

'----------------------------------------------------------------------
' PeriodRows Get/Let
'----------------------------------------------------------------------
' Purpose:  The number of rows to create for the calendar
'----------------------------------------------------------------------
Public Property Get PeriodRows() As Integer
Attribute PeriodRows.VB_Description = "The number of rows that are displayed on the calendar control."
Attribute PeriodRows.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PeriodRows = m_nPeriodRows
End Property 'Get PeriodRows

Public Property Let PeriodRows(ByVal nPeriodRows As Integer)
    m_nPeriodRows = nPeriodRows
    PropertyChanged pnPeriodRows
    If m_bAutoPaint Then Refresh
End Property 'Set PeriodRows

'**********************************************************************
' METHOD IMPLEMENTATION
'**********************************************************************

'----------------------------------------------------------------------
'Routine Name       :   (PUBLIC) AboutBox()
'Version            :   1.00.00
'Last Updated       :   10/23/97
'
'Display an AboutBox for the control
'----------------------------------------------------------------------
Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Displays the controls about box information dialog"
Attribute AboutBox.VB_UserMemId = -552
    
    Dim oLicense    As New CLicense
    
    With oLicense
        .licCLSID = LIC_CLSID
        .licKEY = LIC_KEY
        .ShowAboutBox imgIcon
    End With
    
    Set oLicense = Nothing
    
End Sub 'AboutBox()

'----------------------------------------------------------------------
'Routine Name       :   (PUBLIC) Refresh()
'Version            :   1.00.00
'Last Updated       :   9/2/97
'
'Updates the calendar display using any modified properties
'----------------------------------------------------------------------
Public Sub Refresh()
Attribute Refresh.VB_Description = "Redraws the calendar control"

    Dim bImageFlag      As Boolean
    Dim nRow            As Integer
    Dim nCol            As Integer
    Dim iRow            As Integer
    Dim iCol            As Integer
    Dim nCounter        As Integer
    Dim nCurrentDay     As Integer
    Dim nCurrentPeriod  As Integer
    Dim nHeaderHeight   As Integer      'Used for calculating the Disabled Rectangle Height
    Dim nLineStyle      As Integer
    Dim nOffset         As Integer
    Dim lPreviousFirst  As Long
    Dim lPreviousLast   As Long
    Dim lPreviousYear   As Long
    Dim lresult         As Long
    Dim sCaption        As String
    Dim sWork           As String
    Dim dtCurrentDate   As Date
    Dim oDC             As New CDraw
    Dim vCell           As Variant
    
    On Error Resume Next
    
    'Check to see if the control is visible.
    'Only needs to be drawn if the UserControl is visible
    If m_bActive Then
        'If UserControl.Extender.Visible = False Then Exit Sub
    End If
    
    'No need to continue if this is blank
    If Len(CStr(m_Vars.DateValue)) = 0 Then Exit Sub
    
    'Set our image flag
    If Not m_picImage Is Nothing Then
        'Only is set if this is a valid picture type
        bImageFlag = CBool(m_picImage.Type)
    End If
    
    'Compute the dates based on the calendar period
    Call m_Vars.SetDateRanges(m_Vars.PeriodValue, PeriodYear, Periods, CalendarType, FirstDayOfWeek _
        , YearStartPlacement, ExtraWeekPlacement)
    nCurrentDay = Day(m_Vars.CalendarStart)
    
    'Init our variant that holds the cell location and demensions that will be
    'used for drawing newly selected dates from the calendar
    ReDim m_vaCellLocations(nCalendarRows(), DEF_CALENDAR_COLS)
    ReDim vCell(DEF_CELL_UBOUND)
    
    'Init our drawing object to begin creating the calendar in the memory DC
    oDC.DrawStart UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight, False
    
    With oDC
        Set .Font = UserControl.Font
        'If there is an image then lets tile it on the background
        If bImageFlag Then
            .TileBitmap m_picImage, 0, 0, UserControl.ScaleX(m_picImage.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_picImage.Height, vbHimetric, vbPixels), 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
        
        'Determine the Width of our calendar cells
        'With a flat or noline border the borders are not as thick as the 3D version
        m_nCellWidth = ((UserControl.ScaleWidth) \ DEF_CALENDAR_COLS) - nOffset
        
        'Resize and move our combo boxes, incase the UserControl has been resized or
        'the calendar line type has changed which will affect the size of the control
        lresult = ((m_nCellWidth * DEF_CALENDAR_COLS) - nCalendarRows()) \ 3  'Make the year combo a third of the size of the period combo
        cboPeriod.Move DEF_LEFT_MARGIN, 2, ((m_nCellWidth * DEF_CALENDAR_COLS) - 6) - lresult
        cboYear.Move cboPeriod.Width + 6 + DEF_LEFT_MARGIN, 2, (UserControl.ScaleWidth - (cboPeriod.Width + 6 + DEF_LEFT_MARGIN)) - DEF_LEFT_MARGIN 'lresult
        
        'Set the beginning of the first calendar row
        nRow = DEF_TOP_MARGIN
        
        'Display the days of the week header, if enabled
        If m_bShowDayHeader Then
            Set .Font = DayHeaderFont
            .BackColor = m_oDayHeaderBackColor
            .ForeColor = m_oDayHeaderForeColor
            'If the flat grid line style has been choosen then our cell
            'height is not as tall
            If m_nLineStyle = Flat Then
                .FlatLineColor = FlatLineColor
                nOffset = 3
            Else
                nOffset = 4
            End If
            m_nCellHeight = UserControl.TextHeight("Wednesday") + nOffset
            nHeaderHeight = m_nCellHeight
            
            nCurrentPeriod = m_nFirstCurrentPeriod
            nCol = DEF_LEFT_MARGIN
            'Display the day of the week caption based on the users
            'selection for the formatting
            For iCol = 1 To DEF_CALENDAR_COLS
                Select Case m_nDayHeaderFormat
                Case calOneLetterName
                    sCaption = Left$(Format$(nCurrentPeriod, "Ddd"), 1)
                Case calTwoLetterName
                    sCaption = Left$(Format$(nCurrentPeriod, "Ddd"), 2)
                Case calThreeLetterName
                    sCaption = Format$(nCurrentPeriod, "Ddd")
                Case calFullName
                    sCaption = Format$(nCurrentPeriod, "Dddd")
                End Select
                'Account for a smaller cells when drawing a flat calendar
                If m_nLineStyle = NoLines Then
                    nOffset = 2
                ElseIf m_nLineStyle = Flat Then
                    nOffset = 1
                Else
                    nOffset = 0
                End If
                .Draw3DRect nCol, DEF_TOP_MARGIN, m_nCellWidth + nOffset, m_nCellHeight + nOffset, sCaption, caCenterCenter, m_nLineStyle
                nCol = nCol + m_nCellWidth
                nCurrentPeriod = nCurrentPeriod + 1
                If nCurrentPeriod > 7 Then nCurrentPeriod = 1
            Next
        End If
        'If the days of the week are being displayed then account for them in our
        'cell height calculations
        If m_bShowDayHeader Then
            'Account for smaller cells when the flat calendar style is choosen
            If m_nLineStyle = NoLines Then
                nOffset = 2
            ElseIf m_nLineStyle = Flat Then
                nOffset = 3
            Else
                nOffset = 4
            End If
            m_nCellHeight = ((UserControl.ScaleHeight - DEF_TOP_MARGIN - m_nCellHeight) \ nCalendarRows())
            If m_nLineStyle = Flat Then m_nCellHeight = m_nCellHeight - 1
            nRow = UserControl.TextHeight("Wednesday") + nOffset + DEF_TOP_MARGIN
        Else
            m_nCellHeight = ((UserControl.ScaleHeight - DEF_TOP_MARGIN) \ nCalendarRows())
            nRow = DEF_TOP_MARGIN
        End If
        
        'Create the grid and display the day value
        dtCurrentDate = m_Vars.CalendarStart
        Set .Font = m_DaysFont
        For iRow = 1 To nCalendarRows()
            nCol = DEF_LEFT_MARGIN
            For iCol = 1 To DEF_CALENDAR_COLS
                'Get the next date value that will be displayed
                dtCurrentDate = DateAdd("d", nCounter, m_Vars.CalendarStart)
                'Determine which background and foreground colors to use
                If dtCurrentDate < m_Vars.PeriodStart Then
                    .BackColor = m_oPrePeriodBackColor
                    .ForeColor = m_oPrePeriodForeColor
                ElseIf (dtCurrentDate >= m_Vars.PeriodStart) And (dtCurrentDate <= m_Vars.PeriodEnd) Then
                    .BackColor = m_oCurrentPeriodBackColor
                    .ForeColor = m_oCurrentPeriodForeColor
                Else
                    .BackColor = m_oPostPeriodBackColor
                    .ForeColor = m_oPostPeriodForeColor
                End If
                'Need to set the foreground color here just incase this happens to be
                'the currently active date will have the correct forecolor
                vCell(DEF_CELL_FORE_COLOR) = .ForeColor
                If dtCurrentDate = DateValue Then
                    Set .Font = m_ActiveDayFont
                    .ForeColor = m_oActiveDayForeColor
                End If
                'Account for different line thickness based on the grid style choosen
                If m_nLineStyle = NoLines Then
                    nOffset = 2
                ElseIf m_nLineStyle = Flat Then
                    .FlatLineColor = FlatLineColor
                    nOffset = 1
                Else
'                    If dtCurrentDate = DateValue Then
'                        .Draw3DRect nCol, nRow, m_nCellWidth, m_nCellHeight, Day(dtCurrentDate), m_nDayNumberAlignment, Selected
'                    Else
                        nOffset = 0
'                    End If
                End If
                .Draw3DRect nCol, nRow, m_nCellWidth + nOffset, m_nCellHeight + nOffset, Day(dtCurrentDate), m_nDayNumberAlignment, m_nLineStyle
                'Save our date cell location and size information
                vCell(DEF_CELL_BACK_COLOR) = .BackColor
                vCell(DEF_CELL_LEFT) = nCol
                vCell(DEF_CELL_TOP) = nRow
                'Account for the line thickness differences
                If m_nLineStyle = NoLines Then
                    nOffset = 2
                ElseIf m_nLineStyle = Flat Then
                    nOffset = 1
                Else
                    nOffset = 0
                End If
                vCell(DEF_CELL_WIDTH) = m_nCellWidth + nOffset
                vCell(DEF_CELL_HEIGHT) = m_nCellHeight + nOffset
                m_vaCellLocations(iRow, iCol) = vCell
                If dtCurrentDate = DateValue Then
                    Set .Font = m_DaysFont
                    m_nRow = iRow
                    m_nCol = iCol
                End If
                nCol = nCol + m_nCellWidth
                nCounter = nCounter + 1
            Next
            nRow = nRow + m_nCellHeight
        Next
        
        'Set our Disabled Rect values
        m_udtDisabledRect.Left = DEF_LEFT_MARGIN
        m_udtDisabledRect.Top = DEF_TOP_MARGIN
        If m_nLineStyle = NoLines Then
            nOffset = 4
        ElseIf m_nLineStyle = Flat Then
            nOffset = 2
        Else
            nOffset = 0
        End If
        m_udtDisabledRect.Width = (m_nCellWidth * DEF_CALENDAR_COLS) + nOffset
        m_udtDisabledRect.Height = (m_nCellHeight * nCalendarRows()) + nHeaderHeight + 1
        
        'If the control is disabled then shade it
        If Enabled = False Then
            'Only need to draw the shading if the control
            'is currently enabled
            .ShadeRect m_udtDisabledRect.Left, m_udtDisabledRect.Top, m_udtDisabledRect.Width, m_udtDisabledRect.Height
        End If
        
        'Lets make a memory DC for our repaints, this way we don't have to
        'completely redraw the image using this refresh routine.
        Set m_RefreshDC = Nothing
        m_RefreshDC.Attach UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight, False
        m_RefreshDC.CopyFromHdc .hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        'All done drawing our control so lets display it
        .DrawStop 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End With
    
    'If we are in calWeek mode then we need to check our dropdown
    'list box for the proper entries and update if necessary
    If CalendarType = calWeek Then
        If m_Methods.IsExtraWeek(m_Vars.FirstOfYear, PeriodYear, FirstDayOfWeek, YearStartPlacement) Then
            If cboPeriod.ListCount <> 53 Then UpdateCombos
        Else
            If cboPeriod.ListCount > 52 Then UpdateCombos
        End If
    End If
    'Check to be sure that both ComboBoxes have entries
    If cboPeriod.ListCount = 0 Or cboYear.ListCount = 0 Then _
        Call UpdateCombos
        
End Sub 'Refresh

Private Sub UpdateCombos()

    Dim iIndex          As Integer
    Dim nCount          As Integer
    Dim nWeeks          As Integer
    Dim dtYearBegin     As Date
    Dim dtYearEnd       As Date
    
    On Error GoTo Err_Handler
    
    cboPeriod.Clear
    Select Case m_nCalendarType
    Case calMonth
        For iIndex = 1 To 12
            cboPeriod.AddItem Format$(DateSerial(1997, iIndex, 1), "Mmmm")
        Next
        cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, Format$(m_Vars.DateValue, "Mmmm"))
    Case calPeriod
        'Add the period names from our Periods object and update our
        'NumberOfWeeks array. I choose to store this information in an
        'array rather than the Periods object for speed reasons.
        nCount = m_Vars.Periods.Count
        For iIndex = 1 To nCount
            With m_Vars.Periods(iIndex)
                cboPeriod.AddItem .Name
            End With
        Next
        cboPeriod.ListIndex = CInt(m_Vars.PeriodValue) - 1
    Case calWeek
        If m_Methods.IsExtraWeek(m_Vars.FirstOfYear, PeriodYear, FirstDayOfWeek, YearStartPlacement) Then
            nWeeks = 53
        Else
            nWeeks = 52
        End If
        For iIndex = 1 To nWeeks
            cboPeriod.AddItem "Week " & iIndex
        Next
    End Select
    
    'Fill our Year combobox
    cboYear.Clear
    dtYearBegin = VBA.DateValue("1/1/" & m_nYearBegin)
    dtYearEnd = VBA.DateValue("12/31/" & m_nYearEnd)
    'Check for a valid date range
    If IsInRange(DateValue, CalendarType) = False Then Exit Sub
        
    'Determine how many items we will be adding to the combolist
    nCount = DateDiff("yyyy", dtYearBegin, dtYearEnd) + 1
    
    For iIndex = 0 To nCount - 1
        cboYear.AddItem Format$(DateSerial(Year(DateAdd("yyyy", iIndex, dtYearBegin)), 1, 1), "yyyy")
    Next
    
Exit_Proc:
    Exit Sub
    
Err_Handler:
    GoTo Exit_Proc
End Sub 'UpdateCombos

'----------------------------------------------------------------------
' GetCellLocation()
'----------------------------------------------------------------------
' Purpose:  Row and Col that corresponds to the X and Y position
' Inputs:   X, Y mouse positions
' Outputs:  Row and Col, [DateValue] Returns true if error encountered
'----------------------------------------------------------------------
Private Function GetCellLocation(ByVal x As Single, ByVal y As Single, nRow As Integer, nCol As Integer, Optional dtDateValue As Date) As Boolean

    Dim nTopMargin      As Integer
    Dim nWidth          As Integer
    
    On Error GoTo Err_Handler
    
    'Calculate the top margin
    m_Methods.CopyFont DayHeaderFont, UserControl.Font
    If m_bShowDayHeader Then
        nTopMargin = UserControl.TextHeight("Wednesday") + 4 + DEF_TOP_MARGIN
    Else
        nTopMargin = DEF_TOP_MARGIN
    End If
    
    'Test to be sure that the mouse cursor is located on the calendar
    If nTopMargin >= y Then
        GetCellLocation = True
        GoTo Exit_Proc
    End If
   
    'If the flat grid style is selected then adjust the topmargin
    'If m_nLineStyle = Flat Then nTopMargin = nTopMargin - 1
    
    If m_nLineStyle = Flat Then
        nRow = (y - nTopMargin) \ (((UserControl.ScaleHeight - nTopMargin) \ nCalendarRows()) - 1) + 1
    Else
        nRow = (y - nTopMargin) \ ((UserControl.ScaleHeight - nTopMargin) \ nCalendarRows()) + 1
    End If
    nCol = ((x - DEF_LEFT_MARGIN) \ ((UserControl.ScaleWidth - DEF_LEFT_MARGIN) \ DEF_CALENDAR_COLS)) + 1
    
Exit_Proc:
    'Check to be sure that we have a valid row and column
    If nRow < 1 Or nRow > nCalendarRows() Or nCol < 1 Or nCol > DEF_CALENDAR_COLS Then
        GetCellLocation = True
    Else
        dtDateValue = DateAdd("d", (((nRow - 1) * DEF_CALENDAR_COLS) + nCol) - 1, m_Vars.CalendarStart)
    End If
    Exit Function
    
Err_Handler:
    GetCellLocation = True
    
End Function 'GetCellLocation()

'----------------------------------------------------------------------
' GetDateLocation()
'----------------------------------------------------------------------
' Purpose:  Row and Col that corresponds to the X and Y position
' Inputs:   DateValue
' Outputs:  Row and Col, Returns true if error encountered
'----------------------------------------------------------------------
Private Function GetDateLocation(ByVal dtDateValue As Date, nRow As Integer, nCol As Integer) As Boolean

    On Error GoTo Err_Handler
    
    nRow = (DateDiff("d", m_Vars.CalendarStart, dtDateValue) \ DEF_CALENDAR_COLS) + 1
    nCol = (DateDiff("d", m_Vars.CalendarStart, dtDateValue) Mod DEF_CALENDAR_COLS) + 1
    'Check for a valid row and column
    If nRow < 1 Or nRow > nCalendarRows() Or nCol < 1 Or nCol > DEF_CALENDAR_COLS Then
        GetDateLocation = True
    End If
    
Exit_Proc:
    Exit Function
    
Err_Handler:
    GetDateLocation = True
    
End Function 'GetDateLocation()

'----------------------------------------------------------------------
' SetCalendarStyle
'----------------------------------------------------------------------
' Purpose:  Changes the current Calendar Style to a new one
' Inputs:   New calendar style
' Outputs:  none
'----------------------------------------------------------------------
Private Sub SetCalendarStyle(ByVal nCalendarStyle As CalendarTypes)

    Dim bAutoPaint      As Boolean
    Dim nPeriodValue    As Integer
    Dim nYear           As Integer
    Dim dtDateValue     As Date
    
    If IsInRange(DateValue, nCalendarStyle) = False Then Exit Sub
    
    bAutoPaint = AutoPaint
    AutoPaint = False
    m_bActive = False
    nYear = PeriodYear
    If m_Vars.Isinitialized Then
        If (DateValue < m_Vars.FirstOfYear) Then
            nYear = nYear - 1
        ElseIf (DateValue > m_Vars.EndOfYear) Then
            nYear = nYear + 1
        End If
    End If
    nPeriodValue = m_Methods.DateToPeriod(DateValue, nYear, Periods _
        , CalendarType, FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement)
    m_Vars.PeriodValue = nPeriodValue
    m_Vars.PeriodYear = nYear
    m_Vars.SetDateRanges PeriodValue, PeriodYear, Periods, nCalendarStyle, FirstDayOfWeek _
        , YearStartPlacement, ExtraWeekPlacement
    Call UpdateCombos
    Call RefreshListIndexes
    m_bActive = True
    AutoPaint = bAutoPaint
    
End Sub

'----------------------------------------------------------------------
' FocusRect
'----------------------------------------------------------------------
' Purpose:  draw a focus rect to signify that the calendar
'           area now has focus, or turns off the focus rectangle
'           if one already exists
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Private Sub FocusRect()
    
    Dim vCell As Variant
    
    'Get our cell demensions and adjust for our focus rectangle
    If m_nRow < 1 And m_nCol < 1 Then Exit Sub
    vCell = m_vaCellLocations(m_nRow, m_nCol)
    m_udtFocusArea.Left = vCell(DEF_CELL_LEFT) + 2
    m_udtFocusArea.Top = vCell(DEF_CELL_TOP) + 2
    m_udtFocusArea.Right = vCell(DEF_CELL_LEFT) + vCell(DEF_CELL_WIDTH) - 2
    m_udtFocusArea.Bottom = vCell(DEF_CELL_TOP) + vCell(DEF_CELL_HEIGHT) - 2
    DrawFocusRect m_RefreshDC.hdc, m_udtFocusArea
    'Update our Memory DC copy of the calendar display
    m_RefreshDC.CopyToHdc m_udtFocusArea.Left, m_udtFocusArea.Top, vCell(DEF_CELL_WIDTH) - 2, vCell(DEF_CELL_HEIGHT) - 2
    
End Sub 'FocusRect()

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property
Public Property Get hWndPeriod() As Long
   hWndPeriod = cboPeriod.hwnd
End Property
Public Property Get hWndYear() As Long
   hWndYear = cboYear.hwnd
End Property

'----------------------------------------------------------------------
' ChangeValue
'----------------------------------------------------------------------
' Purpose:  Change to a new date value. Also update the calendar display
' Inputs:   Date to change to the new date
' Outputs:  none
'----------------------------------------------------------------------
Private Sub ChangeValue(ByVal dtDate As Date)

    Dim bAutoPaint  As Boolean
    Dim bCancel     As Boolean
    Dim nCol        As Integer
    Dim nResult     As Integer
    Dim nRow        As Integer
    Dim nPeriodYear As Integer
    Dim dtOldDate   As Date
    Dim oDC         As New CDraw
    Dim vCell       As Variant
    
    'give the developer a chance to cancel the date change
    bCancel = False
    RaiseEvent WillChangeDate(dtDate, bCancel)
    If bCancel Then Exit Sub
    
    'Save the current AutoPaint setting
    bAutoPaint = AutoPaint
    
    'Save the current Date
    dtOldDate = DateValue
    
    'Turn the focus rectangle off for the current
    'location
    Call FocusRect
    'See if the date is for the current period
    If GetDateLocation(dtDate, nRow, nCol) = False Then
        'Same period so we need to move our row and column
        'variables to point to the current location and then
        'the new location within the same period
        m_nLastRow = m_nRow
        m_nLastCol = m_nCol
        m_nRow = nRow
        m_nCol = nCol
    End If
    
    'This portion of code checks to see if the date falls
    'with in the period type year. This takes into account
    'that a period year can start on the previous years last
    'week.
    If (dtDate < m_Vars.PeriodStart) _
      Or (dtDate > m_Vars.PeriodEnd) Then
        'This is a new period so we adjust the dates, period, and
        'period year as necessary and then redraw the calendar
        If (dtDate < m_Vars.FirstOfYear) Then
            nPeriodYear = PeriodYear - 1
        ElseIf (dtDate > m_Vars.EndOfYear) Then
            nPeriodYear = PeriodYear + 1
        Else
            nPeriodYear = PeriodYear
        End If
        'Turn AutoPaint off so that nothing happens until we've
        'set all of our properties as needed
        AutoPaint = False
        nResult = _
            m_Methods.DateToPeriod(DateValue, nPeriodYear, Periods _
            , CalendarType, FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement, dtDate)
        m_bActive = False
        m_Vars.DateValue = dtDate
        m_Vars.PeriodValue = nResult
        m_Vars.PeriodYear = nPeriodYear
        m_Vars.SetDateRanges nResult, nPeriodYear, Periods, CalendarType _
            , FirstDayOfWeek, YearStartPlacement, ExtraWeekPlacement
        'Update the dropdown listboxes
        cboYear.ListIndex = m_Methods.SetComboText(cboYear, CStr(nPeriodYear))
        Select Case CalendarType
        Case CalendarTypes.calMonth
            If cboPeriod.ListIndex <> nResult - 1 Then _
                cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, Format$(DateValue, "Mmmm"))
        Case CalendarTypes.calPeriod
            If cboPeriod.ListIndex <> nResult - 1 Then _
                cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, m_Vars.Periods(nResult).Name)
        Case CalendarTypes.calWeek
            If cboPeriod.ListIndex <> nResult - 1 Then _
                cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, "Week " & nResult)
        End Select
        m_bActive = True
        'Redraws the calendar control using our new modified property
        'values
        AutoPaint = bAutoPaint
        Call FocusRect
        'Fire our date change event
        RaiseEvent DateChange(dtOldDate, DateValue)
        Exit Sub
    End If

    'The date is within the current period so we need to redraw the current location
    'and the new location with the proper fonts and colors. We also move the focus
    'rectangle from the old location to the new location.
    
    'Clear the last active cell and show the new cell as active
    'Init our memory DC
    oDC.DrawStart UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight, True
    With oDC
        'Set font to the DaysFont object
        Set .Font = m_DaysFont
        
        'Reset the last cell location
        vCell = m_vaCellLocations(m_nLastRow, m_nLastCol)
        'This should never be the case, but just incase...
        If IsEmpty(vCell) Then Exit Sub
        .BackColor = vCell(DEF_CELL_BACK_COLOR)
        .ForeColor = vCell(DEF_CELL_FORE_COLOR)
        .FlatLineColor = FlatLineColor
        .Draw3DRect vCell(DEF_CELL_LEFT), vCell(DEF_CELL_TOP), vCell(DEF_CELL_WIDTH) _
          , vCell(DEF_CELL_HEIGHT), Day(m_Vars.DateValue), m_nDayNumberAlignment, m_nLineStyle
        'Now copy the memory image that we created, to the UserControl replacing the existing
        'calendar cell. This is much faster than having to redraw the entire calendar.
        .CopyDC CLng(vCell(DEF_CELL_LEFT)), CLng(vCell(DEF_CELL_TOP)), CLng(vCell(DEF_CELL_WIDTH)) _
          , CLng(vCell(DEF_CELL_HEIGHT))
        
        'Now lets do the same thing for the new active cell
        m_Vars.DateValue = dtDate
        Set .Font = m_ActiveDayFont
        .ForeColor = m_oActiveDayForeColor
        vCell = m_vaCellLocations(m_nRow, m_nCol)
        'Testing one, two, three....
        If IsEmpty(vCell) Then Exit Sub
        .BackColor = vCell(DEF_CELL_BACK_COLOR)
        .Draw3DRect vCell(DEF_CELL_LEFT), vCell(DEF_CELL_TOP), vCell(DEF_CELL_WIDTH) _
          , vCell(DEF_CELL_HEIGHT), Day(m_Vars.DateValue), m_nDayNumberAlignment, m_nLineStyle
        'All done drawing so lets copy the memory image to the UserControl
        .DrawStop CLng(vCell(DEF_CELL_LEFT)), CLng(vCell(DEF_CELL_TOP)), CLng(vCell(DEF_CELL_WIDTH)) _
          , CLng(vCell(DEF_CELL_HEIGHT))
    End With
    'Lets make a memory DC for our repaints, this way we don't have to
    'completely redraw the image using the refresh routine. Zippy do
    'fast.....
    Set m_RefreshDC = Nothing
    m_RefreshDC.Attach UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight, False
    m_RefreshDC.CopyFromHdc oDC.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    'draw a focus rect at the new location
    If ActiveControl.Name <> "ctlFocus" Then
        'If the focus is not currrently at our focus
        'control then we need to change the focus so that
        'we can trap the keyevents
        ctlFocus.SetFocus
    Else
        Call FocusRect
    End If
    
    'Fire our date change event
    RaiseEvent DateChange(dtOldDate, DateValue)
    
End Sub 'ChangeValue()

'----------------------------------------------------------------------
' SetDesignMode
'----------------------------------------------------------------------
' Purpose:  Sets a module level flag to indicate if in design mode
' Inputs:   none
' Outputs:  Modual level flag
'----------------------------------------------------------------------
Private Sub SetDesignMode()
    On Error Resume Next
    m_bDesign = Not Ambient.UserMode
    If Err Then m_bDesign = True
    Err.Clear
End Sub 'SetDesignMode

'----------------------------------------------------------------------
' MouseCapture
'----------------------------------------------------------------------
' Purpose:  Is the only place where setcapture and releasecapture are called
'           setcapture may be called after mouse clicks because VB seems to
'           release capture on my behalf.
' Inputs:   True to start capture, False to release capture
' Outputs:  none
'----------------------------------------------------------------------
Private Sub MouseCapture(bCapture As Boolean)
    If bCapture Then
        SetCapture UserControl.hwnd
    Else
        ReleaseCapture
    End If
End Sub 'MouseCapture

'----------------------------------------------------------------------
' DateTipDisplay
'----------------------------------------------------------------------
' Purpose:  Displays or hides the tooltip window.
' Inputs:   Show/Hide, ToolTip caption
' Outputs:  none
'----------------------------------------------------------------------
Private Sub DateTipDisplay(ByVal bShow As Boolean, Optional ByVal sCaption As String = "")

    If bShow Then
        Call MouseCapture(True)
        m_bToolTipVisible = True
        Set m_ToolTip = New FToolTip
        m_ToolTip.DisplayToolTip sCaption, sCaption
    Else
        Call MouseCapture(False)
        If m_bToolTipVisible Then
            m_ToolTip.HideToolTip
            Set m_ToolTip = Nothing
            m_bToolTipVisible = False
            m_RefreshDC.CopyToHdc 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
    End If

End Sub 'DateTipDisplay

'----------------------------------------------------------------------
' DateTipDisplay
'----------------------------------------------------------------------
' Purpose:  Sets the Periods object structure back to 4
'           week intervals for a total of 13 periods
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Public Sub PeriodDefault()
Attribute PeriodDefault.VB_Description = "Sets the Periods object structure back to 4 week intervals for a total of 13 periods"

    Dim nCount          As Integer
    Dim iPeriodCounter  As Integer
    Dim iIndex          As Integer
    
    'Init our periods
    With m_Vars
        Set .Periods = New CCalendarVBPeriods
        For iIndex = 1 To 13
            iPeriodCounter = iPeriodCounter + 1
            .Periods.Add "Period " & iIndex, 4
        Next
    End With
    m_bAutoPaint = True
    
End Sub

'----------------------------------------------------------------------
'Routine Name       :   (PRIVATE) nCalendarRows()
'Version            :   1.00.00
'Last Updated       :   09/23/1997
'Modifed By         :   Mike Gainer
'
'Returns the number of rows for the calendar based on the calendar type
'----------------------------------------------------------------------
Private Function nCalendarRows() As Integer

    Select Case m_nCalendarType
        Case calMonth, calWeek
            nCalendarRows = DEF_CALENDAR_ROWS
        Case calPeriod
            nCalendarRows = m_nPeriodRows
    End Select

End Function

'----------------------------------------------------------------------
' CheckLicense Method
'----------------------------------------------------------------------
' Purpose: Check to see if the user has a license to use this control
'          in the design enviroment.
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub CheckLicense()

    Dim oLicense    As New CLicense
    Dim sName       As String
    
    With oLicense
        .licCLSID = LIC_CLSID
        .licKEY = LIC_KEY
        If UserControl.Ambient.UserMode = True Then
            'Get the program name and compare with the name
            'stored in the PropertyBag
            sName = .GetProgramName(GetParent(UserControl.hwnd))
            If m_sEnvironment <> sName Then
                If m_sRegistered = "FALSE" Then
                    If .IsRegistered = False Then .ShowAboutBox
                End If
            End If
        Else
            'Were in the design environment so we need to
            'check the registery for a valid license key
            Call .CheckRegistration
        End If
    End With
    
    Set oLicense = Nothing
    
End Sub

'----------------------------------------------------------------------
'Routine Name       :   (Private) IsInRange()
'Version            :   1.00.00
'Last Updated       :   10/29/1997
'Modifed By         :   Mike Gainer
'
'Determines if the date value falls with in the YearBegin and YearEnd
'ranges based on the calendar type.
'
'----------------------------------------------------------------------
Private Function IsInRange(ByVal dtDate As Date, ByVal nCalendarType As CalendarTypes) As Boolean
    
    Dim nYear                   As Integer
    Dim dtLowRange              As Date
    Dim dtHighRange             As Date
    Dim dtHighRangeFirst        As Date
    
    Select Case nCalendarType
    Case calMonth
        If (dtDate < DateSerial(YearBegin, 1, 1)) _
            Or (dtDate > DateSerial(YearEnd, 12, 31)) Then Exit Function
    Case calPeriod, calWeek
        dtLowRange = m_Methods.FirstOfYearDate(YearBegin, FirstDayOfWeek, YearStartPlacement)
        dtHighRangeFirst = m_Methods.FirstOfYearDate(YearEnd, FirstDayOfWeek, YearStartPlacement)
        dtHighRange = m_Methods.EndOfYearDate(dtHighRangeFirst, YearEnd, FirstDayOfWeek, YearStartPlacement)
        If dtDate < dtLowRange Or dtDate > dtHighRange Then Exit Function
    End Select
    
    IsInRange = True
    
End Function

'----------------------------------------------------------------------
'Routine Name       :   (Private) RefreshListIndexes()
'Version            :   1.00.00
'Last Updated       :   10/29/1997
'Modifed By         :   Mike Gainer
'
'Sets the ComboList box listIndexes to reflect the current
'PeriodValue and PeriodYear values.
'----------------------------------------------------------------------
Private Sub RefreshListIndexes()

    On Error Resume Next
    
    m_bActive = False
    cboYear.ListIndex = m_Methods.SetComboText(cboYear, CStr(PeriodYear))
    Select Case CalendarType
    Case CalendarTypes.calMonth
        If cboPeriod.ListIndex <> PeriodValue - 1 Then _
            cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, Format$(DateValue, "Mmmm"))
    Case CalendarTypes.calPeriod
        If cboPeriod.ListIndex <> PeriodValue - 1 Then _
            cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, m_Vars.Periods(PeriodValue).Name)
    Case CalendarTypes.calWeek
        If cboPeriod.ListIndex <> PeriodValue - 1 Then _
            cboPeriod.ListIndex = m_Methods.SetComboText(cboPeriod, "Week " & PeriodValue)
    End Select
    m_bActive = True
    
End Sub

