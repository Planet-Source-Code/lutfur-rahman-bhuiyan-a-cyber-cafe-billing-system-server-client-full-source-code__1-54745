VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalendarVBVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================================
'  Name [CCalendarVBVars.cls]
'
'  Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        10-29-1997
'===========================================================================================
'
'$Runtime Dependencies:
'$DesignTime Dependencies:
'   CCalendarVBPeriods.cls
'
'===========================================================================================
'  Usage Notes:
'
'
'===========================================================================================
'Properties
'   *   Name                *   Name
'
'Events
'
'Methods
'
'===========================================================================================
'Class property variables
Private m_dtFirstOfYear     As Date
Private m_dtEndOfYear       As Date
Private m_dtCalendarStart   As Date
Private m_dtPeriodStart     As Date
Private m_dtPeriodEnd       As Date
Private m_dtNextPeriodYear  As Date
Private m_dtDateValue       As Date
Private m_bIsinitialized    As Boolean
Private m_nPeriodValue      As Integer
Private m_nPeriodYear       As Integer
Private m_Periods           As CCalendarVBPeriods

'----------------------------------------------------------------------
' FirstOfYear Get/Let
'----------------------------------------------------------------------
' Purpose:  The periods beginning date value
'----------------------------------------------------------------------
Friend Property Get FirstOfYear() As Date
    FirstOfYear = m_dtFirstOfYear
End Property 'Get FirstOfYear

Friend Property Let FirstOfYear(ByVal dtFirstOfYear As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtFirstOfYear) Then
        dtFirstOfYear = Date$
    End If
    m_dtFirstOfYear = dtFirstOfYear
End Property 'Let FirstOfYear

'----------------------------------------------------------------------
' EndOfYear Get/Let
'----------------------------------------------------------------------
' Purpose:  The periods ending date value
'----------------------------------------------------------------------
Friend Property Get EndOfYear() As Date
    EndOfYear = m_dtEndOfYear
End Property 'Get EndOfYear

Friend Property Let EndOfYear(ByVal dtEndOfYear As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtEndOfYear) Then
        dtEndOfYear = Date$
    End If
    m_dtEndOfYear = dtEndOfYear
End Property 'Let EndOfYear

'----------------------------------------------------------------------
' CalendarStart Get/Let
'----------------------------------------------------------------------
' Purpose:  The calendars display starting date value
'----------------------------------------------------------------------
Friend Property Get CalendarStart() As Date
    CalendarStart = m_dtCalendarStart
End Property 'Get CalendarStart

Friend Property Let CalendarStart(ByVal dtCalendarStart As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtCalendarStart) Then
        dtCalendarStart = Date$
    End If
    m_dtCalendarStart = dtCalendarStart
End Property 'Let CalendarStart

'----------------------------------------------------------------------
' Periods Get/Let
'----------------------------------------------------------------------
' Purpose:  Periods definition object
'----------------------------------------------------------------------
Friend Property Get Periods() As CCalendarVBPeriods
    Set Periods = m_Periods
End Property 'Get Periods

Friend Property Set Periods(ByVal oPeriods As CCalendarVBPeriods)
    Set m_Periods = oPeriods
End Property 'Let Periods

'----------------------------------------------------------------------
' PeriodStart Get/Let
'----------------------------------------------------------------------
' Purpose:  The start date of the current period
'----------------------------------------------------------------------
Friend Property Get PeriodStart() As Date
    PeriodStart = m_dtPeriodStart
End Property 'Get PeriodStart

Friend Property Let PeriodStart(ByVal dtPeriodStart As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtPeriodStart) Then
        dtPeriodStart = Date$
    End If
    m_dtPeriodStart = dtPeriodStart
End Property 'Let PeriodStart

'----------------------------------------------------------------------
' PeriodEnd Get/Let
'----------------------------------------------------------------------
' Purpose:  The current periods ending date
'----------------------------------------------------------------------
Friend Property Get PeriodEnd() As Date
    PeriodEnd = m_dtPeriodEnd
End Property 'Get PeriodEnd

Friend Property Let PeriodEnd(ByVal dtPeriodEnd As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtPeriodEnd) Then
        dtPeriodEnd = Date$
    End If
    m_dtPeriodEnd = dtPeriodEnd
End Property 'Let PeriodEnd

'----------------------------------------------------------------------
' NextPeriodYear Get/Let
'----------------------------------------------------------------------
' Purpose:  The starting date of the next period year
'----------------------------------------------------------------------
Friend Property Get NextPeriodYear() As Date
    NextPeriodYear = m_dtNextPeriodYear
End Property 'Get NextPeriodYear

Friend Property Let NextPeriodYear(ByVal dtNextPeriodYear As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtNextPeriodYear) Then
        dtNextPeriodYear = Date$
    End If
    m_dtNextPeriodYear = dtNextPeriodYear
End Property 'Let NextPeriodYear

'----------------------------------------------------------------------
' DateValue Get/Let
'----------------------------------------------------------------------
' Purpose:  The value of the currently selected date
'----------------------------------------------------------------------
Friend Property Get DateValue() As Date
    DateValue = m_dtDateValue
End Property 'Get DateValue

Friend Property Let DateValue(ByVal dtDateValue As Date)
    'Make sure that we have a valid date
    If Not IsDate(dtDateValue) Then
        dtDateValue = Date$
    End If
    m_dtDateValue = dtDateValue
End Property 'Let DateValue

'----------------------------------------------------------------------
' PeriodValue Get/Let
'----------------------------------------------------------------------
' Purpose:  The value for the currenly selected period
'----------------------------------------------------------------------
Friend Property Get PeriodValue() As Integer
    PeriodValue = m_nPeriodValue
End Property 'Get PeriodValue

Friend Property Let PeriodValue(ByVal nPeriodValue As Integer)
    m_nPeriodValue = nPeriodValue
End Property 'Let PeriodValue

'----------------------------------------------------------------------
' PeriodYear Get/Let
'----------------------------------------------------------------------
' Purpose:  The value for the current period year
'----------------------------------------------------------------------
Friend Property Get PeriodYear() As Integer
    PeriodYear = m_nPeriodYear
End Property 'Get PeriodYear

Friend Property Let PeriodYear(ByVal nPeriodYear As Integer)
    m_nPeriodYear = nPeriodYear
End Property 'Let PeriodYear

'----------------------------------------------------------------------
' Isinitialized Get
'----------------------------------------------------------------------
' Purpose:  Determines if SetDateRanges has been run, which
'           initializes the object vars.
'----------------------------------------------------------------------
Public Property Get Isinitialized() As Boolean
    Isinitialized = m_bIsinitialized
End Property 'Get Isinitialized

'----------------------------------------------------------------------
' SetDateRanges()
'----------------------------------------------------------------------
' Purpose: Sets our class vars based on the PeriodValue and Year passed. Does
'   not set the DateValue since this var indicates the currently selected
'   date and not a range value.
' Inputs:
' Outputs:
'----------------------------------------------------------------------
Friend Sub SetDateRanges(ByVal nPeriodValue As Integer, ByVal nYear As Integer, ByVal oPeriods As CCalendarVBPeriods, nPeriodType As CalendarTypes, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement, ByVal nExtraWeekPlacement As ExtraWeekPlacements)

    Dim bIsExtraWeek    As Boolean
    Dim nWeek           As Integer
    Dim nWeeks          As Integer
    Dim lPreviousFirst  As Long
    Dim lPreviousLast   As Long
    Dim lPreviousYear   As Long
    Dim dtDate          As Date
    Dim dtWork          As Date
    Dim oMethods        As New CCalendarVBMethods
    
    'Can't have any zero's now can we
'    Debug.Assert nPeriodValue <> 0
    
    dtDate = oMethods.PeriodToDate(nPeriodValue, nYear, oPeriods, nPeriodType _
        , nFirstCurrentPeriod, nCalendarStartWeek, nExtraWeekPlacement)
    PeriodValue = nPeriodValue
    PeriodYear = nYear
    PeriodStart = dtDate
    FirstOfYear = oMethods.FirstOfYearDate(nYear, nFirstCurrentPeriod, nCalendarStartWeek)
    EndOfYear = oMethods.EndOfYearDate(FirstOfYear, nYear, nFirstCurrentPeriod, nCalendarStartWeek)
    NextPeriodYear = DateAdd("d", 1, EndOfYear)
    
    'Set Period end date
    Select Case nPeriodType
    Case calMonth
        PeriodEnd = DateAdd("d", -1, DateAdd("m", 1, PeriodStart))
        'Set calendar start date
        Call oMethods.GetPrevMonthDays(Month(PeriodStart), nFirstCurrentPeriod, nYear, lPreviousFirst, lPreviousLast, lPreviousYear)
        'Check to see if the period start falls on the first day
        If lPreviousLast = -1 Then
            'Falls on the FirstDay so lets add a week to the calendar start
            'date so that the user will be able to select a date from the
            'previous period.
            CalendarStart = DateAdd("d", -7, PeriodStart)
        Else
            CalendarStart = DateSerial(lPreviousYear, Month(DateAdd("m", -1, PeriodStart)), lPreviousFirst)
        End If
    Case calPeriod
        'Check to see if this is a leap week year
        bIsExtraWeek = oMethods.IsExtraWeek(FirstOfYear, PeriodYear, nFirstCurrentPeriod, nCalendarStartWeek)
        If bIsExtraWeek Then
            Select Case nExtraWeekPlacement
            Case ExtraWeekPlacements.calFirstPeriod
                'See if this is the first period
                If nPeriodValue = 1 Then
                    'Add an extra week to the first period
                    PeriodEnd = DateAdd("d", -1, DateAdd("ww", Periods(nPeriodValue).NumberOfWeeks, PeriodStart))
                    PeriodEnd = DateAdd("ww", 1, PeriodEnd)
                Else
                    'Move Period Start up one week since the first period had an extra week
                    'added to it. The length of the period does not change.
                    PeriodEnd = DateAdd("d", -1, DateAdd("ww", Periods(nPeriodValue).NumberOfWeeks, PeriodStart))
                End If
            Case ExtraWeekPlacements.calLastPeriod
                'See if this is the last period
                PeriodEnd = DateAdd("d", -1, DateAdd("ww", Periods(nPeriodValue).NumberOfWeeks, PeriodStart))
                If nPeriodValue = Periods.Count Then PeriodEnd = DateAdd("ww", 1, PeriodEnd)
            End Select
        Else
            PeriodEnd = DateAdd("d", -1, DateAdd("ww", Periods(nPeriodValue).NumberOfWeeks, PeriodStart))
        End If
        CalendarStart = DateAdd("ww", -1, PeriodStart)
    Case calWeek
        PeriodEnd = DateAdd("d", -1, DateAdd("ww", 1, PeriodStart))
        CalendarStart = DateAdd("ww", -2, PeriodStart)
    End Select
    
    m_bIsinitialized = True
    
End Sub 'SetDateRanges

