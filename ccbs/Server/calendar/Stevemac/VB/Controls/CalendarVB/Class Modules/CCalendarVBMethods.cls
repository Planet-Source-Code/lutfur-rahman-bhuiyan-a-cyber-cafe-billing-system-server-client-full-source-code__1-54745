VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalendarVBMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================================
'  Name [CCalendarVBMethods.cls]
'
' Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        10-29-1997
'===========================================================================================
'
'$Runtime Dependencies:
'$DesignTime Dependencies:
'
'===========================================================================================
'  Usage Notes:
'       Contains Methods and Procedures used by CalendarVB control
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

'----------------------------------------------------------------------
' CopyFont
'----------------------------------------------------------------------
' Purpose:  Copies the contents of one StdFont object to another
' Inputs:   source and destination StdFont object
' Outputs:  Destination font
'----------------------------------------------------------------------
Friend Sub CopyFont(fntSource As StdFont, fntDest As StdFont)
    
    'if the destination is nothing, create a new font object
    If fntDest Is Nothing Then Set fntDest = New StdFont
    
    fntDest.Bold = fntSource.Bold
    fntDest.Charset = fntSource.Charset
    fntDest.Italic = fntSource.Italic
    fntDest.Name = fntSource.Name
    fntDest.Size = fntSource.Size
    fntDest.Strikethrough = fntSource.Strikethrough
    fntDest.Underline = fntSource.Underline
    fntDest.Weight = fntSource.Weight
    
End Sub 'CopyFont()


'----------------------------------------------------------------------
' SetComboText
'----------------------------------------------------------------------
' Purpose:  Returns the ListIndex
' Inputs:   ComboBox, Item string to set in ComboBox
' Outputs:  ListIndex
'----------------------------------------------------------------------
Friend Function SetComboText(MyCombo As ComboBox, MyItem As String) As Integer
    Dim i As Integer
    
    For i = 0 To MyCombo.ListCount - 1
        If MyCombo.List(i) = MyItem Then
            SetComboText = i
            Exit Function
        End If
    Next i
    
     ' If the program reaches this point, the string is not in the
     ' list.
    SetComboText = -1
    
End Function 'SetComboText()

'----------------------------------------------------------------------
' PeriodToDate()
'----------------------------------------------------------------------
' Purpose:  Converts from a period value to a date value
'
' Inputs:
' Outputs:  Date value
'----------------------------------------------------------------------
Friend Function PeriodToDate(ByVal nPeriodValue As Integer, ByVal nYear As Integer, ByVal oPeriods As CCalendarVBPeriods, nPeriodType As CalendarTypes, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement, ByVal nExtraWeekPlacement As ExtraWeekPlacements) As Variant
    
    Dim nWeekDay        As Integer
    Dim nNumOfWeeks     As Integer
    Dim dtWork          As Date
    
    Select Case nPeriodType
    Case calMonth
        dtWork = DateSerial(nYear, nPeriodValue, 1)
        PeriodToDate = dtWork
    Case calPeriod
        dtWork = FirstOfYearDate(nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        'Check to see if this is a leap week year
        If IsExtraWeek(dtWork, nYear, nFirstCurrentPeriod, nCalendarStartWeek) Then
            Select Case nExtraWeekPlacement
            Case ExtraWeekPlacements.calFirstPeriod
                'See if this is the first period
                If nPeriodValue > 1 Then
                    nNumOfWeeks = oPeriods.WeeksFromFirstOfYear(nPeriodValue) + 1
                Else
                    nNumOfWeeks = oPeriods.WeeksFromFirstOfYear(nPeriodValue)
                End If
            Case ExtraWeekPlacements.calLastPeriod
                'See if this is the last period
                nNumOfWeeks = oPeriods.WeeksFromFirstOfYear(nPeriodValue)
            End Select
        Else
            nNumOfWeeks = oPeriods.WeeksFromFirstOfYear(nPeriodValue)
        End If
        PeriodToDate = DateAdd("ww", nNumOfWeeks, dtWork)
    Case calWeek
        dtWork = FirstOfYearDate(nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        PeriodToDate = DateAdd("ww", nPeriodValue - 1, dtWork)
    End Select
    
End Function 'PeriodToDate()

'----------------------------------------------------------------------
' IsExtraWeek()
'----------------------------------------------------------------------
' Purpose: Determines if there is an extra week in the Period Year
' Inputs:
' Outputs: True/False
'----------------------------------------------------------------------
Friend Function IsExtraWeek(ByVal dtFirstOfYearDate As Date, ByVal nYear As Integer, ByVal nFirstCurrentPeriod As Integer, ByVal nCalendarStartWeek As Integer) As Boolean

    Dim nWeeks          As Integer
    Dim sWork           As String
    Dim dtEnd           As Date
    Dim dtStart         As Date
    Dim dtStartNext     As Date
    
    dtEnd = DateAdd("d", -1, DateAdd("ww", 52, dtFirstOfYearDate))
    dtStartNext = FirstOfYearDate(nYear + 1, nFirstCurrentPeriod, nCalendarStartWeek)
    If DateDiff("d", dtEnd, dtStartNext) >= 7 Then
        IsExtraWeek = True
    Else
        IsExtraWeek = False
    End If
    
End Function

'----------------------------------------------------------------------
' EndOfYearDate()
'----------------------------------------------------------------------
' Purpose: Returns the Year End Period Date
' Inputs:
' Outputs: Date
'----------------------------------------------------------------------
Friend Function EndOfYearDate(ByVal dtFirstOfYearDate As Date, ByVal nYear As Integer, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement) As Date

    Dim dtNextPeriodYear    As Date
    
    dtNextPeriodYear = FirstOfYearDate(nYear + 1, nFirstCurrentPeriod, nCalendarStartWeek)
    EndOfYearDate = DateAdd("d", -1, dtNextPeriodYear)

End Function 'Get EndOfYearDate

'----------------------------------------------------------------------
' FirstOfYearDate()
'----------------------------------------------------------------------
' Purpose:  Determines the date for the first day of the period year
'           Uses the Calendar Start Week property to determine if the
'           date falls on the last week of the pervious year or the
'           first week of the current year
' Inputs:
' Outputs:  Date for the first day of the period year
'----------------------------------------------------------------------
Friend Function FirstOfYearDate(ByVal nYear As Integer, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement) As Date
    
    Dim dtDate As Date
    
    dtDate = DateSerial(nYear, 1, 1)
    
    If WeekDay(dtDate) = nFirstCurrentPeriod Then
        FirstOfYearDate = dtDate
    Else
        Select Case nCalendarStartWeek
        Case CalYearStartPlacement.calCurrentYear
            If WeekDay(dtDate) > nFirstCurrentPeriod Then
                FirstOfYearDate = DateAdd("d", ((7 - WeekDay(dtDate)) + nFirstCurrentPeriod), dtDate)
            Else
                FirstOfYearDate = DateAdd("d", nFirstCurrentPeriod - WeekDay(dtDate), dtDate)
            End If
        Case CalYearStartPlacement.calPreviousYear
            If WeekDay(dtDate) > nFirstCurrentPeriod Then
                FirstOfYearDate = DateAdd("w", -(WeekDay(dtDate) - nFirstCurrentPeriod), dtDate)
            Else
                FirstOfYearDate = DateAdd("d", -((7 + (WeekDay(dtDate)) - nFirstCurrentPeriod)), dtDate)
            End If
        End Select
    End If
    
End Function 'FirstOfYearDate()

'----------------------------------------------------------------------
' GetPrevMonthDays()
'----------------------------------------------------------------------
' Purpose:  Calculates the first and last day of the previous month
'           that should be displayed before the first day of the
'           of the given month and year
' Inputs:
' Outputs:  first and last day of prev month to display
'----------------------------------------------------------------------
Friend Sub GetPrevMonthDays(ByVal nCurMonth As Long, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCurYear As Integer, nFirst As Long, nLast As Long, lYear As Long)
    Dim lColDayOne  As Long         'column of 1st day of cur month
    Dim dtTemp      As Date         'temp date
    
    'construct a date to do date math
    dtTemp = DateSerial(nCurYear, nCurMonth, 1)
    
    'determine the column of the first day of the current month
    lColDayOne = WeekDay(dtTemp, nFirstCurrentPeriod)
    
    'if the first day of the current month is in column 1, we
    'don't need to paint any days from the prev month, so return
    'zeros and -1 for the first and last value
    If lColDayOne = 1 Then
        nFirst = 0
        nLast = -1
    Else
        'if there are days to paint, calculate the last and
        'first day using date math
        dtTemp = DateAdd("d", -1, dtTemp)
        nLast = VBA.Day(dtTemp)
        
        dtTemp = DateAdd("d", -(lColDayOne - 2), dtTemp)
        nFirst = VBA.Day(dtTemp)
    
        dtTemp = DateAdd("d", -(lColDayOne - 2), dtTemp)
        lYear = VBA.Year(dtTemp)
    End If 'no days to paint
    
End Sub 'GetPrevMonthDays()

'----------------------------------------------------------------------
' DateToPeriod()
'----------------------------------------------------------------------
' Purpose:  Converts from a date value to a period value
'
' Inputs:
' Outputs:  Period value, nPeriodYear will be changed if the date
'           does not exist in the current periodyear.
'----------------------------------------------------------------------
Public Function DateToPeriod(ByVal dtCurrentDate As Date, nPeriodYear As Integer, ByVal oPeriods As CCalendarVBPeriods, nPeriodType As Integer, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement, ByVal nExtraWeekPlacement As ExtraWeekPlacements, Optional vDate As Variant = Null) As Variant
    
    Dim iIndex          As Integer
    Dim nCount          As Integer
    Dim nWeek           As Integer
    Dim nWeeks          As Integer
    Dim nWeekStart      As Integer
    Dim nWeekEnd        As Integer
    Dim nYear           As Integer
    Dim dtEOYDate       As Date
    Dim dtFOYDate       As Date
    Dim dtDate          As Date
    
    'Init our return value so that we never have and
    'invalid return result
    DateToPeriod = 1
    
    If IsNull(vDate) Then
        dtDate = dtCurrentDate
    Else
        dtDate = vDate
    End If
    
    Select Case nPeriodType
    Case calMonth
        If Year(dtDate) <> nPeriodYear Then nPeriodYear = Year(dtDate)
        DateToPeriod = Month(dtDate)
    Case calPeriod
        'Determine if this is the same PeriodYear, if not then change our passed
        'PeriodYear value to the new one.
        dtFOYDate = FirstOfYearDate(nPeriodYear, nFirstCurrentPeriod, nCalendarStartWeek)
        dtEOYDate = EndOfYearDate(dtFOYDate, nPeriodYear, nFirstCurrentPeriod, nCalendarStartWeek)
        If Year(dtDate) < Year(dtFOYDate) Or Year(dtDate) > Year(dtEOYDate) Then _
            nPeriodYear = Year(dtDate)
        'Get the week location
        nWeek = DateDiff("ww", dtFOYDate, dtDate, nFirstCurrentPeriod) + 1
        'Loop through our period definition object and determine which period the
        'week falls within, this will be our period value.
        nCount = oPeriods.Count
        For iIndex = 1 To nCount
            nWeekStart = oPeriods.WeeksFromFirstOfYear(iIndex)
            nWeekEnd = oPeriods.WeeksFromFirstOfYear(iIndex) + oPeriods(iIndex).NumberOfWeeks
            'Account for our extra week
            If IsExtraWeek(dtFOYDate, nPeriodYear, nFirstCurrentPeriod, nCalendarStartWeek) Then
                Select Case nExtraWeekPlacement
                Case ExtraWeekPlacements.calFirstPeriod
                    If iIndex = 1 Then nWeekEnd = nWeekEnd + 1
                Case ExtraWeekPlacements.calLastPeriod
                    If iIndex = nCount Then nWeekEnd = nWeekEnd + 1
                End Select
            End If
            If nWeek >= nWeekStart And nWeek <= nWeekEnd Then
                DateToPeriod = iIndex
                Exit For
            End If
        Next
    Case calWeek
        dtFOYDate = FirstOfYearDate(nPeriodYear, nFirstCurrentPeriod, nCalendarStartWeek)
        dtEOYDate = EndOfYearDate(dtFOYDate, nPeriodYear, nFirstCurrentPeriod, nCalendarStartWeek)
        If Year(dtDate) < Year(dtFOYDate) Or Year(dtDate) > Year(dtEOYDate) Then _
            nPeriodYear = Year(dtDate)
        DateToPeriod = DateDiff("ww", dtFOYDate, dtDate, nFirstCurrentPeriod) + 1
    End Select
    
End Function 'DateToPeriod()

'----------------------------------------------------------------------
' DateToPeriodYear()
'----------------------------------------------------------------------
' Purpose:  Converts from a date value to a period year value
'
' Inputs:
' Outputs:  Year value
'----------------------------------------------------------------------
Public Function DateToPeriodYear(ByVal dtCurrentDate As Date, nPeriodType As Integer, ByVal nFirstCurrentPeriod As DaysOfTheWeek, ByVal nCalendarStartWeek As CalYearStartPlacement) As Long
    
    Dim iIndex          As Integer
    Dim nCount          As Integer
    Dim nWeek           As Integer
    Dim nWeeks          As Integer
    Dim nWeekStart      As Integer
    Dim nWeekEnd        As Integer
    Dim nYear           As Integer
    Dim dtEOYDate       As Date
    Dim dtFOYDate       As Date
    Dim dtDate          As Date
    
    'Init our return value so that we never have and
    'invalid return result
    nYear = Year(dtCurrentDate)
    
    Select Case nPeriodType
    Case calMonth
        DateToPeriodYear = nYear
    Case calPeriod
        'Determine if this is the same PeriodYear, if not then change our passed
        'PeriodYear value to the new one.
        dtFOYDate = FirstOfYearDate(nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        dtEOYDate = EndOfYearDate(dtFOYDate, nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        If dtCurrentDate < dtFOYDate Then
            DateToPeriodYear = nYear - 1
        ElseIf dtCurrentDate > dtEOYDate Then
            DateToPeriodYear = nYear + 1
        Else
            DateToPeriodYear = nYear
        End If
    Case calWeek
        dtFOYDate = FirstOfYearDate(nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        dtEOYDate = EndOfYearDate(dtFOYDate, nYear, nFirstCurrentPeriod, nCalendarStartWeek)
        If dtCurrentDate < dtFOYDate Then
            DateToPeriodYear = nYear - 1
        ElseIf dtCurrentDate > dtEOYDate Then
            DateToPeriodYear = nYear + 1
        Else
            DateToPeriodYear = nYear
        End If
    End Select
    
End Function 'DateToPeriodYear()



