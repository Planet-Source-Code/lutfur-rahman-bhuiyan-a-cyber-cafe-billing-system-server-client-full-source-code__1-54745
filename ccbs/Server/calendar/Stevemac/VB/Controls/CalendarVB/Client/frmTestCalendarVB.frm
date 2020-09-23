VERSION 5.00
Object = "*\A..\..\CALEND~1\CALEND~1.VBP"
Begin VB.Form frmTestCalendar 
   Caption         =   "Calendar Control"
   ClientHeight    =   3180
   ClientLeft      =   3975
   ClientTop       =   2820
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestCalendarVB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   3300
   Begin ctrCalendarVB.CalendarVB calTest 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4789
      LineStyle       =   2
      CurrentPeriodbackColor=   -2147483643
      CurrentPeriodForeColor=   -2147483630
      DayHeaderBackColor=   -2147483633
      DayHeaderForeColor=   -2147483630
      ActiveDayForeColor=   158700
      FlatLineColor   =   12632256
      PrePeriodBackColor=   -2147483648
      PrePeriodforeColor=   -2147483632
      PostPeriodBackColor=   -2147483648
      PostPeriodforeColor=   -2147483632
      ActiveDayFontBold=   -1  'True
      ActiveDayFontItalic=   0   'False
      ActiveDayFontSize=   8.25
      ActiveDayFontName=   "Tahoma"
      DayHeaderFontBold=   0   'False
      DayHeaderFontItalic=   0   'False
      DayHeaderFontSize=   8.25
      DayHeaderFontName=   "Tahoma"
      DaysFontBold    =   0   'False
      DaysFontItalic  =   0   'False
      DaysFontSize    =   8.25
      DaysFontName    =   "Tahoma"
   End
   Begin VB.Label lblSelectedDate 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
End
Attribute VB_Name = "frmTestCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_c() As cFlatControl

Private Sub calTest_DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
   lblSelectedDate.Caption = "Selected: " & Format$(NewDate, "Long Date")
End Sub

Private Sub Form_Load()
   ReDim m_c(0 To 1) As cFlatControl
   Set m_c(0) = New cFlatControl
   m_c(0).hWndAttach calTest.hWndPeriod, calTest.hwnd, True
   Set m_c(1) = New cFlatControl
   m_c(1).hWndAttach calTest.hWndYear, calTest.hwnd, True
   calTest_DateChange Now, calTest.DateValue
End Sub
