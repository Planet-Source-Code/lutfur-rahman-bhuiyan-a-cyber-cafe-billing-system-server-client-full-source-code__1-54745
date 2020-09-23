VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalendarVBPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================
'  Name [CCalnedarVBPeriod.cls]
'
' Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        9/12/97
'=======================================================

'$Runtime Dependencies:
'$DesignTime Dependencies:

'=======================================================
'  Usage Notes:
'       Used for defining a period
'
'=======================================================
'Properties
'   *   Name                *   NumberOfWeeks
'
'Events
'
'Methods
'
'=======================================================
Option Explicit

'Public Class Properties
Private m_nNumberOfWeeks    As Integer
Private m_sName             As String

'----------------------------------------------------------------------
' Name Get/Let
'----------------------------------------------------------------------
' Purpose:  The name of the period
'----------------------------------------------------------------------
Public Property Get Name() As String
    Name = m_sName
End Property 'Get Name

Public Property Let Name(ByVal sName As String)
    m_sName = sName
End Property 'Let Name

'----------------------------------------------------------------------
' Name Get/Let
'----------------------------------------------------------------------
' Purpose:  The number of weeks for this period
'----------------------------------------------------------------------
Public Property Get NumberOfWeeks() As Integer
    NumberOfWeeks = m_nNumberOfWeeks
End Property 'Get NumberOfWeeks

Public Property Let NumberOfWeeks(ByVal nNumberOfWeeks As Integer)
    m_nNumberOfWeeks = nNumberOfWeeks
End Property 'Let NumberOfWeeks
