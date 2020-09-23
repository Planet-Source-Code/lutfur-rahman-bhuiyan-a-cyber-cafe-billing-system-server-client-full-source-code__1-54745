VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPeriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================================
'  Name [CPeriods.cls]
'
' Copyright (c) 1997, CTR Business Systems, Inc.
'              All Rights Reserved
'
' Information Contained Herin is Proprietary and Confidential
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        9/12/97
'===========================================================================================

'$Runtime Dependencies:
'$DesignTime Dependencies:
'   CPeriod.cls

'===========================================================================================
'  Usage Notes:
'       Used for storing and retrieving period objects
'
'===========================================================================================
'Properties
'   *   Name                *   Name
'
'Events
'
'Methods
'   *   Add                 *   Count                   *   Item
'   *   Remove
'===========================================================================================
Option Explicit

Private m_naWeeksFromFirstOfYear() As Integer
Private m_Periods As Collection

Private Sub Class_Initialize()
    ' explicit creating is slightly faster
    ' than auto-instancing
    Set m_Periods = New Collection
End Sub

Public Function Add(ByVal sName As String, Optional ByVal nNumberOfWeeks As Integer = 4) As CPeriod
    On Error Resume Next
    Dim newItem As New CPeriod
    Dim nCount As Integer
    Dim iIndex As Integer
    
    ' add to the private collection
    newItem.Name = sName
    newItem.NumberOfWeeks = nNumberOfWeeks
    m_Periods.Add newItem, sName
    
    'Update our WeeksFromFirstOfYear array
    nCount = m_Periods.Count
    ReDim m_naWeeksFromFirstOfYear(1 To nCount)
    For iIndex = 1 To nCount
        With m_Periods(iIndex)
            If iIndex = 1 Then
                m_naWeeksFromFirstOfYear(iIndex) = 0
            Else
                m_naWeeksFromFirstOfYear(iIndex) = m_naWeeksFromFirstOfYear(iIndex - 1) + m_Periods(iIndex - 1).NumberOfWeeks
            End If
        End With
    Next
    
    ' return the new item to the program
    Set Add = newItem
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'This allows us to use for For Each
    'object enum
    Set NewEnum = m_Periods.[_NewEnum]
End Function

Sub Remove(index As Variant)
    On Error Resume Next
    m_Periods.Remove index
End Sub

Function Item(index As Variant) As CPeriod
Attribute Item.VB_UserMemId = 0
    'This is set as the default
    On Error Resume Next
    Set Item = m_Periods.Item(index)
End Function

Property Get Count() As Long
    Count = m_Periods.Count
End Property

Public Function WeeksFromFirstOfYear(ByVal iIndex As Integer)
    WeeksFromFirstOfYear = m_naWeeksFromFirstOfYear(iIndex)
End Function
