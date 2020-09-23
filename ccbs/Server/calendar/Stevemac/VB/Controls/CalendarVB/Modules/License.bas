Attribute VB_Name = "License"
Option Explicit

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1

Private Const LIC_CLSID = _
 "38B4A9C4-97D0-11D0-87A0-444553540000"

Private Const LIC_KEY = _
 "ebjbdifbginbbicbjcobebbiococbihbnioc"

Private Const LIC_DEMO = "Unregistered"

Public Const GWL_HINSTANCE = (-6)

Private Declare Function RegQueryValue _
    Lib "advapi32.dll" Alias "RegQueryValueA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal lpValue As String, _
    lpcbValue As Long _
    ) As Long
Private Declare Function RegCreateKey _
    Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long _
    ) As Long
Private Declare Function RegSetValue _
    Lib "advapi32.dll" Alias "RegSetValueA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal dwType As Long, _
    ByVal lpData As String, _
    ByVal cbData As Long _
    ) As Long
Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long _
    ) As Long
Declare Function GetModuleFileName _
    Lib "kernel32" Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long _
    ) As Long

Public Function GetRegLicKey() As String
    Dim lRetVal As Long
    Dim lpValue As String * 2048
    Dim lpcbValue As Long
    Dim sLicKey As String

    lpcbValue = 2048
    lRetVal = RegQueryValue(HKEY_CLASSES_ROOT, _
              "LICENSES\" & LIC_CLSID, _
               lpValue, _
              lpcbValue)

    If lRetVal = 0 Then
        sLicKey = Left(lpValue, lpcbValue - 1)
    End If

    GetRegLicKey = sLicKey
End Function

Sub SetRegLicKey(ByVal sLicInfo As String)
    Dim lRetVal As Long
    Dim hKey As Long

    lRetVal = RegCreateKey(HKEY_CLASSES_ROOT, _
              "LICENSES\" & LIC_CLSID, _
               hKey)

    If lRetVal = 0 Then
        lRetVal = RegSetValue(hKey, _
                              "", _
                               REG_SZ, _
                              sLicInfo, _
                              Len(sLicInfo))
    End If

End Sub

Public Function IsRegistered() As Boolean
    Dim sLicKey As String

    sLicKey = GetRegLicKey()

    If sLicKey = LIC_KEY Then
        IsRegistered = True
    End If

End Function

Public Function IsDemo() As Boolean
    Dim sLicKey As String
    sLicKey = GetRegLicKey()

    If sLicKey = LIC_DEMO Then
        IsDemo = True
    End If

End Function

Sub Register()
    Call SetRegLicKey(LIC_KEY)
End Sub

Sub RegisterDemo()
    Call SetRegLicKey(LIC_DEMO)
End Sub

Public Sub CheckRegistration()

    If IsRegistered() = False Then

        If IsDemo() = False Then
            Call ShowAboutBox
            Call RegisterDemo
        End If

    End If

End Sub

Function GetProgramName(ByVal hwndParent As Long)
    Dim hInstance As Long
    Dim sFullName As String * 1024
    Dim EXEName As String
    Dim iMarker As Integer
    Dim lRetVal As Long

    hInstance = GetWindowLong(hwndParent, _
                              GWL_HINSTANCE)
    lRetVal = GetModuleFileName(hInstance, _
                                sFullName, 1024)
    EXEName = Left(sFullName, lRetVal)
    iMarker = ReverseFind(EXEName, "\")

    If iMarker > 0 Then
        EXEName = Mid(EXEName, iMarker + 1)
    End If

    GetProgramName = EXEName
End Function

Function ReverseFind( _
ByVal sVal As String, _
sChar As String _
)
    Dim iMarker As Integer
    Dim bDone As Boolean

    iMarker = Len(sVal)

    While bDone = False And iMarker > 0

        If Mid(sVal, iMarker, 1) = sChar Then
            bDone = True
        Else
            iMarker = iMarker - 1
        End If

    Wend

    ReverseFind = iMarker
End Function

Public Function IsRegistrationCode(ByVal sCode As String) As Boolean
    IsRegistrationCode = (sCode = LIC_KEY)
End Function

Public Sub ShowAboutBox(Optional imgIcon As Picture = Nothing)
    Dim nResult     As Integer
    Dim sDesc       As String
    Dim FAbout      As New FctlAbout
    
    nResult = IsRegistered()
    If nResult = False Then
        sDesc = "To purchase a copy of this product call "
        sDesc = sDesc & "CTR Business Systems, Inc. at (503)293-2414. "
        sDesc = sDesc & "CTR accepts Visa and Mastercard. "
        sDesc = sDesc & "The cost of this product is $99 + S&H."
    Else
        sDesc = App.FileDescription
    End If
    Load FAbout
    FAbout.ShowModal "CalendarVB", App.Major & "." & App.Minor, sDesc, App.CompanyName, App.LegalCopyright, App.LegalTrademarks, imgIcon, nResult
    Unload FAbout
    Set FAbout = Nothing

End Sub
