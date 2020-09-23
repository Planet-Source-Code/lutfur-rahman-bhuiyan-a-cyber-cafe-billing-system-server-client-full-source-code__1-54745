VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================
'  Name [CLicense.cls]
'
'  Copyright © 1997-1999 by CTR Business Systems, Inc.
'
'  Version:     1.00
'  Author:      Mike Gainer
'  Date:        10-28-1997
'===========================================================================================
'
'$Runtime Dependencies:
'$DesignTime Dependencies:
'
'===========================================================================================
'  Usage Notes:
'
'
'===========================================================================================
'Properties
'   *   licCLSID                *   licKEY
'
'Methods
'   *   CheckRegistration       *   GetProgramName              *   GetRegLicKey
'   *   GetWindowLong           *   IsDemo                      *   IsRegistered
'   *   IsRegistrationCode      *   Register                    *   RegisterDemo
'   *   ReverseFind             *   SetRegLicKey                *   ShowAboutBox
'===========================================================================================
Option Explicit

'CLASS CONSTANTS ==============================================================
Private Const GWL_HINSTANCE = (-6)
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1
Private Const ENCRYPT_KEY = "007" & vbCrLf & vbTab & "CTR Business Systems, Inc."
Private Const LIC_CLSID = "A8A58E34-4C7D-11D1-9037-00A0C91EF7D6"
Private Const LIC_KEY = "411-0827594"
Private Const LIC_DEMO = "TRYME"

'API DECLARES =================================================================
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'CLASS PROPERTIES =============================================================
Private m_slicCLSID     As String
Private m_slicKEY       As String

'----------------------------------------------------------------------
' GetRegLicKey Method
'----------------------------------------------------------------------
' Purpose: Returns the License Key from the registery
' Inputs:   None
' Outputs:  License Key
'----------------------------------------------------------------------
Public Function GetRegLicKey() As String
Attribute GetRegLicKey.VB_Description = "Returns the License Key from the registery."
    Dim lRetVal     As Long
    Dim lpValue     As String * 2048
    Dim lpcbValue   As Long
    Dim slicKEY     As String

    lpcbValue = 2048
    lRetVal = RegQueryValue(HKEY_CLASSES_ROOT, "LICENSES\" & licCLSID, lpValue, lpcbValue)

    If lRetVal = 0 Then
        slicKEY = Left(lpValue, lpcbValue - 1)
    End If

    GetRegLicKey = Encrypt(slicKEY, ENCRYPT_KEY)
End Function 'GetRegLicKey()

'----------------------------------------------------------------------
' SetRegLicKey Method
'----------------------------------------------------------------------
' Purpose:  Saves the License Key to the registery
' Inputs:   License Key
' Outputs:  None
'----------------------------------------------------------------------
Public Sub SetRegLicKey(ByVal sLicInfo As String)
Attribute SetRegLicKey.VB_Description = "Saves the License Key to the registery."
    Dim lRetVal     As Long
    Dim hKey        As Long
    Dim sWork       As String
    
    sWork = Encrypt(sLicInfo, ENCRYPT_KEY)
    lRetVal = RegCreateKey(HKEY_CLASSES_ROOT, "LICENSES\" & licCLSID, hKey)

    If lRetVal = 0 Then
        lRetVal = RegSetValue(hKey, "", REG_SZ, sWork, Len(sWork))
    End If

End Sub 'SetRegLicKey()

'----------------------------------------------------------------------
' IsRegistered Method
'----------------------------------------------------------------------
' Purpose:  Checks for a valid License Key in the Registery
' Inputs:   None
' Outputs:  True if valid License Key found, else false.
'----------------------------------------------------------------------
Public Function IsRegistered() As Boolean
Attribute IsRegistered.VB_Description = "Returns true if this is a registered copy, else false."
    Dim slicKEY As String

    slicKEY = GetRegLicKey()

    If slicKEY = licKEY Then
        IsRegistered = True
    End If

End Function 'IsRegistered()

'----------------------------------------------------------------------
' IsDemo Method
'----------------------------------------------------------------------
' Purpose:  Checks to see if this is a demo version
' Inputs:   None
' Outputs:  True if Demo version, else false
'----------------------------------------------------------------------
Public Function IsDemo() As Boolean
Attribute IsDemo.VB_Description = "Return true if this is a demo copy, else false."
    Dim slicKEY As String
    
    slicKEY = GetRegLicKey()

    If slicKEY = LIC_DEMO Then
        IsDemo = True
    End If

End Function 'IsDemo()

'----------------------------------------------------------------------
' Register Method
'----------------------------------------------------------------------
' Purpose:  Saves the License Key to the registery
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Public Sub Register()
Attribute Register.VB_Description = "Performs the registration process."
    Call SetRegLicKey(licKEY)
End Sub 'Register()

'----------------------------------------------------------------------
' RegisterDemo Method
'----------------------------------------------------------------------
' Purpose:  Register as a Demo version in the system registery
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Public Sub RegisterDemo()
Attribute RegisterDemo.VB_Description = "Performs the Demo registration process."
    Call SetRegLicKey(LIC_DEMO)
End Sub 'RegisterDemo()

'----------------------------------------------------------------------
' CheckRegistration Method
'----------------------------------------------------------------------
' Purpose:  Shows AboutBox with Unregistered message if not registered
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Public Sub CheckRegistration()
Attribute CheckRegistration.VB_Description = "Displays AboutBox with Unregistered Info if not regiserted."

    If IsRegistered() = False Then
        If IsDemo() = False Then
            Call ShowAboutBox
            Call RegisterDemo
        End If
    End If

End Sub 'CheckRegistration()

'----------------------------------------------------------------------
' GetProgramName Method
'----------------------------------------------------------------------
' Purpose:  Gets the program name based on the Parent Window Handle
' Inputs:   None
' Outputs:  Application name
'----------------------------------------------------------------------
Public Function GetProgramName(ByVal hwndParent As Long)
Attribute GetProgramName.VB_Description = "Returns the Application Name based on the passed Parent Window Handle."
    Dim hInstance   As Long
    Dim sFullName   As String * 1024
    Dim sExeName    As String
    Dim iMarker     As Integer
    Dim lRetVal     As Long

    hInstance = GetWindowLong(hwndParent, GWL_HINSTANCE)
    lRetVal = GetModuleFileName(hInstance, sFullName, 1024)
    sExeName = Left(sFullName, lRetVal)
    iMarker = ReverseFind(sExeName, "\")

    If iMarker > 0 Then
        sExeName = Mid$(sExeName, iMarker + 1)
    End If

    GetProgramName = sExeName
End Function 'GetProgramName()

'----------------------------------------------------------------------
' ReverseFind Method
'----------------------------------------------------------------------
' Purpose:  Return location of found character. Search from right to
'           left.
' Inputs:   None
' Outputs:  Index location of found character
'----------------------------------------------------------------------
Public Function ReverseFind(ByVal sVal As String, sChar As String) As Integer
Attribute ReverseFind.VB_MemberFlags = "40"
    Dim iMarker     As Integer
    Dim bDone       As Boolean

    iMarker = Len(sVal)

    While bDone = False And iMarker > 0
        If Mid$(sVal, iMarker, 1) = sChar Then
            bDone = True
        Else
            iMarker = iMarker - 1
        End If
    Wend

    ReverseFind = iMarker
End Function 'ReverseFind()

'----------------------------------------------------------------------
' IsRegistrationCode Method
'----------------------------------------------------------------------
' Purpose:  Determines if passed License Key is valid
' Inputs:   License Key value
' Outputs:  True if valid key, else false
'----------------------------------------------------------------------
Public Function IsRegistrationCode(ByVal sCode As String) As Boolean
Attribute IsRegistrationCode.VB_Description = "Returns true if the passed License Key is valid."
    IsRegistrationCode = (sCode = licKEY)
End Function 'IsRegistrationCode()

'----------------------------------------------------------------------
' ShowAboutBox Method
'----------------------------------------------------------------------
' Purpose:  Shows the AboutBox displaying Unregistered info if necessary
' Inputs:   Option icon to be displayed in AboutBox
' Outputs:  None
'----------------------------------------------------------------------
Public Sub ShowAboutBox(Optional imgIcon As Picture = Nothing)
Attribute ShowAboutBox.VB_Description = "Shows the AboutBox displaying the Unregistered Info if necessary."
    Dim nResult     As Integer
    Dim sDesc       As String
    Dim FAbout      As New FctlAbout
    
    nResult = IsRegistered()
'    If nResult = False Then
'        sDesc = "To purchase a copy of this product call "
'        sDesc = sDesc & "Acme at (555)555-5555. "
'        sDesc = sDesc & "Acme accepts Visa and Mastercard. "
'        sDesc = sDesc & "The cost of this product is $BigBucks + S&H."
'    Else
        sDesc = App.FileDescription
'    End If
    Load FAbout
    If Not imgIcon Is Nothing Then
        FAbout.ShowModal "CalendarVB", App.Major & "." & App.Minor & "." & App.Revision, sDesc, App.CompanyName, App.LegalCopyright, App.LegalTrademarks, imgIcon, nResult
    Else
        FAbout.ShowModal "CalendarVB", App.Major & "." & App.Minor & "." & App.Revision, sDesc, App.CompanyName, App.LegalCopyright, App.LegalTrademarks, bIsRegistered:=nResult
    End If
    Unload FAbout
    Set FAbout = Nothing

End Sub 'ShowAboutBox()

'----------------------------------------------------------------------
' licCLSID Get/Let
'----------------------------------------------------------------------
' Purpose:  CLSID used for registration identification
'----------------------------------------------------------------------
Public Property Get licCLSID() As String
Attribute licCLSID.VB_Description = "CLSID code value"
    If Len(m_slicCLSID) = 0 Then m_slicCLSID = LIC_CLSID
    licCLSID = m_slicCLSID
End Property 'Get licCLSID

Public Property Let licCLSID(ByVal slicCLSID As String)
    m_slicCLSID = slicCLSID
End Property 'Let licCLSID

'----------------------------------------------------------------------
' licKEY Get/Let
'----------------------------------------------------------------------
' Purpose:  Registration Key Value
'----------------------------------------------------------------------
Public Property Get licKEY() As String
Attribute licKEY.VB_Description = "License Key Value"
    If Len(m_slicKEY) = 0 Then m_slicKEY = LIC_KEY
    licKEY = m_slicKEY
End Property 'Get licKEY

Public Property Let licKEY(ByVal slicKEY As String)
    m_slicKEY = slicKEY
End Property 'Let licKEY

'----------------------------------------------------------------------
' Encrypt Method
'----------------------------------------------------------------------
' Purpose: XOR String Encryption. THE PASSWORDS ARE CASE SENSITIVE
' Inputs:
'   Text = The text to be encrypted/decypted (ex: "My Text")
'   Key = The password or key (ex: "MyPassword")
' Outputs:
'   Returns either encrypted or decrypted string
'----------------------------------------------------------------------
Private Function Encrypt(ByVal sText As String, ByVal sKey As String) As String

    
    'Declare some varibles
    Dim sEncryptedStr   As String        'This will hold the final encrypted string
    Dim nMask           As String        'This will hold pieces of the key
    Dim nCounter        As Integer       'This will help us rotate through the key
    Dim iIndex          As Integer
    
    'Fill this variable with spaces
    sEncryptedStr = Space$(Len(sText))
    
    'Loop through the text for encrypting
    For iIndex = 1 To Len(sText)
        'Loop through the characters in the key
        nCounter = nCounter + 1
        If nCounter > Len(sKey) Then nCounter = 1
        nMask = Asc(Mid$(sKey, nCounter, 1))
        'XOR the text with the appropriate part of the key
        Mid$(sEncryptedStr, iIndex, 1) = Chr$(Asc(Mid$(sText, iIndex, 1)) Xor nMask)
    Next iIndex
    
    'Send the result back
    Encrypt = sEncryptedStr
    
End Function 'Encrypt()

