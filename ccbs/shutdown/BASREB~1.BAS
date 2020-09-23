Attribute VB_Name = "rEsTaRt"
''''''''''''''''''''''''''''''''''''''''''''
'Written by Marty Forde                    '
''''''''''''''''''''''''''''''''''''''''''''
'About Me:                                 '
'Occupation: Sophmore HighSchool Student   '
'    and VB consultant                     '
'Age: 16                                   '
'Expertise Areas: VB Api programming and   '
'    multimedia programming                '
'E-mail: Marty1149@aol.com                 '
''''''''''''''''''''''''''''''''''''''''''''

''''''''
'Notes:'
''''''''
'To shutdown, reboot, or logoff a machine in VB, you use the ExitWindowEx API function.
'In addition, if you are running NT or 2000, you need to call the AdjustTokenProvileges
'function first. This function encapsulates the necessary tasks (i.e., checking the OS and
'calling the AdjustTokenPrivileges function if necessary).
'
'One of the Enum Values in the declarations, passed to the function, determines
'whether you log off, restart, or shutdown, and whether the shutdown is graceful
'(running applications are given a chance to to exit and possibly cancel the shutdown) or
'forced.

Option Explicit

  Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, _
                    ByVal dwReserved As Long) As Long

  Private Const ENDSESSION_LOGOFF As Long = &H80000000

  Public Enum EShutDownTypes
         [_First] = 0
         EWX_LOGOFF = 0
         EWX_SHUTDOWN = 1&
         EWX_REBOOT = 2&
         EWX_FORCE = 4&
         EWX_POWEROFF = 8&
         EWX_FORCEIFHUNG = 10& ' NT5 only
         EWX_RESET = EWX_LOGOFF Or EWX_FORCE Or EWX_REBOOT
         [_Last] = &H20& - 1
  End Enum
  
  Public Enum EShutDownErrorBaseConstant
         eeSSDErrorBase = vbObjectError Or (1048 + &H210)
  End Enum


  Private Type OSVERSIONINFO
         dwOSVersionInfoSize As Long
         dwMajorVersion As Long
         dwMinorVersion As Long
         dwBuildNumber As Long
         dwPlatformId As Long
         szCSDVersion As String * 128
  End Type

  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
         (lpVersionInformation As OSVERSIONINFO) As Long

  Private Const VER_PLATFORM_WIN32_NT = 2
  Private Const VER_PLATFORM_WIN32_WINDOWS = 1
  Private Const VER_PLATFORM_WIN32s = 0


  Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
  Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
  Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
  Private Const FORMAT_MESSAGE_FROM_STRING = &H400
  Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
  Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
 
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
         (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
          ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
          Arguments As Long) As Long


  Private Type LARGE_INTEGER
         LowPart As Long
         HighPart As Long
  End Type
 
  Private Type LUID
         LowPart As Long
         HighPart As Long
  End Type

  Private Type LUID_AND_ATTRIBUTES
         pLuid As LUID
         Attributes As Long
  End Type

  Private Type TOKEN_PRIVILEGES
         PrivilegeCount As Long
         Privileges(0 To 0) As LUID_AND_ATTRIBUTES
  End Type

  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
  Private Declare Function OpenProcessToken Lib "advapi32.dll" _
         (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
          TokenHandle As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Private Declare Function GetTokenInformation Lib "advapi32.dll" _
         (ByVal TokenHandle As Long, TokenInformationClass As Integer, _
          TokenInformation As Any, ByVal TokenInformationLength As Long, _
          ReturnLength As Long) As Long
  Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
         (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
          NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
          PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
  Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
          Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
          ByVal lpName As String, lpLuid As LUID) As Long
 
  Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
  Private Const SE_PRIVILEGE_ENABLED = &H2

  Private Const READ_CONTROL = &H20000
  Private Const STANDARD_RIGHTS_ALL = &H1F0000
  Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
  Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
  Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
  Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

  Private Const TOKEN_ASSIGN_PRIMARY = &H1
  Private Const TOKEN_DUPLICATE = (&H2)
  Private Const TOKEN_IMPERSONATE = (&H4)
  Private Const TOKEN_QUERY = (&H8)
  Private Const TOKEN_QUERY_SOURCE = (&H10)
  Private Const TOKEN_ADJUST_PRIVILEGES = (&H20)
  Private Const TOKEN_ADJUST_GROUPS = (&H40)
  Private Const TOKEN_ADJUST_DEFAULT = (&H80)
  Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or _
                                    TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or _
                                    TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or _
                                    TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
  Private Const TOKEN_READ = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
  Private Const TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or _
                               TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
  Private Const TOKEN_EXECUTE = (STANDARD_RIGHTS_EXECUTE)
  Private Const TokenDefaultDacl = 6
  Private Const TokenGroups = 2
  Private Const TokenImpersonationLevel = 9
  Private Const TokenOwner = 4
  Private Const TokenPrimaryGroup = 5
  Private Const TokenPrivileges = 3
  Private Const TokenSource = 7
  Private Const TokenStatistics = 10
  Private Const TokenType = 8
  Private Const TokenUser = 1

Public Function WinError(ByVal lLastDLLError As Long) As String


  Dim sBuff As String
  Dim lCount As Long
               

  sBuff = Space(255)
  lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                         0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
  

  If lCount Then
      WinError = Left(sBuff, lCount)
  End If
               
End Function

Public Function IsNT() As Boolean


  Static bOnce As Boolean
  Static bValue As Boolean

  If Not (bOnce) Then
      Dim tVI As OSVERSIONINFO
 
      tVI.dwOSVersionInfoSize = Len(tVI)
 
      If (GetVersionEx(tVI) <> 0) Then
          bValue = (tVI.dwPlatformId = VER_PLATFORM_WIN32_NT)
          bOnce = True
      End If
  End If

  IsNT = bValue
               
End Function

Private Function NTEnableShutDown(ByRef sMsg As String) As Boolean

  Dim tLUID As LUID
  Dim hProcess As Long
  Dim hToken As Long
  Dim tTP As TOKEN_PRIVILEGES
  Dim tTPOld As TOKEN_PRIVILEGES
  Dim lTpOld As Long
  Dim lR As Long

  lR = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, tLUID)
               
  If (lR <> 0) Then
                           
      hProcess = GetCurrentProcess()
      If (hProcess <> 0) Then
          lR = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
          If (lR <> 0) Then
                               
              With tTP
                   .PrivilegeCount = 1
                   With .Privileges(0)
                        .Attributes = SE_PRIVILEGE_ENABLED
                        .pLuid.HighPart = tLUID.HighPart
                        .pLuid.LowPart = tLUID.LowPart
                   End With
              End With
                       
              lR = AdjustTokenPrivileges(hToken, 0, tTP, Len(tTP), tTPOld, lTpOld)
                       
              If (lR <> 0) Then
                  NTEnableShutDown = True
              Else
                  Err.Raise eeSSDErrorBase + 6, App.EXEName & ".mShutDown", _
                  "Can't enable shutdown: You do not have the privileges to " & _
                  "shutdown this system. [" & WinError(Err.LastDllError) & "]"
              End If
                       
              CloseHandle hToken
          Else
              Err.Raise eeSSDErrorBase + 6, App.EXEName & ".mShutDown", _
              "Can't enable shutdown: You do not have the privileges to " & _
              "shutdown this system. [" & WinError(Err.LastDllError) & "]"
          End If
      Else
          Err.Raise eeSSDErrorBase + 5, App.EXEName & ".mShutDown", _
                    "Can't enable shutdown: Can't determine the current process. [" & _
                    WinError(Err.LastDllError) & "]"
      End If
  Else
      Err.Raise eeSSDErrorBase + 4, App.EXEName & ".mShutDown", _
                "Can't enable shutdown: Can't find the SE_SHUTDOWN_NAME privilege value. [" & _
                WinError(Err.LastDllError) & "]"
  End If

End Function

Public Function ShutdownSystem(Optional ByVal eType As EShutDownTypes = EWX_RESET) As Boolean

  Dim lR As Long
  Dim sMsg As String


  If (eType < EShutDownTypes.[_First] And eType > EShutDownTypes.[_Last]) Then
      Err.Raise eeSSDErrorBase + 7, App.EXEName & ".mShutDown", _
                "Invalid parameter to ShutdownSystem: " & _
                eType, vbInformation
      Exit Function
  End If


  If (IsNT) Then
      If Not (NTEnableShutDown(sMsg)) Then
          Exit Function
      End If
  End If

  lR = ExitWindowsEx(eType, &HFFFFFFFF)

  If (lR = 0) Then
      Err.Raise eeSSDErrorBase + 3, App.EXEName & ".mShutDown", _
                "ShutdownSystem failed: " & _
                WinError(Err.LastDllError)
  Else
      ShutdownSystem = True
  End If

End Function



