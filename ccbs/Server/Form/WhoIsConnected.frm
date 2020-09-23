VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connected IP"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "WhoIsConnected.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   660
      Top             =   5115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   360
      Left            =   4860
      TabIndex        =   2
      Top             =   5145
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4875
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   8599
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get"
      Default         =   -1  'True
      Height          =   390
      Left            =   5985
      TabIndex        =   0
      Top             =   5130
      Width           =   1020
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Dim SCounter As Integer
Dim LoginId As String
Dim LoginDomain As String
Dim LoginServer As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const ERROR_SUCCESS            As Long = 0
Private Const MIB_TCP_STATE_CLOSED     As Long = 1
Private Const MIB_TCP_STATE_LISTEN     As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT   As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD   As Long = 4
Private Const MIB_TCP_STATE_ESTAB      As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1  As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2  As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
Private Const MIB_TCP_STATE_CLOSING    As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK   As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT  As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB As Long = 12

Private Declare Function GetTcpTable Lib "iphlpapi.dll" _
  (ByRef pTcpTable As Any, _
   ByRef pdwSize As Long, _
   ByVal bOrder As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   Src As Any, _
   ByVal bcount As Long)
  
Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
  
Private Declare Function inet_ntoa Lib "wsock32" _
  (ByVal addr As Long) As Long

Private Declare Function ntohs Lib "wsock32" _
   (ByVal addr As Long) As Long
   
   Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2

Private Type WSAData
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  iMaxSockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSAData) As Long
  
Private Declare Function WSACleanUp Lib "wsock32" Alias "WSACleanup" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

'Private Declare Sub CopyMemory Lib "kernel32" _
'   Alias "RtlMoveMemory" _
'  (xDest As Any, _
'   xSource As Any, _
'   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
  
  Private Const NERR_SUCCESS As Long = 0&
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const ERROR_MORE_DATA As Long = 234&
Private Const LB_SETTABSTOPS As Long = &H192

'for use on Win NT/2000 only
Private Type WKSTA_USER_INFO_0
  wkui0_username  As Long
End Type

Private Type WKSTA_USER_INFO_1
  wkui1_username As Long
  wkui1_logon_domain As Long
  wkui1_oth_domains As Long
  wkui1_logon_server As Long
End Type

Private Declare Function NetWkstaUserEnum Lib "netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   bufptr As Long, _
   ByVal prefmaxlen As Long, _
   entriesread As Long, _
   totalentries As Long, _
   resume_handle As Long) As Long
        
Private Declare Function NetApiBufferFree Lib "netapi32" _
   (ByVal Buffer As Long) As Long

Private Declare Sub CpMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)
   
Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Private Declare Function SendMessage Lib "User32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

  
  
Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSAData
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Public Sub SocketsCleanup()
   
   If WSACleanUp() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub


Public Function GetHostNameFromIP(ByVal sAddress As String) As String
   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   If SocketsInitialize() Then
     'convert string address to long
      hAddress = inet_addr(sAddress)
      If hAddress <> SOCKET_ERROR Then
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
         If ptrHosent <> 0 Then
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         End If 'If ptrHosent
      SocketsCleanup
      End If 'If hAddress
   End If  'If SocketsInitialize
End Function

   

Public Function GetInetStrFromPtr(Address As Long) As String
  
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function


Private Sub Command2_Click()
End
End Sub



Private Sub Form_Load()
ReDim TabArray(0 To 3) As Long
   nid.cbSize = Len(nid)
   nid.hWnd = form1.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = form1.Icon
   nid.szTip = "Who Is Connected ?..." & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid

    Me.Hide
   TabArray(0) = 58
   TabArray(1) = 140
   TabArray(2) = 171
   TabArray(3) = 217
   With ListView1
      .View = lvwReport
      .ColumnHeaders.Add , , "Local IP Address"
      .ColumnHeaders.Add , , "Local Port"
      .ColumnHeaders.Add , , "Remote IP Address"
      .ColumnHeaders.Add , , "Remote Host Name"
      .ColumnHeaders.Add , , "Login At Remote Host"
      .ColumnHeaders.Add , , "Login Domain"
      .ColumnHeaders.Add , , "Login Server"
      .ColumnHeaders.Add , , "Remote Port"
      .ColumnHeaders.Add , , "State (dec)"
      .ColumnHeaders.Add , , "State Description"
      .ColumnHeaders.Add , , "OS Informations"
   End With
   
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  ListView1.SortKey = ColumnHeader.Index - 1
  ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
  ListView1.Sorted = True
  
End Sub


Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


Private Sub Command1_Click()

   Dim TcpRow As MIB_TCPROW
   Dim buff() As Byte
   Dim cbRequired As Long
   Dim nStructSize As Long
   Dim nRows As Long
   Dim cnt As Long
   Dim tmp As String
   Dim itmx As ListItem
       
   Call GetTcpTable(ByVal 0&, cbRequired, 1)

   If cbRequired > 0 Then
    
      ReDim buff(0 To cbRequired - 1) As Byte
      
      If GetTcpTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
      
        'saves using LenB in the CopyMemory calls below
         nStructSize = LenB(TcpRow)
   
        'first 4 bytes is a long indicating the
        'number of entries in the table
         CopyMemory nRows, buff(0), 4
        'itmx.ListSubItems.Clear
        ListView1.ListItems.Clear
         For cnt = 1 To nRows
         
           'moving past the four bytes obtained
           'above, get one chunk of data and cast
           'into an TcpRow type
            CopyMemory TcpRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
            
           'pass the results to the listview
            With TcpRow
               
               Set itmx = ListView1.ListItems.Add(, , GetInetStrFromPtr(.dwLocalAddr))
              ' itmx.SubItems(1) = ntohs(.dwLocalPort)
              
               itmx.SubItems(2) = GetInetStrFromPtr(.dwRemoteAddr)
               If itmx.SubItems(2) = "0.0.0.0" Then
                    itmx.SubItems(3) = ""
                ElseIf itmx.SubItems(2) = "192.168.100.6" Or itmx.SubItems(2) = "192.168.100.222" Then
                     
                Else
                    itmx.SubItems(3) = GetHostNameFromIP(itmx.SubItems(2))
                End If
               If itmx.SubItems(3) <> "" Then
                    GetUserInfo (itmx.SubItems(3))
                    'MsgBox GetUserInfo(itmx.SubItems(3))
                End If
                itmx.SubItems(4) = LoginId
                If itmx.SubItems(2) <> "0.0.0.0" Then
                    CheckForFlag itmx.SubItems(2), itmx.SubItems(3), itmx.SubItems(4)
                End If
                itmx.SubItems(5) = LoginDomain
                itmx.SubItems(6) = LoginServer
                LoginId = ""
                LoginDomain = ""
                LoginServer = ""
               itmx.SubItems(7) = ntohs(.dwRemotePort)
               
              ' itmx.SubItems(8) = (.dwState)
                
              'the MSDN has a description defined only
              'for the MIB_TCP_STATE_DELETE_TCB member.
               Select Case .dwState
                  Case MIB_TCP_STATE_CLOSED:       tmp = "closed"
                  Case MIB_TCP_STATE_LISTEN:       tmp = "listening"
                  Case MIB_TCP_STATE_SYN_SENT:     tmp = "sent"
                  Case MIB_TCP_STATE_SYN_RCVD:     tmp = "received"
                  Case MIB_TCP_STATE_ESTAB:        tmp = "established"
                  Case MIB_TCP_STATE_FIN_WAIT1:    tmp = "fin wait 1"
                  Case MIB_TCP_STATE_FIN_WAIT2:    tmp = "fin wait 1"
                  Case MIB_TCP_STATE_CLOSE_WAIT:   tmp = "close wait"
                  Case MIB_TCP_STATE_CLOSING:      tmp = "closing"
                  Case MIB_TCP_STATE_LAST_ACK:     tmp = "last ack"
                  Case MIB_TCP_STATE_TIME_WAIT:    tmp = "time wait"
                  Case MIB_TCP_STATE_DELETE_TCB:   tmp = "TCB deleted"
               End Select
               
               'itmx.SubItems(9) = tmp
              ' If itmx.SubItems(3) <> "" Then
              '      itmx.SubItems(10) = GetSystemInfo(itmx.SubItems(3))
              '  End If
               tmp = ""

            End With
         Next
         UpdateSavedData
      End If
   End If
            
End Sub
Sub UpdateSavedData()
Dim found As Boolean
    For i = 0 To 19
    found = False
    If ConnectInfo(i).sIP = "" Then
        Exit For
    End If
        For j = 0 To ListView1.ListItems.Count - 1
            If ConnectInfo(i).sIP = ListView1.ListItems(j + 1).ListSubItems(2).Text Then
                found = True
            End If
        Next j
        If ConnectInfo(i).sIP <> "" Then
            If Not found Then
                For k = i To 18
                    ConnectInfo(k).sIP = ConnectInfo(k + 1).sIP
                    ConnectInfo(k).sHost = ConnectInfo(k + 1).sHost
                    ConnectInfo(k).sUser = ConnectInfo(k + 1).sUser
                Next k
                SCounter = SCounter - 1
            End If
        End If
    Next i
End Sub
Sub CheckForFlag(cIp As String, cHost As String, cUser As String)
Dim isFound As Boolean
    isFound = False
    For i = 0 To 19
        If cIp = ConnectInfo(i).sIP Then
            isFound = True
            Exit For
        End If
    Next i
    If Not isFound Then
        ConnectInfo(SCounter).sIP = cIp
        ConnectInfo(SCounter).sHost = cHost
        ConnectInfo(SCounter).sUser = cUser
        SCounter = SCounter + 1
        mUser = cUser
        mHost = cHost
        mIP = cIp
        frmFlag.Show vbModal
    End If
End Sub
Sub GetUserInfo(bServer As String)
 Dim bufptr          As Long
   Dim dwServer        As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim nStatus         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long
   Dim wui1            As WKSTA_USER_INFO_1
     
   dwServer = StrPtr(bServer)
   
   
   Do
   
      nStatus = NetWkstaUserEnum(dwServer, _
                                 1, _
                                 bufptr, _
                                 MAX_PREFERRED_LENGTH, _
                                 dwEntriesread, _
                                 dwTotalentries, _
                                 dwResumehandle)

     'Administrators local group can successfully
     'execute NetWkstaUserEnum locally and on
     'a remote server.
      If nStatus = NERR_SUCCESS Or _
         nStatus = ERROR_MORE_DATA Then
         
         If dwEntriesread > 0 Then
         
            nStructSize = LenB(wui1)
         
   
            For cnt = 0 To dwEntriesread - 1
            
               CpMemory wui1, ByVal bufptr + (nStructSize * cnt), nStructSize
   
               LoginId = GetPointerToByteStringW(wui1.wkui1_username)
               LoginDomain = GetPointerToByteStringW(wui1.wkui1_logon_domain)
               LoginServer = GetPointerToByteStringW(wui1.wkui1_logon_server)
   
                             
   
            Next
            
         End If
      
      
      End If
   
   Loop While nStatus = ERROR_MORE_DATA
   
  'clean up
   Call NetApiBufferFree(bufptr)

End Sub


Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
  
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
   
      tmplen = lstrlenW(dwData) * 2
      
      If tmplen <> 0 Then
      
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CpMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
         
     End If
     
   End If
    
End Function
Function GetSystemInfo(sysName As String) As String
Dim sysInfo As String
On Error Resume Next
Set SystemSet = GetObject("winmgmts:\\" & sysName).InstancesOf("Win32_OperatingSystem")

For Each System In SystemSet
    sysInfo = System.Caption & " " & System.Manufacturer & " " & System.Version
Next
GetSystemInfo = sysInfo
End Function

Private Sub Timer1_Timer()
Command1_Click
End Sub


Private Sub Form_Terminate()
   'Delete the added icon from the taskbar status area when the
   'program ends.
'   Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove _
   (Button As Integer, _
    Shift As Integer, _
    x As Single, _
    Y As Single)
    On Error Resume Next
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
            frmList.Show
'          Shell_NotifyIcon NIM_MODIFY, nid
 
       'Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
            If MsgBox("Are You Sure You Want To Exit From 'WhoIsConnected'", vbCritical + vbYesNo, "Exit") = vbYes Then
                Shell_NotifyIcon NIM_DELETE, nid
                End
                'form1.Show
            End If
       'Case WM_RBUTTONDBLCLK
    End Select
End Sub
