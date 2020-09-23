Attribute VB_Name = "Module1"
Option Explicit

Declare Function ReleaseCapture _
  Lib "user32" () As Long

Declare Function SendMessage Lib _
  "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const SOCKET_ERROR = -1
Public Const AF_INET = 2
Public Const PF_INET = AF_INET
Public Const MAXGETHOSTSTRUCT = 1024
Public Const SOCK_STREAM = 1
Public Const MSG_PEEK = 2

Private Type SockAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As String * 4
    sin_zero As String * 8
End Type

Private Type T_WSA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Dim WSAData As T_WSA

Type Inet_Address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type

Public IPStruct As Inet_Address

Public Type T_Host
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

' KERNEL32.DLL funtions
Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)

' WSOCK32.DLL functions
Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Declare Function inet_addr Lib "wsock32.dll" (ByVal addr As String) As Long
Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Declare Function GetHostName Lib "wsock32.dll" Alias "gethostname" (ByVal HostName As String, HostLen As Long) As Long
Declare Function WSAStartup Lib "wsock32.dll" (ByVal a As Long, b As T_WSA) As Long
Declare Function WSACleanUp Lib "wsock32.dll" Alias "WSACleanup" () As Integer
Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal afinet As Integer, ByVal socktype As Integer, ByVal protocol As Integer) As Long
Declare Function ConnectWinsock Lib "wsock32.dll" Alias "connect" (ByVal sock As Long, sockstruct As SockAddr, ByVal structlen As Integer) As Integer
Declare Function send Lib "wsock32.dll" (ByVal sock As Long, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function recv Lib "wsock32.dll" (ByVal sock As Long, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function htonl Lib "wsock32.dll" (ByVal a As Long) As Long
Declare Function ntohl Lib "wsock32.dll" (ByVal a As Long) As Long
Declare Function htons Lib "wsock32.dll" (ByVal a As Integer) As Integer
Declare Function ntohs Lib "wsock32.dll" (ByVal a As Integer) As Integer
Declare Function closesocket Lib "wsock32.dll" (ByVal sn As Long) As Integer
 
 Function HostByName(sHost As String) As String
    Dim s As String
    Dim p As Long
    Dim Host As T_Host
    Dim ListAddress As Long
    Dim ListAddr As Long
    Dim Address As Long

    s = String(64, 0)
    sHost = sHost + Right(s, 64 - Len(sHost))
    p = GetHostByName(sHost)
    If p = SOCKET_ERROR Then
        Exit Function
    Else
        If p <> 0 Then
            CopyMemory Host.h_name, ByVal p, Len(Host)
            ListAddress = Host.h_addr_list
            CopyMemory ListAddr, ByVal ListAddress, 4
            CopyMemory Address, ByVal ListAddr, 4
            HostByName = InetAddrLongToString(Address)
        Else
            'HostByName = "No DNS Entry"
        End If
    End If
End Function

Private Function InetAddrStringToLong(Address As String) As Long
    InetAddrStringToLong = inet_addr(Address)
End Function

Private Function InetAddrLongToString(Address As Long) As String
    CopyMemory IPStruct, Address, 4
    InetAddrLongToString = CStr(Asc(IPStruct.Byte4)) + "." + CStr(Asc(IPStruct.Byte3)) + "." + CStr(Asc(IPStruct.Byte2)) + "." + CStr(Asc(IPStruct.Byte1))
End Function

Function HostByAddress(ByVal sAddress As String) As String
    Dim lAddress As Long
    Dim p As Long
    Dim HostName As String
    Dim Host As T_Host

    lAddress = inet_addr(sAddress)
    p = gethostbyaddr(lAddress, 4, PF_INET)
    If p <> 0 Then
        CopyMemory Host, ByVal p, Len(Host)
        HostName = String(256, 0)
        CopyMemory ByVal HostName, ByVal Host.h_name, 256
        If HostName = "" Then HostByAddress = "Unable to Resolve Address"
        HostByAddress = Left(HostName, InStr(HostName, Chr(0)) - 1)
    Else
        'HostByAddress = "No DNS Entry"
    End If
End Function

Private Function ResolveHost(sHost As String) As Long
    Dim lAddress As Long

    lAddress = InetAddrStringToLong(sHost)
    If lAddress = SOCKET_ERROR Then
        ResolveHost = inet_addr(HostByName(sHost))
    Else
        ResolveHost = lAddress
    End If
End Function

Public Function WinsockConnect(ByVal m_RemoteHost As String, m_RemotePort As Long, iSocket As Long) As Boolean
    Dim sock As SockAddr
    Dim sRemoteIP As String
    Dim X As Long
    Dim bAddr(0 To 3) As Byte
    Dim i As Integer

    iSocket = Socket(AF_INET, SOCK_STREAM, 0)
    If iSocket < 1 Then
        WinsockConnect = False
        Exit Function
    End If
    sRemoteIP = ""
    sock.sin_family = AF_INET
    X = ResolveHost(m_RemoteHost)
    CopyMemory bAddr(0), X, 4
    For i = 0 To 3
        sRemoteIP = sRemoteIP & Chr(bAddr(i))
    Next
    sock.sin_addr = sRemoteIP
    sock.sin_port = htons(m_RemotePort)
    sock.sin_zero = String(8, 0)
    If ConnectWinsock(iSocket, sock, Len(sock)) Then
        WinsockConnect = False
        Exit Function
    End If
    WinsockConnect = True
End Function

Public Sub WinsockInit()
    WSAStartup &H101, WSAData
End Sub
 

