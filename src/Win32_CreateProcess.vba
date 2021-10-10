Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle       As Long
End Type

Private Type STARTUPINFO
    cb              As Long
    lpReserved      As String
    lpDesktop       As String
    lpTitle         As String
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Byte
    hStdInput       As LongPtr
    hStdOutput      As LongPtr
    hStdError       As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess    As LongPtr
    hThread     As LongPtr
    dwProcessId As Long
    dwThreadId  As Long
End Type

Private Const CREATE_NEW_CONSOLE = &H10
Private Const CREATE_SUSPENDED = &H4
Private Const DEBUG_ONLY_THIS_PROCESS = &H2

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As LongPtr


Sub AutoOpen()



    Dim secAttrPrc As SECURITY_ATTRIBUTES: secAttrPrc.nLength = Len(secAttrPrc)
    Dim secAttrThr As SECURITY_ATTRIBUTES: secAttrThr.nLength = Len(secAttrThr)

    Dim startInfo  As STARTUPINFO
    Dim procInfo   As PROCESS_INFORMATION

    If CreateProcess( _
         lpApplicationName:=vbNullString, _
         lpCommandLine:="cmd.exe", _
         lpProcessAttributes:=secAttrPrc, _
         lpThreadAttributes:=secAttrThr, _
         bInheritHandles:=False, _
         dwCreationFlags:=0, _
         lpEnvironment:=0, _
         lpCurrentDirectory:=Environ("USERPROFILE"), _
         lpStartupInfo:=startInfo, _
         lpProcessInformation:=procInfo) Then

     Else
        MsgBox "Couldnt create process"
     End If

End Sub
