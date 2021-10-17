Sub AutoOpen()
' Using this technique the new process will be spawned under “wmiprvse.exe” instead of the Office process.
Const SW_NORMAL = 1
strComputer = "."
strCommand = "Notepad.exe" 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")

' Configure the Notepad process to show a window
Set objStartup = objWMIService.Get("Win32_ProcessStartup")
Set objConfig = objStartup.SpawnInstance_
objConfig.ShowWindow = SW_NORMAL

' Create Notepad process
Set objProcess = objWMIService.Get("Win32_Process")
intReturn = objProcess.Create _
    (strCommand, Null, objConfig, intProcessID)
If intReturn <> 0 Then
    Wscript.Echo "Process could not be created." & _
        vbNewLine & "Command line: " & strCommand & _
        vbNewLine & "Return value: " & intReturn
Else
    Wscript.Echo "Process created." & _
        vbNewLine & "Command line: " & strCommand & _
        vbNewLine & "Process ID: " & intProcessID
End If

End Sub
