Sub AutoOpen()
Set shell_object = CreateObject("WScript.Shell")
shell_object.Exec ("calc.exe")
End Sub