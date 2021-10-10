Sub AutoOpen()

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "calc.exe", "", "", "runas", 1

End Sub