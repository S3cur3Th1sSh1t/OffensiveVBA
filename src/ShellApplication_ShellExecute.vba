Sub AutoOpen()

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "calc.exe", "", "", "open", 1

End Sub