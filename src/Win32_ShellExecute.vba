Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
         ByVal hwnd As Long, _
         ByVal lpOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal lpShowCmd As Long) As Long

Sub AutoOpen()
Call ShellExecute(0, "Open", "calc.exe", "", "", 1)

End Sub