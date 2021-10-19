' Credit to https://github.com/susam/aes.vbs/blob/a0cb5f9ffbd90b435622f5cfdb84264e1a319bf2/aes.vbs
' AES-256-CBC with HMAC-SHA-256 in VBScript

Set utf8 = CreateObject("System.Text.UTF8Encoding")
Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform")
Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform")
Set mac = CreateObject("System.Security.Cryptography.HMACSHA256")
Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")
Set mem = CreateObject("System.IO.MemoryStream")


' Return the minimum value between two integer values.
'
' Arguments:
'   a (Long): An integer.
'   b (Long): Another integer.
'
' Return:
'   Long: Minimum of the two integer values.
Function Min(a, b)
    Min = a
    If b < a Then Min = b
End Function


' Convert a byte array to a Base64 string representation of it.
'
' Arguments:
'   bytes (Byte()): Byte array.
'
' Returns:
'   String: Base64 representation of the input byte array.
Function B64Encode(bytes)
    blockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(bytes) - 1 Step blockSize
        length = Min(blockSize, LenB(bytes) - offset)
        b64Block = b64Enc.TransformFinalBlock((bytes), offset, length)
        result = result & utf8.GetString((b64Block))
    Next
    B64Encode = result
End Function


' Convert a Base64 string to a byte array.
'
' Arguments:
'   b64Str (String): Base64 string.
'
' Returns:
'   Byte(): A byte array that the Base64 string decodes to.
Function B64Decode(b64Str)
    bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((bytes), 0, LenB(bytes))
End Function


' Concatenate two byte arrays.
'
' Arguments:
'   a (Byte()): A byte array.
'   b (Byte()): Another byte array.
'
' Returns:
'   Byte(): Concatenated byte arrays.
Function ConcatBytes(a, b)
    mem.SetLength(0)
    mem.Write (a), 0, LenB(a)
    mem.Write (b), 0, LenB(b)
    ConcatBytes = mem.ToArray()
End Function


' Check if two byte arrays are equal.
'
' Arguments:
'   a (Byte()): A byte array.
'   b (Byte()): Another byte array.
'
' Returns:
'   Boolean: True if both byte arrays are equal; False otherwise.
Function EqualBytes(a, b)
    EqualBytes = False
    If LenB(a) <> LenB(b) Then Exit Function
    diff = 0
    For i = 1 to LenB(a)
        diff = diff Or (AscB(MidB(a, i, 1)) Xor AscB(MidB(b, i, 1)))
    Next
    EqualBytes = Not diff
End Function


' Compute message authentication code using HMAC-SHA-256.
'
' Arguments:
'   msgBytes (Byte()): Message to be authenticated.
'   keyBytes (Byte()): Secret key.
'
' Returns:
'   Byte(): Message authenticate code.
Function ComputeMAC(msgBytes, keyBytes)
    mac.Key = keyBytes
    ComputeMAC = mac.ComputeHash_2((msgBytes))
End Function


' Encrypt plaintext and compute MAC for the result.
'
' The length of AES encryption key (aesKey) must be 256 bits (32 bytes).
' It must be provided as a Base64 encoded string. On macOS or Linux,
' enter this command to generate a Base64 encoded 256-bit key:
'
'   head -c32 /dev/urandom | base64
'
' The HMAC secret key (macKey) can be any length but a minimum of
' 256 bits (32 bytes) is recommended as the length of this key. It must
' be provided as a Base64 encoded string.
'
' The return value of this function is composed of the following three
' Base64 encoded strings joined with colons:
'
'   - Message authentication code.
'   - Randomly generated 128-bit initialization vector (IV).
'   - Ciphertext.
'
' Note:
'
'   - A 256-bit key after Base64 encoding contains 44 characters
'     including one '=' character as padding at the end.
'   - A 128-bit IV after Base64 encoding contains 24 characters
'     including two '=' characters as padding at the end.
'
' Arguments:
'   plaintext (String): Text to be encrypted.
'   aesKey (String): AES encryption key encoded as a Base64 string.
'   macKey (String): HMAC secret key encoded as a Base64 string.
'
' Returns:
'   String: MAC, IV, and ciphertext joined with colons.
Function Encrypt(plaintext, aesKey, macKey)
    aes.GenerateIV()
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    Set aesEnc = aes.CreateEncryptor_2((aesKeyBytes), aes.IV)
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    macBytes = ComputeMAC(ConcatBytes(aes.IV, cipherBytes), macKeyBytes)
    Encrypt = B64Encode(macBytes) & ":" & B64Encode(aes.IV) & ":" & _
              B64Encode(cipherBytes)
End Function


' Decrypt ciphertext after authenticating IV and ciphertext using MAC.
'
' MAC, IV, and ciphertext must be encoded in Base64. They are provided
' together as a single string with the Base64 encoded values separated
' by colons. See the comment for Encrypt() function to read more about
' the format.
'
' Arguments:
'   macIVCipherText (String): Colon separated MAC, IV, and ciphertext.
'   aesKey (String): AES encryption key encoded as a Base64 string.
'   macKey (String): HMAC secret key encoded as a Base64 string.
'
' Returns:
'   String: Plaintext that the given ciphertext decrypts to.
Function Decrypt(macIVCiphertext, aesKey, macKey)
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    tokens = Split(macIVCiphertext, ":")
    macBytes = B64Decode(tokens(0))
    ivBytes = B64Decode(tokens(1))
    cipherBytes = B64Decode(tokens(2))
    macActual = ComputeMAC(ConcatBytes(ivBytes, cipherBytes), macKeyBytes)
    If Not EqualBytes(macBytes, macActual) Then
        Err.Raise vbObjectError + 1000, "Decrypt()", "Bad MAC"
    End If
    Set aesDec = aes.CreateDecryptor_2((aesKeyBytes), (ivBytes))
    plainBytes = aesDec.TransformFinalBlock((cipherBytes), 0, LenB(cipherBytes))
    Decrypt = utf8.GetString((plainBytes))
End Function


' Show results of Encrypt() and Decrypt() calls for demo purpose.
Function CryptoDemo
    demoPlaintext = "hello"
    demoAESKey = "CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg="
    demoMACKey = "wDF4W9XQ6wy2DmI/7+ONF+mwCEr9tVgWGLGHUYnguh4="
    encrypted2 = "7BnQ5trOLDk8cecEnVayfSW9Q2fA38FvFkDlwHxbAKA=" & _
                 ":M1ipFnh884qcXYlX9NPjwA==" & _
                 ":ANF8P0PfaUQwvcS2jiIpdQ=="

    encrypted1 = Encrypt(demoPlaintext, demoAESKey, demoMACKey)
    decrypted1 = Decrypt(encrypted1, demoAESKey, demoMACKey)
    decrypted2 = Decrypt(encrypted2, demoAESKey, demoMACKey)

    CryptoDemo = "demoAESKey: " & demoAESKey & vbCrLf & _
                 "demoMACKey: " & demoMACKey & vbCrLF & _
                 "encrypted1: " & encrypted1 & vbCrLf & _
                 "decrypted1: " & decrypted1 & vbCrLf & _
                 "encrypted2: " & encrypted2 & vbCrLf & _
                 "decrypted2: " & decrypted2 & vbCrLf
End Function


' Show interesting properties of cryptography objects used here.
Function CryptoInfo
    Set enc = aes.CreateEncryptor_2(aes.Key, aes.IV)
    Set dec = aes.CreateDecryptor_2(aes.Key, aes.IV)

    CryptoInfo = "aes.BlockSize: " & aes.BlockSize & vbCrLf & _
                 "aes.FeedbackSize: " & aes.FeedbackSize & vbCrLf & _
                 "aes.KeySize: " & aes.KeySize & vbCrLf & _
                 "aes.Mode: " & aes.Mode & vbCrLf & _
                 "aes.Padding: " & aes.Padding & vbCrLf & _
                 "mac.HashName: " & mac.HashName & vbCrLf & _
                 "mac.HashSize: " & mac.HashSize & vbCrLf & _
                 "aesEnc.InputBlockSize: " & enc.InputBlockSize & vbCrLf & _
                 "aesEnc.OutputBlockSize: " & enc.OutputBlockSize & vbCrLf & _
                 "aesDec.InputBlockSize: " & enc.InputBlockSize & vbCrLf & _
                 "aesDec.OutputBlockSize: " & enc.OutputBlockSize & vbCrLf & _
                 "b64Enc.InputBlockSize: " & b64Enc.InputBlockSize & vbCrLf & _
                 "b64Enc.OutputBlockSize: " & b64Enc.OutputBlockSize & vbCrLf & _
                 "b64Dec.InputBlockSize: " & b64Dec.InputBlockSize & vbCrLf & _
                 "b64Dec.OutputBlockSize: " & b64Dec.OutputBlockSize & vbCrLf
End Function
