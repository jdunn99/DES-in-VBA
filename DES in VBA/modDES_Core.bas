Attribute VB_Name = "modDES_Core"
Option Explicit
' DES Core Algorithm
' Encrpytion, Decryption, Cipher Block Chaining, Feistel Rounds

' DES encryption function.
'
' Param hexPlaintext - The plaintext being encrypted
' Param hexKey - The key used for encrpytion
' Param subkeys - Array of 16 48-bit key s
'
' Returns String - The resulting encrypted string in hexadecimal format
Public Function DES_Encrypt(hexPlaintext As String, hexKey As String, subkeys() As BitBuffer) As String
    Dim pt As New BitBuffer
    Dim key As New BitBuffer
    Dim l As New BitBuffer
    Dim R As New BitBuffer
    
    Dim i As Long
    
    ' Initial permutation
    pt.To_Binary (hexPlaintext)
    pt.Permute (Get_IP())
    
    ' Split left and right substrings
    Set l = New BitBuffer
    l.Set_Length (32)
    
    Set R = New BitBuffer
    R.Set_Length 32
    
    For i = 0 To 31
        l.Set_Bit i, pt.Get_Bit(i)
        R.Set_Bit i, pt.Get_Bit(i + 32)
    Next i
    
    ' Perform 16 ronuds of Feistel
    For i = 0 To 15
        Dim prev As BitBuffer
        Set prev = R.Copy()
        
        Set R = l.XOR_Buffer(Feistel(R, subkeys(i)))
        Set l = prev
    Next i
    
    ' Rebuild buffer and convert to hex string
    Dim combined As New BitBuffer
    For i = 0 To 31
        combined.Set_Bit i, R.Get_Bit(i)
        combined.Set_Bit i + 32, l.Get_Bit(i)
    Next i
    combined.Permute (Get_Inverse_IP())
    
    DES_Encrypt = combined.To_Hex()
End Function

' DES decryption function.
'
' Param hexCiphertext - The ciphertext being decrypted
' Param hexKey - The key used for decryption
' Param subkeys - Array of 16 48-bit key s
'
' Returns String - The resulting encrypted string in hexadecimal format
Public Function DES_Decrypt(ByVal hexCiphertext As String, ByVal hexKey As String, ByVal hexIV As String)
    Dim i As Long
    Dim currentHex As String
    Dim decrypted As String
    Dim result As String
    
    Dim current As New BitBuffer
    Dim prev As New BitBuffer
    Dim subkeys() As BitBuffer
    
    subkeys = Generate_Subkeys(hexKey)
    
    Dim j As Long
    Dim temp As BitBuffer
    
    i = LBound(subkeys)
    j = UBound(subkeys)
    
    Do While i < j
        Set temp = subkeys(i)
        Set subkeys(i) = subkeys(j)
        Set subkeys(j) = temp
        i = i + 1
        j = j - 1
    Loop
    
    prev.To_Binary hexIV
    
    For i = 1 To Len(hexCiphertext) Step 16
        currentHex = Mid(hexCiphertext, i, 16)

        decrypted = DES_Encrypt(currentHex, hexKey, subkeys)
        current.To_Binary decrypted
        
        Set current = current.XOR_Buffer(prev)
        result = result & current.To_Hex()
        
        ' FIXED: update previous block
        prev.To_Binary currentHex
    Next i
    
    result = Remove_Padding(result)
    
    Dim hexToStr As String
    For i = 1 To Len(result) Step 2
        currentHex = Mid(result, i, 2)
        j = val("&H" & currentHex)
        hexToStr = hexToStr & Chr(j)
    Next i
    
    DES_Decrypt = hexToStr
End Function

Public Function Encrypt(plaintext As String, key As String, subkeys() As BitBuffer, hexIV As String) As String
    Dim previous As New BitBuffer
    previous.To_Binary hexIV
    
    plaintext = Apply_Padding(String_To_Hex(plaintext))
    
    Dim result As String
    Dim i As Long
    
    For i = 1 To Len(plaintext) Step 16
        Dim chunk As String
        chunk = Mid(plaintext, i, 16)
        
        Dim current As New BitBuffer
        current.To_Binary chunk
        Set current = current.XOR_Buffer(previous)
        
        Dim encrypted As String
        encrypted = DES_Encrypt(current.To_Hex(), key, subkeys)
        
        result = result & encrypted
        previous.To_Binary encrypted
    Next i
    
    Encrypt = result
End Function

Public Function Apply_Padding(text As String) As String
    Dim length As Long
    Dim padBytes As Long
    Dim pad As String
    Dim i As Long
    
    length = Len(text) \ 2
    padBytes = 8 - (length Mod 8)
    pad = Right("0" & Hex(padBytes), 2)
    
    For i = 1 To padBytes
        text = text & pad
    Next i
    
    Apply_Padding = text
End Function


Public Function Feistel(R As BitBuffer, subkey As BitBuffer) As BitBuffer
    Dim expanded As BitBuffer
    Set expanded = R.Copy()
    expanded.Set_Length 48
    
    expanded.Permute (Get_Exp_Dbox())
    Set expanded = expanded.XOR_Buffer(subkey)
    
    Dim result As New BitBuffer
    Dim i As Long
    Dim j As Long
    Dim mask As Long
    Dim row As Long
    Dim col As Long
    Dim val As Long
    
    For i = 0 To 7
        row = expanded.Get_Bit(i * 6) * 2 + expanded.Get_Bit(i * 6 + 5)
        col = expanded.Get_Bit(i * 6 + 1) * 8 + expanded.Get_Bit(i * 6 + 2) * 4 + _
              expanded.Get_Bit(i * 6 + 3) * 2 + expanded.Get_Bit(i * 6 + 4)
        val = Get_SBox_Value(i, row, col)
        
        For j = 0 To 3
            mask = 2 ^ (3 - j)
            If (val And mask) <> 0 Then
                result.Set_Bit i * 4 + j, 1
            Else
                result.Set_Bit i * 4 + j, 0
            End If
        Next j
    Next i
    
    result.Permute (Get_PBox())
    Set Feistel = result
End Function
