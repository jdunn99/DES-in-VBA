Attribute VB_Name = "modDES_Core"
Option Explicit
' DES Core Algorithm
' Encrpytion, Decryption, Cipher Block Chaining, Feistel Rounds

Public Function DES_Bytes(ByRef data() As Byte, ByVal key As String, ByVal hexIV As String, ByRef subkeys() As BitBuffer, ByVal isEncrypt As Boolean) As String
    Apply_Byte_Padding data
    Dim result As String
    
    Dim pt As String
    pt = Byte_To_Hex(data)
    result = Cipher_Block_Chain(pt, key, subkeys, True, hexIV)
    
    DES_Bytes = result
End Function



Public Function DES_Core(ByVal plaintext As String, ByVal key As String, ByVal hexIV As String, ByRef subkeys() As BitBuffer, ByVal isEncrypt As Boolean) As String
    Dim result As String
    Dim i As Long
  
    If isEncrypt Then plaintext = Apply_Padding(String_To_Hex(plaintext))
    
    Debug.Print plaintext
    result = ""
    'result = Cipher_Block_Chain(plaintext, key, subkeys, isEncrypt, hexIV)

    If Not isEncrypt Then
        Dim j As Long
        Dim currentHex As String
        Dim hexToStr As String
        
        result = Remove_Padding(result)
        
        For i = 1 To Len(result) Step 2
            currentHex = Mid(result, i, 2)
            j = val("&H" & currentHex)
            hexToStr = hexToStr & Chr(j)
        Next i

        result = hexToStr
    End If

    DES_Core = result
End Function

Public Function Cipher_Block_Chain_Bytes(ByRef data() As Byte, ByVal key As String, ByRef subkeys() As BitBuffer, ByVal isEncrypt As Boolean, ByVal hexIV As String)
    Dim i As Long
    Dim j As Long
    
    Dim numBlocks As Long
    Dim currentBlock(0 To 7) As Byte
    Dim previousBlock(0 To 7) As Byte
    Dim result() As Byte
    
    
End Function

Public Function Cipher_Block_Chain(ByVal plaintext As String, ByVal key As String, ByRef subkeys() As BitBuffer, ByVal isEncrypt As Boolean, ByVal hexIV As String) As String

    Dim prev As BitBuffer
    Set prev = New BitBuffer
    prev.To_Binary hexIV

    Dim result As String
    Dim i As Long

    For i = 1 To Len(plaintext) Step 16
        Dim chunk As String
        chunk = Mid(plaintext, i, 16)

        If isEncrypt Then
            Dim current As New BitBuffer
            current.To_Binary chunk
            Set current = current.XOR_Buffer(prev)
            
            Dim encrypted As String
            encrypted = Encrypt(current.To_Hex(), key, subkeys)
            
            result = result & encrypted
            prev.To_Binary encrypted

        Else
            Dim decrypted As String
            decrypted = Encrypt(chunk, key, subkeys) ' reversed keys
            
            Dim tempBuf As New BitBuffer
            tempBuf.To_Binary decrypted
            
            Set tempBuf = tempBuf.XOR_Buffer(prev)
            result = result & tempBuf.To_Hex()
            
            prev.To_Binary chunk
        End If
    Next i

    Cipher_Block_Chain = result
End Function

Public Function Encrypt(hexPlaintext As String, hexKey As String, subkeys() As BitBuffer) As String

    Dim plaintext As New BitBuffer
    Dim L As New BitBuffer
    Dim R As New BitBuffer
    Dim i As Long

    plaintext.To_Binary hexPlaintext
    plaintext.Permute Get_IP()

    L.Set_Length 32
    R.Set_Length 32

    For i = 0 To 31
        L.Set_Bit i, plaintext.Get_Bit(i)
        R.Set_Bit i, plaintext.Get_Bit(i + 32)
    Next i

    Dim prev As BitBuffer
    For i = 0 To 15
        Set prev = R.Copy()
        Set R = L.XOR_Buffer(Feistel(R, subkeys(i)))
        Set L = prev
    Next i

    Dim combined As New BitBuffer

    For i = 0 To 31
        combined.Set_Bit i, R.Get_Bit(i)
        combined.Set_Bit i + 32, L.Get_Bit(i)
    Next i

    combined.Permute Get_Inverse_IP()

    Encrypt = combined.To_Hex()
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
