Attribute VB_Name = "DES"
Sub Run_DES()
    Dim plaintext As String
    plaintext = Range("B1").value
    
    Dim key As String
    key = Range("B2").value
    
    Dim initVector As String
    initVector = Range("B4").value
    
    If initVector = "" Then initVector = "0000000000000000"
    
    If Len(key) <> 16 Then
        MsgBox "Key must be 16 characters"
        Exit Sub
    End If
    
    Dim subkeys() As BitBuffer
    subkeys = Generate_Subkeys(key)
    
    Dim result As String
    result = Encrypt(plaintext, key, subkeys, initVector) ' FIXED
    
    Range("B3").value = result
    Range("B5").value = DES_Decrypt(result, key, initVector)
End Sub

Public Function DES_Encrypt(hexPlaintext As String, hexKey As String, subkeys() As BitBuffer) As String
    Dim pt As New BitBuffer
    Dim key As New BitBuffer
    Dim l As New BitBuffer
    Dim R As New BitBuffer
    
    Dim i As Long
    
    pt.To_Binary (hexPlaintext)
    pt.Permute (Get_IP())
    
    Set l = New BitBuffer
    l.Set_Length (32)
    
    Set R = New BitBuffer
    R.Set_Length 32
    
    For i = 0 To 31
        l.Set_Bit i, pt.Get_Bit(i)
        R.Set_Bit i, pt.Get_Bit(i + 32)
    Next i
    
    For i = 0 To 15
        Dim prev As BitBuffer
        Set prev = R.Copy()
        
        Set R = l.XOR_Buffer(Feistel_Function(R, subkeys(i)))
        Set l = prev
    Next i
    
    Dim combined As New BitBuffer
    For i = 0 To 31
        combined.Set_Bit i, R.Get_Bit(i)
        combined.Set_Bit i + 32, l.Get_Bit(i)
    Next i
    combined.Permute (Get_Inverse_IP())
    
    DES_Encrypt = combined.To_Hex()
End Function

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
    
    Dim test As String
    For i = 1 To Len(result) Step 2
        currentHex = Mid(result, i, 2)
        j = val("&H" & currentHex)
        test = test & Chr(j)
    Next i
    
    DES_Decrypt = test
End Function

Public Function Feistel_Function(R As BitBuffer, subkey As BitBuffer) As BitBuffer
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
    Set Feistel_Function = result
End Function

Public Function Generate_Subkeys(hexKey As String) As BitBuffer()
    Dim key As New BitBuffer
    Dim keys(0 To 15) As BitBuffer
    Dim l As New BitBuffer
    Dim R As New BitBuffer
    
    Dim i As Long
    Dim j As Long
    Dim shifts As Variant: shifts = Array(1, 1, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 1)
    
    key.To_Binary (hexKey)
    key.Permute (Get_PC1())
    
    Set l = New BitBuffer
    l.Set_Length 28
    
    Set R = New BitBuffer
    R.Set_Length 28
    
    For i = 0 To 27
        l.Set_Bit (i), key.Get_Bit(i)
        R.Set_Bit (i), key.Get_Bit(i + 28)
    Next i
    
    For i = 0 To 15
        l.Shift_Left (CInt(shifts(i)))
        R.Shift_Left (CInt(shifts(i)))
        
        Dim combined As New BitBuffer
        combined.Set_Length 56
        
        For j = 0 To 27
            combined.Set_Bit (j), l.Get_Bit(j)
            combined.Set_Bit (j + 28), R.Get_Bit(j)
        Next j
        
        combined.Permute (Get_PC2())
        Set keys(i) = combined.Copy()
        keys(i).Set_Length 48
    Next i
    Generate_Subkeys = keys
End Function

Public Function Apply_Padding(ByVal text As String) As String
    Dim length As Long
    Dim additionalBytes As Long
    Dim padding As String
    Dim i As Integer
    
    length = Len(text) / 2
    additionalBytes = 8 - (length Mod 8)
    padding = Right("0" & Hex(additionalBytes), 2)
    
    For i = 1 To additionalBytes
        text = text & padding
    Next i
    
    Apply_Padding = text
End Function

Public Function Remove_Padding(ByVal text As String) As String
    Dim lastByte As String
    Dim bytesToRemove As Long
    
    lastByte = Right(text, 2)
    bytesToRemove = val("&H" & lastByte)
    
    Remove_Padding = Left(text, Len(text) - (bytesToRemove * 2))
End Function

Public Function Encrypt(ByVal plaintext As String, ByVal key As String, subkeys() As BitBuffer, ByVal hexIV As String) As String
    Dim i As Long
    Dim chunk As String
    Dim result As String
    Dim previous As New BitBuffer
    
    previous.To_Binary hexIV
    
    plaintext = String_To_Hex(plaintext)
    plaintext = Apply_Padding(plaintext)

    For i = 1 To Len(plaintext) Step 16
        chunk = Mid(plaintext, i, 16)
        Dim current As New BitBuffer
        current.To_Binary (chunk)
        
        Set current = current.XOR_Buffer(previous)
        
        Dim encryptedHex As String
        encryptedHex = DES_Encrypt(current.To_Hex(), key, subkeys)
        
        result = result & encryptedHex
        
        previous.To_Binary (encryptedHex)
    Next i
    
    Encrypt = result
End Function

Public Function String_To_Hex(text As String) As String
    Dim i As Long
    Dim result As String
    
    For i = 1 To Len(text)
        result = result & Right("0" & Hex(Asc(Mid(text, i, 1))), 2)
    Next i
    
    String_To_Hex = result
End Function
