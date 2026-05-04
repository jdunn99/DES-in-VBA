Attribute VB_Name = "modDES_Core"
Option Explicit
' DES Core Algorithm
' Encrpytion, Decryption, Cipher Block Chaining, Feistel Rounds


Public Function Encrypt_Bytes(ByRef data() As Byte, ByRef subkeys() As Byte) As Byte()
    Dim block() As Byte
    Dim L() As Byte
    Dim r() As Byte
    Dim prev() As Byte
    Dim combined(0 To 7) As Byte
    Dim feistelResult() As Byte
    Dim subkey(0 To 5) As Byte
    
    Dim i As Long
    Dim j As Long
    
    block = Permute(data, Get_IP(), 64)
    ReDim L(0 To 3)
    ReDim r(0 To 3)
    
    For i = 0 To 3
        L(i) = block(i)
        r(i) = block(i + 4)
    Next i
    
    For i = 0 To 15
        For j = 0 To 5
            subkey(j) = subkeys(i * 6 + j)
        Next j
        feistelResult = Feistel(r, subkey)
        prev = r
        r = XOR_Bytes(L, feistelResult)
        L = prev
    Next i
    
    For i = 0 To 3
        combined(i) = r(i)
        combined(i + 4) = L(i)
    Next i
    
    Encrypt_Bytes = Permute(combined, Get_Inverse_IP(), 64)
End Function

Public Function DES_Bytes(ByRef data() As Byte, ByRef subkeys() As Byte, ByVal hexIV As String, ByVal isEncrypt As Boolean) As Byte()
    Dim i As Long, j As Long, length As Long
    Dim currentBlock(0 To 7) As Byte
    Dim prevBlock() As Byte
    Dim temp() As Byte
    Dim resultBytes() As Byte
   
    
    If isEncrypt Then Apply_Byte_Padding data
    length = (UBound(data) + 1) / 8
    ReDim resultBytes(UBound(data))
    
    
    prevBlock = Hex_To_Byte(hexIV)

    For i = 0 To length - 1
        For j = 0 To 7: currentBlock(j) = data((i * 8) + j): Next j
        
        temp = currentBlock
        If isEncrypt Then
            temp = XOR_Bytes(temp, prevBlock)
            temp = Encrypt_Bytes(temp, subkeys)
            prevBlock = temp
        Else
            temp = Encrypt_Bytes(currentBlock, subkeys)
            temp = XOR_Bytes(temp, prevBlock)
        prevBlock = currentBlock
        End If
        For j = 0 To 7: resultBytes((i * 8) + j) = temp(j): Next j
        If i Mod 5000 = 0 Then DoEvents
    Next i

    
    DES_Bytes = resultBytes
End Function


Public Function Run_DES(ByVal plaintext As String, ByRef subkeys() As Byte, ByVal hexIV As String, ByVal isEncrypt As Boolean)
    Dim data() As Byte
    Dim result() As Byte
    Dim resultStr As String
    Dim pt As String
    
    pt = plaintext
    If isEncrypt Then
        pt = String_To_Hex(plaintext)
    End If
    
    data = Hex_To_Byte(pt)
    result = DES_Bytes(data, subkeys, hexIV, isEncrypt)
    resultStr = Byte_To_Hex(result)
    
    If Not isEncrypt Then
        result = Remove_Byte_Padding(result)
        resultStr = Byte_To_Hex(result)
        
        Dim j As Long
        Dim i As Long
        Dim currentHex As String
        Dim hexToStr As String
        
        For i = 1 To Len(resultStr) Step 2
            currentHex = Mid(resultStr, i, 2)
            j = val("&H" & currentHex)
            hexToStr = hexToStr & Chr(j)
        Next i
        
        resultStr = hexToStr
    End If
    
    Run_DES = resultStr
End Function

Public Function Feistel(ByRef r() As Byte, ByRef subkey() As Byte) As Byte()
    ' Expand R from 32 to 48 bits using E D-box
    Dim expanded() As Byte
    expanded = Permute(r, Get_Exp_Dbox(), 48)
    
    ' XOR with subkey
    expanded = XOR_Bytes(expanded, subkey)
    
    ' S-box substitution
    Dim sboxResult(0 To 3) As Byte
    Dim i As Long, j As Long
    Dim row As Long, col As Long, sval As Long
    
    For i = 0 To 7
        row = Get_Bit(expanded, i * 6) * 2 + Get_Bit(expanded, i * 6 + 5)
        col = Get_Bit(expanded, i * 6 + 1) * 8 + Get_Bit(expanded, i * 6 + 2) * 4 + _
              Get_Bit(expanded, i * 6 + 3) * 2 + Get_Bit(expanded, i * 6 + 4)
        sval = Get_SBox_Value(i, row, col)
        
        For j = 0 To 3
            If (sval And (2 ^ (3 - j))) <> 0 Then
                Set_Bit sboxResult, i * 4 + j, 1
            Else
                Set_Bit sboxResult, i * 4 + j, 0
            End If
        Next j
    Next i
    
    ' Apply P-box permutation
    Feistel = Permute(sboxResult, Get_PBox(), 32)
End Function
