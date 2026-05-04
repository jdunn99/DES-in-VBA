Attribute VB_Name = "modKeySchedule"
Public Function Generate_Subkeys(ByRef keyBytes() As Byte) As Byte()
    Dim shifts As Variant
    shifts = Array(1, 1, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 1)

    Dim permuted() As Byte
    permuted = Permute(keyBytes, Get_PC1(), 56)

    Dim L(0 To 27) As Byte
    Dim r(0 To 27) As Byte
    Dim i As Long, j As Long

    For i = 0 To 27
        L(i) = Get_Bit(permuted, i)
        r(i) = Get_Bit(permuted, i + 28)
    Next i

    Dim result(0 To 95) As Byte

    For i = 0 To 15
        Shift_Left L, 28, CInt(shifts(i))
        Shift_Left r, 28, CInt(shifts(i))

        Dim combined(0 To 6) As Byte
        For j = 0 To 27
            Set_Bit combined, j, L(j)
            Set_Bit combined, j + 28, r(j)
        Next j

        Dim subkey() As Byte
        subkey = Permute(combined, Get_PC2(), 48)

        For j = 0 To 5
            result(i * 6 + j) = subkey(j)
        Next j
    Next i

    Generate_Subkeys = result
End Function

Public Function Parse_Subkeys(ByRef keyBytes() As Byte, ByVal isEncrypt As Boolean) As Byte()
    Dim keys() As Byte
    keys = Generate_Subkeys(keyBytes)
    
    If Not isEncrypt Then
        ' Reverse the 16 subkeys in-place
        Dim i As Long, j As Long
        Dim temp As Byte
        Dim a As Long, b As Long
        
        i = 0
        j = 15
        Do While i < j
            ' Swap subkey i and subkey j
            Dim k As Long
            For k = 0 To 5
                a = i * 6 + k
                b = j * 6 + k
                temp = keys(a)
                keys(a) = keys(b)
                keys(b) = temp
            Next k
            i = i + 1
            j = j - 1
        Loop
    End If
    
    Parse_Subkeys = keys
End Function
