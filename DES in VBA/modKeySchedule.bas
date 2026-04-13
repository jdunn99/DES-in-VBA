Attribute VB_Name = "modKeySchedule"
Public Function Generate_Subkeys(hexKey As String) As BitBuffer()
    Dim key As New BitBuffer
    Dim keys(0 To 15) As BitBuffer
    Dim L As New BitBuffer
    Dim R As New BitBuffer
    
    Dim i As Long
    Dim j As Long
    Dim shifts As Variant: shifts = Array(1, 1, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 1)
    
    key.To_Binary (hexKey)
    key.Permute (Get_PC1())
    
    Set L = New BitBuffer
    L.Set_Length 28
    
    Set R = New BitBuffer
    R.Set_Length 28
    
    For i = 0 To 27
        L.Set_Bit (i), key.Get_Bit(i)
        R.Set_Bit (i), key.Get_Bit(i + 28)
    Next i
    
    For i = 0 To 15
        L.Shift_Left (CInt(shifts(i)))
        R.Shift_Left (CInt(shifts(i)))
        
        Dim combined As New BitBuffer
        combined.Set_Length 56
        
        For j = 0 To 27
            combined.Set_Bit (j), L.Get_Bit(j)
            combined.Set_Bit (j + 28), R.Get_Bit(j)
        Next j
        
        combined.Permute (Get_PC2())
        Set keys(i) = combined.Copy()
        keys(i).Set_Length 48
    Next i
    Generate_Subkeys = keys
End Function

Public Function Parse_Subkeys(ByVal key As String, ByVal isEncrypt As Boolean) As BitBuffer()
    Dim keys() As BitBuffer
    keys = Generate_Subkeys(key)

    If Not isEncrypt Then
        Dim i As Long, j As Long
        Dim temp As BitBuffer
        
        i = LBound(keys)
        j = UBound(keys)
        
        Do While i < j
            Set temp = keys(i)
            Set keys(i) = keys(j)
            Set keys(j) = temp
            
            i = i + 1
            j = j - 1
        Loop
    End If

    Parse_Subkeys = keys
End Function
