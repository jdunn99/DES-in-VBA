Attribute VB_Name = "modKeySchedule"
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
