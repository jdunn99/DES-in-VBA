Attribute VB_Name = "modBit_Operations"
Public Function Get_Bit(ByRef data() As Byte, ByVal n As Long) As Byte
    Get_Bit = (data(n \ 8) \ (2 ^ (7 - (n Mod 8)))) And 1
End Function

Public Sub Set_Bit(ByRef data() As Byte, ByVal n As Long, ByVal value As Byte)
    Dim index As Long
    Dim bitIndex As Long

    index = n \ 8
    bitIndex = 7 - (n Mod 8)

    If value = 1 Then
      data(index) = data(index) Or (1 * 2 ^ bitIndex)
    Else
      data(index) = data(index) And CByte(Not CByte(1 * 2 ^ bitIndex))
    End If
End Sub

Public Function Permute(ByRef data() As Byte, ByRef table As Variant, ByVal output As Long) As Byte()
    Dim result() As Byte
    ReDim result(0 To (output \ 8) - 1)

    Dim i As Long
    For i = 0 To output - 1
        Set_Bit result, i, Get_Bit(data, table(i) - 1)
    Next i
    Permute = result
End Function

Public Function XOR_Bytes(ByRef a() As Byte, ByRef b() As Byte) As Byte()
    Dim result() As Byte
    Dim i As Long
    ReDim result(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        result(i) = a(i) Xor b(i)
    Next i
    XOR_Bytes = result
End Function

Public Sub Shift_Left(ByRef bits() As Byte, ByVal length As Long, ByVal count As Long)
    Dim i As Long, c As Long
    Dim first As Byte
    For c = 1 To count
        first = bits(0)
        For i = 0 To length - 2
            bits(i) = bits(i + 1)
        Next i
        bits(length - 1) = first
    Next c
End Sub

