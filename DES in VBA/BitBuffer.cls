VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Bits(0 To 63) As Byte
Private length As Long
Private Sub Class_Initialize()
    length = 64
End Sub

Public Sub Set_Length(l As Long)
    length = l
End Sub

Public Sub Set_Bit(index As Long, value As Byte)
    Bits(index) = IIf(value > 0, 1, 0)
End Sub

Public Function Get_Bit(index As Long) As Byte
    Get_Bit = Bits(index)
End Function

'Shifts n bits to the left
Public Sub Shift_Left(count As Long)
    Dim i As Long
    Dim j As Long
    Dim startBit As Byte
    
    For i = 1 To count
        startBit = Bits(0)
        For j = 0 To length - 2
            Bits(j) = Bits(j + 1)
        Next j
        Bits(length - 1) = startBit
    Next i
End Sub

'Convert hex to binary
Public Sub To_Binary(ByVal str As String)
    Dim i As Long
    Dim j As Long
    Dim digit As Long
    Dim index As Integer
    
    index = 0
    
    For i = 1 To Len(str)
        'Hex to decimal
        digit = val("&H" & Mid(str, i, 1))
        
        'Get the 4 bits
        For j = 3 To 0 Step -1
            If (digit And (2 ^ j)) <> 0 Then
                Bits(index) = 1
            Else
                Bits(index) = 0
            End If
            index = index + 1
        Next j
    Next i
End Sub

'Convert bits to hexadecimal
Public Function To_Hex()
    Dim i As Long
    Dim j As Long
    Dim sum As Long
    Dim result As String
    
    For i = 0 To (length / 4) - 1
        sum = 0
        For j = 0 To 3
            If Bits(i * 4 + j) = 1 Then sum = sum + (2 ^ (3 - j))
        Next j
        result = result & Hex(sum)
    Next i
    To_Hex = result
End Function

Public Sub Permute(table As Variant)
    Dim temp(0 To 63) As Byte
    Dim i As Long
    
    For i = LBound(table) To UBound(table)
        temp(i - LBound(table)) = Bits(table(i) - 1)
    Next i
    For i = 0 To 63
        Bits(i) = temp(i)
    Next i
End Sub

Public Function XOR_Buffer(buffer As BitBuffer) As BitBuffer
    Dim res As New BitBuffer
    Dim i As Long
    res.Set_Length length
    For i = 0 To length - 1
        res.Set_Bit i, (Bits(i) Xor buffer.Get_Bit(i))
    Next i
    Set XOR_Buffer = res
End Function

Public Function Copy() As BitBuffer
    Dim result As New BitBuffer
    result.Set_Length length
    
    Dim i As Long
    For i = 0 To 63
        result.Set_Bit i, Bits(i)
    Next i
    
    Set Copy = result
End Function
