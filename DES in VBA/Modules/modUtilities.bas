Attribute VB_Name = "modUtilities"
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

Public Sub Apply_Byte_Padding(ByRef data() As Byte)
    Dim originalSize As Long: originalSize = UBound(data) + 1
    Dim paddingNeeded As Integer
    Dim newSize As Long
    Dim i As Long
    
    paddingNeeded = 8 - (originalSize Mod 8)
    newSize = originalSize + paddingNeeded
    
    ReDim Preserve data(newSize - 1)

    For i = originalSize To newSize - 1
        data(i) = CByte(paddingNeeded)
    Next i
End Sub

Public Function Remove_Byte_Padding(ByRef data() As Byte) As Byte()
    Dim lastByte As Integer
    Dim newSize As Long
    Dim result() As Byte
    Dim i As Long
    
    lastByte = CInt(data(UBound(data)))
    
    If lastByte < 1 Or lastByte > 8 Then
        Remove_Byte_Padding = data
        Exit Function
    End If
    
    newSize = (UBound(data) + 1) - lastByte
    ReDim result(0 To newSize - 1)
    
    For i = 0 To newSize - 1
        result(i) = data(i)
    Next i
    
    Remove_Byte_Padding = result
End Function

Public Function Remove_Padding(ByVal text As String) As String
    Dim lastByte As String
    Dim bytesToRemove As Long
    
    lastByte = Right(text, 2)
    bytesToRemove = val("&H" & lastByte)
    
    Remove_Padding = Left(text, Len(text) - (bytesToRemove * 2))
End Function

Public Function String_To_Hex(text As String) As String
    Dim i As Long
    Dim result As String
    
    For i = 1 To Len(text)
        result = result & Right("0" & Hex(Asc(Mid(text, i, 1))), 2)
    Next i
    
    String_To_Hex = result
End Function

Public Function Hex_To_Byte(ByVal str As String) As Byte()
    Dim i As Long
    Dim s As String
    
    Dim output() As Byte
    
    s = Replace(str, " ", "")
   
    ReDim output((Len(s) \ 2) - 1)
    
    For i = 0 To UBound(output)
        ' See: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
        output(i) = CByte(val("&H" & Mid(s, i * 2 + 1, 2)))
    Next i
    Hex_To_Byte = output
End Function

Public Function Byte_To_Hex(ByRef bytes() As Byte) As String
    Dim i As Long
    Dim output As String
    
    ' Allocate space
    output = Space$(UBound(bytes) * 2 + 2)
    
    For i = LBound(bytes) To UBound(bytes)
        Mid$(output, (i * 2) + 1, 2) = Right$("0" & Hex$(bytes(i)), 2)
    Next i
    Byte_To_Hex = Trim(output)
End Function

Public Sub Suspend(ByVal status As Boolean)
    With Application
        .ScreenUpdating = Not status
        .Calculation = IIf(Start, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not status
    End With
End Sub

Public Function Reverse(ByRef arr() As Variant) As Variant
  Dim i As Long
  Dim j As Long

  Dim temp As Variant

  Do While i < j
    Set temp = arr(i)
    Set arr(i) = arr(j)
    Set arr(j) = temp

    i = i + 1
    j = j + 1
  Loop

  Reverse = arr
End Function
