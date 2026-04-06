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

