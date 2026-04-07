Attribute VB_Name = "modFile_Handler"

Public Function Read_File(ByVal path As String) As Byte()
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim bytes() As Byte
    Open path For Binary Access Read As #fileNum
        If LOF(fileNum) = 0 Then
            Close #fileNum
            Exit Function
        End If
        ReDim bytes(0 To LOF(fileNum) - 1)
        Get #fileNum, 1, bytes
    Close #fileNum
    
    Read_File = bytes
End Function


Public Sub Write_File(ByVal path As String, ByVal data As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    If Dir(path) <> "" Then Kill path
    Open path For Binary Access Write As #fileNum
        Dim output() As Byte
        output = Hex_To_Byte(data)
        Put #fileNum, , output
    Close #fileNum
End Sub

