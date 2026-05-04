Attribute VB_Name = "modFile_Handler"
Option Explicit

Public FileBuffer As New Collection
Private fileBytes() As Byte
Private originalPath As String

Public Function Load_File() As String
    Dim path As String
    Dim fDialog As FileDialog
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fDialog
        .Title = "Select File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then path = .SelectedItems(1)
    End With
    
    Load_File = path
End Function

Public Sub Process_File(ByVal path As String, ByRef keyBytes() As Byte, ByVal hexIV As String, ByRef subkeys() As Byte, ByVal isEncrypt As Boolean)
    Dim fIn As Integer
    Dim fOut As Integer
    Dim savePath As Variant
    Dim buffer() As Byte
    Dim processed() As Byte
    
    Dim totalBytes As Long
    Dim chunkSize As Long
    Dim processedBytes As Long
    
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="processed_file", _
        fileFilter:="All Files (*.*), *.*", _
        Title:="Save Processed File")
    
    If savePath = "False" Then Exit Sub
    If Dir(savePath) <> "" Then Kill savePath
    
    fIn = FreeFile
    Open path For Binary Access Read As #fIn
    totalBytes = LOF(fIn)
    
    fOut = FreeFile
    Open savePath For Binary Access Write As #fOut
    
    chunkSize = 65536 ' 64KB
    processedBytes = 0
    
    Do While processedBytes < totalBytes
       
        If totalBytes - processedBytes < chunkSize Then
            chunkSize = totalBytes - processedBytes
            ReDim buffer(0 To chunkSize - 1)
        End If
        
        
        Get #fIn, , buffer
        
        processed = DES_Bytes(buffer, subkeys, hexIV, True)
        Put #fOut, , processed
        
        processedBytes = processedBytes + chunkSize
    Loop
    
    Close #fIn
    Close #fOut
    MsgBox "File " & IIf(isEncrypt, "Encrypted", "Decrypted") & " successfully!", vbInformation
    
End Sub
