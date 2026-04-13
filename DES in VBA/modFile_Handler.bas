Attribute VB_Name = "modFile_Handler"
Option Explicit

Public FileBuffer As New Collection
Private fileBytes() As Byte
Private originalPath As String

Public Function Read_File() As Byte()
    Dim fileNum As Integer
    fileNum = FreeFile

    Dim fDialog As fileDialog
    Set fDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    Dim path As String
    Dim fileName As String
    Dim table As ListObject
    Dim newRow As ListRow
    
    With fDialog
        .Title = "Select a file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then
            path = .SelectedItems(1)
            fileName = Mid(path, InStrRev(path, "\") + 1)
            
            Set table = ThisWorkbook.Sheets("Sheet2").ListObjects("FileTable")
            Set newRow = table.ListRows.Add
            
            With newRow
                .Range(1) = fileName
                .Range(2) = path
                .Range(3) = "Processing"
                .Range(4) = "SAVE"
            End With
            
            table.Range.Columns.AutoFit
        Else
            ' Throw an Error
        End If
    End With
    
    Dim bytes() As Byte
    Open path For Binary Access Read As #fileNum
        If LOF(fileNum) = 0 Then
            Close #fileNum
            Exit Function
        End If
        ReDim bytes(0 To LOF(fileNum) - 1)
        Get #fileNum, 1, bytes
    Close #fileNum
    
    
    ' Set the input to the file path
    Set_Shape_Text "FilePathBox", path

    fileBytes = bytes
    originalPath = path
    
    Debug.Print path
    FileBuffer.Add bytes, path
    
    Dim key As String
    Dim iv As String
    Dim subkeys() As BitBuffer
    Dim result As String
    Dim fileHex As String
    
    key = "133457799BBCDFF1"
    iv = ""
  
    If iv = "" Then iv = DEFAULT_IV
    If Not ValidateKey(key) Then Exit Function
    
    subkeys = Parse_Subkeys(key, True)
    fileHex = Byte_To_Hex(bytes)
    
    Debug.Print fileHex

    result = DES_Bytes(bytes, key, iv, subkeys, True)

    
    Update_Status path
    
    Debug.Print "File successfully encrypted"
    
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

Public Sub Save_File(ByVal path As String)
    Dim fileNum As Integer
    Dim fDialog As fileDialog

    Dim bytes() As Byte
    bytes = FileBuffer(path)
    
    If (Not fileBytes) = -1 Then
        MsgBox "No bytes in memory", vbCritical
        Exit Sub
    End If
    
    Dim savePath As Variant ' Must be Variant to handle "Cancel"
    Dim fileFilter As String
    
    ' Define your filters: "Description, *.extension"
    fileFilter = "Text Files (*.txt), *.txt, Encrypted Files (*.enc), *.enc, All Files (*.*), *.*"
    
    ' This opens the TRUE Save As dialog
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="test2.txt", _
        fileFilter:=fileFilter, _
        Title:="Save Processed File")
        
End Sub

Public Sub Update_Status(ByVal name As String)
    Dim table As ListObject
    Dim cell As Range
    
    Set table = ThisWorkbook.Sheets("Sheet2").ListObjects("FileTable")
    Set cell = table.ListColumns("Path").DataBodyRange.Find(What:=name, LookAt:=xlWhole)
    If Not cell Is Nothing Then
        cell.Offset(0, 1).value = "Complete"
        DoEvents
    End If
    
End Sub

