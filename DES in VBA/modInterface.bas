Attribute VB_Name = "modInterface"
Option Explicit

Private Const DEFAULT_IV As String = "0000000000000000"

Private Function Get_Shape_Text(name As String) As String
    Get_Shape_Text = ActiveSheet.Shapes(name).TextFrame2.TextRange.Characters.text
End Function

Private Sub Set_Shape_Text(name As String, value As String)
    ActiveSheet.Shapes(name).TextFrame2.TextRange.Characters.text = value
End Sub

Private Function ValidateKey(key As String) As Boolean
    If Len(key) <> 16 Then
        MsgBox "Key must be 16 characters"
        ValidateKey = False
    Else
        ValidateKey = True
    End If
End Function

Sub Run_DES()
    Dim plaintext As String
    Dim key As String
    Dim iv As String
    Dim subkeys() As BitBuffer
    
    plaintext = Get_Shape_Text("InputBox")
    key = Get_Shape_Text("KeyBox")
    iv = Get_Shape_Text("IVBox")
    
    If iv = "" Then iv = DEFAULT_IV
    If Not ValidateKey(key) Then Exit Sub
    
    subkeys = Generate_Subkeys(key)
    
    Set_Shape_Text "OutputBox", Encrypt(plaintext, key, subkeys, iv)

End Sub

Sub Run_Decrypt()
    Dim ciphertext As String
    Dim key As String
    Dim iv As String
    
    ciphertext = Get_Shape_Text("InputBox")
    key = Get_Shape_Text("KeyBox")
    iv = Get_Shape_Text("IVBox")
    
    If iv = "" Then iv = DEFAULT_IV
    If Not ValidateKey(key) Then Exit Sub
    
    Set_Shape_Text "OutputBox", DES_Decrypt(ciphertext, key, iv)
End Sub

'TODO: Add a better Form
Public Sub DES_RANGE()
    Dim cell As Range
    Dim key As String
    Dim iv As String
    Dim processedValue As String
    Dim subkeys() As BitBuffer
    
    Dim mode As String
    mode = "DECRYPT"
    
    key = InputBox("Enter 16-character Key:", "DES " & mode)
    If Len(key) <> 16 Then
        MsgBox "Key must be 16 characters!"
        Exit Sub
    End If
    
    iv = InputBox("Enter IV (leave blank for default):", "Initialization Vector")
    If iv = "" Then iv = String(16, "0")
    
    If UCase(mode) = "ENCRYPT" Then
        subkeys = Generate_Subkeys(key)
    End If
    
    For Each cell In Selection
        If Not IsEmpty(cell) Then
            Select Case UCase(mode)
                Case "ENCRYPT"
                    processedValue = Encrypt(cell.value, key, subkeys, iv)
                Case "DECRYPT"
                    processedValue = DES_Decrypt(cell.value, key, iv)
                Case Else
                    MsgBox "Invalid mode specified.", vbCritical
                    Exit Sub
            End Select
            
            cell.value = processedValue
        End If
    Next cell
    
    MsgBox "Operation Complete", vbInformation
End Sub

' TODO: Add interface for encryption/decryption and a file picker
Public Sub DES_FILE()
    Dim path As String
    path = "C:\Users\jack\Desktop\test.txt"
    Dim key As String
    Dim iv As String
    Dim subkeys() As BitBuffer
    Dim result As String
    Dim fileHex As String
   
    Dim bytes() As Byte

    bytes = Read_File(path)
    fileHex = Byte_To_Hex(bytes)
    
    key = Get_Shape_Text("KeyBox")
    iv = Get_Shape_Text("IVBox")
    
    If iv = "" Then iv = DEFAULT_IV
    If Not ValidateKey(key) Then Exit Sub
    
    subkeys = Generate_Subkeys(key)
    'result = DES_Decrypt(fileHex, key, iv)
    result = Encrypt(fileHex, key, subkeys, iv)

    Write_File path, result
End Sub

