Attribute VB_Name = "modInterface"
Option Explicit



Public Function Get_Shape_Text(name As String) As String
    Get_Shape_Text = ActiveSheet.Shapes(name).TextFrame2.TextRange.Characters.text
End Function

Public Sub Set_Shape_Text(name As String, value As String)
    ActiveSheet.Shapes(name).TextFrame2.TextRange.Characters.text = value
End Sub

Public Function ValidateKey(key As String) As Boolean
    If Len(key) <> 16 Then
        MsgBox "Key must be 16 characters"
        ValidateKey = False
    Else
        ValidateKey = True
    End If
End Function

Public Sub Encrypt_Button()
    DES_Calculator (True)
End Sub

Public Sub Decrypt_Button()
    DES_Calculator (False)
End Sub

Public Sub DES_Calculator(ByVal isEncrypt As Boolean)
  Dim plaintext As String
  Dim key As String
  Dim iv As String
  Dim subkeys() As BitBuffer

  plaintext = Get_Shape_Text("InputBox")
  key = Get_Shape_Text("KeyBox")
  iv = Get_Shape_Text("IVBox")
  
  If iv = "" Then iv = DEFAULT_IV
  If Not ValidateKey(key) Then Exit Sub
  
  subkeys = Parse_Subkeys(key, isEncrypt)

  Set_Shape_Text "OutputBox", DES_Core(plaintext, key, iv, subkeys, isEncrypt)
End Sub

'TODO: Add a better Form
Public Function DES_RANGE(ByVal data As Variant, ByVal key As String, ByVal isEncrypt As Boolean, Optional ByVal iv As String = DEFAULT_IV) As Variant
  Debug.Print "Running DES RANGE"
  Dim subkeys() As BitBuffer
  Dim result() As Variant
  Dim i As Long
  Dim j As Long
  Dim strData As String
  
  Dim temp As Variant
  temp = data
  
  If Not ValidateKey(key) Then Exit Function
  If Not IsArray(data) Then
      ' Process the single value directly
      Exit Function
  End If
  
  subkeys = Parse_Subkeys(key, isEncrypt)
  
  ReDim result(LBound(temp, 1) To UBound(temp, 1), LBound(temp, 2) To UBound(temp, 2))

For i = LBound(temp, 1) To UBound(temp, 1)
    For j = LBound(temp, 2) To UBound(temp, 2)
      If Not IsEmpty(temp(i, j)) Then
        result(i, j) = DES_Core(CStr(temp(i, j)), key, iv, subkeys, isEncrypt)
      End If
    Next j
  Next i
  
  DES_RANGE = result
End Function

' TODO: Add interface for encryption/decryption and a file picker
Public Sub DES_FILE()
    Read_File
    

    
    'subkeys = Generate_Subkeys(key)
    'result = DES_Decrypt(fileHex, key, iv)
    
    'result = Encrypt(fileHex, key, subkeys, iv)

    'Write_File path, result
End Sub

