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
  Dim resultStr As String
  Dim key As String
  Dim iv As String
  
  Dim result() As Byte
  Dim data() As Byte
  Dim keyBytes() As Byte
  Dim subkeys() As Byte

  plaintext = Get_Shape_Text("InputBox")
  key = Get_Shape_Text("KeyBox")
  iv = Get_Shape_Text("IVBox")
  
  If iv = "" Then iv = DEFAULT_IV
  If Not ValidateKey(key) Then Exit Sub
  
  keyBytes = Hex_To_Byte(key)
  subkeys = Parse_Subkeys(keyBytes, isEncrypt)
  resultStr = Run_DES(plaintext, subkeys, iv, isEncrypt)
  
  Set_Shape_Text "OutputBox", resultStr
End Sub


Public Sub DES_Range(ByRef keyBytes() As Byte, ByVal hexIV As String, ByRef subkeys() As Byte, ByVal isEncrypt As Boolean)
    Dim plaintext As String
    Dim cell As Range
    
    For Each cell In Selection
        If Not IsEmpty(cell.value) Then
            plaintext = CStr(cell.value)
            cell.value = Run_DES(plaintext, subkeys, hexIV, isEncrypt)
        End If
    Next cell
End Sub
