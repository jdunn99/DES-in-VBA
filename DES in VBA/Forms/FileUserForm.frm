VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileUserForm 
   Caption         =   "File Upload"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "FileUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public path As String
Private fileBytes() As Byte

' Load the bytes into memory
Private Sub btnChooseFile_Click()
    path = Load_File
    
    fileName = Mid(path, InStrRev(path, "\") + 1)
    Me.labelFile.Caption = fileName
End Sub

Public Sub btnSubmit_Click()
    Dim key As String, iv As String
    Dim isEncrypt As Boolean
    Dim keyBytes() As Byte
    Dim subkeys() As Byte
    
    key = Me.txtKey.value
    If Not ValidateKey(key) Then Exit Sub
    
    iv = Me.txtIV.value
    If iv = "" Then iv = DEFAULT_IV
    isEncrypt = Me.optEncrypt.value
    
    If path = "" Then
        MsgBox "Please select a source file first."
        Exit Sub
    End If
    
    keyBytes = Hex_To_Byte(key)
    subkeys = Parse_Subkeys(keyBytes, isEncrypt)
    
    Process_File path, keyBytes, iv, subkeys, isEncrypt
    Unload Me
End Sub
