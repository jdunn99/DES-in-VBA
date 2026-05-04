VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RangeUserForm 
   Caption         =   "Encrypt Range"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "RangeUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "rangeUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSubmit_Click()
    Dim key As String
    Dim iv As String
    Dim rangeValue As String
    Dim isEncrypt As Boolean
    
    Dim result() As Variant
    Dim keyBytes() As Byte
    Dim data() As Byte
    Dim subkeys() As Byte
    
    key = Me.txtKey.value
    If Not ValidateKey(key) Then Exit Sub
    
    iv = Me.txtIV.value
    If iv = "" Then iv = DEFAULT_IV

    isEncrypt = Me.optEncrypt.value
    rangeValue = Me.txtRange.value
    
    Range(rangeValue).Select
    
    keyBytes = Hex_To_Byte(key)
    subkeys = Parse_Subkeys(keyBytes, isEncrypt)
    DES_Range keyBytes, iv, subkeys, isEncrypt

    MsgBox "Operation Complete", vbInformation
    Unload Me
End Sub


Private Sub txtRange_AfterUpdate()
    Range(Me.txtRange.value).Select
End Sub

Private Sub UserForm_Initialize()
    ' Set default mode
    Me.optEncrypt.value = True
    
    ' Automatically grab the highlighted range from the sheet
    If TypeName(Selection) = "Range" Then
        Me.txtRange.text = Selection.Address(False, False)
    End If
End Sub
