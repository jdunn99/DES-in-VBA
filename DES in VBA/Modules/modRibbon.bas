Attribute VB_Name = "modRibbon"
Public Sub Load_Range_UserForm()
    Load rangeUserForm
    rangeUserForm.Show
End Sub

Public Sub Ribbon_Decrypt(control As IRibbonControl)
    Load_Range_UserForm
End Sub

Public Sub Ribbon_Encrypt(control As IRibbonControl)
    Load_Range_UserForm
End Sub


