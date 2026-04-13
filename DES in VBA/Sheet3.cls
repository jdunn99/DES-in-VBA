VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim table As ListObject
    Dim col As ListColumn
    
    If Target.Cells.count > 1 Then Exit Sub
    
    Set table = Me.ListObjects("FileTable")
    If table Is Nothing Then Exit Sub
    
    Set col = table.ListColumns("Action")
    
    If Not Intersect(Target, col.DataBodyRange) Is Nothing Then
        Dim status As String
        Dim path As String
        status = Intersect(Target.EntireRow, table.ListColumns("Status").DataBodyRange).value
        If status <> "Complete" Then Exit Sub
        
        path = Intersect(Target.EntireRow, table.ListColumns("Path").DataBodyRange).value
        If path = "" Then Exit Sub
        
        Save_File (path)
    End If
End Sub

