VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal TARGET As Range, Cancel As Boolean)
    If TARGET.Column = 1 And TARGET.row > 1 And ActiveSheet.Name = Summary.strSheetName Then
        If Trim(TARGET.Value) <> "" Then
            Cancel = False
            EditData.CreateSheet Trim(TARGET.Value)
        End If
    End If
End Sub

Private Sub Worksheet_Change(ByVal TARGET As Range)
    If TARGET.Address = EditData.cellType And ActiveSheet.Name = EditData.strSheetName Then
        EditData.hiddenCell TARGET.Value
    End If
End Sub

