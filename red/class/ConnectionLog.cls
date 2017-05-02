VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Ú‘±—š—ð
Public Sub ConnectLog()
    Dim intCnt, intMax As Integer
    Dim intAry() As Variant
    Application.CutCopyMode = False
    With ConnectionLog
        .Rows(2).Insert
        .Range("A2").Value = "COPY"
        ConnectionInfo.Range("E2:E21").Copy
        .Range("B2").PasteSpecial , , , True
        intMax = .Range("A1").End(xlToRight).Column - .Range("A1").Column + 1
        ReDim intAry(intMax - 1)
        For intCnt = 1 To intMax
            intAry(intCnt - 1) = intCnt
        Next
        .Range("A1").Resize(.Range("A1").End(xlDown).Row, .Range("A1").End(xlToRight).Column).RemoveDuplicates intAry, xlYes
    End With
End Sub

'î•ñ•¡»
Private Sub CopyInfo(ByVal TARGET As Range)
    Application.CutCopyMode = False
    TARGET.Offset(0, 1).Resize(, ConnectionLog.Range("A1").End(xlToRight).Column).Copy
    ConnectionInfo.Range("E2").PasteSpecial , , , True
End Sub

'ƒCƒxƒ“ƒg
Private Sub Worksheet_BeforeDoubleClick(ByVal TARGET As Range, Cancel As Boolean)
    If TARGET.Row > 1 And TARGET.Column = 1 And TARGET.Value = "COPY" Then
        CopyInfo TARGET
    End If
End Sub

