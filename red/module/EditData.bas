Attribute VB_Name = "EditData"
Option Explicit

'シート名
Public Const strSheetName As String = "データ編集"

'表示開始位置
Private Const cellList As String = "A1"

'種類セル位置
Public cellType As String

'シート
Private thisSheet As Worksheet

'シート作成
Public Sub CreateSheet(ByVal ID As String)
    
    On Error GoTo ErrHandle

    Dim Sheet As Worksheet
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim maxRow As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strSql As String
    
    For Each Sheet In Worksheets
        If Sheet.Name = strSheetName Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Template.Visible = xlSheetVisible
    Template.Copy Sheets(1)
    Set thisSheet = ActiveSheet
    Template.Visible = xlSheetHidden
    thisSheet.Name = strSheetName
    
    Windows.Application.ScreenUpdating = False
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    strSql = "SELECT * FROM EditData WHERE [基本:ID] = '" & ID & "'"
    
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
    If rsBase.EOF Then
        MsgBox "対象データがありません。", vbOKOnly, "情報"
    Else
        maxRow = rsBase.Fields.Count
        For intRow = 0 To maxRow - 1
            If UBound(Split(rsBase.Fields(intRow).Name, ":")) > 0 Then
                Range(cellList).Offset(intRow, 0).Value = Split(rsBase.Fields(intRow).Name, ":")(0)
                Range(cellList).Offset(intRow, 1).Value = Split(rsBase.Fields(intRow).Name, ":")(1)
                If rsBase.Fields(intRow).Name = "基本:種類" Then
                    cellType = Range(cellList).Offset(intRow, 2).Address
                    Range(cellType).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ThisWorkbook.getTypeList
                End If
            Else
                Range(cellList).Offset(intRow, 1).Value = rsBase.Fields(intRow).Name
            End If
            Range(cellList).Offset(intRow, 2).NumberFormat = "@"
            Range(cellList).Offset(intRow, 2).Value = rsBase.Fields(intRow).Value
        Next
    End If
    rsBase.Close

    conQCSDB.Close
    Set rsBase = Nothing
    Set conQCSDB = Nothing

    ActiveWindow.Zoom = 90
    With Range(cellList).Offset(0, 0).Resize(maxRow, 2)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.ThemeColor = xlThemeColorAccent3
        .Font.Size = 9
        .EntireColumn.AutoFit
    End With
    With Range(cellList).Offset(0, 2).Resize(maxRow, 1)
        .Borders.LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Font.Size = 9
        .EntireColumn.AutoFit
    End With
    With Range(cellList).Offset(maxRow, 0).Resize(maxRow, 3)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    
    hiddenCell Range(cellType).Value
   
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True
    
    Exit Sub

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "システムエラー"

    If Not rsBase Is Nothing Then
        If rsBase.State <> adStateClosed Then
            rsBase.Close
        End If
        Set rsBase = Nothing
    End If

    If Not conQCSDB Is Nothing Then
        If conQCSDB.State = adStateOpen Then
            conQCSDB.Close
        End If
        If conQCSDB.State <> adStateClosed Then
            conQCSDB.RollbackTrans
            conQCSDB.Close
        End If
        Set conQCSDB = Nothing
    End If

    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

End Sub

'セル非表示
Public Sub hiddenCell(ByVal Value As String)
    Dim pattern As New Collection
    Dim intRow As Integer
    Windows.Application.ScreenUpdating = False
    
    pattern.Add "|HW|", "VM"
    
    For intRow = Range(cellList).row To Range(cellList).End(xlDown).row
        If ThisWorkbook.isExists(pattern, Value) Then
            If InStr(pattern(Value), "|" & Cells(intRow, Range(cellList).Column).Value & "|") > 0 Then
                Rows(intRow).EntireRow.Hidden = True
            Else
                Rows(intRow).EntireRow.Hidden = False
            End If
        Else
            Rows(intRow).EntireRow.Hidden = False
        End If
    Next intRow
    Windows.Application.ScreenUpdating = True
End Sub

Public Sub test()
    CreateSheet "00000000000195"
End Sub

