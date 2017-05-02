VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SafeArrayGetDim Lib "oleaut32" (ByRef psa() As Any) As Long

'対象シート
Private sheetMainte As Worksheet
'検索条件入力開始位置
Private cellKensaku As String
'一覧表示開始位置
Private cellList As String
'参照テーブル
Private nameTable As String
'表示列数
Private maxCol As Integer
'表示タイプ(S=文字、N=数値)
Private typeDisp As String
'カラム名
Private nameCols As String
'キー項目フラグ(K=キー、N=キー以外)
Private keyCol As String

'セッター
Public Property Set Sheet(MaintenanceSheet As Worksheet)
    Set sheetMainte = MaintenanceSheet
End Property
Public Property Let KensakuCell(Cell As String)
    cellKensaku = Cell
End Property
Public Property Let ListCell(Cell As String)
    cellList = Cell
End Property
Public Property Let TableName(Name As String)
    nameTable = Name
End Property
Public Property Let ColumnMax(Max As Integer)
    maxCol = Max
End Property
Public Property Let DisplayTypes(DType As String)
    typeDisp = DType
End Property
Public Property Let ColumnNames(Names As String)
    nameCols = Names
End Property
Public Property Let KeyColumns(Keys As String)
    keyCol = Keys
End Property

'指定位置表示タイプ
Private Function DisplayType(index As Integer) As String
    DisplayType = Split(typeDisp, ",")(index)
End Function

'指定位置カラム名
Private Function ColumnName(index As Integer) As String
    ColumnName = Split(nameCols, ",")(index)
End Function

'指定位置キー項目フラグ
Private Function IsKeyColmun(index As Integer) As Boolean
    IsKeyColmun = (Split(keyCol, ",")(index) = "K")
End Function

'一覧クリア
Public Sub ClearList()
    sheetMainte.Activate
    sheetMainte.Rows(sheetMainte.Range(cellList).row & ":" & sheetMainte.Range(cellList).End(xlDown).row).Delete
    sheetMainte.Range(cellList).Offset(0, 0).Activate
End Sub

'一覧表示
Public Sub LoadList()
    
    On Error GoTo ErrHandle

    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strSql As String
    Dim intBool As Integer
    Dim strType(2) As String
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    ClearList
    
    sheetMainte.Activate

    intBool = 0
    strType(0) = "WHERE "
    strType(1) = "AND "
    
    strSql = "SELECT * FROM " & nameTable & " "
    For intCol = 0 To maxCol - 1
        If Trim(Range(cellKensaku).Offset(0, intCol).Value) <> "" Then
            If DisplayType(intCol) = "S" Then
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " LIKE '%" & Trim(Range(cellKensaku).Offset(0, intCol).Value) & "%' "
                intBool = 1
            Else
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " = " & Trim(Range(cellKensaku).Offset(0, intCol).Value) & " "
                intBool = 1
            End If
        End If
    Next
    
    intBool = 0
    strType(0) = "ORDER BY "
    strType(1) = ", "
    
    For intCol = 0 To maxCol - 1
        If IsKeyColmun(intCol) Then
            strSql = strSql & strType(intBool) & ColumnName(intCol) & " "
            intBool = 1
        End If
    Next
    
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
    If rsBase.EOF Then
        MsgBox "対象レコードがありません。", vbOKOnly, "情報"
    Else
        intRow = 0
        Do Until rsBase.EOF
            For intCol = 0 To maxCol - 1
                If DisplayType(intCol) = "S" Then
                    Range(cellList).Offset(intRow, intCol).NumberFormat = "@"
                End If
                Range(cellList).Offset(intRow, intCol).Value = Trim(rsBase.Fields(intCol).Value)
            Next
            intRow = intRow + 1
            rsBase.MoveNext
        Loop

        With Range(cellList).Offset(0, 0).Resize(intRow, maxCol)
            .Borders.LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlHairline
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
    End If
    rsBase.Close

    conQCSDB.Close
    Set rsBase = Nothing
    Set conQCSDB = Nothing

    Range(cellList).Offset(0, 0).Activate
    
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

'修正個所チェック
Public Function CheckList(conQCSDB As ADODB.Connection) As Integer()

    On Error GoTo ErrHandle

    sheetMainte.Activate
    
    Dim rsBase As New ADODB.Recordset
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strSql As String
    Dim intBool As Integer
    Dim aryRow() As Integer
    Dim strType(2) As String
    
    strType(0) = "WHERE "
    strType(1) = "AND "
    
    For intRow = Range(cellList).row To Range(cellList).End(xlDown).row
        
        intBool = 0
        strSql = "SELECT * FROM " & nameTable & " "
        For intCol = 0 To maxCol - 1
            If IsKeyColmun(intCol) Then
                If DisplayType(intCol) = "S" Then
                    strSql = strSql & strType(intBool) & ColumnName(intCol) & " = '" & Trim(Cells(intRow, intCol + 1).Value) & "' "
                    intBool = 1
                Else
                    strSql = strSql & strType(intBool) & ColumnName(intCol) & " = " & Trim(Cells(intRow, intCol + 1).Value) & " "
                    intBool = 1
                End If
            End If
        Next
        
'        Debug.Print strSQL
        Set rsBase = conQCSDB.Execute(strSql, , adCmdText)
        If rsBase.EOF Then
            Range(Cells(intRow, 1), Cells(intRow, maxCol)).Interior.ThemeColor = xlThemeColorAccent6
            If SafeArrayGetDim(aryRow) = 0 Then
                ReDim aryRow(0)
            Else
                ReDim Preserve aryRow(UBound(aryRow) + 1)
            End If
            aryRow(UBound(aryRow)) = intRow
        Else
            intBool = 0
            For intCol = 0 To maxCol - 1
                If IsNull(rsBase.Fields(intCol).Value) Then
                    If Cells(intRow, intCol + 1).Value = "" Then
                        Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent3
                    Else
                        Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent6
                        intBool = 1
                    End If
                Else
                    If DisplayType(intCol) = "S" Then
                        If Cells(intRow, intCol + 1).Value = Trim(rsBase.Fields(intCol).Value) Then
                            Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent3
                        Else
                            Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent6
                            intBool = 1
                        End If
                    Else
                        If Cells(intRow, intCol + 1).Value = rsBase.Fields(intCol).Value Then
                            Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent3
                        Else
                            Cells(intRow, intCol + 1).Interior.ThemeColor = xlThemeColorAccent6
                            intBool = 1
                        End If
                    End If
                End If
                If intBool = 1 Then
                    If SafeArrayGetDim(aryRow) = 0 Then
                        ReDim aryRow(0)
                    Else
                        ReDim Preserve aryRow(UBound(aryRow) + 1)
                    End If
                    aryRow(UBound(aryRow)) = intRow
                End If
            Next
        End If
        rsBase.Close
        Set rsBase = Nothing
        
    Next
    
    CheckList = aryRow
        
    Exit Function

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "システムエラー"

    If Not rsBase Is Nothing Then
        If rsBase.State <> adStateClosed Then
            rsBase.Close
        End If
        Set rsBase = Nothing
    End If

End Function

'更新
Public Sub UpdateTable()

    On Error GoTo ErrHandle

    sheetMainte.Activate
    
    Dim conQCSDB As New ADODB.Connection
    Dim aryRow() As Integer
    Dim aryCol() As Integer
    Dim boolAdd As Boolean
    Dim boolResult As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    Dim cntI, cntJ As Integer
    
    For intCol = 0 To maxCol - 1
        If IsKeyColmun(intCol) Then
            If SafeArrayGetDim(aryCol) = 0 Then
                ReDim aryCol(0)
            Else
                ReDim Preserve aryCol(UBound(aryCol) + 1)
            End If
            aryCol(UBound(aryCol)) = intCol
        End If
    Next

    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual

    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open

    aryRow = CheckList(conQCSDB)
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True
    If SafeArrayGetDim(aryRow) <> 0 Then
        If MsgBox("更新箇所を確認してください。" & vbCrLf & "この内容でDBを更新しますか。", vbYesNo, "情報") = vbYes Then
            Windows.Application.ScreenUpdating = False
            Windows.Application.Calculation = xlCalculationAutomatic
            Windows.Application.Calculation = xlCalculationManual
            conQCSDB.BeginTrans
            For cntI = LBound(aryRow) To UBound(aryRow)
                intRow = aryRow(cntI)
                boolAdd = True
                For cntJ = LBound(aryCol) To UBound(aryCol)
                    intCol = aryCol(cntJ)
                    If Cells(intRow, intCol + 1).Interior.ThemeColor <> xlThemeColorAccent6 Then
                        boolAdd = False
                        Exit For
                    End If
                Next
                boolResult = False
                If boolAdd Then
                    boolResult = InsertRecord(conQCSDB, intRow)
                Else
                    boolResult = UpdateRecord(conQCSDB, intRow)
                End If
                If Not boolResult Then
                    Exit For
                End If
            Next
            If boolResult Then
                conQCSDB.CommitTrans
                MsgBox "DBを更新しました。", vbOKOnly, "情報"
            Else
                conQCSDB.RollbackTrans
            End If
        End If
    Else
        MsgBox "変更箇所がありません。", vbOKOnly, "情報"
    End If

    conQCSDB.Close
    Set conQCSDB = Nothing

    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

    Exit Sub

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "システムエラー"

    If Not conQCSDB Is Nothing Then
        If conQCSDB.State <> adStateClosed Then
            conQCSDB.RollbackTrans
            conQCSDB.Close
        End If
        Set conQCSDB = Nothing
    End If

    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

End Sub

'レコード挿入
Private Function InsertRecord(conQCSDB As ADODB.Connection, row As Integer) As Boolean

    On Error GoTo ErrHandle

    sheetMainte.Activate
    
    Dim rsBase As New ADODB.Recordset
    Dim strSql As String
    Dim dtNow As Date
    Dim intCol As Integer
    Dim strType(2) As String
    Dim intBool As Integer
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    InsertRecord = False

    strSql = "SELECT * FROM " & nameTable & " "
    For intCol = 0 To maxCol - 1
        If IsKeyColmun(intCol) Then
            If DisplayType(intCol) = "S" Then
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " = '" & Trim(Cells(row, intCol + 1).Value) & "' "
                intBool = 1
            Else
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " = " & Trim(Cells(row, intCol + 1).Value) & " "
                intBool = 1
            End If
        End If
    Next

    dtNow = Now

    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockOptimistic, adCmdText
    If rsBase.EOF Then
        rsBase.AddNew
        For intCol = 0 To maxCol - 1
            If DisplayType(intCol) = "S" Then
                rsBase.Fields(intCol).Value = Trim(Cells(row, intCol + 1).Value)
            Else
                rsBase.Fields(intCol).Value = Cells(row, intCol + 1).Value
            End If
        Next
        rsBase.Update
    Else
        MsgBox "更新チェック後にレコードが挿入されました" & vbCrLf & "レコード挿入を見送ります。", vbOKOnly, "情報"
        Exit Function
    End If

    rsBase.Close
    Set rsBase = Nothing

    InsertRecord = True

    Exit Function

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "システムエラー"

    If Not rsBase Is Nothing Then
        If rsBase.State <> adStateClosed Then
            rsBase.Close
        End If
        Set rsBase = Nothing
    End If

End Function

'レコード更新
Private Function UpdateRecord(conQCSDB As ADODB.Connection, row As Integer) As Boolean

    On Error GoTo ErrHandle

    sheetMainte.Activate
    
    Dim rsBase As New ADODB.Recordset
    Dim strSql As String
    Dim dtNow As Date
    Dim intCol As Integer
    Dim strType(2) As String
    Dim intBool As Integer
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    UpdateRecord = False

    strSql = "SELECT * FROM " & nameTable & " "
    For intCol = 0 To maxCol - 1
        If IsKeyColmun(intCol) Then
            If DisplayType(intCol) = "S" Then
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " = '" & Trim(Cells(row, intCol + 1).Value) & "' "
                intBool = 1
            Else
                strSql = strSql & strType(intBool) & ColumnName(intCol) & " = " & Trim(Cells(row, intCol + 1).Value) & " "
                intBool = 1
            End If
        End If
    Next

    dtNow = Now

    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockOptimistic, adCmdText
    If rsBase.EOF Then
        MsgBox "更新チェック後にレコードが削除されました" & vbCrLf & "レコード更新を見送ります。", vbOKOnly, "情報"
        Exit Function
    Else
        For intCol = 0 To maxCol - 1
            If Not IsKeyColmun(intCol) Then
                If DisplayType(intCol) = "S" Then
                    rsBase.Fields(intCol).Value = Trim(Cells(row, intCol + 1).Value)
                Else
                    rsBase.Fields(intCol).Value = Cells(row, intCol + 1).Value
                End If
            End If
        Next
        rsBase.Update
    End If

    rsBase.Close
    Set rsBase = Nothing

    UpdateRecord = True

    Exit Function

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "システムエラー"

    If Not rsBase Is Nothing Then
        If rsBase.State <> adStateClosed Then
            rsBase.Close
        End If
        Set rsBase = Nothing
    End If

End Function


