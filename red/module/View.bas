Attribute VB_Name = "View"
'一覧表示開始位置
Private Const cellList As String = "A1"

'接続情報取得(オブジェクト)
Public Function getConObj() As DatabaseInterface

    Set getConObj = New DatabaseInterfaceImpl
    getConObj.Provider = "SQLOLEDB"
    getConObj.DataSource = "172.127.24.116,2025"
    getConObj.InitialCatalog = "GXKDB"
    getConObj.UserID = "gxkiban"
    getConObj.Password = "gxkiban"

End Function

'接続情報取得(文字)
Public Function getConInfo() As String

    getConInfo = getConObj.GetConnectString

End Function

'テーブル読み込み
Public Sub loadTable(ByRef shtTarget As Worksheet, ByRef strSql As String, Optional ByRef maxCol As Integer = 0, Optional ByRef nameFont As String = "")
    
    On Error GoTo ErrHandle

    shtTarget.Activate
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim intCol As Integer
    Dim intRow As Integer
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
    If rsBase.EOF Then
        MsgBox "対象レコードがありません。", vbOKOnly, "情報"
    Else
        maxCol = rsBase.Fields.Count
        
        intRow = 0
        For intCol = 0 To maxCol - 1
            Range(cellList).Offset(intRow, intCol).Value = rsBase.Fields(intCol).Name
        Next
        
        intRow = 1
        Do Until rsBase.EOF
            For intCol = 0 To maxCol - 1
                If Left(Trim(rsBase.Fields(intCol).Value), 1) <> "=" Then
                    Range(cellList).Offset(intRow, intCol).NumberFormat = "@"
                End If
                Range(cellList).Offset(intRow, intCol).Value = Trim(rsBase.Fields(intCol).Value)
            Next
            intRow = intRow + 1
            rsBase.MoveNext
        Loop
    End If
    rsBase.Close

    conQCSDB.Close
    Set rsBase = Nothing
    Set conQCSDB = Nothing

    If nameFont <> "" Then
        Range("A1", Cells(Range("A1").End(xlDown).row, Range("A1").End(xlToRight).Column)).Select
        Selection.Font.Name = nameFont
    End If
    
    ActiveWindow.Zoom = 90
    Rows(1).RowHeight = 50
    With Range("A1").Offset(0, 0).Resize(1, Range("A1").End(xlToRight).Column)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.ThemeColor = xlThemeColorAccent3
        .Font.Size = 9
    End With
    Range("A1").Resize(1, Range("A1").End(xlToRight).Column).WrapText = True
    With Range("A1").Offset(1, 0).Resize(Range("A1").End(xlDown).row - 1, Range("A1").End(xlToRight).Column)
        .Borders.LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Font.Size = 9
        .EntireColumn.ColumnWidth = 1
        .EntireColumn.AutoFit
    End With
    Range("A1").Offset(Range("A1").End(xlDown).row, 0).Resize(1, Range("A1").End(xlToRight).Column).Borders(xlEdgeTop).LineStyle = xlContinuous
   
    Range("A1").Select
    Selection.AutoFilter
   
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
Public Sub CheckList(ByRef shtTarget As Worksheet, ByRef strSQLBase As String, Optional ByVal strKey As String = "object_id")

    On Error GoTo ErrHandle

    shtTarget.Activate
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim intRow As Integer
    Dim intCol As Integer
    Dim maxCol As Integer
    Dim strSql As String
    
    maxCol = Range(cellList).End(xlToRight).Column
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    For intRow = Range(cellList).row + 1 To Range(cellList).End(xlDown).row
        
        strSql = strSQLBase & " WHERE " & strKey & " = '" & Trim(Cells(intRow, 1).Value) & "' "
        
        Set rsBase = conQCSDB.Execute(strSql, , adCmdText)
        If rsBase.EOF Then
            Range(Cells(intRow, 1), Cells(intRow, maxCol)).Interior.ThemeColor = xlThemeColorAccent6
        Else
            For intCol = 1 To maxCol
                If IsNull(rsBase.Fields(intCol - 1).Value) Then
                    If Cells(intRow, intCol).Value = "" Then
                        Cells(intRow, intCol).Interior.pattern = xlNone
                    Else
                        Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                    End If
                Else
                    If Cells(intRow, intCol).Value = Trim(rsBase.Fields(intCol - 1).Value) Then
                        Cells(intRow, intCol).Interior.pattern = xlNone
                    Else
                        Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                    End If
                End If
            Next
        End If
        rsBase.Close
        Set rsBase = Nothing
        
    Next
    
    conQCSDB.Close
    Set conQCSDB = Nothing

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

'オブジェクト一覧取得
Public Function listObject(where As String) As String
    
    On Error GoTo ErrHandle

    listObject = ""
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim strSql As String
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    strSql = "SELECT * FROM ObjectList " & where
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rsBase.EOF
        listObject = listObject & "," & rsBase.Fields("リスト").Value
        rsBase.MoveNext
    Loop
    
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

End Function

