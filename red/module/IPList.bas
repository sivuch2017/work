Attribute VB_Name = "IPList"
Option Explicit

'シート名
Public Const strSheetName As String = "IP管理"

'対象テーブル
Private Const TABLES As String = "Address"

'更新可能カラム名
Private Const typeUpd As String = "ホスト名,ID,依存先ID"

'カラム⇒テーブル
Private Const ColToTable As String = "Address,Address,Address"

'更新可能カラム名(テーブル側)
Private Const tableUpd As String = "host_name,object_id,dependence_place"

'更新可能フラグ
Private Function boolUpdCol(ColName As String) As Boolean
    Dim index As Integer
    boolUpdCol = False
    For index = 0 To UBound(Split(typeUpd, ","))
        If Split(typeUpd, ",")(index) = ColName Then
            boolUpdCol = True
            Exit For
        End If
    Next
End Function

'更新可能カラム名取得
Private Function tableCol(ColName As String) As String
    Dim index As Integer
    tableCol = ""
    For index = 0 To UBound(Split(typeUpd, ","))
        If Split(typeUpd, ",")(index) = ColName Then
            tableCol = Split(tableUpd, ",")(index)
            Exit For
        End If
    Next
End Function

'テーブル名取得
Private Function TableName(ColName As String) As String
    Dim index As Integer
    TableName = ""
    For index = 0 To UBound(Split(typeUpd, ","))
        If Split(typeUpd, ",")(index) = ColName Then
            TableName = Split(ColToTable, ",")(index)
            Exit For
        End If
    Next
End Function

'一覧を作成
Public Sub createIpListSheet()
    
    Dim sht As Worksheet
    Dim maxCol As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strList As String
    
    For Each sht In Worksheets
        If sht.Name = strSheetName Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Set sht = Worksheets.Add
    sht.Name = strSheetName
    
    Windows.Application.ScreenUpdating = False
    
    loadTable sht, "SELECT * FROM dbo.IPList ORDER BY sort_ip", maxCol
    
    Columns(2).EntireColumn.Hidden = True
    
    For intCol = 1 To maxCol
        If Columns(intCol).EntireColumn.ColumnWidth > 48 Then
            Columns(intCol).EntireColumn.ColumnWidth = 48
        End If
        Range("A1").Offset(0, intCol - 1).WrapText = True
        If boolUpdCol(Range("A1").Offset(0, intCol - 1).Value) Then
        Else
            With Range("A1").Offset(1, intCol - 1).Resize(Range("A1").End(xlDown).row - 1, 1)
                .Interior.ThemeColor = xlThemeColorDark2
            End With
        End If
    Next
    
    Columns(10).GROUP
    Columns(11).GROUP
    For intRow = 2 To Range("A1").End(xlDown).row
        If Cells(intRow, 3).Value = "" And Cells(intRow, 6).Value = "" Then
            Rows(intRow).GROUP
        End If
    Next
    
    Range("F2").Select
    ActiveWindow.FreezePanes = True
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

    Windows.Application.ScreenUpdating = True

End Sub

'データベースと一覧を比較
Public Sub checkIPListSheet()
    
    Dim shtWork As Worksheet
    Dim shtIPList As Worksheet
    Dim celTgt As Range
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtIPList = shtWork
        End If
    Next
    
    If shtIPList Is Nothing Then
        MsgBox "対象シートが読み込まれていません。", vbOKOnly, "警告"
        Exit Sub
    End If
    
    CheckList shtIPList, "SELECT * FROM dbo.IPList", "IPアドレス"
    For Each celTgt In Range("A1").Offset(1, 0).Resize(Range("A1").End(xlDown).row - 1, Range("A1").End(xlToRight).Column)
        If celTgt.Interior.ThemeColor <> xlThemeColorAccent6 And Not boolUpdCol(Cells(1, celTgt.Column).Value) Then
            celTgt.Interior.ThemeColor = xlThemeColorDark2
        End If
    Next
    
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=0
    MsgBox "更新箇所を確認した後、もう一度更新処理を実施して下さい。", vbOKOnly, "情報"
    
End Sub

'データベース更新
Public Sub updateSheet()

    On Error GoTo ErrHandle

    Dim shtWork As Worksheet
    Dim shtIPList As Worksheet
    Dim intCol As Integer
    Dim intRow As Integer
    Dim boolKey As Boolean
    Dim boolUpd As Boolean
    Dim boolResult As Boolean
    Dim conQCSDB As New ADODB.Connection
    Dim hashTables As Collection
    Dim Name As Variant
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtIPList = shtWork
        End If
    Next
    
    If shtIPList Is Nothing Then
        MsgBox "対象シートが読み込まれていません。", vbOKOnly, "警告"
        Exit Sub
    End If
    
    shtIPList.Activate
    
    boolKey = False
    boolUpd = False
    For intCol = 1 To Range("A1").End(xlToRight).Column
        For intRow = 2 To Range("A1").End(xlDown).row
            If Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                If boolUpdCol(Cells(1, intCol).Value) Then
                    boolUpd = True
                Else
                    boolKey = True
                End If
            End If
        Next
    Next
    
    If boolKey Then
        MsgBox "更新不可項目が変更されています。", vbOKOnly, "警告"
        Exit Sub
    End If
    
    If Not boolUpd Then
        checkIPListSheet
        Exit Sub
    End If
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    conQCSDB.BeginTrans
    For intRow = 2 To Range("A1").End(xlDown).row
        boolResult = True
        Set hashTables = New Collection
        For Each Name In Split(TABLES, ",")
            hashTables.Add New Collection, Name
            hashTables(Name).Add New Collection, "KEY"
            hashTables(Name).Add New Collection, "VALUE"
        Next
        For intCol = 1 To Range("A1").End(xlToRight).Column
            If Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                If hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Count = 0 Then
                    hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Add "ip_address"
                    hashTables(TableName(Cells(1, intCol).Value)).Item("VALUE").Add Trim(Cells(intRow, 1)), "ip_address"
                End If
                hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Add tableCol(Cells(1, intCol).Value)
                hashTables(TableName(Cells(1, intCol).Value)).Item("VALUE").Add Trim(Cells(intRow, intCol)), tableCol(Cells(1, intCol).Value)
            End If
        Next
        For Each Name In Split(TABLES, ",")
            If hashTables(Name).Item("KEY").Count > 0 Then
                boolResult = UpdateRecord(conQCSDB, Name, hashTables(Name))
            End If
        Next
        If Not boolResult Then
            Exit For
        End If
    Next
    If boolResult Then
        conQCSDB.CommitTrans
        MsgBox "DBを更新しました。", vbOKOnly, "情報"
        createIpListSheet
    Else
        conQCSDB.RollbackTrans
    End If

    conQCSDB.Close
    Set conQCSDB = Nothing
    
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

End Sub

'レコード更新
Private Function UpdateRecord(conQCSDB As ADODB.Connection, table As Variant, values As Collection) As Boolean

'    On Error GoTo ErrHandle

    Dim rsBase As New ADODB.Recordset
    Dim strSql As String
    Dim dtNow As Date
    Dim intCol As Integer
    Dim strType(2) As String
    Dim intBool As Integer
    Dim Key As Variant
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    UpdateRecord = False

    strSql = "SELECT * FROM " & table & " WHERE ip_address = '" & values("VALUE").Item("ip_address") & "' "
    Debug.Print strSql
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockOptimistic, adCmdText
    If rsBase.EOF Then
        MsgBox "更新チェック後にレコードが削除されました" & vbCrLf & "レコード更新を見送ります。", vbOKOnly, "情報"
        Exit Function
    Else
        For Each Key In values("KEY")
            rsBase.Fields(Key).Value = values("VALUE").Item(Key)
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

'稼働チェック
Public Sub CheckIP()

    Dim shtWork As Worksheet
    Dim shtIPList As Worksheet
    Dim intCol As Integer
    Dim intRow As Integer
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtIPList = shtWork
        End If
    Next
    
    If shtIPList Is Nothing Then
        MsgBox "対象シートが読み込まれていません。", vbOKOnly, "警告"
        Exit Sub
    End If
    
    shtIPList.Activate
    
    For intRow = 1 To Range("A1").End(xlDown).row
        If ThisWorkbook.pingIp(Cells(intRow, 1).Value) Then
            Cells(intRow, 1).Font.Bold = True
        Else
            Cells(intRow, 1).Font.Bold = False
        End If
    Next

End Sub
