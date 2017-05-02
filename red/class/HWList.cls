VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HWList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.0
Option Explicit

Implements DataListInterface

Private Super As DataListInterface

'対象シートオブジェクト
Property Set DataListInterface_SheetObject(ByRef Sheet As Worksheet)
    Set Super.SheetObject = Sheet
End Property
Property Get DataListInterface_SheetObject() As Worksheet
    Set DataListInterface_SheetObject = Super.SheetObject
End Property

'表示開始位置
Public Property Let DataListInterface_StartCol(ByVal Address As String)
    Super.StartCol = Address
End Property
Public Property Get DataListInterface_StartCol() As String
    DataListInterface_StartCol = Super.StartCol
End Property

'読み出しSQL
Public Property Let DataListInterface_QuerySql(ByVal Sql As String)
    Super.QuerySQL = Sql
End Property
Public Property Get DataListInterface_QuerySql() As String
    DataListInterface_QuerySql = Super.QuerySQL
End Property

'SQLパラメータ
Public Property Let DataListInterface_QueryParam(ByVal Param As String)
    Super.QueryParam = Param
End Property
Public Property Get DataListInterface_QueryParam() As String
    DataListInterface_QueryParam = Super.QueryParam
End Property

'隠蔽対象カラム名(カンマ区切り)
Public Property Let DataListInterface_HiddenColNames(ByVal Names As String)
    Super.HiddenColNames = Names
End Property
Public Property Get DataListInterface_HiddenColNames() As String
    DataListInterface_HiddenColNames = Super.HiddenColNames
End Property

'グループ化対象カラム名(カンマ区切り)
Public Property Let DataListInterface_GroupColNames(ByVal Names As String)
    Super.GroupColNames = Names
End Property
Public Property Get DataListInterface_GroupColNames() As String
    DataListInterface_GroupColNames = Super.GroupColNames
End Property

'グループ化対象ロー条件(パイプ、カンマ区切り)
'カラム名,条件(EQ or NE),値|カラム名,条件(EQ or NE),値...
Public Property Let DataListInterface_GroupRow(ByVal NameAndValue As String)
    Super.GroupRow = NameAndValue
End Property
Public Property Get DataListInterface_GroupRow() As String
    DataListInterface_GroupRow = Super.GroupRow
End Property

'更新対象カラム名(カンマ区切り)
Public Property Let DataListInterface_UpdColNames(ByVal Names As String)
    Super.UpdColNames = Names
End Property
Public Property Get DataListInterface_UpdColNames() As String
    DataListInterface_UpdColNames = Super.UpdColNames
End Property

'キーカラム名(カンマ区切り)
Public Property Let DataListInterface_KeyColNames(ByVal Names As String)
    Super.KeyColNames = Names
End Property
Public Property Get DataListInterface_KeyColNames() As String
    DataListInterface_KeyColNames = Super.KeyColNames
End Property

'表題背景色
Public Property Let DataListInterface_TitleThemeColor(ByVal ThemeColor As Variant)
    Super.TitleThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_TitleThemeColor() As Variant
    DataListInterface_TitleThemeColor = Super.TitleThemeColor
End Property

'データ背景色
Public Property Let DataListInterface_DataThemeColor(ByVal ThemeColor As Variant)
    Super.DataThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_DataThemeColor() As Variant
    DataListInterface_DataThemeColor = Super.DataThemeColor
End Property

'接続先
Public Property Set DataListInterface_Connector(ByRef Database As DatabaseInterface)
    Set Super.Connector = Database
End Property
Public Property Get DataListInterface_Connector() As DatabaseInterface
    Set DataListInterface_Connector = Super.Connector
End Property

'カラムに対するテーブル名(カンマ区切り)
Public Property Let DataListInterface_ColToTable(ByVal Names As String)
    Super.ColToTable = Names
End Property
Public Property Get DataListInterface_ColToTable() As String
    DataListInterface_ColToTable = Super.ColToTable
End Property

'カラム別名コレクション
Public Property Set DataListInterface_ColAlias(ByRef NameHash As Collection)
    Set Super.ColAlias = NameHash
End Property
Public Property Get DataListInterface_ColAlias() As Collection
    Set DataListInterface_ColAlias = Super.ColAlias
End Property

'カラムに対するテーブル名を取得
Public Function DataListInterface_GetTableNameByCol(Name As String) As String
    DataListInterface_GetTableNameByCol = Super.GetTableNameByCol(Name)
End Function

'グループ対象か
Public Function DataListInterface_IsGroupCol(Name As String) As Boolean
    DataListInterface_IsGroupCol = Super.IsGroupCol(Name)
End Function

'グループ対象か
Public Function DataListInterface_IsGroupRow(Name As String, Optional Value As Variant) As Boolean
    DataListInterface_IsGroupRow = Super.IsGroupRow(Name, Value)
End Function

'更新可能フラグ
Public Function DataListInterface_IsUpdCol(Name As String) As Boolean
    DataListInterface_IsUpdCol = Super.IsUpdCol(Name)
End Function

'キーカラムフラグ
Public Function DataListInterface_IsKeyCol(Name As String) As Boolean
    DataListInterface_IsKeyCol = Super.IsKeyCol(Name)
End Function

'特殊文字を含むか文字列か？
Public Function DataListInterface_IsInMeta(Value As String) As Boolean
    DataListInterface_IsInMeta = Super.IsInMeta(Value)
End Function

'一覧を作成
Public Sub DataListInterface_CreateList()
    Super.CreateList
End Sub

'枠固定
Public Sub DataListInterface_FreezePanes()
    Super.FreezePanes
End Sub

'更新チェック
Public Function DataListInterface_Check() As Integer

'    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim intRow As Integer
    Dim intMaxRow As Integer
    Dim intCol As Integer
    Dim strSQLWk As String
    Dim varCompA As Variant
    Dim varCompB As Variant
    Dim intBool As Integer
    Dim strType(2) As String
    Dim bolComp As Boolean
    
    DataListInterface_Check = 0
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    If Super.QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Function
    End If

    With Super.SheetObject
        If .Range(Super.StartCol).Offset(1, 0).Value = "" Then
            intMaxRow = 0
        Else
            intMaxRow = .Range(Super.StartCol).End(xlDown).Row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        For intRow = .Range(Super.StartCol).Row + 1 To intMaxRow
            intBool = 0
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                If Super.IsKeyCol(.Cells(.Range(Super.StartCol).Row, intCol).Value) Then
'                    If Super.IsInMeta(.Cells(intRow, intCol).Value) Then
                        strSQLWk = strSQLWk & strType(intBool) & .Cells(.Range(Super.StartCol).Row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                        intBool = 1
'                    Else
'                        strSQLWk = strSQLWk & strType(intBool) & .Cells(.Range(Super.StartCol).Row, intCol).Value & " = " & Trim(.Cells(intRow, intCol).Value) & " "
'                        intBool = 1
'                    End If
                End If
            Next

            Set objCon = Super.Connector.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open strSQLWk, objCon, adOpenStatic, adLockReadOnly, adCmdText
            
            If objRec.EOF Then
                '新規行
                For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                    If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                        .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                        DataListInterface_Check = 1
                    End If
                Next
            Else
                For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                    If IsNull(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value) Then
                        varCompA = ""
                    Else
                        If Super.IsInMeta(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value) Then
                            varCompA = Trim(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value)
                        Else
                            varCompA = objRec.Fields(intCol - .Range(Super.StartCol).Column).Value
                            If IsNumeric(varCompA) Then
                                varCompA = Trim(str(varCompA))
                            End If
                        End If
                    End If
                    If Super.IsInMeta(.Cells(intRow, intCol).Value) Then
                        varCompB = Trim(.Cells(intRow, intCol).Value)
                    Else
                        varCompB = .Cells(intRow, intCol).Value
                    End If
                    If varCompA = varCompB Then
                        If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlColorIndexNone
                        End If
                    Else
                        If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                            DataListInterface_Check = 1
                        Else
                            bolComp = True
                        End If
                    End If
                Next
            End If

            objRec.Close
            objCon.Close
            Set objRec = Nothing
            Set objCon = Nothing

        Next
    End With

    If bolComp Then
        MsgBox "更新不可項目が変更されていますが無視します。", vbOKOnly
    End If

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    If Not objCon Is Nothing Then
        If objCon.State = adStateOpen Then
            objCon.Close
        End If
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(Check):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'更新
Public Sub DataListInterface_Update()

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim intCol As Integer
    Dim intRow As Integer
    Dim hashTables As Collection
    Dim strTable As String
    Dim strKey1, strKey2 As String
    Dim varTmp As Variant
    Dim boolSetKey, boolSetValue As Boolean

    If Super.QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Set hashTables = New Collection
    For Each varTmp In GetTableNameList
        hashTables.Add New Collection, varTmp
        hashTables(varTmp).Add New Collection, "KEY"
        hashTables(varTmp).Add New Collection, "VALUE"
    Next
    
    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    With DataListInterface_SheetObject
        boolSetKey = False
        For intRow = .Range(DataListInterface_StartCol).Row + 1 To .Range(DataListInterface_StartCol).End(xlDown).Row
            boolSetValue = False
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If DataListInterface_IsKeyCol(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value) Then
                    If boolSetKey Then
                    Else
                        For Each varTmp In GetTableNameList
                            hashTables(varTmp).Item("VALUE").Add Trim(.Cells(intRow, intCol)), .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value
                            Next
                        boolSetKey = True
                    End If
                Else
                    If .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                        hashTables(GetTableName(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value)).Item("KEY").Add .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value
                        hashTables(GetTableName(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value)).Item("VALUE").Add Trim(.Cells(intRow, intCol)), .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value
                        boolSetValue = True
                    End If
                End If
            Next
            If boolSetValue Then
                If Not UpdateRecord(hashTables) Then
                    objCon.RollbackTrans
                    objCon.Close
                    Set objCon = Nothing
                    Exit Sub
                End If
            End If
        Next
    End With
    objCon.CommitTrans

    objCon.Close
    Set objCon = Nothing

    Exit Sub

ErrHandle:

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "HWList(Update):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'削除
Public Sub DataListInterface_Delete(DeleteRow As Integer)

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim intCol As Integer
    Dim intRow As Integer
    Dim varName As Variant
    Dim strKey1, strKey2 As String
    Dim strSQLORG, strSql As String
    Dim hashWork As Collection
    Dim varKey As Variant

    If Super.QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    With Super.SheetObject
        intRow = DeleteRow
        Set hashWork = New Collection
        For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
            If Super.IsKeyCol(.Cells(.Range(Super.StartCol).Row, intCol).Value) Then
                hashWork.Add Trim(.Cells(intRow, intCol)), .Cells(.Range(Super.StartCol).Row, intCol).Value
            End If
        Next
        If hashWork.Count < 1 Then
            MsgBox "指定された行にキー項目が存在しません"
            objCon.RollbackTrans
            objCon.Close
            Set objCon = Nothing
            Exit Sub
        End If
        
        strSQLORG = "DELETE " & Mid(Super.QuerySQL, InStr(UCase(Super.QuerySQL), "FROM")) & " WHERE "
        strSql = strSQLORG
        For Each varKey In Split(Super.KeyColNames, ",")
            If strSql <> strSQLORG Then
                strSql = strSql & " AND "
            End If
            strSql = strSql & varKey & " = '" & hashWork.Item(varKey) & "'"
        Next
        Super.Connector.ConnectDatabase.Execute strSql
    End With
    objCon.CommitTrans

    objCon.Close
    Set objCon = Nothing

    Exit Sub

ErrHandle:

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(Delete):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'初期化
Private Sub Class_Initialize()
    Set Super = New DataListInterfaceImpl
End Sub

'更新用SQL取得
Private Function GetBaseSql() As String
    Dim strBuf As String
    Dim intIndex As Integer
    
    strBuf = Super.QuerySQL
    intIndex = InStr(UCase(strBuf), "ORDER")
    If intIndex <> 0 Then
        strBuf = Left(strBuf, intIndex - 1)
    End If
    
    intIndex = InStr(UCase(strBuf), "WHERE")
    If intIndex <> 0 Then
        strBuf = Left(strBuf, intIndex - 1)
    End If
    
    GetBaseSql = strBuf
End Function

'行データをDBに反映
Private Function UpdateRecord(ByRef hashData As Collection) As Boolean

    On Error GoTo ErrHandle

    Dim objRec As New ADODB.Recordset
    Dim hashWork As Collection
    Dim strSql As String
    Dim strObjId As String
    Dim varKey As Variant
    Dim strTable As String
    
    UpdateRecord = False

    Dim wrk As String
    collectionPrint hashData, wrk
    Debug.Print wrk
    UpdateRecord = True
    
'    strTable = GetTableName
'    If hashData(strTable).Item("KEY").Count > 0 Then
'        strSQL = Super.QuerySQL & " WHERE "
'        For Each varKey In Split(Super.KeyColNames, ",")
'            If strSQL <> Super.QuerySQL & " WHERE " Then
'                strSQL = strSQL & " AND "
'            End If
'            strSQL = strSQL & varKey & " = '" & hashData(strTable).Item("VALUE").Item(varKey) & "'"
'        Next
'        objRec.Open strSQL, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
'        If objRec.EOF Then
'            objRec.AddNew
'            For Each varKey In Split(Super.KeyColNames, ",")
'                If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
'                    If varKey = "object_id" And hashData(strTable).Item("VALUE").Item(varKey) = "*" Then
'                        objRec.Fields(varKey).Value = GetObjectId
'                    Else
'                        objRec.Fields(varKey).Value = hashData(strTable).Item("VALUE").Item(varKey)
'                    End If
'                End If
'            Next
'            For Each varKey In hashData(strTable).Item("KEY")
'                If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
'                    objRec.Fields(varKey).Value = hashData(strTable).Item("VALUE").Item(varKey)
'                End If
'            Next
'        Else
'            For Each varKey In hashData(strTable).Item("KEY")
'                If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
'                    objRec.Fields(varKey).Value = hashData(strTable).Item("VALUE").Item(varKey)
'                Else
'                    objRec.Fields(varKey).Value = Null
'                End If
'            Next
'        End If
'        objRec.Update
'        objRec.Close
'    End If
'
'    Set objRec = Nothing
'
'    UpdateRecord = True

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    Err.Raise Err.Number, "SimpleList(UpdateRecord):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'テーブル名取得
Public Function GetTableName(Optional ByVal Name As String) As String
    Dim varItem As Variant
    If IsMissing(Name) Or DataListInterface_ColToTable = "" Then
        'BUG FROM区の最初がテーブル名であることが前提
        GetTableName = Split(Mid(DataListInterface_QuerySql, InStr(UCase(DataListInterface_QuerySql), "FROM")), " ")(1)
    Else
        GetTableName = DataListInterface_GetTableNameByCol(Name)
        If GetTableName = "" Then
            GetTableName = "HardWare"
        End If
    End If
End Function

'カラム名取得
Public Function GetColName(ByVal Name As String) As String
    On Error GoTo ErrHandle
    If Super.ColAlias Is Nothing Then
        GetColName = Name
    Else
        GetColName = Super.ColAlias.Item(Name)
    End If
    Exit Function
ErrHandle:
    GetColName = Name
End Function

'空きobject_id取得
Private Function GetObjectId() As String
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT RIGHT('00000000000000'+CONVERT(VARCHAR,MIN(CONVERT(BIGINT,object_id))+1),14) AS object_id FROM ObjectMaster WHERE CONVERT(BIGINT,object_id)+1 NOT IN (SELECT object_id FROM ObjectMaster)"
    objRec.Open strSql, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    GetObjectId = objRec.Fields("object_id").Value
    objRec.Close
End Function

'更新対象テーブルのコレクション
Public Function GetTableNameList() As Collection
    Dim varTmp As Variant
    Set GetTableNameList = New Collection
    If DataListInterface_ColToTable = "" Then
        varTmp = GetTableName
        GetTableNameList.Add varTmp, varTmp
    Else
        On Error Resume Next
        For Each varTmp In Split(DataListInterface_ColToTable, ",")
            GetTableNameList.Add varTmp, varTmp
        Next
    End If
End Function

Public Sub collectionPrint(ByVal Hash As Collection, ByRef out As String)
    Dim var As Variant
    out = out & "{"
    For Each var In Hash
        If TypeName(var) = "Collection" Then
            collectionPrint var, out
        Else
            out = out & var & ","
        End If
    Next
    out = out & "}"
End Sub
