VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VMList"
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

    On Error GoTo ErrHandle

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
    Dim bolComp, bolHost, bolVMS As Boolean
    Dim strWrk As String
    
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
            intMaxRow = .Range(Super.StartCol).End(xlDown).row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        bolHost = False
        For intRow = .Range(Super.StartCol).row + 1 To intMaxRow
            intBool = 0
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                If Super.IsKeyCol(.Cells(.Range(Super.StartCol).row, intCol).Value) Then
                        strSQLWk = strSQLWk & strType(intBool) & .Cells(.Range(Super.StartCol).row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                        intBool = 1
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
            
            If .Cells(intRow, GetColByName("IPアドレス")).Interior.ThemeColor = xlThemeColorAccent6 And .Cells(intRow, GetColByName("ホスト名")).Interior.ThemeColor = xlThemeColorAccent6 Then
                bolHost = True
            End If
            
            If .Cells(intRow, GetColByName("VMサーバー名")).Interior.ThemeColor = xlThemeColorAccent6 And GetObjectIdByName(Trim(.Cells(intRow, GetColByName("VMサーバー名")).Value)) = "" Then
                bolVMS = True
            End If

        Next
    End With

    If bolComp Or bolHost Then
        strWrk = ""
        If bolComp Then
            strWrk = "更新不可項目の変更は無視されます。"
        End If
        If bolHost Then
            If strWrk <> "" Then
                strWrk = strWrk & vbCrLf
            End If
            strWrk = strWrk & "既存のIPアドレスへ変更する場合、ホスト名の変更は無視されます。"
        End If
        If bolVMS Then
            If strWrk <> "" Then
                strWrk = strWrk & vbCrLf
            End If
            strWrk = strWrk & "存在しないVMサーバー名の変更は無視されます。"
        End If
        MsgBox strWrk, vbOKOnly
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
    
    Err.Raise Err.Number, "VMList(Check):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'更新
Public Sub DataListInterface_Update()

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim intCol As Integer
    Dim intRow As Integer
    Dim hashTables, hashTmp As Collection
    Dim strTable As String
    Dim strKey1, strKey2 As String
    Dim varTmp As Variant
    Dim boolSetKey, boolSetValue As Boolean

    If Super.QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Set hashTables = New Collection
    
    With DataListInterface_SheetObject
        For intRow = .Range(DataListInterface_StartCol).row + 1 To .Range(DataListInterface_StartCol).End(xlDown).row
            Set hashTmp = Nothing
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                    strTable = DataListInterface_GetTableNameByCol(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                    If hashTmp Is Nothing Then
                        Set hashTmp = New Collection
                        hashTmp.Add New Collection, strTable
                        hashTmp(strTable).Add New Collection, "COL_NAME"
                        hashTmp(strTable).Add New Collection, "COL_VALUE"
                    ElseIf Not IsCollectionKey(hashTmp, strTable) Then
                        hashTmp.Add New Collection, strTable
                        hashTmp(strTable).Add New Collection, "COL_NAME"
                        hashTmp(strTable).Add New Collection, "COL_VALUE"
                    End If
                    If .Cells(.Range(DataListInterface_StartCol).row, intCol).Value = "VMサーバー名" Then
                        hashTmp(strTable).Item("COL_NAME").Add GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                        hashTmp(strTable).Item("COL_VALUE").Add GetObjectIdByName(Trim(.Cells(intRow, intCol).Value)), GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                    Else
                        hashTmp(strTable).Item("COL_NAME").Add GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                        hashTmp(strTable).Item("COL_VALUE").Add Trim(.Cells(intRow, intCol).Value), GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                    End If
                    If Not IsCollectionKey(hashTmp(strTable), "KEY_VALUE") Then
                        If strTable = "Address" Then
                            hashTmp(strTable).Add Trim(.Cells(intRow, GetColByName("IPアドレス")).Value), "KEY_VALUE"
                        ElseIf strTable = "Server" Then
                            hashTmp(strTable).Add Trim(.Cells(intRow, GetColByName("object_id")).Value), "KEY_VALUE"
                        Else
                            hashTmp(strTable).Add Trim(.Cells(intRow, GetColByName("object_id")).Value), "KEY_VALUE"
                            hashTmp(strTable).Item("COL_NAME").Add "type_id"
                            hashTmp(strTable).Item("COL_VALUE").Add "02", "type_id"
                        End If
                    End If
                End If
            Next
            If Not hashTmp Is Nothing Then
                hashTables.Add hashTmp
            End If
        Next
    End With

        If UpdateRecord(hashTables) Then
        End If
    
    Exit Sub

ErrHandle:

    Err.Raise Err.Number, "VMList(Update):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'削除
Public Sub DataListInterface_Delete(DeleteRow As Integer)

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim strObjId As String
    Dim strSql As String

    If Super.QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    With Super.SheetObject
        'キー固定
        strObjId = .Cells(DeleteRow, GetColByName("object_id")).Value
        'Address更新
        strSql = "UPDATE Address SET object_id = null WHERE object_id = '" & strObjId & "'"
        Super.Connector.ConnectDatabase.Execute strSql
        'Server削除
        strSql = "DELETE Server WHERE object_id = '" & strObjId & "'"
        Super.Connector.ConnectDatabase.Execute strSql
        'ObjectMaster削除
        strSql = "DELETE ObjectMaster WHERE object_id = '" & strObjId & "'"
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
    
    Err.Raise Err.Number, "VMList(Delete):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'初期化
Private Sub Class_Initialize()
    Set Super = New DataListInterfaceImpl
End Sub

'項目名から桁位置を取得
Private Function GetColByName(Name As String) As Integer
    Dim intCol As Integer
    GetColByName = -1
    With Super.SheetObject
        For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
            If .Cells(.Range(Super.StartCol).row, intCol).Value = Name Then
                GetColByName = intCol
                Exit For
            End If
        Next
    End With
End Function

'ビュー項目名からテーブル項目名を取得
Private Function GetTableColNameByName(Name As String) As String
    Select Case Name
        Case "object_id"
            GetTableColNameByName = "object_id"
        Case "dependence_place"
            GetTableColNameByName = "dependence_place"
        Case "VMサーバー名"
            GetTableColNameByName = "dependence_place"
        Case "VM名"
            GetTableColNameByName = "vm_name"
        Case "ホスト名"
            GetTableColNameByName = "host_name"
        Case "IPアドレス"
            GetTableColNameByName = "ip_address"
        Case "貸出依頼者"
            GetTableColNameByName = "rental_client"
        Case "担当者"
            GetTableColNameByName = "person"
        Case "内容"
            GetTableColNameByName = "contents"
        Case "状況"
            GetTableColNameByName = "situation"
        Case "予定"
            GetTableColNameByName = "future"
        Case "管理外"
            GetTableColNameByName = "outside_flag"
        Case "備考"
            GetTableColNameByName = "memo"
    End Select
End Function

'Collection.Contains
Function IsCollectionKey(Hash As Collection, Key) As Boolean
    On Error GoTo ErrHandle
    Hash.Item (Key)
    IsCollectionKey = True
    Exit Function
ErrHandle:
    IsCollectionKey = False
End Function

'行データをDBに反映
Private Function UpdateRecord(ByVal hashData As Collection) As Boolean

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim objRec As New ADODB.Recordset
    Dim hashTmp As Collection
    Dim strSql As String
    Dim strObjId As String
    Dim varKey As Variant
    Dim strTable As Variant
    Dim strKeyColName As String
    
    UpdateRecord = False

    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    
    For Each hashTmp In hashData
        strObjId = GetObjectId
        '依存関係があるため更新順を指定
        For Each strTable In Array("ObjectMaster", "Server", "Address")
            If IsCollectionKey(hashTmp, strTable) Then
                If strTable = "Address" Then
                    strKeyColName = "ip_address"
                Else
                    strKeyColName = "object_id"
                End If
                strSql = "SELECT * FROM " & strTable & " WHERE " & strKeyColName & " = '" & hashTmp(strTable).Item("KEY_VALUE") & "'"
                objRec.Open strSql, objCon, adOpenStatic, adLockOptimistic, adCmdText
                If objRec.EOF Then
                    objRec.AddNew
                    For Each varKey In hashTmp(strTable).Item("COL_NAME")
                        If hashTmp(strTable).Item("COL_VALUE").Item(varKey) <> "" Then
                            Debug.Print strTable & ":" & varKey & ":" & hashTmp(strTable).Item("COL_VALUE").Item(varKey)
                            objRec.Fields(varKey).Value = hashTmp(strTable).Item("COL_VALUE").Item(varKey)
                        End If
                    Next
                    If strTable = "Address" Then
                        objRec.Fields("object_id") = strObjId
                    End If
                    If hashTmp(strTable).Item("KEY_VALUE") = "*" Then
                        'object_id自動採番が指定されている場合
                        objRec.Fields(strKeyColName).Value = strObjId
                    Else
                        objRec.Fields(strKeyColName).Value = hashTmp(strTable).Item("KEY_VALUE")
                    End If
                Else
                    For Each varKey In hashTmp(strTable).Item("COL_NAME")
                        If strTable = "Address" And varKey = "host_name" Then
                        Else
                            If hashTmp(strTable).Item("COL_VALUE").Item(varKey) <> "" Then
                                Debug.Print strTable & ":" & varKey & ":" & hashTmp(strTable).Item("COL_VALUE").Item(varKey)
                                objRec.Fields(varKey).Value = hashTmp(strTable).Item("COL_VALUE").Item(varKey)
                            Else
                                objRec.Fields(varKey).Value = Null
                            End If
                        End If
                    Next
                End If
                objRec.Update
                objRec.Close
            End If
        Next
    Next
    
    Set objRec = Nothing

    objCon.CommitTrans
    objCon.Close
    Set objCon = Nothing
    
    UpdateRecord = True

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "VMList(UpdateRecord):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

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

'object_id取得
Private Function GetObjectIdByName(HostName As String) As String
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT [object_id] FROM Address WHERE [host_name]=N'" & HostName & "'"
    objRec.Open strSql, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    If objRec.EOF Then
        GetObjectIdByName = ""
    Else
        GetObjectIdByName = objRec.Fields("object_id").Value
    End If
    objRec.Close
End Function
