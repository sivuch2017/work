VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.1
Option Explicit

Implements DataListInterface

Private Super As DataListInterface

Const NOEDIT = xlThemeColorDark2
Const TARGET = xlThemeColorAccent6
Const TABLEKEYS = "TABLEKEYS"
Const TABLEVALS = "TABLEVALS"

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
    If DataListInterface_ColToTable = "" Then
        'BUG FROM区の最初がテーブル名であることが前提
        DataListInterface_GetTableNameByCol = Split(Mid(DataListInterface_QuerySql, InStr(UCase(DataListInterface_QuerySql), "FROM")), " ")(1)
    Else
        DataListInterface_GetTableNameByCol = Super.GetTableNameByCol(Name)
    End If
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
    Dim intRow, intMaxRow, intCol As Integer
    Dim strSQLWk, strWrk As String
    Dim varCompA, varCompB As Variant
    Dim bolComp As Boolean
    
    DataListInterface_Check = 0
    
    If DataListInterface_QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Function
    End If

    With DataListInterface_SheetObject
        If .Range(DataListInterface_StartCol).Offset(1, 0).Value = "" Then
            intMaxRow = 0
        Else
            intMaxRow = .Range(DataListInterface_StartCol).End(xlDown).Row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        For intRow = .Range(DataListInterface_StartCol).Row + 1 To intMaxRow
            strWrk = "WHERE "
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If DataListInterface_IsKeyCol(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value) Then
                    strSQLWk = strSQLWk & strWrk & .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                    strWrk = "AND "
                End If
            Next

            Set objCon = DataListInterface_Connector.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open strSQLWk, objCon, adOpenStatic, adLockReadOnly, adCmdText
            
            If objRec.EOF Then
                '新規行
                For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                    If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                        .Cells(intRow, intCol).Interior.ThemeColor = TARGET
                        DataListInterface_Check = 1
                    End If
                Next
            Else
                '既存行
                For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                    'DB側値セット
                    If IsNull(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value) Then
                        varCompA = ""
                    Else
                        If DataListInterface_IsInMeta(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value) Then
                            varCompA = Trim(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value)
                        Else
                            varCompA = objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value
                            If IsNumeric(varCompA) Then
                                varCompA = Trim(str(varCompA))
                            End If
                        End If
                    End If
                    'セル側値セット
                    If DataListInterface_IsInMeta(.Cells(intRow, intCol).Value) Then
                        varCompB = Trim(.Cells(intRow, intCol).Value)
                    Else
                        varCompB = .Cells(intRow, intCol).Value
                    End If
                    '値比較
                    If varCompA = varCompB Then
                        If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlColorIndexNone
                        End If
                    Else
                        If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                            .Cells(intRow, intCol).Interior.ThemeColor = TARGET
                            DataListInterface_Check = 1
                        Else
                            '数式を除く
                            If Left(varCompA, 1) <> "=" Then
                                bolComp = True
                            End If
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
        strWrk = ""
        If bolComp Then
            strWrk = "更新不可項目の変更は無視されます。"
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
    
    Err.Raise Err.Number, "SimpleList(Check):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'更新
Public Sub DataListInterface_Update()

    On Error GoTo ErrHandle

    Dim intCol, intRow As Integer
    Dim hashTables As Collection
    Dim hashTmp As Object
    Dim strTable As String
    Dim varTmp As Variant

    If DataListInterface_QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Set hashTables = New Collection
    
    With DataListInterface_SheetObject
        '行ループ
        For intRow = .Range(DataListInterface_StartCol).Row + 1 To .Range(DataListInterface_StartCol).End(xlDown).Row
            '行ハッシュ初期化
            Set hashTmp = Nothing
            '桁ループ
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If .Cells(intRow, intCol).Interior.ThemeColor = TARGET Then
                    '更新色検知した場合の処理
                    strTable = DataListInterface_GetTableNameByCol(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value)
                    If hashTmp Is Nothing Then
                        '行ハッシュ作成
                        Set hashTmp = CreateObject("Scripting.Dictionary")
                    End If
                    If Not hashTmp.Exists(strTable) Then
                        '行ハッシュにテーブルハッシュ作成
                        hashTmp.Add strTable, CreateObject("Scripting.Dictionary")
                        hashTmp(strTable).Add TABLEVALS, CreateObject("Scripting.Dictionary")
                    End If
                    'テーブルハッシュのデータ部にキーと値をセット
                    hashTmp(strTable).Item(TABLEVALS).Add GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value), Trim(.Cells(intRow, intCol).Value)
                    If Not hashTmp(strTable).Exists(TABLEKEYS) Then
                        'テーブルハッシュにキー部が無い場合
                        hashTmp(strTable).Add TABLEKEYS, CreateObject("Scripting.Dictionary")
                        For Each varTmp In Split(DataListInterface_KeyColNames, ",")
                            'テーブルハッシュのキー部にキーと値をセット
                            hashTmp(strTable).Item(TABLEKEYS).Add GetTableColNameByName(varTmp), Trim(.Cells(intRow, GetColByName(varTmp)).Value)
                        Next
                    End If
                End If
            Next
            If Not hashTmp Is Nothing Then
                '行ハッシュ追加
                hashTables.Add hashTmp
            End If
        Next
    End With
    
    '更新処理
    UpdateRecord hashTables
    
    Exit Sub

ErrHandle:

    Err.Raise Err.Number, "SimpleList(Update):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'削除
Public Sub DataListInterface_Delete(DeleteRow As Integer)

' テーブル取得方法を思いつくまで保留
MsgBox "メンテ中"

'    On Error GoTo ErrHandle
'
'    Dim objCon As ADODB.Connection
'    Dim strObjId, strSQL As String
'    Dim hashTmp As Object
'    Dim strTable As Variant
'
'    If DataListInterface_QueryParam <> "" Then
'        MsgBox "プロシージャーには未対応", vbOKOnly
'        Exit Sub
'    End If
'
'    Set hashTmp = CreateObject("Scripting.Dictionary")
'    With DataListInterface_SheetObject
'        For Each strTable In Split(DataListInterface_ColToTable, ",")
'            If Not hashTmp.Exists(strTable) Then
'                strSQL = "DELETE " & strTable & " WHERE "
'                strDmy = ""
'                For Each varKey In Split(DataListInterface_KeyColNames, ",")
'                    strSQL = strSQL & strDmy & GetTableColNameByName(varKey) & " = '" & .Cells(DeleteRow, GetColByName(varKey)).Value & "' "
'                    strDmy = "AND "
'                Next
'                hashTmp.Add strTable, strSQL
'            End If
'        Next
'    End With
'
'    Set objCon = DataListInterface_Connector.ConnectDatabase
'    objCon.BeginTrans
'    For Each strTable In TableNameSort(hashTmp.Keys)
'        objCon.Execute hashTmp(strTable)
'    Next
'    objCon.CommitTrans
'
'    objCon.Close
'    Set objCon = Nothing
'
'    Exit Sub
'
'ErrHandle:
'
'    If Not objCon Is Nothing Then
'        If objCon.State <> adStateClosed Then
'            objCon.RollbackTrans
'            objCon.Close
'        End If
'        Set objCon = Nothing
'    End If
'
'    Err.Raise Err.Number, "SimpleList(Delete):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'初期化
Private Sub Class_Initialize()
    Set Super = New DataListInterfaceImpl
End Sub

'更新用SQL取得
Private Function GetBaseSql() As String
    Dim strBuf As String
    Dim intIndex As Integer
    
    strBuf = DataListInterface_QuerySql
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

'更新データをDBに反映
Private Sub UpdateRecord(ByVal hashData As Collection)

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim objRec As New ADODB.Recordset
    Dim hashTmp As Object
    Dim varTable, varKey As Variant
    Dim strObjId, strSql, strDmy As String
    
    Set objCon = DataListInterface_Connector.ConnectDatabase
    objCon.BeginTrans
    
    For Each hashTmp In hashData
        strObjId = ""
        For Each varTable In hashTmp.Keys
            strSql = "SELECT * FROM " & varTable & " WHERE "
            strDmy = ""
            For Each varKey In hashTmp(varTable).Item(TABLEKEYS).Keys
                strSql = strSql & strDmy & varKey & " = '" & hashTmp(varTable).Item(TABLEKEYS).Item(varKey) & "' "
                strDmy = "AND "
            Next
            objRec.Open strSql, objCon, adOpenStatic, adLockOptimistic, adCmdText
            If objRec.EOF Then
                objRec.AddNew
                For Each varKey In hashTmp(varTable).Item(TABLEKEYS)
                    If hashTmp(varTable).Item(TABLEKEYS).Item(varKey) <> "" Then
                        If hashTmp(varTable).Item(TABLEKEYS).Item(varKey) <> "*" Then
                            Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEKEYS).Item(varKey)
                            objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEKEYS).Item(varKey)
                        Else
                            If strObjId = "" Then
                                strObjId = GetObjectId
                            End If
                            Debug.Print varTable & ":" & varKey & ":" & strObjId
                            objRec.Fields(varKey).Value = strObjId
                        End If
                    End If
                Next
                For Each varKey In hashTmp(varTable).Item(TABLEVALS)
                    If hashTmp(varTable).Item(TABLEVALS).Item(varKey) <> "" Then
                        Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                        objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                    End If
                Next
            Else
                For Each varKey In hashTmp(varTable).Item(TABLEVALS)
                    If hashTmp(varTable).Item(TABLEVALS).Item(varKey) <> "" Then
                        Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                        objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                    Else
                        objRec.Fields(varKey).Value = Null
                    End If
                Next
            End If
            objRec.Update
            objRec.Close
        Next
    Next
    
    Set objRec = Nothing

    objCon.CommitTrans
    objCon.Close
    Set objCon = Nothing
    
    Exit Sub

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
    
    Err.Raise Err.Number, "SimpleList(UpdateRecord):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'空きobject_id取得
Private Function GetObjectId() As String
    
    On Error GoTo ErrHandle
    
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT RIGHT('00000000000000'+CONVERT(VARCHAR,MIN(CONVERT(BIGINT,object_id))+1),14) AS object_id FROM ObjectMaster WHERE CONVERT(BIGINT,object_id)+1 NOT IN (SELECT object_id FROM ObjectMaster)"
    objRec.Open strSql, DataListInterface_Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    If objRec.EOF Then
    Else
        GetObjectId = objRec.Fields("object_id").Value
    End If
    
    objRec.Close
    Set objRec = Nothing

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    Err.Raise Err.Number, "SimpleList(GetObjectId):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'項目名から桁位置を取得
Private Function GetColByName(ByVal Name As String) As Integer
    Dim intCol As Integer
    GetColByName = -1
    With DataListInterface_SheetObject
        For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
            If .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value = Name Then
                GetColByName = intCol
                Exit For
            End If
        Next
    End With
End Function

'ビュー項目名からテーブル項目名を取得
Private Function GetTableColNameByName(ByVal Name As String) As String
    Select Case Name
        Case Else
            GetTableColNameByName = Name
    End Select
End Function

'テーブル名更新順
Private Sub TableNameSort(ByRef NameArray() As Variant)
    Dim varSwap As Variant
    Dim i, j As Integer
    For i = LBound(NameArray) To UBound(NameArray)
        For j = UBound(NameArray) To i Step -1
            If NameArray(i) = "ObjectMaster" Then
                varSwap = NameArray(i)
                NameArray(i) = NameArray(j)
                NameArray(j) = varSwap
            End If
        Next j
    Next i
End Sub


