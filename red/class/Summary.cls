VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 2.0
Option Explicit

Implements DataListInterface

Private Super As DataListInterface
Private Cha

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
    Dim intRow, intMaxRow, intCol As Integer
    Dim strSQLWk, strWrk As String
    Dim varCompA, varCompB As Variant
    Dim bolComp, bolHard As Boolean
    
    DataListInterface_Check = -1
    
    If DataListInterface_QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Function
    End If

    With DataListInterface_SheetObject
        If .Range(DataListInterface_StartCol).Offset(1, 0).Value = "" Then
            intMaxRow = 0
        Else
            intMaxRow = .Range(DataListInterface_StartCol).End(xlDown).row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        bolHard = False
        For intRow = .Range(DataListInterface_StartCol).row + 1 To intMaxRow
            strWrk = "WHERE "
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If DataListInterface_IsKeyCol(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value) Then
                    strSQLWk = strSQLWk & strWrk & .Cells(.Range(DataListInterface_StartCol).row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                    strWrk = "AND "
                End If
            Next

            Set objCon = DataListInterface_Connector.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open strSQLWk, objCon, adOpenStatic, adLockReadOnly, adCmdText
            
            If objRec.EOF Then
                '新規行
                If Trim(.Cells(intRow, GetColByName("ID")).Value) = "*" Then
                    For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                        'セル側値セット
                        If DataListInterface_IsInMeta(.Cells(intRow, intCol).Value) Then
                            varCompB = Trim(.Cells(intRow, intCol).Value)
                        Else
                            varCompB = .Cells(intRow, intCol).Value
                        End If
                        '値比較
                        If varCompB = "" Then
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
                Else
                    MsgBox "削除された行が存在します。", vbOKOnly
                End If
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

    If bolComp Or bolHard Then
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
    
    Err.Raise Err.Number, "Summary(Check):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'更新
Public Sub DataListInterface_Update()

    Dim intCol, intRow As Integer
    Dim recSummary As SummaryRec

    If DataListInterface_QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    With DataListInterface_SheetObject
        '行ループ
        For intRow = .Range(DataListInterface_StartCol).row + 1 To .Range(DataListInterface_StartCol).End(xlDown).row
            'レコード初期化
            Set recSummary = Nothing
            '桁ループ
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If .Cells(intRow, intCol).Interior.ThemeColor = TARGET Then
                    '更新色検知した場合の処理
                    If recSummary Is Nothing Then
                        'レコード作成
                        Set recSummary = New SummaryRec
                        recSummary.ID = Trim(.Cells(intRow, GetColByName("ID")).Value)
                        Set recSummary.DB情報 = DataListInterface_Connector
                    End If
                    Select Case Trim(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                        Case "ID"
                        Case "P"
                            recSummary.P = Trim(.Cells(intRow, intCol).Value)
                        Case "VM名"
                            recSummary.VM名 = Trim(.Cells(intRow, intCol).Value)
                        Case "依存先"
                            recSummary.依存先 = Trim(.Cells(intRow, intCol).Value)
                        Case "貸出依頼者"
                            recSummary.貸出依頼者 = Trim(.Cells(intRow, intCol).Value)
                        Case "担当者"
                            recSummary.担当者 = Trim(.Cells(intRow, intCol).Value)
                        Case "Type"
                            recSummary.CodeType = Trim(.Cells(intRow, intCol).Value)
                        Case "ホスト名"
                            recSummary.ホスト名 = Trim(.Cells(intRow, intCol).Value)
                        Case "IPアドレス"
                            recSummary.IPアドレス = Trim(.Cells(intRow, intCol).Value)
                        Case "内容"
                            recSummary.内容 = Trim(.Cells(intRow, intCol).Value)
                        Case "状況"
                            recSummary.状況 = Trim(.Cells(intRow, intCol).Value)
                        Case "予定"
                            recSummary.予定 = Trim(.Cells(intRow, intCol).Value)
                        Case "備考"
                            recSummary.備考 = Trim(.Cells(intRow, intCol).Value)
                        Case "管理外"
                            recSummary.管理外 = Trim(.Cells(intRow, intCol).Value)
                        Case "ラックNO"
                            recSummary.ラックNO = Trim(.Cells(intRow, intCol).Value)
                        Case "保守契約"
                            MsgBox "DEBUG:未対応(無視します):" & Trim(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                            'recSummary.保守契約 = Trim(.Cells(intRow, intCol).Value)
                        Case "シリアル番号"
                            recSummary.シリアル番号 = Trim(.Cells(intRow, intCol).Value)
                        Case "製品番号"
                            recSummary.製品番号 = Trim(.Cells(intRow, intCol).Value)
                        Case "メーカー"
                            recSummary.メーカー = Trim(.Cells(intRow, intCol).Value)
                        Case "マシン"
                            recSummary.マシン = Trim(.Cells(intRow, intCol).Value)
                        Case "資産番号"
                            MsgBox "DEBUG:未対応(無視します):" & Trim(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                            'recSummary.資産番号 = Trim(.Cells(intRow, intCol).Value)
                        Case "現物設置フロア"
                            MsgBox "DEBUG:未対応(HardWareのみ更新):" & Trim(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                            recSummary.現物設置フロア = Trim(.Cells(intRow, intCol).Value)
                        Case Else
                            MsgBox "DEBUG:未定義:" & Trim(.Cells(.Range(DataListInterface_StartCol).row, intCol).Value)
                    End Select
                End If
            Next
            If Not recSummary Is Nothing Then
                recSummary.Update
            End If
        Next
    End With
    
End Sub

'削除
Public Sub DataListInterface_Delete(DeleteRow As Integer)

    Dim objCon As ADODB.Connection
    Dim strObjId As String
    Dim strSql As String
    Dim strDmy As String
    Dim varKey As Variant

    If DataListInterface_QueryParam <> "" Then
        MsgBox "プロシージャーには未対応", vbOKOnly
        Exit Sub
    End If
    
    Dim recSummary As SummaryRec
    Set recSummary = New SummaryRec
    Set recSummary.DB情報 = Super.Connector
    recSummary.ID = DataListInterface_SheetObject.Cells(DeleteRow, GetColByName("ID")).Value
    recSummary.Delete
    
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

'項目名から桁位置を取得
Private Function GetColByName(ByVal Name As String) As Integer
    Dim intCol As Integer
    GetColByName = -1
    With DataListInterface_SheetObject
        For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
            If .Cells(.Range(DataListInterface_StartCol).row, intCol).Value = Name Then
                GetColByName = intCol
                Exit For
            End If
        Next
    End With
End Function

