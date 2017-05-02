VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataListInterfaceImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.1
Option Explicit

Implements DataListInterface

Private ThisSheetObject As Worksheet
Private ThisStartCol As String
Private ThisQuerySql As String
Private ThisQueryParam As String
Private ThisHiddenColNames As String
Private ThisGroupColNames As String
Private ThisGroupRow As String
Private ThisUpdColNames As String
Private ThisKeyColNames As String
Private ThisTitleThemeColor As Variant
Private ThisDataThemeColor As Variant
Private ThisConnector As DatabaseInterface
Private ThisColToTable As String
Private ThisColAlias As Collection

'特殊文字
Private Const MetaChar As String = " ""'`#%><!.[]*$;:?^()+\"

'対象シートオブジェクト
Property Set DataListInterface_SheetObject(ByRef Sheet As Worksheet)
    Set ThisSheetObject = Sheet
End Property
Property Get DataListInterface_SheetObject() As Worksheet
    Set DataListInterface_SheetObject = ThisSheetObject
End Property

'表示開始位置
Public Property Let DataListInterface_StartCol(ByVal Address As String)
    ThisStartCol = Address
End Property
Public Property Get DataListInterface_StartCol() As String
    DataListInterface_StartCol = ThisStartCol
End Property

'読み出しSQL
Public Property Let DataListInterface_QuerySql(ByVal Sql As String)
    ThisQuerySql = Sql
End Property
Public Property Get DataListInterface_QuerySql() As String
    DataListInterface_QuerySql = ThisQuerySql
End Property

'SQLパラメータ
Public Property Let DataListInterface_QueryParam(ByVal Param As String)
    ThisQueryParam = Param
End Property
Public Property Get DataListInterface_QueryParam() As String
    DataListInterface_QueryParam = ThisQueryParam
End Property

'隠蔽対象カラム名(カンマ区切り)
Public Property Let DataListInterface_HiddenColNames(ByVal Names As String)
    ThisHiddenColNames = Names
End Property
Public Property Get DataListInterface_HiddenColNames() As String
    DataListInterface_HiddenColNames = ThisHiddenColNames
End Property

'グループ化対象カラム名(カンマ区切り)
Public Property Let DataListInterface_GroupColNames(ByVal Names As String)
    ThisGroupColNames = Names
End Property
Public Property Get DataListInterface_GroupColNames() As String
    DataListInterface_GroupColNames = ThisGroupColNames
End Property

'グループ化対象ロー条件(パイプ、カンマ区切り)
'カラム名,条件(EQ or NE),値|カラム名,条件(EQ or NE),値...
Public Property Let DataListInterface_GroupRow(ByVal NameAndValue As String)
    ThisGroupRow = NameAndValue
End Property
Public Property Get DataListInterface_GroupRow() As String
    DataListInterface_GroupRow = ThisGroupRow
End Property

'更新対象カラム名(カンマ区切り)
Public Property Let DataListInterface_UpdColNames(ByVal Names As String)
    ThisUpdColNames = Names
End Property
Public Property Get DataListInterface_UpdColNames() As String
    DataListInterface_UpdColNames = ThisUpdColNames
End Property

'キーカラム名(カンマ区切り)
Public Property Let DataListInterface_KeyColNames(ByVal Names As String)
    ThisKeyColNames = Names
End Property
Public Property Get DataListInterface_KeyColNames() As String
    DataListInterface_KeyColNames = ThisKeyColNames
End Property

'表題背景色
Public Property Let DataListInterface_TitleThemeColor(ByVal ThemeColor As Variant)
    ThisTitleThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_TitleThemeColor() As Variant
    If ThisTitleThemeColor = 0 Then
        DataListInterface_TitleThemeColor = xlColorIndexNone
    Else
        DataListInterface_TitleThemeColor = ThisTitleThemeColor
    End If
End Property

'データ背景色
Public Property Let DataListInterface_DataThemeColor(ByVal ThemeColor As Variant)
    ThisDataThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_DataThemeColor() As Variant
    If ThisDataThemeColor = 0 Then
        DataListInterface_DataThemeColor = xlColorIndexNone
    Else
        DataListInterface_DataThemeColor = ThisDataThemeColor
    End If
End Property

'接続先
Public Property Set DataListInterface_Connector(ByRef Database As DatabaseInterface)
    Set ThisConnector = Database
End Property
Public Property Get DataListInterface_Connector() As DatabaseInterface
    Set DataListInterface_Connector = ThisConnector
End Property

'カラムに対するテーブル名(カンマ区切り)
Public Property Let DataListInterface_ColToTable(ByVal Names As String)
    ThisColToTable = Names
End Property
Public Property Get DataListInterface_ColToTable() As String
    DataListInterface_ColToTable = ThisColToTable
End Property

'カラム別名コレクション
Public Property Set DataListInterface_ColAlias(ByRef NameHash As Collection)
    Set ThisColAlias = NameHash
End Property
Public Property Get DataListInterface_ColAlias() As Collection
    Set DataListInterface_ColAlias = ThisColAlias
End Property

'カラムに対するテーブル名を取得
Public Function DataListInterface_GetTableNameByCol(Name As String) As String
    Dim intCnt As Variant
    DataListInterface_GetTableNameByCol = ""
    For intCnt = 0 To UBound(Split(ThisUpdColNames, ","))
        If Split(ThisUpdColNames, ",")(intCnt) = Name Then
            DataListInterface_GetTableNameByCol = Split(ThisColToTable, ",")(intCnt)
            Exit For
        End If
    Next
End Function

'グループ対象か
Public Function DataListInterface_IsGroupCol(Name As String) As Boolean
    Dim varTmp As Variant
    DataListInterface_IsGroupCol = False
    For Each varTmp In Split(ThisGroupColNames, ",")
        If varTmp = Name Then
            DataListInterface_IsGroupCol = True
            Exit For
        End If
    Next
End Function

'グループ対象か
Public Function DataListInterface_IsGroupRow(Name As String, Optional Value As Variant) As Boolean
    Dim strItem As Variant
    DataListInterface_IsGroupRow = False
    If ThisGroupRow <> "" Then
        For Each strItem In Split(ThisGroupRow, "|")
            If IsMissing(Value) Then
                If Split(strItem, ",")(0) = Name Then
                    DataListInterface_IsGroupRow = True
                    Exit Function
                End If
            Else
                If Split(strItem, ",")(0) = Name Then
                    Select Case Split(strItem, ",")(1)
                        Case "EQ"
                            If Split(strItem, ",")(2) = Value Then
                                DataListInterface_IsGroupRow = True
                                Exit Function
                            End If
                        Case "NE"
                            If Split(strItem, ",")(2) <> Value Then
                                DataListInterface_IsGroupRow = True
                                Exit Function
                            End If
                    End Select
                End If
            End If
        Next
    End If
End Function

'更新可能フラグ
Public Function DataListInterface_IsUpdCol(Name As String) As Boolean
    Dim varTmp As Variant
    DataListInterface_IsUpdCol = False
    For Each varTmp In Split(ThisUpdColNames, ",")
        If varTmp = Name Then
            DataListInterface_IsUpdCol = True
            Exit For
        End If
    Next
End Function

'キーカラムフラグ
Public Function DataListInterface_IsKeyCol(Name As String) As Boolean
    DataListInterface_IsKeyCol = (InStr(ThisKeyColNames, Name) > 0)
End Function

'特殊文字を含むか文字列か？
Public Function DataListInterface_IsInMeta(Value As String) As Boolean
    Dim intCnt As Integer
    DataListInterface_IsInMeta = False
    If Len(Value) > 1 And Left(Value, 1) = "0" Then
        DataListInterface_IsInMeta = True
    Else
        For intCnt = 1 To Len(MetaChar)
            If InStr(Value, Mid(MetaChar, intCnt, 1)) > 0 Then
                DataListInterface_IsInMeta = True
                Exit For
            End If
        Next
    End If
End Function

'一覧を作成
Public Sub DataListInterface_CreateList()

    Dim intRow, intMaxRow As Integer
    Dim intCol As Integer

    '設定チェック
    If ThisSheetObject Is Nothing Or ThisStartCol = "" Or ThisQuerySql = "" Or ThisConnector Is Nothing Then
        Err.Raise 10000, "DataListInterfaceImpl(CreateList):", "必要な設定がされていません。"
    End If
    
    'データ貼り付け
    LoadData
    
    'シート調整
    With ThisSheetObject
        If .Range(ThisStartCol).Offset(1, 0).Value <> "" Then
            intRow = .Range(ThisStartCol).End(xlDown).row - .Range(ThisStartCol).row + 1
        Else
            intRow = 1
        End If
        If ThisUpdColNames <> "" Then
            intRow = intRow + 1
        End If
        intCol = .Range(ThisStartCol).End(xlToRight).Column - .Range(ThisStartCol).Column + 1
        
        'タイトル列
        With .Range(ThisStartCol).Resize(1, intCol)
            .Interior.ThemeColor = DataListInterface_TitleThemeColor
            .Interior.TintAndShade = 0.8
            .Borders.LineStyle = xlContinuous
            .Font.Size = 9
            .VerticalAlignment = xlTop
            .WrapText = True
            .AutoFilter
            .EntireColumn.ColumnWidth = 1
        End With
        
        'データ列
        If intRow <> 1 Then
            With .Range(ThisStartCol).Offset(1, 0).Resize(intRow - 1, intCol)
                .Select
                .Interior.ThemeColor = DataListInterface_DataThemeColor
                .Interior.TintAndShade = 0.8
                .Borders.LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Font.Size = 9
                .EntireColumn.AutoFit
            End With
        End If
        
        .Range(ThisStartCol).Resize(1, intCol).EntireRow.AutoFit
        
        '最右列+1
        With .Range(ThisStartCol).Offset(0, intCol).Resize(intRow, 1)
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        
        '最下行+1
        With .Range(ThisStartCol).Offset(intRow, 0).Resize(1, intCol)
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        
        For intCol = .Range(ThisStartCol).Column To .Range(ThisStartCol).End(xlToRight).Column
            '更新不可列
            If ThisUpdColNames <> "" Then
                If DataListInterface_IsUpdCol(.Cells(.Range(ThisStartCol).row, intCol).Value) Then
                Else
                    If DataListInterface_IsKeyCol(.Cells(.Range(ThisStartCol).row, intCol).Value) Then
                        intMaxRow = intRow - 2
                    Else
                        intMaxRow = intRow - 1
                    End If
                    If intMaxRow > 0 Then
                        .Cells(.Range(ThisStartCol).row + 1, intCol).Resize(intMaxRow, 1).Interior.ThemeColor = xlThemeColorDark2
                    End If
                End If
            End If
            '隠蔽
            If IsHiddenCol(.Cells(.Range(ThisStartCol).row, intCol).Value) Then
                .Columns(intCol).EntireColumn.Hidden = True
            End If
            'グループ化
            If DataListInterface_IsGroupCol(.Cells(.Range(ThisStartCol).row, intCol).Value) Then
                .Columns(intCol).GROUP
            End If
            'サイズ制限
            If .Columns(intCol).EntireColumn.ColumnWidth > 48 Then
                .Columns(intCol).EntireColumn.ColumnWidth = 48
            End If
        Next
        
        '行グループ化
        For intCol = .Range(ThisStartCol).Column To .Range(ThisStartCol).End(xlToRight).Column
            If DataListInterface_IsGroupRow(.Cells(.Range(ThisStartCol).row, intCol).Value) Then
                If .Range(ThisStartCol).Offset(1, 0).Value <> "" Then
                    For intRow = .Range(ThisStartCol).row + 1 To .Range(ThisStartCol).End(xlDown).row
                        If DataListInterface_IsGroupRow(.Cells(.Range(ThisStartCol).row, intCol).Value, .Cells(intRow, intCol).Value) Then
                            If .Rows(intRow).OutlineLevel < 2 Then
                                .Rows(intRow).GROUP
                            End If
                        End If
                    Next
                End If
            End If
        Next
        
        'カーソル移動
        SendKeys "^{HOME}"
    End With

End Sub

'枠固定
Public Sub DataListInterface_FreezePanes()
    ThisSheetObject.Range(ThisStartCol).Offset(1, 0).Select
    ActiveWindow.FreezePanes = True
End Sub

'更新チェック
Public Function DataListInterface_Check() As Integer
End Function

'更新
Public Sub DataListInterface_Update()
End Sub

'削除
Public Sub DataListInterface_Delete(DeleteRow As Integer)
End Sub

'データ取得
Private Sub LoadData()
    
    On Error GoTo ErrHandle

    ThisSheetObject.Activate
    
    Dim objCon As ADODB.Connection
    Dim objCat As ADOX.catalog
    Dim objCmd As ADODB.Command
    Dim objRec As ADODB.Recordset
    Dim intCol As Integer
    Dim intRow As Integer
    Dim aryStr As Variant
    
    Set objCon = ThisConnector.ConnectDatabase
    Set objCat = New ADOX.catalog
    objCat.ActiveConnection = objCon
    
    If ThisQueryParam = "" Then
        Set objRec = New ADODB.Recordset
        objRec.Open ThisQuerySql, objCon, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Set objCmd = objCat.Procedures(ThisQuerySql).Command
        Set objRec = objCmd.Execute(Parameters:=ThisQueryParam)
    End If
    
    intRow = 0
    For intCol = 0 To objRec.Fields.Count - 1
        ThisSheetObject.Range(ThisStartCol).Offset(intRow, intCol).Value = objRec.Fields(intCol).Name
    Next
    
    If objRec.EOF Then
        MsgBox "対象レコードがありません。", vbOKOnly, "情報"
    Else
        intRow = 1
        Do Until objRec.EOF
            For intCol = 0 To objRec.Fields.Count - 1
                'NULLチェック
                If IsNull(objRec.Fields(intCol).Value) Then
                    ThisSheetObject.Range(ThisStartCol).Offset(intRow, intCol).Value = ""
                Else
                    '数式チェック
                    If Left(Trim(objRec.Fields(intCol).Value), 1) <> "=" Then
                        ThisSheetObject.Range(ThisStartCol).Offset(intRow, intCol).NumberFormat = "@"
                    End If
                    'リンク型チェック
                    If IsNull(objRec.Fields(intCol).Properties("BASETABLENAME").Value) Then
                        ThisSheetObject.Range(ThisStartCol).Offset(intRow, intCol).Value = Trim(objRec.Fields(intCol).Value)
                    Else
                        If ThisConnector.IsLinkType(objRec.Fields(intCol).Properties("BASETABLENAME").Value, objRec.Fields(intCol).Name) Then
                            aryStr = Split(objRec.Fields(intCol).Value, "#")
                            ThisSheetObject.Hyperlinks.Add Range(ThisStartCol).Offset(intRow, intCol), aryStr(1), , , aryStr(0)
                        Else
                            ThisSheetObject.Range(ThisStartCol).Offset(intRow, intCol).Value = Trim(objRec.Fields(intCol).Value)
                        End If
                    End If
                End If
            Next
            intRow = intRow + 1
            objRec.MoveNext
        Loop
    End If
    
    objRec.Close
    objCon.Close
    Set objRec = Nothing
    Set objCmd = Nothing
    Set objCon = Nothing
    
    Exit Sub

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    Set objCmd = Nothing
    
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
    
    Err.Raise Err.Number, "DataListInterfaceImpl(LoadData):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'漢字変換
Function ToSJIS(ByVal SourceString As String)
    Dim buf() As Byte
    Dim code As String
    If SourceString <> "" Then
        buf = SourceString
'        If (buf(0) = 255) And (buf(1) = 254) Then
            code = "unicode"
'        ElseIf (buf(0) = 239) And (buf(1) = 187) And (buf(2) = 191) Then
'            code = "utf-8"
'        Else
'            code = "shift-jis"
'        End If
        With CreateObject("ADODB.Stream")
            .Open
            .Type = adTypeText
            .Charset = code
            .WriteText SourceString
            .Position = 0
            .Type = adTypeText
            .Charset = "shift-jis"
            ToSJIS = .ReadText()
        End With
    Else
        ToSJIS = ""
    End If
End Function

'隠蔽対象か
Public Function IsHiddenCol(Name As String) As Boolean
    Dim varTmp As Variant
    IsHiddenCol = False
    For Each varTmp In Split(ThisHiddenColNames, ",")
        If varTmp = Name Then
            IsHiddenCol = True
            Exit For
        End If
    Next
End Function

