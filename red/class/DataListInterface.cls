VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataListInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.0
Option Explicit

'対象シートオブジェクト
Property Set SheetObject(ByRef Sheet As Worksheet)
End Property
Property Get SheetObject() As Worksheet
End Property

'表示開始位置
Public Property Let StartCol(ByVal Address As String)
End Property
Public Property Get StartCol() As String
End Property

'読み出しSQL
Public Property Let QuerySQL(ByVal Sql As String)
End Property
Public Property Get QuerySQL() As String
End Property

'SQLパラメータ
Public Property Let QueryParam(ByVal Param As String)
End Property
Public Property Get QueryParam() As String
End Property

'隠蔽対象カラム名(カンマ区切り)
Public Property Let HiddenColNames(ByVal Names As String)
End Property
Public Property Get HiddenColNames() As String
End Property

'グループ化対象カラム名(カンマ区切り)
Public Property Let GroupColNames(ByVal Names As String)
End Property
Public Property Get GroupColNames() As String
End Property

'グループ化対象ロー条件(パイプ、カンマ区切り)
'カラム名,条件(EQ or NE),値|カラム名,条件(EQ or NE),値...
Public Property Let GroupRow(ByVal NameAndValue As String)
End Property
Public Property Get GroupRow() As String
End Property

'更新対象カラム名(カンマ区切り)
Public Property Let UpdColNames(ByVal Names As String)
End Property
Public Property Get UpdColNames() As String
End Property

'キーカラム名(カンマ区切り)
Public Property Let KeyColNames(ByVal Names As String)
End Property
Public Property Get KeyColNames() As String
End Property

'表題背景色
Public Property Let TitleThemeColor(ByVal ThemeColor As Variant)
End Property
Public Property Get TitleThemeColor() As Variant
End Property

'データ背景色
Public Property Let DataThemeColor(ByVal ThemeColor As Variant)
End Property
Public Property Get DataThemeColor() As Variant
End Property

'接続先
Public Property Set Connector(ByRef Database As DatabaseInterface)
End Property
Public Property Get Connector() As DatabaseInterface
End Property

'カラムに対するテーブル名(カンマ区切り)
Public Property Let ColToTable(ByVal Names As String)
End Property
Public Property Get ColToTable() As String
End Property

'カラム別名コレクション
Public Property Set ColAlias(ByRef NameHash As Collection)
End Property
Public Property Get ColAlias() As Collection
End Property

'カラムに対するテーブル名を取得
Public Function GetTableNameByCol(Name As String) As String
End Function

'グループ対象か
Public Function IsGroupCol(Name As String) As Boolean
End Function

'グループ対象か
Public Function IsGroupRow(Name As String, Optional Value As Variant) As Boolean
End Function

'更新可能フラグ
Public Function IsUpdCol(Name As String) As Boolean
End Function

'キーカラムフラグ
Public Function IsKeyCol(Name As String) As Boolean
End Function

'特殊文字を含むか文字列か？
Public Function IsInMeta(Value As String) As Boolean
End Function

'一覧を作成
Public Sub CreateList()
End Sub

'枠固定
Public Sub FreezePanes()
End Sub

'更新チェック
Public Function Check() As Integer
End Function

'更新
Public Sub Update()
End Sub

'削除
Public Sub Delete(DeleteRow As Integer)
End Sub


