VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewControllerInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'対象シートオブジェクト
Public Property Set SheetObject(ByRef Sheet As Worksheet)
End Property
Public Property Get SheetObject() As Worksheet
End Property

'DataListオブジェクト
Public Property Set ArrayDataList(ByRef ColArray As Collection)
End Property
Public Property Get ArrayDataList() As Collection
End Property

'シートを最後に追加するか
Public Property Let BoolAddAfterSheet(ByVal Flag As Boolean)
End Property
Public Property Get BoolAddAfterSheet() As Boolean
End Property

'DataListオブジェクトを追加
Public Sub AddDataList(ByRef objDataList As DataListInterface)
End Sub

'シートを作成
Public Sub CreateSheet(ByVal SheeName As String)
End Sub

'変更箇所チェック
Public Sub CheckSheet()
End Sub

'DB更新
Public Sub UpdateDatabase()
End Sub

'レコード削除
Public Sub DeleteRecord(ByVal DeleteRow As Integer)
End Sub

