VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'初期化
Sub Init(ByRef DataHash As Object)
End Sub

'ディクショナリ内の項目数を取得します
Property Get Count() As Long
End Property

'指定した列がディクショナリに含まれているかどうかを示します
Function Exists(ByVal Key As String) As Boolean
End Function
    
'指定した列に対する値の取得
Property Get Item(ByVal Key As String)
End Property

'指定した列に対する項目を設定
Property Let Item(ByVal Key As String, ByVal Val As Variant)
End Property

'レコード内のすべての項目を含む配列を取得します
Function Items() As Variant
End Function

'レコード内のすべてのキーを含む配列を取得します
Function Keys() As Variant
End Function

'レコード削除
Sub DeleteRecord(ByRef Con As ADODB.Connection)
End Sub

'レコード更新(レコードが無い場合は追加)
Sub UpdateRecord(ByRef Con As ADODB.Connection)
End Sub

