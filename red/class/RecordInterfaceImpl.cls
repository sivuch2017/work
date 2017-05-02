VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordInterfaceImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements RecordInterface

'CompareMethod列挙体
Public Enum CompareMethod
    BinaryCompare = 0
    DatabaseCompare = 2
    TextCompare = 1
End Enum

'データ
Private objDataHash As Object

'コンストラクタ
Public Sub Class_Initialize()
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'初期化
Public Sub RecordInterface_Init(ByRef DataHash As Object)
    Set objDataHash = DataHash
End Sub

'ディクショナリ内の項目数を取得します
Public Property Get RecordInterface_Count() As Long
    RecordInterface_Count = objDataHash.Count
End Property

'指定した列がディクショナリに含まれているかどうかを示します
Public Function RecordInterface_Exists(ByVal Key As String) As Boolean
    RecordInterface_Exists = objDataHash.Exists(Key)
End Function
    
'指定した列に対する値の取得
Public Property Get RecordInterface_Item(ByVal Key As String)
    If IsObject(objDataHash.Item(Key)) Then
        Set RecordInterface_Item = objDataHash.Item(Key)
    Else
        RecordInterface_Item = objDataHash.Item(Key)
    End If
End Property

'指定した列に対する項目を設定
Public Property Let RecordInterface_Item(ByVal Key As String, ByVal Val As Variant)
    If objDataHash.Exists(Key) Then
        objDataHash.Item(Key) = Val
    Else
        Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.Item", "存在しない列を指定(" & Key & ")"
    End If
End Property

'レコード内のすべての項目を含む配列を取得します
Public Function RecordInterface_Items() As Variant
    RecordInterface_Items = objDataHash.Items
End Function

'レコード内のすべてのキーを含む配列を取得します
Public Function RecordInterface_Keys() As Variant
    RecordInterface_Keys = objDataHash.Keys
End Function

'レコード削除
Public Sub RecordInterface_DeleteRecord(ByRef Con As ADODB.Connection)
    Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.DeleteRecord", "未実装"
End Sub

'レコード更新(レコードが無い場合は追加)
Public Sub RecordInterface_UpdateRecord(ByRef Con As ADODB.Connection)
    Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.UpdateRecord", "未実装"
End Sub
