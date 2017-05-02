VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements RecordInterface

'親クラス
Private Super As RecordInterface

'固定情報
Private Const TBL_NAME As String = "Property"
Private Const CLS_NAME As String = "PropertyRecord"

'コンストラクタ
Sub Class_Initialize()
    Set Super = New RecordInterfaceImpl
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'初期化
Private Sub RecordInterface_Init(ByRef DataHash As Object)
    DataHash("property_number") = Empty
    DataHash("unit_code") = Empty
    DataHash("unit_name") = Empty
    DataHash("property_name_1") = Empty
    DataHash("property_name_2") = Empty
    DataHash("quantity") = Empty
    DataHash("date_of_acquisition") = Empty
    DataHash("building_name") = Empty
    DataHash("actual_thing_collation") = Empty
    DataHash("person") = Empty
    DataHash("floor") = Empty
    DataHash("classification") = Empty
    DataHash("number_of_sheets") = Empty
    Super.Init DataHash
End Sub

'ディクショナリ内の項目数を取得します
Property Get RecordInterface_Count() As Long
    RecordInterface_Count = Super.Count
End Property

'指定した列がディクショナリに含まれているかどうかを示します
Function RecordInterface_Exists(ByVal Key As String) As Boolean
    RecordInterface_Exists = Super.Exists(Key)
End Function
    
'指定した列に対する値の取得
Property Get RecordInterface_Item(ByVal Key As String)
    If IsObject(Super.Item(Key)) Then
        Set RecordInterface_Item = Super.Item(Key)
    Else
        RecordInterface_Item = Super.Item(Key)
    End If
End Property

'指定した列に対する項目を設定
Property Let RecordInterface_Item(ByVal Key As String, ByVal Val As Variant)
    Super.Item(Key) = Val
End Property

'レコード内のすべての項目を含む配列を取得します
Function RecordInterface_Items() As Variant
    RecordInterface_Items = Super.Items
End Function

'レコード内のすべてのキーを含む配列を取得します
Function RecordInterface_Keys() As Variant
    RecordInterface_Keys = Super.Keys
End Function

'レコード削除
Sub RecordInterface_DeleteRecord(ByRef Con As ADODB.Connection)
    Dim objRec As ADODB.Recordset
    If IsEmpty(Super.Item("property_number")) Then
    Else
        Set objRec = New ADODB.Recordset
        objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
        objRec.Filter = "property_number = '" & Super.Item("property_number") & "'"
        'キーを指定しているので複数レコードはない
        If Not objRec.EOF Then
            objRec.Delete
        End If
        objRec.Close
        Set objRec = Nothing
    End If
End Sub

'レコード更新(レコードが無い場合は追加)
Sub RecordInterface_UpdateRecord(ByRef Con As ADODB.Connection)
    Dim objRec As ADODB.Recordset
    Dim varKey As Variant
    
    If IsEmpty(Super.Item("property_number")) Then
    Else
        If IsUpdate Then
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "property_number = '" & Super.Item("property_number") & "'"
            'キーを指定しているので複数レコードはない
            If objRec.EOF Then
                objRec.AddNew
            End If
            For Each varKey In Super.Keys
                If Not IsEmpty(Super.Item(varKey)) Then
                    objRec.Fields(varKey).Value = Super.Item(varKey)
                End If
            Next
            objRec.Update
            objRec.Close
            Set objRec = Nothing
        End If
    End If
End Sub

'更新対象か
Private Function IsUpdate() As Boolean
    Dim varKey As Variant
    IsUpdate = False
    For Each varKey In Super.Keys
        If Not IsEmpty(Super.Item(varKey)) Then
            If varKey <> "property_number" Then
                IsUpdate = True
            End If
        End If
    Next
End Function

