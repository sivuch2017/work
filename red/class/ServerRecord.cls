VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServerRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements RecordInterface

'親クラス
Private Super As RecordInterface

'固定情報
Private Const TBL_NAME As String = "Server"
Private Const CLS_NAME As String = "ServerRecord"

'コンストラクタ
Sub Class_Initialize()
    Set Super = New RecordInterfaceImpl
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'初期化
Private Sub RecordInterface_Init(ByRef DataHash As Object)
    DataHash("object_id") = Empty
    DataHash("ip_address") = Empty
    DataHash("hard_id") = Empty
    DataHash("vm_name") = Empty
    DataHash("dependence_place") = Empty
    DataHash("stop_sequence") = Empty
    DataHash("starting_order") = Empty
    DataHash("stop_procedure_sheet") = Empty
    DataHash("start_procedure_sheet") = Empty
    DataHash("gx_server_flag") = Empty
    DataHash("cpu") = Empty
    DataHash("memory") = Empty
    DataHash("capacity") = Empty
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
    If IsEmpty(Super.Item("object_id")) Then
        If IsEmpty(Super.Item("dependence_place")) Then
            Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is empty."
        Else
            '依存先から対応オブジェクトを削除
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "dependence_place = '" & Super.Item("dependence_place") & "'"
            Do Until objRec.EOF
                objRec.Fields("dependence_place").Value = Null
                objRec.Update
                objRec.MoveNext
            Loop
            objRec.Close
            Set objRec = Nothing
        End If
    Else
        Set objRec = New ADODB.Recordset
        objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
        objRec.Filter = "object_id = '" & Super.Item("object_id") & "'"
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
    Dim blnInData As Boolean
    
    If IsEmpty(Super.Item("object_id")) Then
        Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is empty."
    End If
    If Super.Item("object_id") = "*" Then
        Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is *."
    End If
    
    If IsUpdate Then
        Set objRec = New ADODB.Recordset
        objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
        objRec.Filter = "object_id = '" & Super.Item("object_id") & "'"
        'object_id、gx_server_flag以外の項目にデータが入っている、もしくは新規レコードの場合に立てるフラグ
        blnInData = False
        'キーを指定しているので複数レコードはない
        If objRec.EOF Then
            objRec.AddNew
            blnInData = True
        End If
        For Each varKey In Super.Keys
            If Not IsEmpty(Super.Item(varKey)) Then
                objRec.Fields(varKey).Value = Super.Item(varKey)
            End If
            If varKey <> "object_id" And varKey <> "gx_server_flag" And Not IsNull(objRec.Fields(varKey).Value) Then
                blnInData = True
            End If
        Next
        '更新対象の規定値以外の項目が全てNULLだった場合削除に変更する
        If blnInData Then
            objRec.Update
        Else
            objRec.Delete
        End If
        objRec.Close
        Set objRec = Nothing
    End If
End Sub

'更新対象か
Private Function IsUpdate() As Boolean
    Dim varKey As Variant
    IsUpdate = False
    For Each varKey In Super.Keys
        If Not IsEmpty(Super.Item(varKey)) Then
            If varKey <> "object_id" Then
                IsUpdate = True
            End If
        End If
    Next
End Function


