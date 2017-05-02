VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShareFolderUserLinkRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements RecordInterface

'親クラス
Private Super As RecordInterface

'固定情報
Private Const TBL_NAME As String = "ShareFolderUserLink"
Private Const CLS_NAME As String = "ShareFolderUserLinkRecord"

'コンストラクタ
Sub Class_Initialize()
    Set Super = New RecordInterfaceImpl
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'初期化
Private Sub RecordInterface_Init(ByRef DataHash As Object)
    DataHash("object_id") = Empty
    DataHash("link_id") = Empty
    DataHash("write_flag") = Empty
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
        If IsEmpty(Super.Item("link_id")) Then
            Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is empty."
        Else
            '依存先から対応オブジェクトを削除
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "link_id = '" & Super.Item("link_id") & "'"
            Do Until objRec.EOF
                objRec.Delete
                objRec.MoveNext
            Loop
            objRec.Close
            Set objRec = Nothing
        End If
    Else
        If IsEmpty(Super.Item("link_id")) Then
            '依存先から対応オブジェクトを削除
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "object_id = '" & Super.Item("object_id") & "'"
            Do Until objRec.EOF
                objRec.Delete
                objRec.MoveNext
            Loop
            objRec.Close
            Set objRec = Nothing
        Else
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "object_id = '" & Super.Item("object_id") & "' And link_id = '" & Super.Item("link_id") & "'"
            'キーを指定しているので複数レコードはない
            If Not objRec.EOF Then
                objRec.Delete
            End If
            objRec.Close
            Set objRec = Nothing
        End If
    End If
End Sub

'レコード更新(レコードが無い場合は追加)
Sub RecordInterface_UpdateRecord(ByRef Con As ADODB.Connection)
    Dim objRec As ADODB.Recordset
    Dim varKey As Variant
    
    If IsEmpty(Super.Item("object_id")) Or IsEmpty(Super.Item("link_id")) Then
        Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is empty."
    End If
    If Super.Item("object_id") = "*" Then
        Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is *."
    End If
    
    If IsUpdate Then
        Set objRec = New ADODB.Recordset
        objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "object_id = '" & Super.Item("object_id") & "' And link_id = '" & Super.Item("link_id") & "'"
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
End Sub

'更新対象か
Private Function IsUpdate() As Boolean
    Dim varKey As Variant
    IsUpdate = False
    For Each varKey In Super.Keys
        If Not IsEmpty(Super.Item(varKey)) Then
            If varKey <> "object_id" And varKey <> "link_id" Then
                IsUpdate = True
            End If
        End If
    Next
End Function

