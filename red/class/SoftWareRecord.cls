VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SoftWareRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements RecordInterface

'�e�N���X
Private Super As RecordInterface

'�Œ���
Private Const TBL_NAME As String = "SoftWare"
Private Const CLS_NAME As String = "SoftWareRecord"

'�R���X�g���N�^
Sub Class_Initialize()
    Set Super = New RecordInterfaceImpl
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'������
Private Sub RecordInterface_Init(ByRef DataHash As Object)
    DataHash("object_id") = Empty
    DataHash("maintenance_id") = Empty
    DataHash("software_name") = Empty
    DataHash("maker") = Empty
    DataHash("serial_number") = Empty
    DataHash("dependence_place") = Empty
    Super.Init DataHash
End Sub

'�f�B�N�V���i�����̍��ڐ����擾���܂�
Property Get RecordInterface_Count() As Long
    RecordInterface_Count = Super.Count
End Property

'�w�肵���񂪃f�B�N�V���i���Ɋ܂܂�Ă��邩�ǂ����������܂�
Function RecordInterface_Exists(ByVal Key As String) As Boolean
    RecordInterface_Exists = Super.Exists(Key)
End Function
    
'�w�肵����ɑ΂���l�̎擾
Property Get RecordInterface_Item(ByVal Key As String)
    If IsObject(Super.Item(Key)) Then
        Set RecordInterface_Item = Super.Item(Key)
    Else
        RecordInterface_Item = Super.Item(Key)
    End If
End Property

'�w�肵����ɑ΂��鍀�ڂ�ݒ�
Property Let RecordInterface_Item(ByVal Key As String, ByVal Val As Variant)
    Super.Item(Key) = Val
End Property

'���R�[�h���̂��ׂĂ̍��ڂ��܂ޔz����擾���܂�
Function RecordInterface_Items() As Variant
    RecordInterface_Items = Super.Items
End Function

'���R�[�h���̂��ׂẴL�[���܂ޔz����擾���܂�
Function RecordInterface_Keys() As Variant
    RecordInterface_Keys = Super.Keys
End Function

'���R�[�h�폜
Sub RecordInterface_DeleteRecord(ByRef Con As ADODB.Connection)
    Dim objRec As ADODB.Recordset
    If IsEmpty(Super.Item("object_id")) Then
        If IsEmpty(Super.Item("dependence_place")) Then
            Err.Raise vbObjectError + 512 + 1, CLS_NAME & ".UpdateRecord", "object_id is empty."
        Else
            '�ˑ��悩��Ή��I�u�W�F�N�g���폜
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
        '�L�[���w�肵�Ă���̂ŕ������R�[�h�͂Ȃ�
        If Not objRec.EOF Then
            objRec.Delete
        End If
        objRec.Close
        Set objRec = Nothing
    End If
End Sub

'���R�[�h�X�V(���R�[�h�������ꍇ�͒ǉ�)
Sub RecordInterface_UpdateRecord(ByRef Con As ADODB.Connection)
    Dim objRec As ADODB.Recordset
    Dim varKey As Variant
    
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
        '�L�[���w�肵�Ă���̂ŕ������R�[�h�͂Ȃ�
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

'�X�V�Ώۂ�
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
