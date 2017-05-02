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

'�e�N���X
Private Super As RecordInterface

'�Œ���
Private Const TBL_NAME As String = "Property"
Private Const CLS_NAME As String = "PropertyRecord"

'�R���X�g���N�^
Sub Class_Initialize()
    Set Super = New RecordInterfaceImpl
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'������
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
    If IsEmpty(Super.Item("property_number")) Then
    Else
        Set objRec = New ADODB.Recordset
        objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
        objRec.Filter = "property_number = '" & Super.Item("property_number") & "'"
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
    
    If IsEmpty(Super.Item("property_number")) Then
    Else
        If IsUpdate Then
            Set objRec = New ADODB.Recordset
            objRec.Open TBL_NAME, Con, adOpenKeyset, adLockOptimistic
            objRec.Filter = "property_number = '" & Super.Item("property_number") & "'"
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
    End If
End Sub

'�X�V�Ώۂ�
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
