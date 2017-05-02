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

'CompareMethod�񋓑�
Public Enum CompareMethod
    BinaryCompare = 0
    DatabaseCompare = 2
    TextCompare = 1
End Enum

'�f�[�^
Private objDataHash As Object

'�R���X�g���N�^
Public Sub Class_Initialize()
    RecordInterface_Init CreateObject("Scripting.Dictionary")
End Sub

'������
Public Sub RecordInterface_Init(ByRef DataHash As Object)
    Set objDataHash = DataHash
End Sub

'�f�B�N�V���i�����̍��ڐ����擾���܂�
Public Property Get RecordInterface_Count() As Long
    RecordInterface_Count = objDataHash.Count
End Property

'�w�肵���񂪃f�B�N�V���i���Ɋ܂܂�Ă��邩�ǂ����������܂�
Public Function RecordInterface_Exists(ByVal Key As String) As Boolean
    RecordInterface_Exists = objDataHash.Exists(Key)
End Function
    
'�w�肵����ɑ΂���l�̎擾
Public Property Get RecordInterface_Item(ByVal Key As String)
    If IsObject(objDataHash.Item(Key)) Then
        Set RecordInterface_Item = objDataHash.Item(Key)
    Else
        RecordInterface_Item = objDataHash.Item(Key)
    End If
End Property

'�w�肵����ɑ΂��鍀�ڂ�ݒ�
Public Property Let RecordInterface_Item(ByVal Key As String, ByVal Val As Variant)
    If objDataHash.Exists(Key) Then
        objDataHash.Item(Key) = Val
    Else
        Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.Item", "���݂��Ȃ�����w��(" & Key & ")"
    End If
End Property

'���R�[�h���̂��ׂĂ̍��ڂ��܂ޔz����擾���܂�
Public Function RecordInterface_Items() As Variant
    RecordInterface_Items = objDataHash.Items
End Function

'���R�[�h���̂��ׂẴL�[���܂ޔz����擾���܂�
Public Function RecordInterface_Keys() As Variant
    RecordInterface_Keys = objDataHash.Keys
End Function

'���R�[�h�폜
Public Sub RecordInterface_DeleteRecord(ByRef Con As ADODB.Connection)
    Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.DeleteRecord", "������"
End Sub

'���R�[�h�X�V(���R�[�h�������ꍇ�͒ǉ�)
Public Sub RecordInterface_UpdateRecord(ByRef Con As ADODB.Connection)
    Err.Raise vbObjectError + 512 + 1, "RecordInterfaceImpl.UpdateRecord", "������"
End Sub