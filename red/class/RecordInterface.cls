VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'������
Sub Init(ByRef DataHash As Object)
End Sub

'�f�B�N�V���i�����̍��ڐ����擾���܂�
Property Get Count() As Long
End Property

'�w�肵���񂪃f�B�N�V���i���Ɋ܂܂�Ă��邩�ǂ����������܂�
Function Exists(ByVal Key As String) As Boolean
End Function
    
'�w�肵����ɑ΂���l�̎擾
Property Get Item(ByVal Key As String)
End Property

'�w�肵����ɑ΂��鍀�ڂ�ݒ�
Property Let Item(ByVal Key As String, ByVal Val As Variant)
End Property

'���R�[�h���̂��ׂĂ̍��ڂ��܂ޔz����擾���܂�
Function Items() As Variant
End Function

'���R�[�h���̂��ׂẴL�[���܂ޔz����擾���܂�
Function Keys() As Variant
End Function

'���R�[�h�폜
Sub DeleteRecord(ByRef Con As ADODB.Connection)
End Sub

'���R�[�h�X�V(���R�[�h�������ꍇ�͒ǉ�)
Sub UpdateRecord(ByRef Con As ADODB.Connection)
End Sub
