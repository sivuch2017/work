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

'�ΏۃV�[�g�I�u�W�F�N�g
Public Property Set SheetObject(ByRef Sheet As Worksheet)
End Property
Public Property Get SheetObject() As Worksheet
End Property

'DataList�I�u�W�F�N�g
Public Property Set ArrayDataList(ByRef ColArray As Collection)
End Property
Public Property Get ArrayDataList() As Collection
End Property

'�V�[�g���Ō�ɒǉ����邩
Public Property Let BoolAddAfterSheet(ByVal Flag As Boolean)
End Property
Public Property Get BoolAddAfterSheet() As Boolean
End Property

'DataList�I�u�W�F�N�g��ǉ�
Public Sub AddDataList(ByRef objDataList As DataListInterface)
End Sub

'�V�[�g���쐬
Public Sub CreateSheet(ByVal SheeName As String)
End Sub

'�ύX�ӏ��`�F�b�N
Public Sub CheckSheet()
End Sub

'DB�X�V
Public Sub UpdateDatabase()
End Sub

'���R�[�h�폜
Public Sub DeleteRecord(ByVal DeleteRow As Integer)
End Sub
