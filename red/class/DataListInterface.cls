VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataListInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.0
Option Explicit

'�ΏۃV�[�g�I�u�W�F�N�g
Property Set SheetObject(ByRef Sheet As Worksheet)
End Property
Property Get SheetObject() As Worksheet
End Property

'�\���J�n�ʒu
Public Property Let StartCol(ByVal Address As String)
End Property
Public Property Get StartCol() As String
End Property

'�ǂݏo��SQL
Public Property Let QuerySQL(ByVal Sql As String)
End Property
Public Property Get QuerySQL() As String
End Property

'SQL�p�����[�^
Public Property Let QueryParam(ByVal Param As String)
End Property
Public Property Get QueryParam() As String
End Property

'�B���ΏۃJ������(�J���}��؂�)
Public Property Let HiddenColNames(ByVal Names As String)
End Property
Public Property Get HiddenColNames() As String
End Property

'�O���[�v���ΏۃJ������(�J���}��؂�)
Public Property Let GroupColNames(ByVal Names As String)
End Property
Public Property Get GroupColNames() As String
End Property

'�O���[�v���Ώۃ��[����(�p�C�v�A�J���}��؂�)
'�J������,����(EQ or NE),�l|�J������,����(EQ or NE),�l...
Public Property Let GroupRow(ByVal NameAndValue As String)
End Property
Public Property Get GroupRow() As String
End Property

'�X�V�ΏۃJ������(�J���}��؂�)
Public Property Let UpdColNames(ByVal Names As String)
End Property
Public Property Get UpdColNames() As String
End Property

'�L�[�J������(�J���}��؂�)
Public Property Let KeyColNames(ByVal Names As String)
End Property
Public Property Get KeyColNames() As String
End Property

'�\��w�i�F
Public Property Let TitleThemeColor(ByVal ThemeColor As Variant)
End Property
Public Property Get TitleThemeColor() As Variant
End Property

'�f�[�^�w�i�F
Public Property Let DataThemeColor(ByVal ThemeColor As Variant)
End Property
Public Property Get DataThemeColor() As Variant
End Property

'�ڑ���
Public Property Set Connector(ByRef Database As DatabaseInterface)
End Property
Public Property Get Connector() As DatabaseInterface
End Property

'�J�����ɑ΂���e�[�u����(�J���}��؂�)
Public Property Let ColToTable(ByVal Names As String)
End Property
Public Property Get ColToTable() As String
End Property

'�J�����ʖ��R���N�V����
Public Property Set ColAlias(ByRef NameHash As Collection)
End Property
Public Property Get ColAlias() As Collection
End Property

'�J�����ɑ΂���e�[�u�������擾
Public Function GetTableNameByCol(Name As String) As String
End Function

'�O���[�v�Ώۂ�
Public Function IsGroupCol(Name As String) As Boolean
End Function

'�O���[�v�Ώۂ�
Public Function IsGroupRow(Name As String, Optional Value As Variant) As Boolean
End Function

'�X�V�\�t���O
Public Function IsUpdCol(Name As String) As Boolean
End Function

'�L�[�J�����t���O
Public Function IsKeyCol(Name As String) As Boolean
End Function

'���ꕶ�����܂ނ������񂩁H
Public Function IsInMeta(Value As String) As Boolean
End Function

'�ꗗ���쐬
Public Sub CreateList()
End Sub

'�g�Œ�
Public Sub FreezePanes()
End Sub

'�X�V�`�F�b�N
Public Function Check() As Integer
End Function

'�X�V
Public Sub Update()
End Sub

'�폜
Public Sub Delete(DeleteRow As Integer)
End Sub

