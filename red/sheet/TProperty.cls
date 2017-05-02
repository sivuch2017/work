VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SafeArrayGetDim Lib "oleaut32" (ByRef psa() As Any) As Long

'�����������͊J�n�ʒu
Private Const cellKensaku As String = "A5"
'�ꗗ�\���J�n�ʒu
Private Const cellList As String = "A10"
'�Q�ƃe�[�u��
Private Const nameTable As String = "dbo.Property"
'�\����
Private Const maxCol As Integer = 13
'�\���^�C�v(S=�����AN=���l)
Private Const typeDisp As String = "S,S,S,S,S,N,S,S,S,S,S,S,N"
'�J������
Private Const nameCols As String = "property_number,unit_code,unit_name,property_name_1,property_name_2,quantity,date_of_acquisition,building_name,actual_thing_collation,person,floor,classification,number_of_sheets"
'�L�[���ڃt���O(K=�L�[�AN=�L�[�ȊO)
Private Const keyCol As String = "K,N,N,N,N,N,N,N,N,N,N,N,N"

'�ݒ�N���X
Private objTableMaintenance As TableMaintenance

'������
Public Sub InitSheet()
    Set objTableMaintenance = New TableMaintenance
    Set objTableMaintenance.Sheet = Me
    objTableMaintenance.KensakuCell = cellKensaku
    objTableMaintenance.ListCell = cellList
    objTableMaintenance.TableName = nameTable
    objTableMaintenance.ColumnMax = maxCol
    objTableMaintenance.DisplayTypes = typeDisp
    objTableMaintenance.ColumnNames = nameCols
    objTableMaintenance.KeyColumns = keyCol
End Sub

'�ꗗ�N���A
Public Sub ClearList()
    If objTableMaintenance Is Nothing Then
        InitSheet
    End If
    objTableMaintenance.ClearList
End Sub

'�ꗗ��������
Private Sub ClearButton_Click()
    ClearList
End Sub

'�ꗗ�\������
Private Sub ListButton_Click()
    If objTableMaintenance Is Nothing Then
        InitSheet
    Else
        objTableMaintenance.LoadList
    End If
End Sub

'�X�V����
Private Sub UpdateButton_Click()
    If objTableMaintenance Is Nothing Then
        InitSheet
    Else
        objTableMaintenance.UpdateTable
    End If
End Sub
