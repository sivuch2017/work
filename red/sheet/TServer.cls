VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TServer"
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
Private Const nameTable As String = "dbo.Server"
'�\����
Private Const maxCol As Integer = 10
'�\���^�C�v(S=�����AN=���l)
Private Const typeDisp As String = "S,S,S,S,S,N,N,S,S,S"
'�J������
Private Const nameCols As String = "object_id,ip_address,hard_id,vm_name,dependence_place,stop_sequence,starting_order,stop_procedure_sheet,start_procedure_sheet,gx_server_flag"
'�L�[���ڃt���O(K=�L�[�AN=�L�[�ȊO)
Private Const keyCol As String = "K,N,N,N,N,N,N,N,N,N"

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
