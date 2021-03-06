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

'検索条件入力開始位置
Private Const cellKensaku As String = "A5"
'一覧表示開始位置
Private Const cellList As String = "A10"
'参照テーブル
Private Const nameTable As String = "dbo.Server"
'表示列数
Private Const maxCol As Integer = 10
'表示タイプ(S=文字、N=数値)
Private Const typeDisp As String = "S,S,S,S,S,N,N,S,S,S"
'カラム名
Private Const nameCols As String = "object_id,ip_address,hard_id,vm_name,dependence_place,stop_sequence,starting_order,stop_procedure_sheet,start_procedure_sheet,gx_server_flag"
'キー項目フラグ(K=キー、N=キー以外)
Private Const keyCol As String = "K,N,N,N,N,N,N,N,N,N"

'設定クラス
Private objTableMaintenance As TableMaintenance

'初期化
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

'一覧クリア
Public Sub ClearList()
    If objTableMaintenance Is Nothing Then
        InitSheet
    End If
    objTableMaintenance.ClearList
End Sub

'一覧消去押下
Private Sub ClearButton_Click()
    ClearList
End Sub

'一覧表示押下
Private Sub ListButton_Click()
    If objTableMaintenance Is Nothing Then
        InitSheet
    Else
        objTableMaintenance.LoadList
    End If
End Sub

'更新押下
Private Sub UpdateButton_Click()
    If objTableMaintenance Is Nothing Then
        InitSheet
    Else
        objTableMaintenance.UpdateTable
    End If
End Sub

