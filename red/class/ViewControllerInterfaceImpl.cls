VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewControllerInterfaceImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ViewControllerInterface

Dim ThisSheetObject As Worksheet
Dim ThisArrayDataList As Collection
Dim ThisCheckStatus As Boolean
Dim ThisBoolAddAfterSheet As Boolean

'�R���X�g���N�^
Private Sub Class_Initialize()
    ThisBoolAddAfterSheet = True
End Sub

'�ΏۃV�[�g�I�u�W�F�N�g
Public Property Set ViewControllerInterface_SheetObject(ByRef Sheet As Worksheet)
    Set ThisSheetObject = Sheet
End Property
Public Property Get ViewControllerInterface_SheetObject() As Worksheet
    Set ViewControllerInterface_SheetObject = ThisSheetObject
End Property

'DataList�I�u�W�F�N�g
Public Property Set ViewControllerInterface_ArrayDataList(ByRef ColArray As Collection)
    Set ThisArrayDataList = ColArray
End Property
Public Property Get ViewControllerInterface_ArrayDataList() As Collection
    Set ViewControllerInterface_ArrayDataList = ThisArrayDataList
End Property

'DataList�I�u�W�F�N�g��ǉ�
Public Sub ViewControllerInterface_AddDataList(ByRef objDataList As DataListInterface)
    If ThisArrayDataList Is Nothing Then
        Set ThisArrayDataList = New Collection
    End If
    ThisArrayDataList.Add objDataList
End Sub

'�V�[�g���Ō�ɒǉ����邩
Public Property Let ViewControllerInterface_BoolAddAfterSheet(ByVal Flag As Boolean)
    ThisBoolAddAfterSheet = Flag
End Property
Public Property Get ViewControllerInterface_BoolAddAfterSheet() As Boolean
    ViewControllerInterface_BoolAddAfterSheet = ThisBoolAddAfterSheet
End Property

'�V�[�g���쐬
Public Sub ViewControllerInterface_CreateSheet(ByVal SheetName As String)
    
    Dim wsTmp As Worksheet
    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '�ݒ�`�F�b�N
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(CreateSheet):", "�K�v�Ȑݒ肪����Ă��܂���B"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(CreateSheet):", "�K�v�Ȑݒ肪����Ă��܂���B"
        End If
    End If
    
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    '�X�V�`�F�b�N�N���A
    ThisCheckStatus = False
    
    '�����V�[�g�폜
    For Each wsTmp In Worksheets
        If wsTmp.Name = SheetName Then
            Application.DisplayAlerts = False
            wsTmp.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    '�V�K�V�[�g�쐬
    Template.Visible = xlSheetVisible
    If ViewControllerInterface_BoolAddAfterSheet Then
        Template.Copy After:=Sheets(Worksheets.Count)
    Else
        Template.Copy Sheets(1)
    End If
    Set ViewControllerInterface_SheetObject = ActiveSheet
    Template.Visible = xlSheetHidden
    ThisSheetObject.Name = SheetName

    For intIndex = 1 To ThisArrayDataList.Count
        Set objDataList = ThisArrayDataList(intIndex)
        Set objDataList.SheetObject = ThisSheetObject
        objDataList.CreateList
    Next

    '�V�[�g����
    ActiveWindow.Zoom = 90
    Set objDataList = ThisArrayDataList(ThisArrayDataList.Count)
    objDataList.FreezePanes
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

End Sub

'�ύX�ӏ��`�F�b�N
Public Sub ViewControllerInterface_CheckSheet()
    
    On Error GoTo ErrHandle

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    Dim intRtn As Integer
    
    ThisCheckStatus = False
    
    '�ݒ�`�F�b�N
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(CheckSheet):", "�K�v�Ȑݒ肪����Ă��܂���B"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(CheckSheet):", "�K�v�Ȑݒ肪����Ă��܂���B"
        End If
    End If
    
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    For intIndex = 1 To ThisArrayDataList.Count
        Set objDataList = ThisArrayDataList(intIndex)
        intRtn = objDataList.Check
        If intRtn > 0 Then
            ThisCheckStatus = True
        End If
        If intRtn < 0 Then
            ThisCheckStatus = False
            Exit For
        End If
    Next
    
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

    Exit Sub

ErrHandle:

    ThisCheckStatus = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True
    
    Err.Raise Err.Number, "ViewControllerInterfaceImpl(CheckSheet):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'DB�X�V
Public Sub ViewControllerInterface_UpdateDatabase()

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '�X�V�`�F�b�N�ς݂�
    If ThisCheckStatus Then
        
        '�ݒ�`�F�b�N
        If ThisArrayDataList Is Nothing Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(UpdateDatabase):", "�K�v�Ȑݒ肪����Ă��܂���B"
        Else
            If ThisArrayDataList.Count < 1 Then
                Err.Raise 10000, "ViewControllerInterfaceImpl(UpdateDatabase):", "�K�v�Ȑݒ肪����Ă��܂���B"
            End If
        End If
        
        Windows.Application.ScreenUpdating = False
        Windows.Application.Calculation = xlCalculationAutomatic
        Windows.Application.Calculation = xlCalculationManual
        
        For intIndex = 1 To ThisArrayDataList.Count
            Set objDataList = ThisArrayDataList(intIndex)
            objDataList.Update
        Next
        
        Windows.Application.Calculation = xlCalculationAutomatic
        Windows.Application.ScreenUpdating = True
        
        ViewControllerInterface_CreateSheet ThisSheetObject.Name
    
    Else
        ViewControllerInterface_CheckSheet
        If ThisCheckStatus Then
            MsgBox "�X�V�ӏ����m�F������A������x�X�V���������{���ĉ������B", vbOKOnly, "���"
        Else
            MsgBox "�X�V�ӏ��͂���܂���B", vbOKOnly, "���"
        End If
    End If

End Sub

'���R�[�h�폜
Public Sub ViewControllerInterface_DeleteRecord(ByVal DeleteRow As Integer)

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '�ݒ�`�F�b�N
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(DeleteRecord):", "�K�v�Ȑݒ肪����Ă��܂���B"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(DeleteRecord):", "�K�v�Ȑݒ肪����Ă��܂���B"
        End If
    End If
    
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    For intIndex = 1 To ThisArrayDataList.Count
        Set objDataList = ThisArrayDataList(intIndex)
        objDataList.Delete (DeleteRow)
    Next
    
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True
    
    ViewControllerInterface_CreateSheet ThisSheetObject.Name

End Sub
