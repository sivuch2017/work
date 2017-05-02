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

'コンストラクタ
Private Sub Class_Initialize()
    ThisBoolAddAfterSheet = True
End Sub

'対象シートオブジェクト
Public Property Set ViewControllerInterface_SheetObject(ByRef Sheet As Worksheet)
    Set ThisSheetObject = Sheet
End Property
Public Property Get ViewControllerInterface_SheetObject() As Worksheet
    Set ViewControllerInterface_SheetObject = ThisSheetObject
End Property

'DataListオブジェクト
Public Property Set ViewControllerInterface_ArrayDataList(ByRef ColArray As Collection)
    Set ThisArrayDataList = ColArray
End Property
Public Property Get ViewControllerInterface_ArrayDataList() As Collection
    Set ViewControllerInterface_ArrayDataList = ThisArrayDataList
End Property

'DataListオブジェクトを追加
Public Sub ViewControllerInterface_AddDataList(ByRef objDataList As DataListInterface)
    If ThisArrayDataList Is Nothing Then
        Set ThisArrayDataList = New Collection
    End If
    ThisArrayDataList.Add objDataList
End Sub

'シートを最後に追加するか
Public Property Let ViewControllerInterface_BoolAddAfterSheet(ByVal Flag As Boolean)
    ThisBoolAddAfterSheet = Flag
End Property
Public Property Get ViewControllerInterface_BoolAddAfterSheet() As Boolean
    ViewControllerInterface_BoolAddAfterSheet = ThisBoolAddAfterSheet
End Property

'シートを作成
Public Sub ViewControllerInterface_CreateSheet(ByVal SheetName As String)
    
    Dim wsTmp As Worksheet
    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '設定チェック
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(CreateSheet):", "必要な設定がされていません。"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(CreateSheet):", "必要な設定がされていません。"
        End If
    End If
    
    Windows.Application.ScreenUpdating = False
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.Calculation = xlCalculationManual
    
    '更新チェッククリア
    ThisCheckStatus = False
    
    '同名シート削除
    For Each wsTmp In Worksheets
        If wsTmp.Name = SheetName Then
            Application.DisplayAlerts = False
            wsTmp.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    '新規シート作成
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

    'シート調整
    ActiveWindow.Zoom = 90
    Set objDataList = ThisArrayDataList(ThisArrayDataList.Count)
    objDataList.FreezePanes
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    
    Windows.Application.Calculation = xlCalculationAutomatic
    Windows.Application.ScreenUpdating = True

End Sub

'変更箇所チェック
Public Sub ViewControllerInterface_CheckSheet()
    
    On Error GoTo ErrHandle

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    Dim intRtn As Integer
    
    ThisCheckStatus = False
    
    '設定チェック
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(CheckSheet):", "必要な設定がされていません。"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(CheckSheet):", "必要な設定がされていません。"
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

'DB更新
Public Sub ViewControllerInterface_UpdateDatabase()

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '更新チェック済みか
    If ThisCheckStatus Then
        
        '設定チェック
        If ThisArrayDataList Is Nothing Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(UpdateDatabase):", "必要な設定がされていません。"
        Else
            If ThisArrayDataList.Count < 1 Then
                Err.Raise 10000, "ViewControllerInterfaceImpl(UpdateDatabase):", "必要な設定がされていません。"
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
            MsgBox "更新箇所を確認した後、もう一度更新処理を実施して下さい。", vbOKOnly, "情報"
        Else
            MsgBox "更新箇所はありません。", vbOKOnly, "情報"
        End If
    End If

End Sub

'レコード削除
Public Sub ViewControllerInterface_DeleteRecord(ByVal DeleteRow As Integer)

    Dim intIndex As Integer
    Dim objDataList As DataListInterface
    
    '設定チェック
    If ThisArrayDataList Is Nothing Then
        Err.Raise 10000, "ViewControllerInterfaceImpl(DeleteRecord):", "必要な設定がされていません。"
    Else
        If ThisArrayDataList.Count < 1 Then
            Err.Raise 10000, "ViewControllerInterfaceImpl(DeleteRecord):", "必要な設定がされていません。"
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

