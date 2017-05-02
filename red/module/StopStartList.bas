Attribute VB_Name = "StopStartList"
Option Explicit

'�V�[�g��
Public Const strSheetName As String = "��d�Ή�"

'�Ώۃe�[�u��
Private Const TABLES As String = "Server,ObjectMaster"

'�X�V�\�J������
Private Const typeUpd As String = "��~��,��~P,�N����,�N��P,��"

'�J�����˃e�[�u��
Private Const ColToTable As String = "Server,Server,Server,Server,ObjectMaster"

'�X�V�\�J������(�e�[�u����)
Private Const tableUpd As String = "stop_sequence,stop_procedure_sheet,starting_order,start_procedure_sheet,situation"

'�X�V�\�t���O
Private Function boolUpdCol(ColName As String) As Boolean
    boolUpdCol = (InStr(typeUpd, ColName) > 0)
End Function

'�X�V�\�J�������擾
Private Function tableCol(ColName As String) As String
    Dim index As Integer
    tableCol = ""
    For index = 0 To UBound(Split(typeUpd, ","))
        If Split(typeUpd, ",")(index) = ColName Then
            tableCol = Split(tableUpd, ",")(index)
            Exit For
        End If
    Next
End Function

'�e�[�u�����擾
Private Function TableName(ColName As String) As String
    Dim index As Integer
    TableName = ""
    For index = 0 To UBound(Split(typeUpd, ","))
        If Split(typeUpd, ",")(index) = ColName Then
            TableName = Split(ColToTable, ",")(index)
            Exit For
        End If
    Next
End Function

'�ꗗ���쐬
Public Sub CreateSheet()
    
    Dim sht As Worksheet
    Dim maxCol As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strId As String
    
    For Each sht In Worksheets
        If sht.Name = strSheetName Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Set sht = Worksheets.Add
    sht.Name = strSheetName
    
    Windows.Application.ScreenUpdating = False
    
    loadTable sht, "SELECT * FROM StopStartList WHERE ��<>N'�j����' OR �� IS NULL ORDER BY ��~��, IP�\�[�g�p", maxCol, "Courier New"
    
    Columns(1).EntireColumn.Hidden = True
   
    For intCol = 1 To maxCol
        If Columns(intCol).EntireColumn.ColumnWidth > 48 Then
            Columns(intCol).EntireColumn.ColumnWidth = 48
        End If
        Range("A1").Offset(0, intCol - 1).WrapText = True
        If boolUpdCol(Range("A1").Offset(0, intCol - 1).Value) Then
        Else
            With Range("A1").Offset(1, intCol - 1).Resize(Range("A1").End(xlDown).row - 1, 1)
                .Interior.ThemeColor = xlThemeColorDark2
            End With
        End If
    Next
    
    Columns(1).GROUP
    Columns(13).GROUP
    For intRow = 2 To Range("A1").End(xlDown).row
        If Cells(intRow, 5).Value = "TIB" Or Cells(intRow, 8).Value = "" Then
            Rows(intRow).GROUP
        End If
    Next
    
    Range("H2").Select
    ActiveWindow.FreezePanes = True
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

    Windows.Application.ScreenUpdating = True

End Sub

'�f�[�^�x�[�X�ƈꗗ���r
Public Sub CheckSheet()
    
    Dim shtWork As Worksheet
    Dim shtList As Worksheet
    Dim celTgt As Range
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtList = shtWork
        End If
    Next
    
    If shtList Is Nothing Then
        MsgBox "�ΏۃV�[�g���ǂݍ��܂�Ă��܂���B", vbOKOnly, "�x��"
        Exit Sub
    End If
    
    CheckList shtList, "SELECT * FROM StopStartList"
    For Each celTgt In Range("A1").Offset(1, 0).Resize(Range("A1").End(xlDown).row - 1, Range("A1").End(xlToRight).Column)
        If celTgt.Interior.ThemeColor <> xlThemeColorAccent6 And Not boolUpdCol(Cells(1, celTgt.Column).Value) Then
            celTgt.Interior.ThemeColor = xlThemeColorDark2
        End If
    Next
    
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=0
    MsgBox "�X�V�ӏ����m�F������A������x�X�V���������{���ĉ������B", vbOKOnly, "���"
    
End Sub

'�f�[�^�x�[�X�X�V
Public Sub updateSheet()

    On Error GoTo ErrHandle

    Dim shtWork As Worksheet
    Dim shtList As Worksheet
    Dim intCol As Integer
    Dim intRow As Integer
    Dim boolKey As Boolean
    Dim boolUpd As Boolean
    Dim boolResult As Boolean
    Dim conQCSDB As New ADODB.Connection
    Dim hashTables As Collection
    Dim Name As Variant
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtList = shtWork
        End If
    Next
    
    If shtList Is Nothing Then
        MsgBox "�ΏۃV�[�g���ǂݍ��܂�Ă��܂���B", vbOKOnly, "�x��"
        Exit Sub
    End If
    
    shtList.Activate
    
    boolKey = False
    boolUpd = False
    For intCol = 1 To Range("A1").End(xlToRight).Column
        For intRow = 2 To Range("A1").End(xlDown).row
            If Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                If boolUpdCol(Cells(1, intCol).Value) Then
                    boolUpd = True
                Else
                    boolKey = True
                End If
            End If
        Next
    Next
    
    If boolKey Then
        MsgBox "�X�V�s���ڂ��ύX����Ă��܂��B", vbOKOnly, "�x��"
        Exit Sub
    End If
    
    If Not boolUpd Then
        CheckSheet
        Exit Sub
    End If
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    conQCSDB.BeginTrans
    For intRow = 2 To Range("A1").End(xlDown).row
        boolResult = True
        Set hashTables = New Collection
        For Each Name In Split(TABLES, ",")
            hashTables.Add New Collection, Name
            hashTables(Name).Add New Collection, "KEY"
            hashTables(Name).Add New Collection, "VALUE"
        Next
        For intCol = 1 To Range("A1").End(xlToRight).Column
            If Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                If hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Count = 0 Then
                    hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Add "object_id"
                    hashTables(TableName(Cells(1, intCol).Value)).Item("VALUE").Add Trim(Cells(intRow, 1)), "object_id"
                End If
                hashTables(TableName(Cells(1, intCol).Value)).Item("KEY").Add tableCol(Cells(1, intCol).Value)
                hashTables(TableName(Cells(1, intCol).Value)).Item("VALUE").Add Trim(Cells(intRow, intCol)), tableCol(Cells(1, intCol).Value)
            End If
        Next
        For Each Name In Split(TABLES, ",")
            If hashTables(Name).Item("KEY").Count > 0 Then
                boolResult = UpdateRecord(conQCSDB, Name, hashTables(Name))
            End If
        Next
        If Not boolResult Then
            Exit For
        End If
    Next
    If boolResult Then
        conQCSDB.CommitTrans
        MsgBox "DB���X�V���܂����B", vbOKOnly, "���"
        CreateSheet
    Else
        conQCSDB.RollbackTrans
    End If

    conQCSDB.Close
    Set conQCSDB = Nothing
    
    Exit Sub

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "�V�X�e���G���["

    If Not conQCSDB Is Nothing Then
        If conQCSDB.State <> adStateClosed Then
            conQCSDB.RollbackTrans
            conQCSDB.Close
        End If
        Set conQCSDB = Nothing
    End If

End Sub

'���R�[�h�X�V
Private Function UpdateRecord(conQCSDB As ADODB.Connection, table As Variant, values As Collection) As Boolean

    On Error GoTo ErrHandle

    Dim rsBase As New ADODB.Recordset
    Dim strSql As String
    Dim dtNow As Date
    Dim intCol As Integer
    Dim strType(2) As String
    Dim intBool As Integer
    Dim Key As Variant
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    UpdateRecord = False

    strSql = "SELECT * FROM " & table & " WHERE object_id = '" & values("VALUE").Item("object_id") & "' "
    Debug.Print strSql
    rsBase.Open strSql, conQCSDB, adOpenStatic, adLockOptimistic, adCmdText
    If rsBase.EOF Then
        MsgBox "�X�V�`�F�b�N��Ƀ��R�[�h���폜����܂���" & vbCrLf & "���R�[�h�X�V��������܂��B", vbOKOnly, "���"
        Exit Function
    Else
        For Each Key In values("KEY")
            If rsBase.Fields(Key).Type = adInteger Then
                If values("VALUE").Item(Key) <> "" Then
                    rsBase.Fields(Key).Value = CInt(values("VALUE").Item(Key))
                Else
                    rsBase.Fields(Key).Value = Null
                End If
            Else
                rsBase.Fields(Key).Value = values("VALUE").Item(Key)
            End If
        Next
        rsBase.Update
    End If

    rsBase.Close
    Set rsBase = Nothing

    UpdateRecord = True

    Exit Function

ErrHandle:

    Debug.Print Err.Source & vbCrLf & Err.Description & vbCrLf
    MsgBox Err.Source & vbCrLf & Err.Description & vbCrLf, vbOKOnly, "�V�X�e���G���["

    If Not rsBase Is Nothing Then
        If rsBase.State <> adStateClosed Then
            rsBase.Close
        End If
        Set rsBase = Nothing
    End If

End Function

'�ғ��`�F�b�N
Public Sub CheckIP()

    Dim shtWork As Worksheet
    Dim shtList As Worksheet
    Dim intCol As Integer
    Dim intRow As Integer
    Dim objCounter As EXCELStatusBar
    
    For Each shtWork In Worksheets
        If shtWork.Name = strSheetName Then
            Set shtList = shtWork
        End If
    Next
    
    If shtList Is Nothing Then
        MsgBox "�ΏۃV�[�g���ǂݍ��܂�Ă��܂���B", vbOKOnly, "�x��"
        Exit Sub
    End If
    
    shtList.Activate
    
    Set objCounter = New EXCELStatusBar
    objCounter.Init Application.WorksheetFunction.RoundUp(Range("A1").End(xlDown).row / 10, 0)
    For intRow = 1 To Range("A1").End(xlDown).row
        DoEvents
        If intRow Mod 10 = 0 Then
            objCounter.CountUp
        End If
        If ThisWorkbook.pingIp(Cells(intRow, 3).Value) Then
            Cells(intRow, 3).Font.Color = vbRed
        Else
            Cells(intRow, 3).Font.Color = vbNormal
        End If
    Next
    Set objCounter = Nothing

End Sub

