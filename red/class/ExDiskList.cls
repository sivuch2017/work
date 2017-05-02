VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExDiskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DataListInterface

Private Super As DataListInterface
Private Const TABLES As String = "OtherObject,Server,HardWare,ObjectMaster"

'�ΏۃV�[�g�I�u�W�F�N�g
Property Set DataListInterface_SheetObject(ByRef Sheet As Worksheet)
    Set Super.SheetObject = Sheet
End Property
Property Get DataListInterface_SheetObject() As Worksheet
    Set DataListInterface_SheetObject = Super.SheetObject
End Property

'�\���J�n�ʒu
Public Property Let DataListInterface_StartCol(ByVal Address As String)
    Super.StartCol = Address
End Property
Public Property Get DataListInterface_StartCol() As String
    DataListInterface_StartCol = Super.StartCol
End Property

'�ǂݏo��SQL
Public Property Let DataListInterface_QuerySql(ByVal Sql As String)
    Super.QuerySQL = Sql
End Property
Public Property Get DataListInterface_QuerySql() As String
    DataListInterface_QuerySql = Super.QuerySQL
End Property

'SQL�p�����[�^
Public Property Let DataListInterface_QueryParam(ByVal Param As String)
    Super.QueryParam = Param
End Property
Public Property Get DataListInterface_QueryParam() As String
    DataListInterface_QueryParam = Super.QueryParam
End Property

'�B���ΏۃJ������(�J���}��؂�)
Public Property Let DataListInterface_HiddenColNames(ByVal Names As String)
    Super.HiddenColNames = Names
End Property
Public Property Get DataListInterface_HiddenColNames() As String
    DataListInterface_HiddenColNames = Super.HiddenColNames
End Property

'�O���[�v���ΏۃJ������(�J���}��؂�)
Public Property Let DataListInterface_GroupColNames(ByVal Names As String)
    Super.GroupColNames = Names
End Property
Public Property Get DataListInterface_GroupColNames() As String
    DataListInterface_GroupColNames = Super.GroupColNames
End Property

'�O���[�v���Ώۃ��[����(�J���}��؂�)
Public Property Let DataListInterface_GroupRow(ByVal NameAndValue As String)
    Super.GroupRow = NameAndValue
End Property
Public Property Get DataListInterface_GroupRow() As String
    DataListInterface_GroupRow = Super.GroupRow
End Property

'�X�V�ΏۃJ������(�J���}��؂�)
Public Property Let DataListInterface_UpdColNames(ByVal Names As String)
    Super.UpdColNames = Names
End Property
Public Property Get DataListInterface_UpdColNames() As String
    DataListInterface_UpdColNames = Super.UpdColNames
End Property

'�L�[�J������(�J���}��؂�)
Public Property Let DataListInterface_KeyColNames(ByVal Names As String)
    Super.KeyColNames = Names
End Property
Public Property Get DataListInterface_KeyColNames() As String
    DataListInterface_KeyColNames = Super.KeyColNames
End Property

'�\��w�i�F
Public Property Let DataListInterface_TitleThemeColor(ByVal ThemeColor As Variant)
    Super.TitleThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_TitleThemeColor() As Variant
    DataListInterface_TitleThemeColor = Super.TitleThemeColor
End Property

'�f�[�^�w�i�F
Public Property Let DataListInterface_DataThemeColor(ByVal ThemeColor As Variant)
    Super.DataThemeColor = ThemeColor
End Property
Public Property Get DataListInterface_DataThemeColor() As Variant
    DataListInterface_DataThemeColor = Super.DataThemeColor
End Property

'�ڑ���
Public Property Set DataListInterface_Connector(ByRef Database As DatabaseInterface)
    Set Super.Connector = Database
End Property
Public Property Get DataListInterface_Connector() As DatabaseInterface
    Set DataListInterface_Connector = Super.Connector
End Property

'�ꗗ���쐬
Public Sub DataListInterface_CreateList()
    Super.CreateList
End Sub

'�g�Œ�
Public Sub DataListInterface_FreezePanes()
    Super.FreezePanes
End Sub

'�X�V�`�F�b�N
Public Function DataListInterface_Check() As Integer

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim intRow As Integer
    Dim intMaxRow As Integer
    Dim intCol As Integer
    Dim strSQLWk As String
    Dim varCompA As Variant
    Dim varCompB As Variant
    Dim intBool As Integer
    Dim strType(2) As String
    Dim bolComp As Boolean
    
    DataListInterface_Check = 0
    
    strType(0) = "WHERE "
    strType(1) = "AND "

    If Super.QueryParam <> "" Then
        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
        Exit Function
    End If

    With Super.SheetObject
        If .Range(Super.StartCol).Offset(1, 0).Value = "" Then
            intMaxRow = 0
        Else
            intMaxRow = .Range(Super.StartCol).End(xlDown).row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        For intRow = .Range(Super.StartCol).row + 1 To intMaxRow
            intBool = 0
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                If Super.IsKeyCol(.Cells(.Range(Super.StartCol).row, intCol).Value) Then
                    strSQLWk = strSQLWk & strType(intBool) & .Cells(.Range(Super.StartCol).row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                    intBool = 1
                End If
            Next

            Set objCon = Super.Connector.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open strSQLWk, objCon, adOpenStatic, adLockReadOnly, adCmdText
            
            If objRec.EOF Then
                '�V�K�s
                For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                    If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                        .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                        DataListInterface_Check = 1
                    End If
                Next
            Else
                For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                    If IsNull(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value) Then
                        varCompA = ""
                    Else
                        If Super.IsInMeta(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value) Then
                            varCompA = Trim(objRec.Fields(intCol - .Range(Super.StartCol).Column).Value)
                        Else
                            varCompA = objRec.Fields(intCol - .Range(Super.StartCol).Column).Value
                            If IsNumeric(varCompA) Then
                                varCompA = Trim(str(varCompA))
                            End If
                        End If
                    End If
                    If Super.IsInMeta(.Cells(intRow, intCol).Value) Then
                        varCompB = Trim(.Cells(intRow, intCol).Value)
                    Else
                        varCompB = .Cells(intRow, intCol).Value
                    End If
                    If varCompA = varCompB Then
                        If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlColorIndexNone
                        End If
                    Else
                        If .Cells(intRow, intCol).Interior.ThemeColor <> xlThemeColorDark2 Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6
                            DataListInterface_Check = 1
                        Else
                            bolComp = True
                        End If
                    End If
                Next
            End If

            objRec.Close
            objCon.Close
            Set objRec = Nothing
            Set objCon = Nothing

        Next
    End With

    If bolComp Then
        MsgBox "�X�V�s���ڂ��ύX����Ă��܂����������܂��B", vbOKOnly
    End If

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    If Not objCon Is Nothing Then
        If objCon.State = adStateOpen Then
            objCon.Close
        End If
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(Check):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'�X�V
Public Sub DataListInterface_Update()

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim intCol As Integer
    Dim intRow As Integer
    Dim hashTables As Collection
    Dim strTable As Variant
    Dim strKey1, strKey2 As String

    If Super.QueryParam <> "" Then
        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
        Exit Sub
    End If
    
    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    With Super.SheetObject
        For intRow = .Range(Super.StartCol).row + 1 To .Range(Super.StartCol).End(xlDown).row
            Set hashTables = New Collection
            For Each strTable In Split(TABLES, ",")
                hashTables.Add New Collection, strTable
                hashTables(strTable).Add New Collection, "KEY"
                hashTables(strTable).Add New Collection, "VALUE"
            Next
            For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
                If Super.IsKeyCol(.Cells(.Range(Super.StartCol).row, intCol).Value) Then
                    For Each strTable In Split(TABLES, ",")
                        hashTables(strTable).Item("VALUE").Add Trim(.Cells(intRow, intCol)), .Cells(.Range(Super.StartCol).row, intCol).Value
                    Next
                Else
                    strTable = GetColToTableName(.Cells(.Range(Super.StartCol).row, intCol).Value)
                    If .Cells(intRow, intCol).Interior.ThemeColor = xlThemeColorAccent6 Then
                        hashTables(strTable).Item("KEY").Add .Cells(.Range(Super.StartCol).row, intCol).Value
                        hashTables(strTable).Item("VALUE").Add Trim(.Cells(intRow, intCol)), .Cells(.Range(Super.StartCol).row, intCol).Value
                    End If
                End If
            Next
            If Not UpdateRecord(hashTables) Then
                objCon.RollbackTrans
                objCon.Close
                Set objCon = Nothing
                Exit Sub
            End If
        Next
    End With
    objCon.CommitTrans

    objCon.Close
    Set objCon = Nothing

    Exit Sub

ErrHandle:

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(Update):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'�폜
Public Sub DataListInterface_Delete(DeleteRow As Integer)

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim intCol As Integer
    Dim intRow As Integer
    Dim varName As Variant
    Dim strKey1, strKey2 As String
    Dim strSQLORG, strSql As String
    Dim hashWork As Collection
    Dim varKey As Variant

    If Super.QueryParam <> "" Then
        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
        Exit Sub
    End If
    
    Set objCon = Super.Connector.ConnectDatabase
    objCon.BeginTrans
    With Super.SheetObject
        intRow = DeleteRow
        Set hashWork = New Collection
        For intCol = .Range(Super.StartCol).Column To .Range(Super.StartCol).End(xlToRight).Column
            If Super.IsKeyCol(.Cells(.Range(Super.StartCol).row, intCol).Value) Then
                hashWork.Add Trim(.Cells(intRow, intCol)), .Cells(.Range(Super.StartCol).row, intCol).Value
            End If
        Next
        If hashWork.Count < 1 Then
            MsgBox "�w�肳�ꂽ�s�ɃL�[���ڂ����݂��܂���"
            objCon.RollbackTrans
            objCon.Close
            Set objCon = Nothing
            Exit Sub
        End If
        
        strSQLORG = "DELETE OtherObject WHERE "
        strSql = strSQLORG
        For Each varKey In Split(Super.KeyColNames, ",")
            If strSql <> strSQLORG Then
                strSql = strSql & " AND "
            End If
            strSql = strSql & varKey & " = '" & hashWork.Item(varKey) & "'"
        Next
        Super.Connector.ConnectDatabase.Execute strSql
    End With
    objCon.CommitTrans

    objCon.Close
    Set objCon = Nothing

    Exit Sub

ErrHandle:

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(Delete):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'�O���[�v�Ώۂ�
Public Function DataListInterface_IsGroupCol(Name As String) As Boolean
    DataListInterface_IsGroupCol = Super.IsGroupCol(Name)
End Function

'�O���[�v�Ώۂ�
Public Function DataListInterface_IsGroupRow(Name As String, Optional Value As Variant) As Boolean
    DataListInterface_IsGroupRow = Super.IsGroupRow(Name, Value)
End Function

'�X�V�\�t���O
Public Function DataListInterface_IsUpdCol(Name As String) As Boolean
    DataListInterface_IsUpdCol = Super.IsUpdCol(Name)
End Function

'�L�[�J�����t���O
Public Function DataListInterface_IsKeyCol(Name As String) As Boolean
    DataListInterface_IsKeyCol = Super.IsKeyCol(Name)
End Function

'���ꕶ�����܂ނ������񂩁H
Public Function DataListInterface_IsInMeta(Value As String) As Boolean
    DataListInterface_IsInMeta = Super.IsInMeta(Value)
End Function

'������
Private Sub Class_Initialize()
    Set Super = New DataListInterfaceImpl
End Sub

'�X�V�pSQL�擾
Private Function GetBaseSql() As String
    Dim strBuf As String
    Dim intIndex As Integer
    
    strBuf = Super.QuerySQL
    intIndex = InStr(UCase(strBuf), "ORDER")
    If intIndex <> 0 Then
        strBuf = Left(strBuf, intIndex - 1)
    End If
    
    intIndex = InStr(UCase(strBuf), "WHERE")
    If intIndex <> 0 Then
        strBuf = Left(strBuf, intIndex - 1)
    End If
    
    GetBaseSql = strBuf
End Function

'�s�f�[�^��DB�ɔ��f
Private Function UpdateRecord(ByRef hashData As Collection) As Boolean

    On Error GoTo ErrHandle

    Dim objRec As New ADODB.Recordset
    Dim hashWork As Collection
    Dim strSql As String
    Dim strObjId As String
    Dim varKey As Variant
    Dim strTable As Variant
    Dim strFloor, strRack, strPosition As Variant
    
    UpdateRecord = False
    strObjId = ""

    For Each strTable In Split(TABLES, ",")
        If hashData(strTable).Item("KEY").Count > 0 Then
            strSql = "SELECT * FROM " & strTable & " WHERE "
            For Each varKey In Split(Super.KeyColNames, ",")
                If strSql <> "SELECT * FROM " & strTable & " WHERE " Then
                    strSql = strSql & " AND "
                End If
                strSql = strSql & ChangeColName(varKey) & " = '" & hashData(strTable).Item("VALUE").Item(varKey) & "'"
            Next
            objRec.Open strSql, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
            If objRec.EOF Then
                objRec.AddNew
                For Each varKey In Split(Super.KeyColNames, ",")
                    If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
                        If varKey = "object_id" And hashData(strTable).Item("VALUE").Item(varKey) = "*" Then
                            If strObjId = "" Then
                                strObjId = GetObjectId
                            End If
                            objRec.Fields(ChangeColName(varKey)).Value = strObjId
                        Else
                            objRec.Fields(ChangeColName(varKey)).Value = hashData(strTable).Item("VALUE").Item(varKey)
                        End If
                    End If
                Next
                For Each varKey In hashData(strTable).Item("KEY")
                    If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
                        If ChangeColName(varKey) = "row_type" Then
                            objRec.Fields(ChangeColName(varKey)).Value = GetCode(ChangeColName(varKey), hashData(strTable).Item("VALUE").Item(varKey))
                        Else
                            If varKey = "�ݒu�ꏊ" Then
                                SplitPosition hashData(strTable).Item("VALUE").Item(varKey), strFloor, strRack, strPosition
                                objRec.Fields(ChangeColName("floor")).Value = strFloor
                                objRec.Fields(ChangeColName("rack_number")).Value = strRack
                                objRec.Fields(ChangeColName("position")).Value = strPosition
                            Else
                                objRec.Fields(ChangeColName(varKey)).Value = hashData(strTable).Item("VALUE").Item(varKey)
                            End If
                        End If
                    End If
                Next
            Else
                For Each varKey In hashData(strTable).Item("KEY")
                    If hashData(strTable).Item("VALUE").Item(varKey) <> "" Then
                        If ChangeColName(varKey) = "row_type" Then
                            objRec.Fields(ChangeColName(varKey)).Value = GetCode(ChangeColName(varKey), hashData(strTable).Item("VALUE").Item(varKey))
                        Else
                            If varKey = "�ݒu�ꏊ" Then
                                SplitPosition hashData(strTable).Item("VALUE").Item(varKey), strFloor, strRack, strPosition
                                objRec.Fields(ChangeColName("floor")).Value = strFloor
                                objRec.Fields(ChangeColName("rack_number")).Value = strRack
                                objRec.Fields(ChangeColName("position")).Value = strPosition
                            Else
                                objRec.Fields(ChangeColName(varKey)).Value = hashData(strTable).Item("VALUE").Item(varKey)
                            End If
                        End If
                    Else
                        objRec.Fields(ChangeColName(varKey)).Value = Null
                    End If
                Next
            End If
            objRec.Update
            objRec.Close
        End If
    Next
    
    Set objRec = Nothing

    UpdateRecord = True

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    Err.Raise Err.Number, "SimpleList(UpdateRecord):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'�e�[�u�����擾
Public Function GetTableName(ByVal QuerySQL As String) As String
    Dim varItem As Variant
    'BUG FROM��̍ŏ����e�[�u�����ł��邱�Ƃ��O��
    GetTableName = Split(Mid(QuerySQL, InStr(UCase(QuerySQL), "FROM")), " ")(1)
End Function

'��object_id�擾
Private Function GetObjectId() As String
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT RIGHT('00000000000000'+CONVERT(VARCHAR,MIN(CONVERT(BIGINT,object_id))+1),14) AS object_id FROM ObjectMaster WHERE CONVERT(BIGINT,object_id)+1 NOT IN (SELECT object_id FROM ObjectMaster)"
    objRec.Open strSql, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    GetObjectId = objRec.Fields("object_id").Value
    objRec.Close
End Function

'Code�`�F�b�N
Private Function GetCode(ByVal Key As String, ByVal Name As String) As String
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    GetCode = ""
    strSql = "SELECT code FROM Code WHERE column_id = '" & Key & "' AND value = '" & Name & "'"
    objRec.Open strSql, Super.Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    If objRec.EOF Then
    Else
        GetCode = objRec.Fields("code").Value
    End If
    objRec.Close
End Function

'�J�������ϊ�
Private Function ChangeColName(ByVal Name As String)
    Select Case Name
        Case "IP�A�h���X"
            ChangeColName = "ip_address"
        Case "�}�V��"
            ChangeColName = "hardware_name"
        Case "���i�ԍ�"
            ChangeColName = "product_number"
        Case "���Y�R�[�h"
            ChangeColName = "property_number"
        Case "�V���A���ԍ�"
            ChangeColName = "serial_number"
        Case "���(NAS/USB)"
            ChangeColName = "row_type"
        Case "���e"
            ChangeColName = "contents"
        Case "�S����"
            ChangeColName = "person"
        Case "�e��(GB)"
            ChangeColName = "capacity"
        Case "���g�p(GB)"
            ChangeColName = "availability"
        Case Else
            ChangeColName = Name
    End Select
End Function

'�Ώۃe�[�u��
Private Function GetColToTableName(ByVal ColName As String)
    Select Case ColName
        Case "dependence_place"
            GetColToTableName = "HardWare"
        Case "�z�X�g��"
            GetColToTableName = "Address"
        Case "IP�A�h���X"
            GetColToTableName = "Server"
        Case "�}�V��"
            GetColToTableName = "HardWare"
        Case "���i�ԍ�"
            GetColToTableName = "HardWare"
        Case "���Y�R�[�h"
            GetColToTableName = "HardWare"
        Case "�V���A���ԍ�"
            GetColToTableName = "HardWare"
        Case "���(NAS/USB)"
            GetColToTableName = "OtherObject"
        Case "�ݒu�ꏊ"
            GetColToTableName = "HardWare"
        Case "���e"
            GetColToTableName = "ObjectMaster"
        Case "�S����"
            GetColToTableName = "ObjectMaster"
        Case "�e��(GB)"
            GetColToTableName = "OtherObject"
        Case "���g�p(GB)"
            GetColToTableName = "OtherObject"
    End Select
End Function

'�ݒu�ꏊ���
Private Sub SplitPosition(ByVal StrPos As String, ByRef Floor As Variant, ByRef Rack As Variant, ByRef Position As Variant)
    Dim longPos As Long
    longPos = InStr(StrPos, "-")
    If (longPos > 0) Then
        Floor = Left(StrPos, longPos - 1)
        longPos = longPos + 1
    Else
        Floor = ""
        longPos = 1
    End If
    Rack = Mid(StrPos, longPos, 3)
    If Val(Rack) > 142 And Val(Rack) < 158 Then
        longPos = longPos + 3
    Else
        Rack = Null
    End If
    Position = Mid(StrPos, longPos)
    If Position = "0" Then
        Position = ""
    End If
End Sub

'�J�����ɑ΂���e�[�u����(�J���}��؂�)
Public Property Let DataListInterface_ColToTable(ByVal Names As String)
    Super.ColToTable = Names
End Property
Public Property Get DataListInterface_ColToTable() As String
    DataListInterface_ColToTable = Super.ColToTable
End Property

'�J�����ʖ��R���N�V����
Public Property Set DataListInterface_ColAlias(ByRef NameHash As Collection)
    Set Super.ColAlias = NameHash
End Property
Public Property Get DataListInterface_ColAlias() As Collection
    Set DataListInterface_ColAlias = Super.ColAlias
End Property

'�J�����ɑ΂���e�[�u�������擾
Public Function DataListInterface_GetTableNameByCol(Name As String) As String
    Dim intCnt As Variant
    DataListInterface_GetTableNameByCol = ""
    For intCnt = 0 To UBound(Split(Super.UpdColNames, ","))
        If Split(Super.UpdColNames, ",")(intCnt) = Name Then
            DataListInterface_GetTableNameByCol = Split(Super.ColToTable, ",")(intCnt)
            Exit For
        End If
    Next
End Function

