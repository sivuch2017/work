VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 1.1
Option Explicit

Implements DataListInterface

Private Super As DataListInterface

Const NOEDIT = xlThemeColorDark2
Const TARGET = xlThemeColorAccent6
Const TABLEKEYS = "TABLEKEYS"
Const TABLEVALS = "TABLEVALS"

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

'�O���[�v���Ώۃ��[����(�p�C�v�A�J���}��؂�)
'�J������,����(EQ or NE),�l|�J������,����(EQ or NE),�l...
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
    If DataListInterface_ColToTable = "" Then
        'BUG FROM��̍ŏ����e�[�u�����ł��邱�Ƃ��O��
        DataListInterface_GetTableNameByCol = Split(Mid(DataListInterface_QuerySql, InStr(UCase(DataListInterface_QuerySql), "FROM")), " ")(1)
    Else
        DataListInterface_GetTableNameByCol = Super.GetTableNameByCol(Name)
    End If
End Function

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
    Dim intRow, intMaxRow, intCol As Integer
    Dim strSQLWk, strWrk As String
    Dim varCompA, varCompB As Variant
    Dim bolComp As Boolean
    
    DataListInterface_Check = 0
    
    If DataListInterface_QueryParam <> "" Then
        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
        Exit Function
    End If

    With DataListInterface_SheetObject
        If .Range(DataListInterface_StartCol).Offset(1, 0).Value = "" Then
            intMaxRow = 0
        Else
            intMaxRow = .Range(DataListInterface_StartCol).End(xlDown).Row
        End If
        intRow = intMaxRow + 1
        bolComp = False
        For intRow = .Range(DataListInterface_StartCol).Row + 1 To intMaxRow
            strWrk = "WHERE "
            strSQLWk = GetBaseSql & " "
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If DataListInterface_IsKeyCol(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value) Then
                    strSQLWk = strSQLWk & strWrk & .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value & " = '" & Trim(.Cells(intRow, intCol).Value) & "' "
                    strWrk = "AND "
                End If
            Next

            Set objCon = DataListInterface_Connector.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open strSQLWk, objCon, adOpenStatic, adLockReadOnly, adCmdText
            
            If objRec.EOF Then
                '�V�K�s
                For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                    If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                        .Cells(intRow, intCol).Interior.ThemeColor = TARGET
                        DataListInterface_Check = 1
                    End If
                Next
            Else
                '�����s
                For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                    'DB���l�Z�b�g
                    If IsNull(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value) Then
                        varCompA = ""
                    Else
                        If DataListInterface_IsInMeta(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value) Then
                            varCompA = Trim(objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value)
                        Else
                            varCompA = objRec.Fields(intCol - .Range(DataListInterface_StartCol).Column).Value
                            If IsNumeric(varCompA) Then
                                varCompA = Trim(str(varCompA))
                            End If
                        End If
                    End If
                    '�Z�����l�Z�b�g
                    If DataListInterface_IsInMeta(.Cells(intRow, intCol).Value) Then
                        varCompB = Trim(.Cells(intRow, intCol).Value)
                    Else
                        varCompB = .Cells(intRow, intCol).Value
                    End If
                    '�l��r
                    If varCompA = varCompB Then
                        If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                            .Cells(intRow, intCol).Interior.ThemeColor = xlColorIndexNone
                        End If
                    Else
                        If .Cells(intRow, intCol).Interior.ThemeColor <> NOEDIT Then
                            .Cells(intRow, intCol).Interior.ThemeColor = TARGET
                            DataListInterface_Check = 1
                        Else
                            '����������
                            If Left(varCompA, 1) <> "=" Then
                                bolComp = True
                            End If
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
        strWrk = ""
        If bolComp Then
            strWrk = "�X�V�s���ڂ̕ύX�͖�������܂��B"
        End If
        MsgBox strWrk, vbOKOnly
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

    Dim intCol, intRow As Integer
    Dim hashTables As Collection
    Dim hashTmp As Object
    Dim strTable As String
    Dim varTmp As Variant

    If DataListInterface_QueryParam <> "" Then
        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
        Exit Sub
    End If
    
    Set hashTables = New Collection
    
    With DataListInterface_SheetObject
        '�s���[�v
        For intRow = .Range(DataListInterface_StartCol).Row + 1 To .Range(DataListInterface_StartCol).End(xlDown).Row
            '�s�n�b�V��������
            Set hashTmp = Nothing
            '�����[�v
            For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
                If .Cells(intRow, intCol).Interior.ThemeColor = TARGET Then
                    '�X�V�F���m�����ꍇ�̏���
                    strTable = DataListInterface_GetTableNameByCol(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value)
                    If hashTmp Is Nothing Then
                        '�s�n�b�V���쐬
                        Set hashTmp = CreateObject("Scripting.Dictionary")
                    End If
                    If Not hashTmp.Exists(strTable) Then
                        '�s�n�b�V���Ƀe�[�u���n�b�V���쐬
                        hashTmp.Add strTable, CreateObject("Scripting.Dictionary")
                        hashTmp(strTable).Add TABLEVALS, CreateObject("Scripting.Dictionary")
                    End If
                    '�e�[�u���n�b�V���̃f�[�^���ɃL�[�ƒl���Z�b�g
                    hashTmp(strTable).Item(TABLEVALS).Add GetTableColNameByName(.Cells(.Range(DataListInterface_StartCol).Row, intCol).Value), Trim(.Cells(intRow, intCol).Value)
                    If Not hashTmp(strTable).Exists(TABLEKEYS) Then
                        '�e�[�u���n�b�V���ɃL�[���������ꍇ
                        hashTmp(strTable).Add TABLEKEYS, CreateObject("Scripting.Dictionary")
                        For Each varTmp In Split(DataListInterface_KeyColNames, ",")
                            '�e�[�u���n�b�V���̃L�[���ɃL�[�ƒl���Z�b�g
                            hashTmp(strTable).Item(TABLEKEYS).Add GetTableColNameByName(varTmp), Trim(.Cells(intRow, GetColByName(varTmp)).Value)
                        Next
                    End If
                End If
            Next
            If Not hashTmp Is Nothing Then
                '�s�n�b�V���ǉ�
                hashTables.Add hashTmp
            End If
        Next
    End With
    
    '�X�V����
    UpdateRecord hashTables
    
    Exit Sub

ErrHandle:

    Err.Raise Err.Number, "SimpleList(Update):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'�폜
Public Sub DataListInterface_Delete(DeleteRow As Integer)

' �e�[�u���擾���@���v�����܂ŕۗ�
MsgBox "�����e��"

'    On Error GoTo ErrHandle
'
'    Dim objCon As ADODB.Connection
'    Dim strObjId, strSQL As String
'    Dim hashTmp As Object
'    Dim strTable As Variant
'
'    If DataListInterface_QueryParam <> "" Then
'        MsgBox "�v���V�[�W���[�ɂ͖��Ή�", vbOKOnly
'        Exit Sub
'    End If
'
'    Set hashTmp = CreateObject("Scripting.Dictionary")
'    With DataListInterface_SheetObject
'        For Each strTable In Split(DataListInterface_ColToTable, ",")
'            If Not hashTmp.Exists(strTable) Then
'                strSQL = "DELETE " & strTable & " WHERE "
'                strDmy = ""
'                For Each varKey In Split(DataListInterface_KeyColNames, ",")
'                    strSQL = strSQL & strDmy & GetTableColNameByName(varKey) & " = '" & .Cells(DeleteRow, GetColByName(varKey)).Value & "' "
'                    strDmy = "AND "
'                Next
'                hashTmp.Add strTable, strSQL
'            End If
'        Next
'    End With
'
'    Set objCon = DataListInterface_Connector.ConnectDatabase
'    objCon.BeginTrans
'    For Each strTable In TableNameSort(hashTmp.Keys)
'        objCon.Execute hashTmp(strTable)
'    Next
'    objCon.CommitTrans
'
'    objCon.Close
'    Set objCon = Nothing
'
'    Exit Sub
'
'ErrHandle:
'
'    If Not objCon Is Nothing Then
'        If objCon.State <> adStateClosed Then
'            objCon.RollbackTrans
'            objCon.Close
'        End If
'        Set objCon = Nothing
'    End If
'
'    Err.Raise Err.Number, "SimpleList(Delete):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'������
Private Sub Class_Initialize()
    Set Super = New DataListInterfaceImpl
End Sub

'�X�V�pSQL�擾
Private Function GetBaseSql() As String
    Dim strBuf As String
    Dim intIndex As Integer
    
    strBuf = DataListInterface_QuerySql
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

'�X�V�f�[�^��DB�ɔ��f
Private Sub UpdateRecord(ByVal hashData As Collection)

    On Error GoTo ErrHandle

    Dim objCon As ADODB.Connection
    Dim objRec As New ADODB.Recordset
    Dim hashTmp As Object
    Dim varTable, varKey As Variant
    Dim strObjId, strSql, strDmy As String
    
    Set objCon = DataListInterface_Connector.ConnectDatabase
    objCon.BeginTrans
    
    For Each hashTmp In hashData
        strObjId = ""
        For Each varTable In hashTmp.Keys
            strSql = "SELECT * FROM " & varTable & " WHERE "
            strDmy = ""
            For Each varKey In hashTmp(varTable).Item(TABLEKEYS).Keys
                strSql = strSql & strDmy & varKey & " = '" & hashTmp(varTable).Item(TABLEKEYS).Item(varKey) & "' "
                strDmy = "AND "
            Next
            objRec.Open strSql, objCon, adOpenStatic, adLockOptimistic, adCmdText
            If objRec.EOF Then
                objRec.AddNew
                For Each varKey In hashTmp(varTable).Item(TABLEKEYS)
                    If hashTmp(varTable).Item(TABLEKEYS).Item(varKey) <> "" Then
                        If hashTmp(varTable).Item(TABLEKEYS).Item(varKey) <> "*" Then
                            Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEKEYS).Item(varKey)
                            objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEKEYS).Item(varKey)
                        Else
                            If strObjId = "" Then
                                strObjId = GetObjectId
                            End If
                            Debug.Print varTable & ":" & varKey & ":" & strObjId
                            objRec.Fields(varKey).Value = strObjId
                        End If
                    End If
                Next
                For Each varKey In hashTmp(varTable).Item(TABLEVALS)
                    If hashTmp(varTable).Item(TABLEVALS).Item(varKey) <> "" Then
                        Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                        objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                    End If
                Next
            Else
                For Each varKey In hashTmp(varTable).Item(TABLEVALS)
                    If hashTmp(varTable).Item(TABLEVALS).Item(varKey) <> "" Then
                        Debug.Print varTable & ":" & varKey & ":" & hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                        objRec.Fields(varKey).Value = hashTmp(varTable).Item(TABLEVALS).Item(varKey)
                    Else
                        objRec.Fields(varKey).Value = Null
                    End If
                Next
            End If
            objRec.Update
            objRec.Close
        Next
    Next
    
    Set objRec = Nothing

    objCon.CommitTrans
    objCon.Close
    Set objCon = Nothing
    
    Exit Sub

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    If Not objCon Is Nothing Then
        If objCon.State <> adStateClosed Then
            objCon.RollbackTrans
            objCon.Close
        End If
        Set objCon = Nothing
    End If
    
    Err.Raise Err.Number, "SimpleList(UpdateRecord):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'��object_id�擾
Private Function GetObjectId() As String
    
    On Error GoTo ErrHandle
    
    Dim objRec As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT RIGHT('00000000000000'+CONVERT(VARCHAR,MIN(CONVERT(BIGINT,object_id))+1),14) AS object_id FROM ObjectMaster WHERE CONVERT(BIGINT,object_id)+1 NOT IN (SELECT object_id FROM ObjectMaster)"
    objRec.Open strSql, DataListInterface_Connector.ConnectDatabase, adOpenStatic, adLockOptimistic, adCmdText
    If objRec.EOF Then
    Else
        GetObjectId = objRec.Fields("object_id").Value
    End If
    
    objRec.Close
    Set objRec = Nothing

    Exit Function

ErrHandle:

    If Not objRec Is Nothing Then
        If objRec.State <> adStateClosed Then
            objRec.Close
        End If
        Set objRec = Nothing
    End If

    Err.Raise Err.Number, "SimpleList(GetObjectId):" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'���ږ����猅�ʒu���擾
Private Function GetColByName(ByVal Name As String) As Integer
    Dim intCol As Integer
    GetColByName = -1
    With DataListInterface_SheetObject
        For intCol = .Range(DataListInterface_StartCol).Column To .Range(DataListInterface_StartCol).End(xlToRight).Column
            If .Cells(.Range(DataListInterface_StartCol).Row, intCol).Value = Name Then
                GetColByName = intCol
                Exit For
            End If
        Next
    End With
End Function

'�r���[���ږ�����e�[�u�����ږ����擾
Private Function GetTableColNameByName(ByVal Name As String) As String
    Select Case Name
        Case Else
            GetTableColNameByName = Name
    End Select
End Function

'�e�[�u�����X�V��
Private Sub TableNameSort(ByRef NameArray() As Variant)
    Dim varSwap As Variant
    Dim i, j As Integer
    For i = LBound(NameArray) To UBound(NameArray)
        For j = UBound(NameArray) To i Step -1
            If NameArray(i) = "ObjectMaster" Then
                varSwap = NameArray(i)
                NameArray(i) = NameArray(j)
                NameArray(j) = varSwap
            End If
        Next j
    Next i
End Sub

