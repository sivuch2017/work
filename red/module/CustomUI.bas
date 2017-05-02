Attribute VB_Name = "CustomUI"
Option Explicit

'�V�[�g���F�r���[�I�u�W�F�N�g�̃n�b�V��
Private ArrayWorkSheet As Object

'���{��
Private objRibbon As IRibbonUI

'�V�[�g��ID
Private Enum Names
    ���z�}�V�� = 8
    �\�t�g�E�F�A = 9
End Enum

'�V�[�g��
Private Property Get NameById(ByVal ID As Long) As String
    Select Case ID
        Case Names.���z�}�V��: NameById = "���z�}�V��"
        Case Names.�\�t�g�E�F�A: NameById = "�\�t�g�E�F�A"
    End Select
End Property

Sub Ribbon_onLoad(ribbon As IRibbonUI)
  Set objRibbon = ribbon
  Call objRibbon.ActivateTab("CustomTab")
End Sub

Sub ActivateTab()
    If objRibbon Is Nothing Then
        '�����I�ɂ�IRibbonUI���Đݒ肷�郍�W�b�N��ǉ�
        'https://social.msdn.microsoft.com/Forums/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba?forum=exceldev
    Else
        Call objRibbon.ActivateTab("CustomTab")
    End If
End Sub

'�V�[�g���݃`�F�b�N
Private Function ExistsSheet(ByVal Name As String) As Boolean
    Dim objSheet As Worksheet
    ExistsSheet = False
    For Each objSheet In Worksheets
        If objSheet.Name = Name Then
            ExistsSheet = True
        End If
    Next objSheet
End Function

'�ėp�X�V����
Private Sub �X�V(ByVal ID As Long)
    On Error GoTo ErrHandle
    Dim strName As String
    strName = NameById(ID)
    If ArrayWorkSheet Is Nothing Then
        MsgBox "�����ێ��f�[�^���j������܂����B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B", vbOKOnly, "�x��"
    ElseIf ArrayWorkSheet.Exists(strName) And ExistsSheet(strName) Then
        ArrayWorkSheet.Item(strName).SheetObject.Activate
        If MsgBox("�u" + strName + "�v�V�[�g�̓��e�ōX�V���܂��B" + vbCrLf + "��낵���ł����B", vbOKCancel) = vbOK Then
            ArrayWorkSheet.Item(strName).UpdateDatabase
        End If
    Else
        MsgBox "�u" + strName + "�v�f�[�^���ǂݍ��܂�Ă��܂���B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B", vbOKOnly, "�x��"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'�ėp�폜����
Private Sub �폜(ByVal ID As Long)
    On Error GoTo ErrHandle
    Dim strName As String
    strName = NameById(ID)
    If ArrayWorkSheet Is Nothing Then
        MsgBox "�����ێ��f�[�^���j������܂����B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B", vbOKOnly, "�x��"
    ElseIf ArrayWorkSheet.Exists(strName) And ExistsSheet(strName) Then
        ArrayWorkSheet.Item(strName).SheetObject.Activate
        If MsgBox("�u" + strName + "�v�V�[�g " & ActiveCell.row & "�s�ڂ̃f�[�^���폜���܂��B" + vbCrLf + "��낵���ł����B", vbOKCancel) = vbOK Then
            ArrayWorkSheet.Item(strName).DeleteRecord ActiveCell.row
        End If
    Else
        MsgBox "�u" + strName + "�v�f�[�^���ǂݍ��܂�Ă��܂���B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B", vbOKOnly, "�x��"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Button11(ByVal control As IRibbonControl)
    ThisWorkbook.clearSheets
End Sub

Sub Button12(ByVal control As IRibbonControl)
    '�錾
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New Summary
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM Summary WHERE ��<>N'�j����' OR �� IS NULL"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "P,VM��,�ˑ���,�ݏo�˗���,�S����,Type,�z�X�g��,IP�A�h���X,���e,��,�\��,���b�NNO,�ێ�_��,�V���A���ԍ�,���i�ԍ�,���l,���[�J�[,�}�V��,���Y�ԍ�,�����ݒu�t���A,�Ǘ��O"
    objList.GroupColNames = "ID,VM��,�ˑ���,�ݏo�˗���,�S����,�ێ�,�ێ瑋��,�ێ�_��,�V���A���ԍ�,���i�ԍ�,���l,���[�J�[,���j�b�g�R�[�h,���j�b�g��,���Y���P(���Y����),���Y���Q(���[�J�[���E�^��),����,�擾��,�r����,�����ƍ��p���(�t���@��E�ݒu�ꏊ��),�����Ǘ��S���ҁi�g�p�ҁj��,�����ݒu�t���A,�敪,���x������,IP�\�[�g�p"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "�J����"
    
    '�I�u�W�F�N�g�ێ�
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button13(ByVal control As IRibbonControl)
    ThisWorkbook.openMaintenance
End Sub

Sub Button14(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "�J����" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "�������������N���A����܂����B" & vbCrLf & "�ύX�ӏ���ޔ����A�ꗗ��ǂݒ����Ă��������B"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "�u�J�����v�V�[�g��p�̋@�\�ł�"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'GX�T�[�o�[�ꗗ
Sub Button21(ByVal control As IRibbonControl)
    Dim objDB As DatabaseInterface
    Dim objList As DataListInterface
    Dim objSheet As ViewControllerInterface
    Dim strViewName As String
    Dim strSql As String

    'SQL�쐬
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "    GXServer.ID, "
    strSql = strSql & "    GXServer.�z�X�g��, "
    strSql = strSql & "    GXServer.IP�A�h���X, "
    strSql = strSql & "    GXServer.���, "
    strSql = strSql & "    GXServer.VM��, "
    strSql = strSql & "    GXServer.VM�T�[�o�[��, "
    strSql = strSql & "    GXServer.�ݏo�˗���, "
    strSql = strSql & "    GXServer.�S����, "
    strSql = strSql & "    GXServer.���e, "
    strSql = strSql & "    GXServer.��, "
    strSql = strSql & "    GXServer.�\��, "
    strSql = strSql & "    GXServer.�}�V��, "
    strSql = strSql & "    GXServer.���蓖��CPU, "
    strSql = strSql & "    GXServer.���蓖�ă�����, "
    strSql = strSql & "    GXServer.�f�B�X�N�e��, "
    strSql = strSql & "    SoftWare.software_name AS OS, "
    strSql = strSql & "    STUFF((SELECT CAST(',' AS VARCHAR(max)) + software_name FROM SoftWare WHERE dependence_place = GXServer.ID ORDER BY SoftWare.software_name FOR XML PATH('')),1,1,'') AS �\�t�g, "
    strSql = strSql & "    GXServer.���l, "
    strSql = strSql & "    GXServer.�Ώۃt���O, "
    strSql = strSql & "    GXServer.IP�\�[�g�p "
    strSql = strSql & "FROM "
    strSql = strSql & "    GXServer "
    strSql = strSql & "    LEFT JOIN "
    strSql = strSql & "    SoftWare ON "
    strSql = strSql & "    GXServer.ID = SoftWare.object_id "
    strSql = strSql & "WHERE "
    strSql = strSql & "    ��<>N'�j����' OR "
    strSql = strSql & "    �� IS NULL "
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New GXServer
    objList.StartCol = "A1"
    objList.QuerySQL = strSql
    objList.KeyColNames = "ID"
    objList.UpdColNames = "�z�X�g��,IP�A�h���X,���,VM��,VM�T�[�o�[��,�ݏo�˗���,�S����,���e,��,�\��,�}�V��,���蓖��CPU,���蓖�ă�����,�f�B�X�N�e��,���l,�Ώۃt���O"
    objList.GroupColNames = "���蓖��CPU,���蓖�ă�����,�f�B�X�N�e��,�Ώۃt���O,IP�\�[�g�p"
    objList.GroupRow = "�Ώۃt���O,EQ,0"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "GX���ꗗ"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button22(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "GX���ꗗ" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "�������������N���A����܂����B" & vbCrLf & "�ύX�ӏ���ޔ����A�ꗗ��ǂݒ����Ă��������B"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "�uGX���ꗗ�v�V�[�g��p�̋@�\�ł�"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Button31(ByVal control As IRibbonControl)
    Dim objDB As DatabaseInterface
    Dim objList As DataListInterface
    Dim objSheet As ViewControllerInterface
    Dim strViewName As String
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New HardWareList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM HWList WHERE ���<>N'�j����' OR ��� IS NULL ORDER BY �Ǘ��O,���b�NNo,[GROUP],�ڑ���,���"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "�ڑ���,���,���[�J�[,�}�V��,���i�ԍ�,CPU,������,�f�B�X�N�e��,�V���A���ԍ�,�t���A,���b�NNo,�ʒu,���e,���Y�R�[�h,�ێ�_��,���x���ԍ�,���x������,�Ǘ��O,�v���_�N�g,���,�\��,�ݏo�˗���,�S����,���l"
    objList.HiddenColNames = "GROUP"
    objList.GroupRow = "�ڑ���,NE,"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "�n�[�h�E�F�A"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button32(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "�n�[�h�E�F�A" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "�������������N���A����܂����B" & vbCrLf & "�ύX�ӏ���ޔ����A�ꗗ��ǂݒ����Ă��������B"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "�u�n�[�h�E�F�A�v�V�[�g��p�̋@�\�ł�"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Button41(ByVal control As IRibbonControl)
    IPList.createIpListSheet
End Sub

Sub Button42(ByVal control As IRibbonControl)
    IPList.updateSheet
End Sub

Sub Button43(ByVal control As IRibbonControl)
    IPList.CheckIP
End Sub

Sub Button51(ByVal control As IRibbonControl)
    StopStartList.CreateSheet
End Sub

Sub Button52(ByVal control As IRibbonControl)
    StopStartList.updateSheet
End Sub

Sub Button53(ByVal control As IRibbonControl)
    StopStartList.CheckIP
End Sub

Sub Button61(ByVal control As IRibbonControl)
    '�錾
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New Summary
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM Summary WHERE ��=N'�j����'"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "P,VM��,�ˑ���,�ݏo�˗���,�S����,Type,�z�X�g��,IP�A�h���X,���e,��,�\��,���b�NNO,�ێ�_��,�V���A���ԍ�,���i�ԍ�,���l,���[�J�[,�}�V��,���Y�ԍ�,�����ݒu�t���A,�Ǘ��O"
    objList.GroupColNames = "ID,VM��,�ˑ���,�ݏo�˗���,�S����,�ێ�,�ێ瑋��,�ێ�_��,�V���A���ԍ�,���i�ԍ�,���l,���[�J�[,���j�b�g�R�[�h,���j�b�g��,���Y���P(���Y����),���Y���Q(���[�J�[���E�^��),����,�擾��,�r����,�����ƍ��p���(�t���@��E�ݒu�ꏊ��),�����Ǘ��S���ҁi�g�p�ҁj��,�����ݒu�t���A,�敪,���x������,IP�\�[�g�p"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "�j���ψꗗ"
    
    '�I�u�W�F�N�g�ێ�
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button62(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "�j���ψꗗ" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "�������������N���A����܂����B" & vbCrLf & "�ύX�ӏ���ޔ����A�ꗗ��ǂݒ����Ă��������B"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "�u�j���ψꗗ�v�V�[�g��p�̋@�\�ł�"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Button71(ByVal control As IRibbonControl)
    Dim objDB As DatabaseInterface
    Dim objList As DataListInterface
    Dim objSheet As ViewControllerInterface
    Dim strViewName As String
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New ExDiskList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM ExDiskList"
    objList.GroupColNames = "object_id,dependence_place,IP�\�[�g�p"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "�O���f�B�X�N"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button81(ByVal control As IRibbonControl)
    '�錾
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New VMList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM VMList"
    objList.GroupColNames = "object_id,dependence_place,�Ǘ��O"
    objList.KeyColNames = "object_id"
    objList.UpdColNames = "VM�T�[�o�[��,VM��,�z�X�g��,IP�A�h���X,�ݏo�˗���,�S����,���e,��,�\��,�Ǘ��O,���l"
    objList.ColToTable = "Server,Server,Address,Server,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.���z�}�V��)
    
    '�f�[�^�ێ�
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button82(ByVal control As IRibbonControl)
    �X�V Names.���z�}�V��
End Sub

Sub Button83(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "���z�}�V��" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "�����ێ��f�[�^���j������܂����B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).DeleteRecord ActiveCell.row
                End If
            Next
        End If
    Else
        MsgBox "�u���z�}�V���v�V�[�g��p�̋@�\�ł�"
    End If
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Button84(ByVal control As IRibbonControl)
    Dim intCol As Integer
    Dim intRow As Integer
    
    If ArrayWorkSheet Is Nothing Then
        MsgBox "�����ێ��f�[�^���j������܂����B" + vbCrLf + "�C������ޔ����ēǂݍ��݂��Ă��������B", vbOKOnly, "�x��"
        Exit Sub
    ElseIf ArrayWorkSheet.Exists("���z�}�V��") Then
    Else
        MsgBox "�ΏۃV�[�g���ǂݍ��܂�Ă��܂���B", vbOKOnly, "�x��"
        Exit Sub
    End If
    
    ArrayWorkSheet.Item("���z�}�V��").SheetObject.Activate
    
    For intRow = 2 To Range("A1").End(xlDown).row
        If ThisWorkbook.pingIp(Cells(intRow, 6).Value) Then
            Cells(intRow, 6).Font.Color = RGB(0, 0, 0)
        Else
            Cells(intRow, 6).Font.Color = RGB(255, 0, 0)
        End If
    Next
End Sub

Sub Button85()
    '�錾
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New VMList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM VMHostList UNION ALL SELECT * FROM VMGuestList ORDER BY dependence_place, is_guest"
    objList.GroupColNames = "object_id,dependence_place,is_guest,�Ǘ��O"
'    objList.KeyColNames = "object_id"
'    objList.UpdColNames = "VM�T�[�o�[��,VM��,�z�X�g��,IP�A�h���X,�ݏo�˗���,�S����,���e,��,�\��,�Ǘ��O,���l"
'    objList.ColToTable = "Server,Server,Address,Server,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.���z�}�V��)
    
    '�f�[�^�ێ�
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

'�\�t�g�E�F�A�ꗗ
Sub Button91(ByVal control As IRibbonControl)
    '�錾
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '�ڑ�
    Set objDB = getConObj
    
    '���X�g
    Set objList = New SoftwareList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM SWList ORDER BY ������ID,�ˑ���"
    objList.GroupColNames = "ID,�ˑ���,�V���A���ԍ�,������ID"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "�ˑ���,�\�t�g�E�F�A��,���[�J�[,�V���A���ԍ�"
    objList.ColToTable = "SoftWare,SoftWare,SoftWare,SoftWare"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    '�V�[�g
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.�\�t�g�E�F�A)
    
    '�f�[�^�ێ�
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button92(ByVal control As IRibbonControl)
    �X�V Names.�\�t�g�E�F�A
End Sub

Sub Button93(ByVal control As IRibbonControl)
    �폜 Names.�\�t�g�E�F�A
End Sub

