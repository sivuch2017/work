VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SoftwareListRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ver 20160408
Option Explicit

'�v���p�e�B(�f�[�^��)
Private varID As Variant             'Software
Private var�ˑ��� As Variant         'ObjectMaster,Software
Private var�\�t�g�E�F�A�� As Variant 'Software
Private var���[�J�[ As Variant       'Software
Private var�V���A���ԍ� As Variant   'Software

Private recObjectMatsr As RecordInterface
Private recSoftWare As RecordInterface

'�v���p�e�B(�A�g��)
Private diCon As DatabaseInterface

'�����ێ�
Private objCode As Object
Private blnNew As Boolean

'�R���X�g���N�^
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recSoftWare = New SoftWareRecord
    blnNew = False
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recSoftWare = Nothing
    Set diCon = Nothing
End Sub

'�v���p�e�B ��������
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recSoftWare.Item("object_id") = varID
End Property

Public Property Get ID() As Variant
    ID = varID
End Property

Public Property Let �ˑ���(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        var�ˑ��� = Null
        recSoftWare.Item("dependence_place") = Null
    Else
        var�ˑ��� = varValue
        recSoftWare.Item("dependence_place") = varValue
    End If
End Property

Public Property Get �ˑ���() As Variant
    �ˑ��� = var�ˑ���
End Property

Public Property Let �\�t�g�E�F�A��(ByVal varValue As Variant)
    If varValue = "" Then
        var�\�t�g�E�F�A�� = Null
    Else
        var�\�t�g�E�F�A�� = varValue
    End If
    recSoftWare.Item("software_name") = var�\�t�g�E�F�A��
End Property

Public Property Get �\�t�g�E�F�A��() As Variant
    �\�t�g�E�F�A�� = var�\�t�g�E�F�A��
End Property

Public Property Let ���[�J�[(ByVal varValue As Variant)
    If varValue = "" Then
        var���[�J�[ = Null
    Else
        var���[�J�[ = varValue
    End If
    recSoftWare.Item("maker") = var���[�J�[
End Property

Public Property Let �V���A���ԍ�(ByVal varValue As Variant)
    If varValue = "" Then
        var�V���A���ԍ� = Null
    Else
        var�V���A���ԍ� = varValue
    End If
    recSoftWare.Item("serial_number") = var�V���A���ԍ�
End Property

Public Property Get �V���A���ԍ�() As Variant
    �V���A���ԍ� = var�V���A���ԍ�
End Property

Public Property Get ���[�J�[() As Variant
    ���[�J�[ = var���[�J�[
End Property

Public Property Set DB���(ByVal diValue As DatabaseInterface)
    Set diCon = diValue
End Property

Public Property Get DB���() As DatabaseInterface
    Set DB��� = diCon
End Property
'�v���p�e�B �����܂�

'Code�ϊ�
Private Function GetTypeId(ByVal varName) As String
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    If objCode.Count > 0 Then
        If objCode.Exists(varName) Then
            GetTypeId = objCode(varName)
        Else
            GetTypeId = "99"
        End If
    Else
        If diCon Is Nothing Then
            Select Case varName
                Case "�{��"
                    GetTypeId = "01"
                Case "VM"
                    GetTypeId = "02"
                Case "����"
                    GetTypeId = "03"
                Case "����"
                    GetTypeId = "04"
                Case "IP"
                    GetTypeId = "05"
                Case "TIB"
                    GetTypeId = "06"
                Case "�\�t�g"
                    GetTypeId = "07"
                Case "USER"
                    GetTypeId = "08"
                Case "���L�t�H���_"
                    GetTypeId = "09"
                Case Else
                    GetTypeId = "99"
            End Select
        Else
            Set objCon = diCon.ConnectDatabase
            Set objRec = New ADODB.Recordset
            objRec.Open "SELECT value,code FROM Code WHERE column_id='type_id'", objCon, adOpenStatic, adLockReadOnly, adCmdText
            Do Until objRec.EOF
                objCode.Add objRec.Fields("value").Value, objRec.Fields("code").Value
                objRec.MoveNext
            Loop
            objRec.Close
            objCon.Close
            Set objRec = Nothing
            Set objCon = Nothing
            If objCode.Exists(varName) Then
                GetTypeId = objCode(varName)
            Else
                GetTypeId = "99"
            End If
        End If
    End If
End Function

'�X�V
Public Sub Update()
    On Error GoTo ErrHandle
    Dim objCon As ADODB.Connection
    Dim strSql As String
    Dim varSwap As Variant
    Dim blnNew As Boolean
    Dim blnUpdType As Boolean
    Dim recUpdate As RecordInterface
    Dim strType
    blnNew = False
    blnUpdType = False
    'DB�ڑ���񂪐ݒ肳��Ă��Ȃ��ꍇ�̓G���[
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Update", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    'object_id����̏ꍇ�̓G���[
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Update", "ID���ݒ肳��Ă��܂���B"
    End If
    'type_id�擾
    strType = GetTypeId("�\�t�g")
    '�g�����U�N�V�����J�n
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '�V�K�̏ꍇ�AObjectMaster�ɓo�^���z�z���ꂽID���Đݒ�
        If varID = "*" Then
            recObjectMatsr.Item("type_id") = strType
            recObjectMatsr.Item("contents") = var�\�t�g�E�F�A��
            recObjectMatsr.UpdateRecord objCon
            ID = recObjectMatsr.Item("object_id")
        Else
            'ID�w���ObjectMaster��object_id�������ꍇ�̓G���[
            If ExistsObjectId(varID) = False Then
                Err.Raise vbObjectError + 512 + 1, "Update", "�w��ID���}�X�^�[�ɑ��݂��܂���B"
            End If
        End If
        '�ˑ���w���ObjectMaster��object_id�������ꍇ�̓G���[
        If Not IsEmpty(var�ˑ���) Then
            If ExistsObjectId(var�ˑ���) = False Then
                Err.Raise vbObjectError + 512 + 1, "Update", "�w��ˑ��悪�}�X�^�[�ɑ��݂��܂���B"
            End If
        End If
        'SoftWare���ڂ��ݒ肳��Ă����SoftWare���X�V
        recSoftWare.UpdateRecord objCon
    '�g�����U�N�V�����I��
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
    Err.Raise vbObjectError + 512 + 1, "Update", Err.Description
End Sub

'�폜
Public Sub Delete()
    On Error GoTo ErrHandle
    Dim objCon As ADODB.Connection
    Dim recUpdate As RecordInterface
    Dim strType As String
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "ID���ݒ肳��Ă��܂���B"
    End If
    If varID = "*" Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "�Ώۂɒǉ��pID���܂܂�Ă��܂��B"
    End If
    
    '��ʎ擾
    strType = GetTypeNameByObjectId
    
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        'SoftWare
        recSoftWare.DeleteRecord objCon
        'ObjectMatsr
        If strType = "�\�t�g" Then
            recObjectMatsr.DeleteRecord objCon
        End If
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
    Err.Raise vbObjectError + 512 + 1, "Delete", Err.Description
End Sub

'����type_id�擾
Private Function GetTypeNameByObjectId() As String
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    GetTypeNameByObjectId = ""
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetTypeNameByObjectId", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetTypeNameByObjectId", "ID���ݒ肳��Ă��܂���B"
    End If
    Set objCon = diCon.ConnectDatabase
    Set objRec = New ADODB.Recordset
    objRec.Open "SELECT value FROM ObjectMaster LEFT JOIN Code ON ObjectMaster.type_id = Code.code AND Code.column_id = 'type_id' WHERE ObjectMaster.object_id = '" & varID & "'", objCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not objRec.EOF Then
        GetTypeNameByObjectId = objRec.Fields("value").Value
    End If
    objRec.Close
    objCon.Close
    Set objRec = Nothing
    Set objCon = Nothing
End Function

'����ip_address�擾
Private Function GetIpByObjectId() As String
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    GetIpByObjectId = ""
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetIpByObjectId", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetIpByObjectId", "ID���ݒ肳��Ă��܂���B"
    End If
    Set objCon = diCon.ConnectDatabase
    If objCon.State <> adStateClosed Then
        blnOpen = True
    Else
        blnOpen = False
    End If
    Set objRec = New ADODB.Recordset
    objRec.Open "SELECT ip_address FROM Server WHERE object_id = '" & varID & "'", objCon, adOpenStatic, adLockReadOnly, adCmdText
    If objRec.EOF Then
        objRec.Close
        objRec.Open "SELECT ip_address FROM Address WHERE object_id = '" & varID & "'", objCon, adOpenStatic, adLockReadOnly, adCmdText
        If Not objRec.EOF Then
            GetIpByObjectId = objRec.Fields("ip_address").Value
        End If
    Else
        GetIpByObjectId = objRec.Fields("ip_address").Value
    End If
    objRec.Close
    Set objRec = Nothing
    If Not blnOpen Then
        objCon.Close
        Set objCon = Nothing
    End If
End Function

'�z�X�g������object_id���擾
Private Function GetObjectIdByHostName() As Variant
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    GetObjectIdByHostName = Null
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetObjectIdByHostName", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetObjectIdByHostName", "ID���ݒ肳��Ă��܂���B"
    End If
    Set objCon = diCon.ConnectDatabase
    If objCon.State <> adStateClosed Then
        blnOpen = True
    Else
        blnOpen = False
    End If
    Set objRec = New ADODB.Recordset
    objRec.Open "SELECT TOP(1) ISNULL(object_id, dependence_place) AS id FROM Address WHERE host_name = '" & var�ˑ��� & "' AND (object_id IS NOT NULL OR dependence_place IS NOT NULL)", objCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not objRec.EOF Then
        GetObjectIdByHostName = objRec.Fields("id").Value
    End If
    objRec.Close
    Set objRec = Nothing
    If Not blnOpen Then
        objCon.Close
        Set objCon = Nothing
    End If
End Function

'ObjectMaster��object_id�����݂��邩
Private Function ExistsObjectId(ByVal varObjId As Variant) As Boolean
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    ExistsObjectId = False
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "ExistsObjectId", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    Set objCon = diCon.ConnectDatabase
    If objCon.State <> adStateClosed Then
        blnOpen = True
    Else
        blnOpen = False
    End If
    Set objRec = New ADODB.Recordset
    objRec.Open "SELECT TOP(1) object_id AS id FROM ObjectMaster WHERE object_id = '" & varObjId & "'", objCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not objRec.EOF Then
        ExistsObjectId = True
    End If
    objRec.Close
    Set objRec = Nothing
    If Not blnOpen Then
        objCon.Close
        Set objCon = Nothing
    End If
End Function