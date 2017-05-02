VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GXServerRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�v���p�e�B(�f�[�^��)
Private varID As Variant                'Server,ObjectMaster,HardWare
Private var�z�X�g�� As Variant          'Address
Private varIP�A�h���X As Variant        'Server,Address
Private var��� As Variant              'ObjectMaster
Private varVM�� As Variant              'Server
Private varVM�T�[�o�[�� As Variant      'Server,Address
Private var�ݏo�˗��� As Variant        'ObjectMaster
Private var�S���� As Variant            'ObjectMaster
Private var���e As Variant              'ObjectMaster
Private var�� As Variant              'ObjectMaster
Private var�\�� As Variant              'ObjectMaster
Private var�}�V�� As Variant            'HardWare
Private var���蓖��CPU As Variant       'Server
Private var���蓖�ă����� As Variant    'Server
Private var�f�B�X�N�e�� As Variant      'Server
Private var���l As Variant              'ObjectMaster
Private var�Ώۃt���O As Variant        'Server
Private varIP�\�[�g�p As Variant        'Server

'�v���p�e�B(�A�g��)
Private diCon As DatabaseInterface

'�����ێ�(���R�[�h)
Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recServer As RecordInterface
Private recAddress As RecordInterface

'�����ێ�(�ޔ�)
Private objCode As Object
Private blnNew As Boolean

'�R���X�g���N�^
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recHardWare = New HardWareRecord
    Set recServer = New ServerRecord
    Set recAddress = New AddressRecord
    blnNew = False
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recHardWare = Nothing
    Set recServer = Nothing
    Set recAddress = Nothing
    Set diCon = Nothing
End Sub

'�v���p�e�B ID
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recHardWare.Item("object_id") = varID
    recServer.Item("object_id") = varID
End Property
Public Property Get ID() As Variant
    ID = varID
End Property

'�v���p�e�B �z�X�g��
Public Property Let �z�X�g��(ByVal varValue As Variant)
    If varValue = "" Then
        var�z�X�g�� = Null
    Else
        var�z�X�g�� = varValue
    End If
    recAddress.Item("host_name") = var�z�X�g��
End Property
Public Property Get �z�X�g��() As Variant
    �z�X�g�� = var�z�X�g��
End Property

'�v���p�e�B IP�A�h���X
Public Property Let IP�A�h���X(ByVal varValue As Variant)
    If varValue = "" Then
        varIP�A�h���X = Null
    Else
        varIP�A�h���X = varValue
    End If
    recServer.Item("ip_address") = varIP�A�h���X
    recAddress.Item("ip_address") = varIP�A�h���X
End Property
Public Property Get IP�A�h���X() As Variant
    IP�A�h���X = varIP�A�h���X
End Property

'�v���p�e�B ���
Public Property Let ���(ByVal varValue As Variant)
    If varValue = "" Then
        var��� = Null
    Else
        var��� = varValue
    End If
    recObjectMatsr.Item("type_id") = GetTypeId(var���)
End Property
Public Property Get ���() As Variant
    ��� = var���
End Property

'�v���p�e�B VM��
Public Property Let VM��(ByVal varValue As Variant)
    If varValue = "" Then
        varVM�� = Null
    Else
        varVM�� = varValue
    End If
    recServer.Item("vm_name") = varVM��
End Property
Public Property Get VM��() As Variant
    VM�� = varVM��
End Property

'�v���p�e�B VM�T�[�o�[��
Public Property Let VM�T�[�o�[��(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        varVM�T�[�o�[�� = Null
        recServer.Item("dependence_place") = Null
    Else
        varVM�T�[�o�[�� = varValue
        varResult = GetObjectIdByHostName(varVM�T�[�o�[��)
        If IsNull(varResult) Then
            MsgBox varID & ":�w�肳�ꂽ�ˑ���͑��݂��Ȃ��̂Ŗ�������܂��B"
        ElseIf varResult = varID Then
            MsgBox varID & ":�ˑ��悪�������g�Ȃ̂Ŗ�������܂��B"
        Else
            recServer.Item("dependence_place") = varResult
        End If
    End If
End Property
Public Property Get VM�T�[�o�[��() As Variant
    VM�T�[�o�[�� = varVM�T�[�o�[��
End Property

'�v���p�e�B �ݏo�˗���
Public Property Let �ݏo�˗���(ByVal varValue As Variant)
    If varValue = "" Then
        var�ݏo�˗��� = Null
    Else
        var�ݏo�˗��� = varValue
    End If
    recObjectMatsr.Item("rental_client") = var�ݏo�˗���
End Property
Public Property Get �ݏo�˗���() As Variant
    �ݏo�˗��� = var�ݏo�˗���
End Property

'�v���p�e�B �S����
Public Property Let �S����(ByVal varValue As Variant)
    If varValue = "" Then
        var�S���� = Null
    Else
        var�S���� = varValue
    End If
    recObjectMatsr.Item("person") = var�S����
End Property
Public Property Get �S����() As Variant
    �S���� = var�S����
End Property

'�v���p�e�B ���e
Public Property Let ���e(ByVal varValue As Variant)
    If varValue = "" Then
        var���e = Null
    Else
        var���e = varValue
    End If
    recObjectMatsr.Item("contents") = var���e
End Property
Public Property Get ���e() As Variant
    ���e = var���e
End Property

'�v���p�e�B ��
Public Property Let ��(ByVal varValue As Variant)
    If varValue = "" Then
        var�� = Null
    Else
        var�� = varValue
    End If
    recObjectMatsr.Item("situation") = var��
End Property
Public Property Get ��() As Variant
    �� = var��
End Property

'�v���p�e�B �\��
Public Property Let �\��(ByVal varValue As Variant)
    If varValue = "" Then
        var�\�� = Null
    Else
        var�\�� = varValue
    End If
    recObjectMatsr.Item("future") = var�\��
End Property
Public Property Get �\��() As Variant
    �\�� = var�\��
End Property

'�v���p�e�B �}�V��
Public Property Let �}�V��(ByVal varValue As Variant)
    If varValue = "" Then
        var�}�V�� = Null
    Else
        var�}�V�� = varValue
    End If
    recHardWare.Item("hardware_name") = var�}�V��
End Property
Public Property Get �}�V��() As Variant
    �}�V�� = var�}�V��
End Property

'�v���p�e�B ���蓖��CPU
Public Property Let ���蓖��CPU(ByVal varValue As Variant)
    If varValue = "" Then
        var���蓖��CPU = Null
    Else
        var���蓖��CPU = varValue
    End If
    recServer.Item("cpu") = var���蓖��CPU
End Property
Public Property Get ���蓖��CPU() As Variant
    ���蓖��CPU = var���蓖��CPU
End Property

'�v���p�e�B ���蓖�ă�����
Public Property Let ���蓖�ă�����(ByVal varValue As Variant)
    If varValue = "" Then
        var���蓖�ă����� = Null
    Else
        var���蓖�ă����� = varValue
    End If
    recServer.Item("memory") = var���蓖�ă�����
End Property
Public Property Get ���蓖�ă�����() As Variant
    ���蓖�ă����� = var���蓖�ă�����
End Property

'�v���p�e�B �f�B�X�N�e��
Public Property Let �f�B�X�N�e��(ByVal varValue As Variant)
    If varValue = "" Then
        var�f�B�X�N�e�� = Null
    Else
        var�f�B�X�N�e�� = varValue
    End If
    recServer.Item("capacity") = var�f�B�X�N�e��
End Property
Public Property Get �f�B�X�N�e��() As Variant
    �f�B�X�N�e�� = var�f�B�X�N�e��
End Property

'�v���p�e�B ���l
Public Property Let ���l(ByVal varValue As Variant)
    If varValue = "" Then
        var���l = Null
    Else
        var���l = varValue
    End If
    recObjectMatsr.Item("memo") = var���l
End Property
Public Property Get ���l() As Variant
    ���l = var���l
End Property

'�v���p�e�B �Ώۃt���O
Public Property Let �Ώۃt���O(ByVal varValue As Variant)
    If varValue = "" Then
        var�Ώۃt���O = Null
    Else
        var�Ώۃt���O = varValue
    End If
    recServer.Item("gx_server_flag") = var�Ώۃt���O
End Property
Public Property Get �Ώۃt���O() As Variant
    �Ώۃt���O = var�Ώۃt���O
End Property

'�v���p�e�B IP�\�[�g�p
Public Property Let IP�\�[�g�p(ByVal varValue As Variant)
    varIP�\�[�g�p = varValue
End Property
Public Property Get IP�\�[�g�p() As Variant
    IP�\�[�g�p = varIP�\�[�g�p
End Property

'�v���p�e�B DB���
Public Property Set DB���(ByVal diValue As DatabaseInterface)
    Set diCon = diValue
End Property
Public Property Get DB���() As DatabaseInterface
    Set DB��� = diCon
End Property

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
                Case "SYBASE"
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
    If IsEmpty(var���) Then
        var��� = GetTypeNameByObjectId
    ElseIf var��� <> GetTypeNameByObjectId Then
        blnUpdType = True
    End If
    '�g�����U�N�V�����J�n
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '�ݒ�type���{�́A�����A���ӂ̏ꍇ
        If var��� = "�{��" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                '�Ȃ�
            End If
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            recObjectMatsr.UpdateRecord objCon
            '�V�K�̏ꍇ�A�z�z���ꂽID���Đݒ�
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'HardWare���ڂ��ݒ肳��Ă����HardWare���X�V
            recHardWare.UpdateRecord objCon
            'Server���ڂ��ݒ肳��Ă����Server���X�V
            recServer.UpdateRecord objCon
            'Address���ڂ��ݒ肳��Ă����object_id�Ɋ֘A���鍀�ڂ��N���A��Address���X�V
            If Not IsEmpty(recAddress.Item("ip_address")) Then
                '��U�N���A
                Set recUpdate = New AddressRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                If Not IsNull(recAddress.Item("ip_address")) Then
                    recAddress.Item("object_id") = Null
                    recAddress.Item("dependence_place") = varID
                    recAddress.UpdateRecord objCon
                End If
            '�z�X�g���̂ݕύX����Ă���ꍇ
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '�ݒ�type��VM�ATIB�̏ꍇ
        ElseIf var��� = "VM" Or var��� = "TIB" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                'HardWare
                recHardWare.DeleteRecord objCon
                Set recUpdate = New HardWareRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
            End If
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            recObjectMatsr.UpdateRecord objCon
            '�V�K�̏ꍇ�A�z�z���ꂽID���Đݒ�
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'Server���ڂ��ݒ肳��Ă����Server���X�V
            recServer.UpdateRecord objCon
            'Address���ڂ��ݒ肳��Ă����object_id�Ɋ֘A���鍀�ڂ��N���A��Address���X�V
            If Not IsEmpty(recAddress.Item("ip_address")) Then
                '��U�N���A
                Set recUpdate = New AddressRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                If Not IsNull(recAddress.Item("ip_address")) Then
                    recAddress.Item("object_id") = Null
                    recAddress.Item("dependence_place") = varID
                    recAddress.UpdateRecord objCon
                End If
            '�z�X�g���̂ݕύX����Ă���ꍇ
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '�ݒ�type������ȊO�ꍇ
        Else
            Err.Raise vbObjectError + 512 + 1, "Update", "�ݒ�\��ނł͂���܂���B"
        End If
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
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "DB��񂪐ݒ肳��Ă��܂���B"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "ID���ݒ肳��Ă��܂���B"
    End If
    If varID = "*" Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "�Ώۂɒǉ��pID���܂܂�Ă��܂��B"
    End If
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
    
    'Address
    Set recUpdate = New AddressRecord
    recUpdate.Item("object_id") = varID
    recUpdate.DeleteRecord objCon
    recUpdate.Item("object_id") = Empty
    recUpdate.Item("dependence_place") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'Server
    recServer.DeleteRecord objCon
    Set recUpdate = New ServerRecord
    recUpdate.Item("dependence_place") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'HardWare
    recHardWare.DeleteRecord objCon
    Set recUpdate = New HardWareRecord
    recUpdate.Item("dependence_place") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'ObjectMatsr
    recObjectMatsr.DeleteRecord objCon
    
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
Private Function GetObjectIdByHostName(ByVal ServerName As String) As Variant
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
    objRec.Open "SELECT TOP(1) ISNULL(object_id, dependence_place) AS id FROM Address WHERE host_name = '" & ServerName & "' AND (object_id IS NOT NULL OR dependence_place IS NOT NULL)", objCon, adOpenStatic, adLockReadOnly, adCmdText
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
