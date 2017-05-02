VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�v���p�e�B(�f�[�^��)
Private varID As Variant                'ObjectMaster,Server,HardWare,Software,OtherObject,ServerUser
Private varNo As Variant                '
Private varP As Variant                 'ObjectMaster
Private varVM�� As Variant              'Server
Private var�ˑ��� As Variant            'Server,HardWare,Address,SoftWare
Private var�ݏo�˗��� As Variant        'ObjectMaster
Private var�S���� As Variant            'ObjectMaster
Private varType As Variant              'ObjectMaster
Private var�z�X�g�� As Variant          'Address
Private varIP�A�h���X As Variant        'Server
Private var���e As Variant              'ObjectMaster
Private var�� As Variant              'ObjectMaster
Private var�\�� As Variant              'ObjectMaster
Private var���b�NNO As Variant          'HardWare
Private var�ێ� As Variant              '
Private var�ێ瑋�� As Variant          '
Private var�ێ�_�� As Variant          'HardWare
Private var�V���A���ԍ� As Variant      'HardWare,SoftWare
Private var���i�ԍ� As Variant          'HardWare
Private var���l As Variant              'ObjectMaster
Private var���[�J�[ As Variant          'HardWare,SoftWare
Private var�}�V�� As Variant            'HardWare
Private var���j�b�g�R�[�h As Variant    '
Private var���j�b�g�� As Variant        '
Private var���Y�ԍ� As Variant          'HardWare
Private var���Y��1 As Variant           '
Private var���Y��2 As Variant           '
Private var���� As Variant              '
Private var�擾�� As Variant            '
Private var�r���� As Variant            '
Private var�����ƍ��p��� As Variant    '
Private var�����Ǘ��S���Җ� As Variant  '
Private var�����ݒu�t���A As Variant    'HardWare
Private var�敪 As Variant              '
Private var���x������ As Variant        '
Private varIP�\�[�g�p As Variant        '
Private var�Ǘ��O As Variant            'ObjectMaster

Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recSoftWare As RecordInterface
Private recServer As RecordInterface
Private recAddress As RecordInterface
Private recOtherObject As RecordInterface
Private recServerUser As RecordInterface
Private recShareFolder As RecordInterface
Private recShareFolderUserLink As RecordInterface

'�v���p�e�B(�A�g��)
Private diCon As DatabaseInterface

'�����ێ�
Private objCode As Object
Private blnNew As Boolean

'�R���X�g���N�^
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recHardWare = New HardWareRecord
    Set recSoftWare = New SoftWareRecord
    Set recServer = New ServerRecord
    Set recAddress = New AddressRecord
    Set recOtherObject = New OtherObjectRecord
    Set recServerUser = New ServerUserRecord
    Set recShareFolder = New ShareFolderRecord
    Set recShareFolderUserLink = New ShareFolderUserLinkRecord
    blnNew = False
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recHardWare = Nothing
    Set recSoftWare = Nothing
    Set recServer = Nothing
    Set recAddress = Nothing
    Set recOtherObject = Nothing
    Set recServerUser = Nothing
    Set recShareFolder = Nothing
    Set recShareFolderUserLink = Nothing
    Set diCon = Nothing
End Sub

'�v���p�e�B ��������
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recHardWare.Item("object_id") = varID
    recSoftWare.Item("object_id") = varID
    recServer.Item("object_id") = varID
    recOtherObject.Item("object_id") = varID
    recServerUser.Item("object_id") = varID
    recShareFolder.Item("object_id") = varID
    recShareFolderUserLink.Item("object_id") = varID
    recAddress.Item("object_id") = varID
End Property

Public Property Get ID() As Variant
    ID = varID
End Property

Public Property Let No(ByVal varValue As Variant)
    varNo = varValue
End Property

Public Property Get No() As Variant
    No = varNo
End Property

Public Property Let P(ByVal varValue As Variant)
    If varValue = "" Then
        varP = Null
    Else
        varP = varValue
    End If
    recObjectMatsr.Item("product") = varP
End Property

Public Property Get P() As Variant
    P = varP
End Property

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

Public Property Let �ˑ���(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        var�ˑ��� = Null
        recHardWare.Item("dependence_place") = Null
        recSoftWare.Item("dependence_place") = Null
        recServer.Item("dependence_place") = Null
        recAddress.Item("dependence_place") = Null
    Else
        var�ˑ��� = varValue
        varResult = GetObjectIdByHostName
        If IsNull(varResult) Then
            MsgBox varID & ":�w�肳�ꂽ�ˑ���͑��݂��Ȃ��̂Ŗ�������܂��B"
        ElseIf varResult = varID Then
            MsgBox varID & ":�ˑ��悪�������g�Ȃ̂Ŗ�������܂��B"
        Else
            recHardWare.Item("dependence_place") = varResult
            recSoftWare.Item("dependence_place") = varResult
            recServer.Item("dependence_place") = varResult
            recAddress.Item("dependence_place") = varResult
        End If
    End If
End Property

Public Property Get �ˑ���() As Variant
    �ˑ��� = var�ˑ���
End Property

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

Public Property Let CodeType(ByVal varValue As Variant)
    If varValue = "" Then
        varType = Null
    Else
        varType = varValue
    End If
    recObjectMatsr.Item("type_id") = GetTypeId(varType)
End Property

Public Property Get CodeType() As Variant
    CodeType = varType
End Property

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

Public Property Let ���b�NNO(ByVal varValue As Variant)
    If varValue = "" Then
        var���b�NNO = Null
    Else
        var���b�NNO = varValue
    End If
    recHardWare.Item("rack_number") = var���b�NNO
End Property

Public Property Get ���b�NNO() As Variant
    ���b�NNO = var���b�NNO
End Property

Public Property Let �ێ�(ByVal varValue As Variant)
    var�ێ� = varValue
End Property

Public Property Get �ێ�() As Variant
    �ێ� = var�ێ�
End Property

Public Property Let �ێ瑋��(ByVal varValue As Variant)
    var�ێ瑋�� = varValue
End Property

Public Property Get �ێ瑋��() As Variant
    �ێ瑋�� = var�ێ瑋��
End Property

Public Property Let �ێ�_��(ByVal varValue As Variant)
    var�ێ�_�� = varValue
    recHardWare.Item("maintenance_id") = varValue
End Property

Public Property Get �ێ�_��() As Variant
    �ێ�_�� = var�ێ�_��
End Property

Public Property Let �V���A���ԍ�(ByVal varValue As Variant)
    If varValue = "" Then
        var�V���A���ԍ� = Null
    Else
        var�V���A���ԍ� = varValue
    End If
    recHardWare.Item("serial_number") = var�V���A���ԍ�
    recSoftWare.Item("serial_number") = var�V���A���ԍ�
End Property

Public Property Get �V���A���ԍ�() As Variant
    �V���A���ԍ� = var�V���A���ԍ�
End Property

Public Property Let ���i�ԍ�(ByVal varValue As Variant)
    If varValue = "" Then
        var���i�ԍ� = Null
    Else
        var���i�ԍ� = varValue
    End If
    recHardWare.Item("product_number") = var���i�ԍ�
End Property

Public Property Get ���i�ԍ�() As Variant
    ���i�ԍ� = var���i�ԍ�
End Property

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

Public Property Let ���[�J�[(ByVal varValue As Variant)
    If varValue = "" Then
        var���[�J�[ = Null
    Else
        var���[�J�[ = varValue
    End If
    recHardWare.Item("maker") = var���[�J�[
End Property

Public Property Get ���[�J�[() As Variant
    ���[�J�[ = var���[�J�[
End Property

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

Public Property Let ���j�b�g�R�[�h(ByVal varValue As Variant)
    var���j�b�g�R�[�h = varValue
End Property

Public Property Get ���j�b�g�R�[�h() As Variant
    ���j�b�g�R�[�h = var���j�b�g�R�[�h
End Property

Public Property Let ���j�b�g��(ByVal varValue As Variant)
    var���j�b�g�� = varValue
End Property

Public Property Get ���j�b�g��() As Variant
    ���j�b�g�� = var���j�b�g��
End Property

Public Property Let ���Y�ԍ�(ByVal varValue As Variant)
    var���Y�ԍ� = varValue
    recHardWare.Item("property_number") = varValue
End Property

Public Property Get ���Y�ԍ�() As Variant
    ���Y�ԍ� = var���Y�ԍ�
End Property

Public Property Let ���Y��1(ByVal varValue As Variant)
    var���Y��1 = varValue
End Property

Public Property Get ���Y��1() As Variant
    ���Y��1 = var���Y��1
End Property

Public Property Let ���Y��2(ByVal varValue As Variant)
    var���Y��2 = varValue
End Property

Public Property Get ���Y��2() As Variant
    ���Y��2 = var���Y��2
End Property

Public Property Let ����(ByVal varValue As Variant)
    var���� = varValue
End Property

Public Property Get ����() As Variant
    ���� = var����
End Property

Public Property Let �擾��(ByVal varValue As Variant)
    var�擾�� = varValue
End Property

Public Property Get �擾��() As Variant
    �擾�� = var�擾��
End Property

Public Property Let �r����(ByVal varValue As Variant)
    var�r���� = varValue
End Property

Public Property Get �r����() As Variant
    �r���� = var�r����
End Property

Public Property Let �����ƍ��p���(ByVal varValue As Variant)
    var�����ƍ��p��� = varValue
End Property

Public Property Get �����ƍ��p���() As Variant
    �����ƍ��p��� = var�����ƍ��p���
End Property

Public Property Let �����Ǘ��S���Җ�(ByVal varValue As Variant)
    var�����Ǘ��S���Җ� = varValue
End Property

Public Property Get �����Ǘ��S���Җ�() As Variant
    �����Ǘ��S���Җ� = var�����Ǘ��S���Җ�
End Property

Public Property Let �����ݒu�t���A(ByVal varValue As Variant)
    If varValue = "" Then
        var�����ݒu�t���A = Null
    Else
        var�����ݒu�t���A = varValue
    End If
    recHardWare.Item("floor") = var�����ݒu�t���A
End Property

Public Property Get �����ݒu�t���A() As Variant
    �����ݒu�t���A = var�����ݒu�t���A
End Property

Public Property Let �敪(ByVal varValue As Variant)
    var�敪 = varValue
End Property

Public Property Get �敪() As Variant
    �敪 = var�敪
End Property

Public Property Let ���x������(ByVal varValue As Variant)
    var���x������ = varValue
End Property

Public Property Get ���x������() As Variant
    ���x������ = var���x������
End Property

Public Property Let IP�\�[�g�p(ByVal varValue As Variant)
    varIP�\�[�g�p = varValue
End Property

Public Property Get IP�\�[�g�p() As Variant
    IP�\�[�g�p = varIP�\�[�g�p
End Property

Public Property Let �Ǘ��O(ByVal varValue As Variant)
    var�Ǘ��O = varValue
    recObjectMatsr.Item("outside_flag") = varValue
End Property

Public Property Get �Ǘ��O() As Variant
    �Ǘ��O = var�Ǘ��O
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
    If IsEmpty(varType) Then
        varType = GetTypeNameByObjectId
    ElseIf varType <> GetTypeNameByObjectId Then
        blnUpdType = True
    End If
    '�g�����U�N�V�����J�n
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '�ݒ�type���{�́A�����A���ӂ̏ꍇ
        If varType = "�{��" Or varType = "����" Or varType = "����" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                'ShareFolderUserLink
                Set recUpdate = New ShareFolderUserLinkRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("link_id") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ShareFolder
                recShareFolder.DeleteRecord objCon
                Set recUpdate = New ShareFolderRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ServerUser
                recServerUser.DeleteRecord objCon
                Set recUpdate = New ServerUserRecord
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
            'HardWare���ڂ��ݒ肳��Ă����HardWare���X�V
            recHardWare.UpdateRecord objCon
            'Server���ڂ��ݒ肳��Ă����Server���X�V
            recServer.Item("dependence_place") = Empty
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
        ElseIf varType = "VM" Or varType = "TIB" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                'ShareFolderUserLink
                Set recUpdate = New ShareFolderUserLinkRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("link_id") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ShareFolder
                recShareFolder.DeleteRecord objCon
                Set recUpdate = New ShareFolderRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ServerUser
                recServerUser.DeleteRecord objCon
                Set recUpdate = New ServerUserRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'OtherObject
                recOtherObject.DeleteRecord objCon
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
        '�ݒ�type��IP�̏ꍇ
        ElseIf varType = "IP" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                'ShareFolderUserLink
                Set recUpdate = New ShareFolderUserLinkRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("link_id") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ShareFolder
                recShareFolder.DeleteRecord objCon
                Set recUpdate = New ShareFolderRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ServerUser
                recServerUser.DeleteRecord objCon
                Set recUpdate = New ServerUserRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'OtherObject
                recOtherObject.DeleteRecord objCon
                'Server
                recServer.DeleteRecord objCon
                Set recUpdate = New ServerRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'SoftWare
                recSoftWare.DeleteRecord objCon
                Set recUpdate = New SoftWareRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
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
                    recAddress.UpdateRecord objCon
                End If
            '�z�X�g���̂ݕύX����Ă���ꍇ
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '�ݒ�type��USER�̏ꍇ
        ElseIf varType = "USER" Then
            MsgBox varID & ":�������Ȃ̂Ŗ�������܂��B"
            'type_id���ύX����Ă���ꍇ
                'OtherObject����Ώۃ��R�[�h���폜
                'Address����ΏۃN���A
                'Address����ˑ��N���A
                'Server����Ώۃ��R�[�h���폜
                'Server����ˑ��N���A
                'SoftWare����Ώۃ��R�[�h���폜
                'SoftWare����ˑ��N���A
                'HardWare����Ώۃ��R�[�h���폜
                'HardWare����ˑ��N���A
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            'ServerUser�ɒǉ�
        '�ݒ�type�����L�t�H���_�̏ꍇ
        ElseIf varType = "���L�t�H���_" Then
            MsgBox varID & ":�������Ȃ̂Ŗ�������܂��B"
            'type_id���ύX����Ă���ꍇ
                'OtherObject����Ώۃ��R�[�h���폜
                'Address����ΏۃN���A
                'Address����ˑ��N���A
                'Server����Ώۃ��R�[�h���폜
                'Server����ˑ��N���A
                'SoftWare����Ώۃ��R�[�h���폜
                'SoftWare����ˑ��N���A
                'HardWare����Ώۃ��R�[�h���폜
                'HardWare����ˑ��N���A
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            'ShareFolder�ɒǉ�
        '�ݒ�type��SYBASE�̏ꍇ
        ElseIf varType = "SYBASE" Then
            'type_id���ύX����Ă���ꍇ���֌W�ȃe�[�u������Ώۂ��폜
            If blnUpdType Then
                'ShareFolderUserLink
                Set recUpdate = New ShareFolderUserLinkRecord
                recUpdate.Item("object_id") = varID
                recUpdate.DeleteRecord objCon
                recUpdate.Item("object_id") = Empty
                recUpdate.Item("link_id") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ShareFolder
                recShareFolder.DeleteRecord objCon
                Set recUpdate = New ShareFolderRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'ServerUser
                recServerUser.DeleteRecord objCon
                Set recUpdate = New ServerUserRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
                'OtherObject
                recOtherObject.DeleteRecord objCon
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
            End If
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            recObjectMatsr.UpdateRecord objCon
            '�V�K�̏ꍇ�A�z�z���ꂽID���Đݒ�
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'SoftWare���ڂ��ݒ肳��Ă����SoftWare���X�V
            recSoftWare.UpdateRecord objCon
        '�ݒ�type������ȊO�ꍇ
        Else
            'ObjectMaster���ڂ��ݒ肳��Ă����ObjectMaster���X�V
            recObjectMatsr.UpdateRecord objCon
            '�V�K�̏ꍇ�A�z�z���ꂽID���Đݒ�
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
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
    
    'ShareFolderUserLink
    Set recUpdate = New ShareFolderUserLinkRecord
    recUpdate.Item("object_id") = varID
    recUpdate.DeleteRecord objCon
    recUpdate.Item("object_id") = Empty
    recUpdate.Item("link_id") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'ShareFolder
    recShareFolder.DeleteRecord objCon
    Set recUpdate = New ShareFolderRecord
    recUpdate.Item("dependence_place") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'ServerUser
    recServerUser.DeleteRecord objCon
    Set recUpdate = New ServerUserRecord
    recUpdate.Item("dependence_place") = varID
    recUpdate.DeleteRecord objCon
    Set recUpdate = Nothing
    'OtherObject
    recOtherObject.DeleteRecord objCon
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
    'SoftWare
    recSoftWare.DeleteRecord objCon
    Set recUpdate = New SoftWareRecord
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
