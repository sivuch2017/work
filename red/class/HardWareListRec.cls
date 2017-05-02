VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HardWareListRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�v���p�e�B(�f�[�^��)
Private varID As Variant                'HardWare,ObjectMaster
Private varGROUP As Variant             '
Private var�ڑ��� As Variant            'HardWare
Private var��� As Variant              'ObjectMaster
Private var���[�J�[ As Variant          'HardWare
Private var�}�V�� As Variant            'HardWare
Private var���i�ԍ� As Variant          'HardWare
Private varCPU As Variant               'HardWare
Private var������ As Variant            'HardWare
Private var�f�B�X�N�e�� As Variant      'HardWare
Private var�V���A���ԍ� As Variant      'HardWare
Private var�t���A As Variant            'HardWare
Private var���b�NNO As Variant          'HardWare
Private var�ʒu As Variant              'HardWare
Private var���e As Variant              'ObjectMaster
Private var���Y�R�[�h As Variant        'HardWare,Property
Private var�ێ�_�� As Variant          'HardWare
Private var���x���ԍ� As Variant        'HardWare
Private var���x������ As Variant        'HardWare,Property
Private var�Ǘ��O As Variant            'ObjectMaster
Private var�v���_�N�g As Variant        'ObjectMaster
Private var��� As Variant              'ObjectMaster
Private var�\�� As Variant              'ObjectMaster
Private var�ݏo�˗��� As Variant        'ObjectMaster
Private var�S���� As Variant            'ObjectMaster
Private var���l As Variant              'ObjectMaster

'�v���p�e�B(�A�g��)
Private diCon As DatabaseInterface

'�����ێ�(���R�[�h)
Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recProperty As RecordInterface

'�����ێ�(�ޔ�)
Private objCode As Object
Private blnNew As Boolean

'�R���X�g���N�^
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recHardWare = New HardWareRecord
    Set recProperty = New PropertyRecord
    blnNew = False
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recHardWare = Nothing
    Set recProperty = Nothing
    Set diCon = Nothing
End Sub

'�v���p�e�B ID
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recHardWare.Item("object_id") = varID
End Property
Public Property Get ID() As Variant
    ID = varID
End Property

'�v���p�e�B GROUP
Public Property Let GROUP(ByVal varValue As Variant)
    varGROUP = varValue
End Property
Public Property Get GROUP() As Variant
    GROUP = varGROUP
End Property

'�v���p�e�B �ڑ���
Public Property Let �ڑ���(ByVal varValue As Variant)
    If varValue = "" Then
        var�ڑ��� = Null
    Else
        var�ڑ��� = varValue
    End If
    recHardWare.Item("dependence_place") = var�ڑ���
End Property
Public Property Get �ڑ���() As Variant
    �ڑ��� = var�ڑ���
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

'�v���p�e�B ���[�J�[
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

'�v���p�e�B ���i�ԍ�
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

'�v���p�e�B CPU
Public Property Let CPU(ByVal varValue As Variant)
    If varValue = "" Then
        varCPU = Null
    Else
        varCPU = varValue
    End If
    recHardWare.Item("cpu") = varCPU
End Property
Public Property Get CPU() As Variant
    CPU = varCPU
End Property

'�v���p�e�B ������
Public Property Let ������(ByVal varValue As Variant)
    If varValue = "" Then
        var������ = Null
    Else
        var������ = varValue
    End If
    recHardWare.Item("memory") = var������
End Property
Public Property Get ������() As Variant
    ������ = var������
End Property

'�v���p�e�B �f�B�X�N�e��
Public Property Let �f�B�X�N�e��(ByVal varValue As Variant)
    If varValue = "" Then
        var�f�B�X�N�e�� = Null
    Else
        var�f�B�X�N�e�� = varValue
    End If
    recHardWare.Item("capacity") = var�f�B�X�N�e��
End Property
Public Property Get �f�B�X�N�e��() As Variant
    �f�B�X�N�e�� = var�f�B�X�N�e��
End Property

'�v���p�e�B �V���A���ԍ�
Public Property Let �V���A���ԍ�(ByVal varValue As Variant)
    If varValue = "" Then
        var�V���A���ԍ� = Null
    Else
        var�V���A���ԍ� = varValue
    End If
    recHardWare.Item("serial_number") = var�V���A���ԍ�
End Property
Public Property Get �V���A���ԍ�() As Variant
    �V���A���ԍ� = var�V���A���ԍ�
End Property

'�v���p�e�B �t���A
Public Property Let �t���A(ByVal varValue As Variant)
    If varValue = "" Then
        var�t���A = Null
    Else
        var�t���A = varValue
    End If
    recHardWare.Item("floor") = var�t���A
End Property
Public Property Get �t���A() As Variant
    �t���A = var�t���A
End Property

'�v���p�e�B ���b�NNO
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

'�v���p�e�B �ʒu
Public Property Let �ʒu(ByVal varValue As Variant)
    If varValue = "" Then
        var�ʒu = Null
    Else
        var�ʒu = varValue
    End If
    recHardWare.Item("position") = var�ʒu
End Property
Public Property Get �ʒu() As Variant
    �ʒu = var�ʒu
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

'�v���p�e�B ���Y�R�[�h
Public Property Let ���Y�R�[�h(ByVal varValue As Variant)
    If varValue = "" Then
        var���Y�R�[�h = Null
    Else
        var���Y�R�[�h = varValue
        recProperty.Item("property_number") = var���Y�R�[�h
    End If
    recHardWare.Item("property_number") = var���Y�R�[�h
End Property
Public Property Get ���Y�R�[�h() As Variant
    ���Y�R�[�h = var���Y�R�[�h
End Property

'�v���p�e�B �ێ�_��
Public Property Let �ێ�_��(ByVal varValue As Variant)
    var�ێ�_�� = varValue
    recHardWare.Item("maintenance_id") = varValue
End Property
Public Property Get �ێ�_��() As Variant
    �ێ�_�� = var�ێ�_��
End Property

'�v���p�e�B ���x���ԍ�
Public Property Let ���x���ԍ�(ByVal varValue As Variant)
    If varValue = "" Then
        var���x���ԍ� = Null
    Else
        var���x���ԍ� = varValue
    End If
    recHardWare.Item("label_number") = var���x���ԍ�
End Property
Public Property Get ���x���ԍ�() As Variant
    ���x���ԍ� = var���x���ԍ�
End Property

'�v���p�e�B ���x������
Public Property Let ���x������(ByVal varValue As Variant)
    If varValue = "" Then
        var���x������ = Null
    Else
        var���x������ = varValue
    End If
    recProperty.Item("number_of_sheets") = var���x������
End Property
Public Property Get ���x������() As Variant
    ���x������ = var���x������
End Property

'�v���p�e�B �Ǘ��O
Public Property Let �Ǘ��O(ByVal varValue As Variant)
    var�Ǘ��O = varValue
    recObjectMatsr.Item("outside_flag") = varValue
End Property
Public Property Get �Ǘ��O() As Variant
    �Ǘ��O = var�Ǘ��O
End Property

'�v���p�e�B �v���_�N�g
Public Property Let �v���_�N�g(ByVal varValue As Variant)
    If varValue = "" Then
        var�v���_�N�g = Null
    Else
        var�v���_�N�g = varValue
    End If
    recObjectMatsr.Item("product") = var�v���_�N�g
End Property
Public Property Get �v���_�N�g() As Variant
    �v���_�N�g = var�v���_�N�g
End Property

'�v���p�e�B ���
Public Property Let ���(ByVal varValue As Variant)
    If varValue = "" Then
        var��� = Null
    Else
        var��� = varValue
    End If
    recObjectMatsr.Item("situation") = var���
End Property
Public Property Get ���() As Variant
    ��� = var���
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
        If var��� = "�{��" Or var��� = "����" Or var��� = "����" Then
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
            'Property���ڂ��ݒ肳��Ă����Property���X�V
            recProperty.UpdateRecord objCon
            'HardWare���ڂ��ݒ肳��Ă����HardWare���X�V
            recHardWare.UpdateRecord objCon
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
