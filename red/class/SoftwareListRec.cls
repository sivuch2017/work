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

'プロパティ(データ部)
Private varID As Variant             'Software
Private var依存先 As Variant         'ObjectMaster,Software
Private varソフトウェア名 As Variant 'Software
Private varメーカー As Variant       'Software
Private varシリアル番号 As Variant   'Software

Private recObjectMatsr As RecordInterface
Private recSoftWare As RecordInterface

'プロパティ(連携部)
Private diCon As DatabaseInterface

'内部保持
Private objCode As Object
Private blnNew As Boolean

'コンストラクタ
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recSoftWare = New SoftWareRecord
    blnNew = False
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recSoftWare = Nothing
    Set diCon = Nothing
End Sub

'プロパティ ここから
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recSoftWare.Item("object_id") = varID
End Property

Public Property Get ID() As Variant
    ID = varID
End Property

Public Property Let 依存先(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        var依存先 = Null
        recSoftWare.Item("dependence_place") = Null
    Else
        var依存先 = varValue
        recSoftWare.Item("dependence_place") = varValue
    End If
End Property

Public Property Get 依存先() As Variant
    依存先 = var依存先
End Property

Public Property Let ソフトウェア名(ByVal varValue As Variant)
    If varValue = "" Then
        varソフトウェア名 = Null
    Else
        varソフトウェア名 = varValue
    End If
    recSoftWare.Item("software_name") = varソフトウェア名
End Property

Public Property Get ソフトウェア名() As Variant
    ソフトウェア名 = varソフトウェア名
End Property

Public Property Let メーカー(ByVal varValue As Variant)
    If varValue = "" Then
        varメーカー = Null
    Else
        varメーカー = varValue
    End If
    recSoftWare.Item("maker") = varメーカー
End Property

Public Property Let シリアル番号(ByVal varValue As Variant)
    If varValue = "" Then
        varシリアル番号 = Null
    Else
        varシリアル番号 = varValue
    End If
    recSoftWare.Item("serial_number") = varシリアル番号
End Property

Public Property Get シリアル番号() As Variant
    シリアル番号 = varシリアル番号
End Property

Public Property Get メーカー() As Variant
    メーカー = varメーカー
End Property

Public Property Set DB情報(ByVal diValue As DatabaseInterface)
    Set diCon = diValue
End Property

Public Property Get DB情報() As DatabaseInterface
    Set DB情報 = diCon
End Property
'プロパティ ここまで

'Code変換
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
                Case "本体"
                    GetTypeId = "01"
                Case "VM"
                    GetTypeId = "02"
                Case "内部"
                    GetTypeId = "03"
                Case "周辺"
                    GetTypeId = "04"
                Case "IP"
                    GetTypeId = "05"
                Case "TIB"
                    GetTypeId = "06"
                Case "ソフト"
                    GetTypeId = "07"
                Case "USER"
                    GetTypeId = "08"
                Case "共有フォルダ"
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

'更新
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
    'DB接続情報が設定されていない場合はエラー
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Update", "DB情報が設定されていません。"
    End If
    'object_idが空の場合はエラー
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Update", "IDが設定されていません。"
    End If
    'type_id取得
    strType = GetTypeId("ソフト")
    'トランザクション開始
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '新規の場合、ObjectMasterに登録し配布されたIDを再設定
        If varID = "*" Then
            recObjectMatsr.Item("type_id") = strType
            recObjectMatsr.Item("contents") = varソフトウェア名
            recObjectMatsr.UpdateRecord objCon
            ID = recObjectMatsr.Item("object_id")
        Else
            'ID指定でObjectMasterにobject_idが無い場合はエラー
            If ExistsObjectId(varID) = False Then
                Err.Raise vbObjectError + 512 + 1, "Update", "指定IDがマスターに存在しません。"
            End If
        End If
        '依存先指定でObjectMasterにobject_idが無い場合はエラー
        If Not IsEmpty(var依存先) Then
            If ExistsObjectId(var依存先) = False Then
                Err.Raise vbObjectError + 512 + 1, "Update", "指定依存先がマスターに存在しません。"
            End If
        End If
        'SoftWare項目が設定されていればSoftWareを更新
        recSoftWare.UpdateRecord objCon
    'トランザクション終了
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

'削除
Public Sub Delete()
    On Error GoTo ErrHandle
    Dim objCon As ADODB.Connection
    Dim recUpdate As RecordInterface
    Dim strType As String
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "DB情報が設定されていません。"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "IDが設定されていません。"
    End If
    If varID = "*" Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "対象に追加用IDが含まれています。"
    End If
    
    '種別取得
    strType = GetTypeNameByObjectId
    
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        'SoftWare
        recSoftWare.DeleteRecord objCon
        'ObjectMatsr
        If strType = "ソフト" Then
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

'既存type_id取得
Private Function GetTypeNameByObjectId() As String
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    GetTypeNameByObjectId = ""
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetTypeNameByObjectId", "DB情報が設定されていません。"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetTypeNameByObjectId", "IDが設定されていません。"
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

'既存ip_address取得
Private Function GetIpByObjectId() As String
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    GetIpByObjectId = ""
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetIpByObjectId", "DB情報が設定されていません。"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetIpByObjectId", "IDが設定されていません。"
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

'ホスト名からobject_idを取得
Private Function GetObjectIdByHostName() As Variant
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    GetObjectIdByHostName = Null
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "GetObjectIdByHostName", "DB情報が設定されていません。"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "GetObjectIdByHostName", "IDが設定されていません。"
    End If
    Set objCon = diCon.ConnectDatabase
    If objCon.State <> adStateClosed Then
        blnOpen = True
    Else
        blnOpen = False
    End If
    Set objRec = New ADODB.Recordset
    objRec.Open "SELECT TOP(1) ISNULL(object_id, dependence_place) AS id FROM Address WHERE host_name = '" & var依存先 & "' AND (object_id IS NOT NULL OR dependence_place IS NOT NULL)", objCon, adOpenStatic, adLockReadOnly, adCmdText
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

'ObjectMasterにobject_idが存在するか
Private Function ExistsObjectId(ByVal varObjId As Variant) As Boolean
    Dim objCon As ADODB.Connection
    Dim objRec As ADODB.Recordset
    Dim blnOpen As Boolean
    ExistsObjectId = False
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "ExistsObjectId", "DB情報が設定されていません。"
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
