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

'プロパティ(データ部)
Private varID As Variant                'Server,ObjectMaster,HardWare
Private varホスト名 As Variant          'Address
Private varIPアドレス As Variant        'Server,Address
Private var種類 As Variant              'ObjectMaster
Private varVM名 As Variant              'Server
Private varVMサーバー名 As Variant      'Server,Address
Private var貸出依頼者 As Variant        'ObjectMaster
Private var担当者 As Variant            'ObjectMaster
Private var内容 As Variant              'ObjectMaster
Private var状況 As Variant              'ObjectMaster
Private var予定 As Variant              'ObjectMaster
Private varマシン As Variant            'HardWare
Private var割り当てCPU As Variant       'Server
Private var割り当てメモリ As Variant    'Server
Private varディスク容量 As Variant      'Server
Private var備考 As Variant              'ObjectMaster
Private var対象フラグ As Variant        'Server
Private varIPソート用 As Variant        'Server

'プロパティ(連携部)
Private diCon As DatabaseInterface

'内部保持(レコード)
Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recServer As RecordInterface
Private recAddress As RecordInterface

'内部保持(退避)
Private objCode As Object
Private blnNew As Boolean

'コンストラクタ
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recHardWare = New HardWareRecord
    Set recServer = New ServerRecord
    Set recAddress = New AddressRecord
    blnNew = False
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recHardWare = Nothing
    Set recServer = Nothing
    Set recAddress = Nothing
    Set diCon = Nothing
End Sub

'プロパティ ID
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recHardWare.Item("object_id") = varID
    recServer.Item("object_id") = varID
End Property
Public Property Get ID() As Variant
    ID = varID
End Property

'プロパティ ホスト名
Public Property Let ホスト名(ByVal varValue As Variant)
    If varValue = "" Then
        varホスト名 = Null
    Else
        varホスト名 = varValue
    End If
    recAddress.Item("host_name") = varホスト名
End Property
Public Property Get ホスト名() As Variant
    ホスト名 = varホスト名
End Property

'プロパティ IPアドレス
Public Property Let IPアドレス(ByVal varValue As Variant)
    If varValue = "" Then
        varIPアドレス = Null
    Else
        varIPアドレス = varValue
    End If
    recServer.Item("ip_address") = varIPアドレス
    recAddress.Item("ip_address") = varIPアドレス
End Property
Public Property Get IPアドレス() As Variant
    IPアドレス = varIPアドレス
End Property

'プロパティ 種類
Public Property Let 種類(ByVal varValue As Variant)
    If varValue = "" Then
        var種類 = Null
    Else
        var種類 = varValue
    End If
    recObjectMatsr.Item("type_id") = GetTypeId(var種類)
End Property
Public Property Get 種類() As Variant
    種類 = var種類
End Property

'プロパティ VM名
Public Property Let VM名(ByVal varValue As Variant)
    If varValue = "" Then
        varVM名 = Null
    Else
        varVM名 = varValue
    End If
    recServer.Item("vm_name") = varVM名
End Property
Public Property Get VM名() As Variant
    VM名 = varVM名
End Property

'プロパティ VMサーバー名
Public Property Let VMサーバー名(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        varVMサーバー名 = Null
        recServer.Item("dependence_place") = Null
    Else
        varVMサーバー名 = varValue
        varResult = GetObjectIdByHostName(varVMサーバー名)
        If IsNull(varResult) Then
            MsgBox varID & ":指定された依存先は存在しないので無視されます。"
        ElseIf varResult = varID Then
            MsgBox varID & ":依存先が自分自身なので無視されます。"
        Else
            recServer.Item("dependence_place") = varResult
        End If
    End If
End Property
Public Property Get VMサーバー名() As Variant
    VMサーバー名 = varVMサーバー名
End Property

'プロパティ 貸出依頼者
Public Property Let 貸出依頼者(ByVal varValue As Variant)
    If varValue = "" Then
        var貸出依頼者 = Null
    Else
        var貸出依頼者 = varValue
    End If
    recObjectMatsr.Item("rental_client") = var貸出依頼者
End Property
Public Property Get 貸出依頼者() As Variant
    貸出依頼者 = var貸出依頼者
End Property

'プロパティ 担当者
Public Property Let 担当者(ByVal varValue As Variant)
    If varValue = "" Then
        var担当者 = Null
    Else
        var担当者 = varValue
    End If
    recObjectMatsr.Item("person") = var担当者
End Property
Public Property Get 担当者() As Variant
    担当者 = var担当者
End Property

'プロパティ 内容
Public Property Let 内容(ByVal varValue As Variant)
    If varValue = "" Then
        var内容 = Null
    Else
        var内容 = varValue
    End If
    recObjectMatsr.Item("contents") = var内容
End Property
Public Property Get 内容() As Variant
    内容 = var内容
End Property

'プロパティ 状況
Public Property Let 状況(ByVal varValue As Variant)
    If varValue = "" Then
        var状況 = Null
    Else
        var状況 = varValue
    End If
    recObjectMatsr.Item("situation") = var状況
End Property
Public Property Get 状況() As Variant
    状況 = var状況
End Property

'プロパティ 予定
Public Property Let 予定(ByVal varValue As Variant)
    If varValue = "" Then
        var予定 = Null
    Else
        var予定 = varValue
    End If
    recObjectMatsr.Item("future") = var予定
End Property
Public Property Get 予定() As Variant
    予定 = var予定
End Property

'プロパティ マシン
Public Property Let マシン(ByVal varValue As Variant)
    If varValue = "" Then
        varマシン = Null
    Else
        varマシン = varValue
    End If
    recHardWare.Item("hardware_name") = varマシン
End Property
Public Property Get マシン() As Variant
    マシン = varマシン
End Property

'プロパティ 割り当てCPU
Public Property Let 割り当てCPU(ByVal varValue As Variant)
    If varValue = "" Then
        var割り当てCPU = Null
    Else
        var割り当てCPU = varValue
    End If
    recServer.Item("cpu") = var割り当てCPU
End Property
Public Property Get 割り当てCPU() As Variant
    割り当てCPU = var割り当てCPU
End Property

'プロパティ 割り当てメモリ
Public Property Let 割り当てメモリ(ByVal varValue As Variant)
    If varValue = "" Then
        var割り当てメモリ = Null
    Else
        var割り当てメモリ = varValue
    End If
    recServer.Item("memory") = var割り当てメモリ
End Property
Public Property Get 割り当てメモリ() As Variant
    割り当てメモリ = var割り当てメモリ
End Property

'プロパティ ディスク容量
Public Property Let ディスク容量(ByVal varValue As Variant)
    If varValue = "" Then
        varディスク容量 = Null
    Else
        varディスク容量 = varValue
    End If
    recServer.Item("capacity") = varディスク容量
End Property
Public Property Get ディスク容量() As Variant
    ディスク容量 = varディスク容量
End Property

'プロパティ 備考
Public Property Let 備考(ByVal varValue As Variant)
    If varValue = "" Then
        var備考 = Null
    Else
        var備考 = varValue
    End If
    recObjectMatsr.Item("memo") = var備考
End Property
Public Property Get 備考() As Variant
    備考 = var備考
End Property

'プロパティ 対象フラグ
Public Property Let 対象フラグ(ByVal varValue As Variant)
    If varValue = "" Then
        var対象フラグ = Null
    Else
        var対象フラグ = varValue
    End If
    recServer.Item("gx_server_flag") = var対象フラグ
End Property
Public Property Get 対象フラグ() As Variant
    対象フラグ = var対象フラグ
End Property

'プロパティ IPソート用
Public Property Let IPソート用(ByVal varValue As Variant)
    varIPソート用 = varValue
End Property
Public Property Get IPソート用() As Variant
    IPソート用 = varIPソート用
End Property

'プロパティ DB情報
Public Property Set DB情報(ByVal diValue As DatabaseInterface)
    Set diCon = diValue
End Property
Public Property Get DB情報() As DatabaseInterface
    Set DB情報 = diCon
End Property

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
                Case "SYBASE"
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
    If IsEmpty(var種類) Then
        var種類 = GetTypeNameByObjectId
    ElseIf var種類 <> GetTypeNameByObjectId Then
        blnUpdType = True
    End If
    'トランザクション開始
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '設定typeが本体、内部、周辺の場合
        If var種類 = "本体" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
            If blnUpdType Then
                'なし
            End If
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'HardWare項目が設定されていればHardWareを更新
            recHardWare.UpdateRecord objCon
            'Server項目が設定されていればServerを更新
            recServer.UpdateRecord objCon
            'Address項目が設定されていればobject_idに関連する項目をクリアしAddressを更新
            If Not IsEmpty(recAddress.Item("ip_address")) Then
                '一旦クリア
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
            'ホスト名のみ変更されている場合
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '設定typeがVM、TIBの場合
        ElseIf var種類 = "VM" Or var種類 = "TIB" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
            If blnUpdType Then
                'HardWare
                recHardWare.DeleteRecord objCon
                Set recUpdate = New HardWareRecord
                recUpdate.Item("dependence_place") = varID
                recUpdate.DeleteRecord objCon
                Set recUpdate = Nothing
            End If
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'Server項目が設定されていればServerを更新
            recServer.UpdateRecord objCon
            'Address項目が設定されていればobject_idに関連する項目をクリアしAddressを更新
            If Not IsEmpty(recAddress.Item("ip_address")) Then
                '一旦クリア
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
            'ホスト名のみ変更されている場合
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '設定typeがそれ以外場合
        Else
            Err.Raise vbObjectError + 512 + 1, "Update", "設定可能種類ではありません。"
        End If
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
    If diCon Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "DB情報が設定されていません。"
    End If
    If IsEmpty(varID) Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "IDが設定されていません。"
    End If
    If varID = "*" Then
        Err.Raise vbObjectError + 512 + 1, "Delete", "対象に追加用IDが含まれています。"
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
Private Function GetObjectIdByHostName(ByVal ServerName As String) As Variant
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

