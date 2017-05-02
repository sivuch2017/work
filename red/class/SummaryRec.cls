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

'プロパティ(データ部)
Private varID As Variant                'ObjectMaster,Server,HardWare,Software,OtherObject,ServerUser
Private varNo As Variant                '
Private varP As Variant                 'ObjectMaster
Private varVM名 As Variant              'Server
Private var依存先 As Variant            'Server,HardWare,Address,SoftWare
Private var貸出依頼者 As Variant        'ObjectMaster
Private var担当者 As Variant            'ObjectMaster
Private varType As Variant              'ObjectMaster
Private varホスト名 As Variant          'Address
Private varIPアドレス As Variant        'Server
Private var内容 As Variant              'ObjectMaster
Private var状況 As Variant              'ObjectMaster
Private var予定 As Variant              'ObjectMaster
Private varラックNO As Variant          'HardWare
Private var保守 As Variant              '
Private var保守窓口 As Variant          '
Private var保守契約 As Variant          'HardWare
Private varシリアル番号 As Variant      'HardWare,SoftWare
Private var製品番号 As Variant          'HardWare
Private var備考 As Variant              'ObjectMaster
Private varメーカー As Variant          'HardWare,SoftWare
Private varマシン As Variant            'HardWare
Private varユニットコード As Variant    '
Private varユニット名 As Variant        '
Private var資産番号 As Variant          'HardWare
Private var資産名1 As Variant           '
Private var資産名2 As Variant           '
Private var数量 As Variant              '
Private var取得日 As Variant            '
Private varビル名 As Variant            '
Private var現物照合用情報 As Variant    '
Private var現物管理担当者名 As Variant  '
Private var現物設置フロア As Variant    'HardWare
Private var区分 As Variant              '
Private varラベル枚数 As Variant        '
Private varIPソート用 As Variant        '
Private var管理外 As Variant            'ObjectMaster

Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recSoftWare As RecordInterface
Private recServer As RecordInterface
Private recAddress As RecordInterface
Private recOtherObject As RecordInterface
Private recServerUser As RecordInterface
Private recShareFolder As RecordInterface
Private recShareFolderUserLink As RecordInterface

'プロパティ(連携部)
Private diCon As DatabaseInterface

'内部保持
Private objCode As Object
Private blnNew As Boolean

'コンストラクタ
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

'デストラクタ
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

'プロパティ ここから
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

Public Property Let 依存先(ByVal varValue As Variant)
    Dim varResult As Variant
    If varValue = "" Then
        var依存先 = Null
        recHardWare.Item("dependence_place") = Null
        recSoftWare.Item("dependence_place") = Null
        recServer.Item("dependence_place") = Null
        recAddress.Item("dependence_place") = Null
    Else
        var依存先 = varValue
        varResult = GetObjectIdByHostName
        If IsNull(varResult) Then
            MsgBox varID & ":指定された依存先は存在しないので無視されます。"
        ElseIf varResult = varID Then
            MsgBox varID & ":依存先が自分自身なので無視されます。"
        Else
            recHardWare.Item("dependence_place") = varResult
            recSoftWare.Item("dependence_place") = varResult
            recServer.Item("dependence_place") = varResult
            recAddress.Item("dependence_place") = varResult
        End If
    End If
End Property

Public Property Get 依存先() As Variant
    依存先 = var依存先
End Property

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

Public Property Let ラックNO(ByVal varValue As Variant)
    If varValue = "" Then
        varラックNO = Null
    Else
        varラックNO = varValue
    End If
    recHardWare.Item("rack_number") = varラックNO
End Property

Public Property Get ラックNO() As Variant
    ラックNO = varラックNO
End Property

Public Property Let 保守(ByVal varValue As Variant)
    var保守 = varValue
End Property

Public Property Get 保守() As Variant
    保守 = var保守
End Property

Public Property Let 保守窓口(ByVal varValue As Variant)
    var保守窓口 = varValue
End Property

Public Property Get 保守窓口() As Variant
    保守窓口 = var保守窓口
End Property

Public Property Let 保守契約(ByVal varValue As Variant)
    var保守契約 = varValue
    recHardWare.Item("maintenance_id") = varValue
End Property

Public Property Get 保守契約() As Variant
    保守契約 = var保守契約
End Property

Public Property Let シリアル番号(ByVal varValue As Variant)
    If varValue = "" Then
        varシリアル番号 = Null
    Else
        varシリアル番号 = varValue
    End If
    recHardWare.Item("serial_number") = varシリアル番号
    recSoftWare.Item("serial_number") = varシリアル番号
End Property

Public Property Get シリアル番号() As Variant
    シリアル番号 = varシリアル番号
End Property

Public Property Let 製品番号(ByVal varValue As Variant)
    If varValue = "" Then
        var製品番号 = Null
    Else
        var製品番号 = varValue
    End If
    recHardWare.Item("product_number") = var製品番号
End Property

Public Property Get 製品番号() As Variant
    製品番号 = var製品番号
End Property

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

Public Property Let メーカー(ByVal varValue As Variant)
    If varValue = "" Then
        varメーカー = Null
    Else
        varメーカー = varValue
    End If
    recHardWare.Item("maker") = varメーカー
End Property

Public Property Get メーカー() As Variant
    メーカー = varメーカー
End Property

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

Public Property Let ユニットコード(ByVal varValue As Variant)
    varユニットコード = varValue
End Property

Public Property Get ユニットコード() As Variant
    ユニットコード = varユニットコード
End Property

Public Property Let ユニット名(ByVal varValue As Variant)
    varユニット名 = varValue
End Property

Public Property Get ユニット名() As Variant
    ユニット名 = varユニット名
End Property

Public Property Let 資産番号(ByVal varValue As Variant)
    var資産番号 = varValue
    recHardWare.Item("property_number") = varValue
End Property

Public Property Get 資産番号() As Variant
    資産番号 = var資産番号
End Property

Public Property Let 資産名1(ByVal varValue As Variant)
    var資産名1 = varValue
End Property

Public Property Get 資産名1() As Variant
    資産名1 = var資産名1
End Property

Public Property Let 資産名2(ByVal varValue As Variant)
    var資産名2 = varValue
End Property

Public Property Get 資産名2() As Variant
    資産名2 = var資産名2
End Property

Public Property Let 数量(ByVal varValue As Variant)
    var数量 = varValue
End Property

Public Property Get 数量() As Variant
    数量 = var数量
End Property

Public Property Let 取得日(ByVal varValue As Variant)
    var取得日 = varValue
End Property

Public Property Get 取得日() As Variant
    取得日 = var取得日
End Property

Public Property Let ビル名(ByVal varValue As Variant)
    varビル名 = varValue
End Property

Public Property Get ビル名() As Variant
    ビル名 = varビル名
End Property

Public Property Let 現物照合用情報(ByVal varValue As Variant)
    var現物照合用情報 = varValue
End Property

Public Property Get 現物照合用情報() As Variant
    現物照合用情報 = var現物照合用情報
End Property

Public Property Let 現物管理担当者名(ByVal varValue As Variant)
    var現物管理担当者名 = varValue
End Property

Public Property Get 現物管理担当者名() As Variant
    現物管理担当者名 = var現物管理担当者名
End Property

Public Property Let 現物設置フロア(ByVal varValue As Variant)
    If varValue = "" Then
        var現物設置フロア = Null
    Else
        var現物設置フロア = varValue
    End If
    recHardWare.Item("floor") = var現物設置フロア
End Property

Public Property Get 現物設置フロア() As Variant
    現物設置フロア = var現物設置フロア
End Property

Public Property Let 区分(ByVal varValue As Variant)
    var区分 = varValue
End Property

Public Property Get 区分() As Variant
    区分 = var区分
End Property

Public Property Let ラベル枚数(ByVal varValue As Variant)
    varラベル枚数 = varValue
End Property

Public Property Get ラベル枚数() As Variant
    ラベル枚数 = varラベル枚数
End Property

Public Property Let IPソート用(ByVal varValue As Variant)
    varIPソート用 = varValue
End Property

Public Property Get IPソート用() As Variant
    IPソート用 = varIPソート用
End Property

Public Property Let 管理外(ByVal varValue As Variant)
    var管理外 = varValue
    recObjectMatsr.Item("outside_flag") = varValue
End Property

Public Property Get 管理外() As Variant
    管理外 = var管理外
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
    If IsEmpty(varType) Then
        varType = GetTypeNameByObjectId
    ElseIf varType <> GetTypeNameByObjectId Then
        blnUpdType = True
    End If
    'トランザクション開始
    Set objCon = diCon.ConnectDatabase
    objCon.BeginTrans
        '設定typeが本体、内部、周辺の場合
        If varType = "本体" Or varType = "内部" Or varType = "周辺" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
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
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'HardWare項目が設定されていればHardWareを更新
            recHardWare.UpdateRecord objCon
            'Server項目が設定されていればServerを更新
            recServer.Item("dependence_place") = Empty
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
        ElseIf varType = "VM" Or varType = "TIB" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
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
        '設定typeがIPの場合
        ElseIf varType = "IP" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
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
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
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
                    recAddress.UpdateRecord objCon
                End If
            'ホスト名のみ変更されている場合
            ElseIf Not IsEmpty(recAddress.Item("host_name")) Then
                recAddress.Item("ip_address") = GetIpByObjectId
                recAddress.UpdateRecord objCon
            End If
        '設定typeがUSERの場合
        ElseIf varType = "USER" Then
            MsgBox varID & ":未実装なので無視されます。"
            'type_idが変更されている場合
                'OtherObjectから対象レコードを削除
                'Addressから対象クリア
                'Addressから依存クリア
                'Serverから対象レコードを削除
                'Serverから依存クリア
                'SoftWareから対象レコードを削除
                'SoftWareから依存クリア
                'HardWareから対象レコードを削除
                'HardWareから依存クリア
            'ObjectMaster項目が設定されていればObjectMasterを更新
            'ServerUserに追加
        '設定typeが共有フォルダの場合
        ElseIf varType = "共有フォルダ" Then
            MsgBox varID & ":未実装なので無視されます。"
            'type_idが変更されている場合
                'OtherObjectから対象レコードを削除
                'Addressから対象クリア
                'Addressから依存クリア
                'Serverから対象レコードを削除
                'Serverから依存クリア
                'SoftWareから対象レコードを削除
                'SoftWareから依存クリア
                'HardWareから対象レコードを削除
                'HardWareから依存クリア
            'ObjectMaster項目が設定されていればObjectMasterを更新
            'ShareFolderに追加
        '設定typeがSYBASEの場合
        ElseIf varType = "SYBASE" Then
            'type_idが変更されている場合無関係なテーブルから対象を削除
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
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
            'SoftWare項目が設定されていればSoftWareを更新
            recSoftWare.UpdateRecord objCon
        '設定typeがそれ以外場合
        Else
            'ObjectMaster項目が設定されていればObjectMasterを更新
            recObjectMatsr.UpdateRecord objCon
            '新規の場合、配布されたIDを再設定
            If varID = "*" Then
                ID = recObjectMatsr.Item("object_id")
            End If
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

