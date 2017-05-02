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

'プロパティ(データ部)
Private varID As Variant                'HardWare,ObjectMaster
Private varGROUP As Variant             '
Private var接続先 As Variant            'HardWare
Private var種類 As Variant              'ObjectMaster
Private varメーカー As Variant          'HardWare
Private varマシン As Variant            'HardWare
Private var製品番号 As Variant          'HardWare
Private varCPU As Variant               'HardWare
Private varメモリ As Variant            'HardWare
Private varディスク容量 As Variant      'HardWare
Private varシリアル番号 As Variant      'HardWare
Private varフロア As Variant            'HardWare
Private varラックNO As Variant          'HardWare
Private var位置 As Variant              'HardWare
Private var内容 As Variant              'ObjectMaster
Private var資産コード As Variant        'HardWare,Property
Private var保守契約 As Variant          'HardWare
Private varラベル番号 As Variant        'HardWare
Private varラベル枚数 As Variant        'HardWare,Property
Private var管理外 As Variant            'ObjectMaster
Private varプロダクト As Variant        'ObjectMaster
Private var状態 As Variant              'ObjectMaster
Private var予定 As Variant              'ObjectMaster
Private var貸出依頼者 As Variant        'ObjectMaster
Private var担当者 As Variant            'ObjectMaster
Private var備考 As Variant              'ObjectMaster

'プロパティ(連携部)
Private diCon As DatabaseInterface

'内部保持(レコード)
Private recObjectMatsr As RecordInterface
Private recHardWare As RecordInterface
Private recProperty As RecordInterface

'内部保持(退避)
Private objCode As Object
Private blnNew As Boolean

'コンストラクタ
Private Sub Class_Initialize()
    Set objCode = CreateObject("Scripting.Dictionary")
    Set recObjectMatsr = New ObjectMasterRecord
    Set recHardWare = New HardWareRecord
    Set recProperty = New PropertyRecord
    blnNew = False
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set objCode = Nothing
    Set recObjectMatsr = Nothing
    Set recHardWare = Nothing
    Set recProperty = Nothing
    Set diCon = Nothing
End Sub

'プロパティ ID
Public Property Let ID(ByVal varValue As Variant)
    varID = varValue
    recObjectMatsr.Item("object_id") = varID
    recHardWare.Item("object_id") = varID
End Property
Public Property Get ID() As Variant
    ID = varID
End Property

'プロパティ GROUP
Public Property Let GROUP(ByVal varValue As Variant)
    varGROUP = varValue
End Property
Public Property Get GROUP() As Variant
    GROUP = varGROUP
End Property

'プロパティ 接続先
Public Property Let 接続先(ByVal varValue As Variant)
    If varValue = "" Then
        var接続先 = Null
    Else
        var接続先 = varValue
    End If
    recHardWare.Item("dependence_place") = var接続先
End Property
Public Property Get 接続先() As Variant
    接続先 = var接続先
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

'プロパティ メーカー
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

'プロパティ 製品番号
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

'プロパティ CPU
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

'プロパティ メモリ
Public Property Let メモリ(ByVal varValue As Variant)
    If varValue = "" Then
        varメモリ = Null
    Else
        varメモリ = varValue
    End If
    recHardWare.Item("memory") = varメモリ
End Property
Public Property Get メモリ() As Variant
    メモリ = varメモリ
End Property

'プロパティ ディスク容量
Public Property Let ディスク容量(ByVal varValue As Variant)
    If varValue = "" Then
        varディスク容量 = Null
    Else
        varディスク容量 = varValue
    End If
    recHardWare.Item("capacity") = varディスク容量
End Property
Public Property Get ディスク容量() As Variant
    ディスク容量 = varディスク容量
End Property

'プロパティ シリアル番号
Public Property Let シリアル番号(ByVal varValue As Variant)
    If varValue = "" Then
        varシリアル番号 = Null
    Else
        varシリアル番号 = varValue
    End If
    recHardWare.Item("serial_number") = varシリアル番号
End Property
Public Property Get シリアル番号() As Variant
    シリアル番号 = varシリアル番号
End Property

'プロパティ フロア
Public Property Let フロア(ByVal varValue As Variant)
    If varValue = "" Then
        varフロア = Null
    Else
        varフロア = varValue
    End If
    recHardWare.Item("floor") = varフロア
End Property
Public Property Get フロア() As Variant
    フロア = varフロア
End Property

'プロパティ ラックNO
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

'プロパティ 位置
Public Property Let 位置(ByVal varValue As Variant)
    If varValue = "" Then
        var位置 = Null
    Else
        var位置 = varValue
    End If
    recHardWare.Item("position") = var位置
End Property
Public Property Get 位置() As Variant
    位置 = var位置
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

'プロパティ 資産コード
Public Property Let 資産コード(ByVal varValue As Variant)
    If varValue = "" Then
        var資産コード = Null
    Else
        var資産コード = varValue
        recProperty.Item("property_number") = var資産コード
    End If
    recHardWare.Item("property_number") = var資産コード
End Property
Public Property Get 資産コード() As Variant
    資産コード = var資産コード
End Property

'プロパティ 保守契約
Public Property Let 保守契約(ByVal varValue As Variant)
    var保守契約 = varValue
    recHardWare.Item("maintenance_id") = varValue
End Property
Public Property Get 保守契約() As Variant
    保守契約 = var保守契約
End Property

'プロパティ ラベル番号
Public Property Let ラベル番号(ByVal varValue As Variant)
    If varValue = "" Then
        varラベル番号 = Null
    Else
        varラベル番号 = varValue
    End If
    recHardWare.Item("label_number") = varラベル番号
End Property
Public Property Get ラベル番号() As Variant
    ラベル番号 = varラベル番号
End Property

'プロパティ ラベル枚数
Public Property Let ラベル枚数(ByVal varValue As Variant)
    If varValue = "" Then
        varラベル枚数 = Null
    Else
        varラベル枚数 = varValue
    End If
    recProperty.Item("number_of_sheets") = varラベル枚数
End Property
Public Property Get ラベル枚数() As Variant
    ラベル枚数 = varラベル枚数
End Property

'プロパティ 管理外
Public Property Let 管理外(ByVal varValue As Variant)
    var管理外 = varValue
    recObjectMatsr.Item("outside_flag") = varValue
End Property
Public Property Get 管理外() As Variant
    管理外 = var管理外
End Property

'プロパティ プロダクト
Public Property Let プロダクト(ByVal varValue As Variant)
    If varValue = "" Then
        varプロダクト = Null
    Else
        varプロダクト = varValue
    End If
    recObjectMatsr.Item("product") = varプロダクト
End Property
Public Property Get プロダクト() As Variant
    プロダクト = varプロダクト
End Property

'プロパティ 状態
Public Property Let 状態(ByVal varValue As Variant)
    If varValue = "" Then
        var状態 = Null
    Else
        var状態 = varValue
    End If
    recObjectMatsr.Item("situation") = var状態
End Property
Public Property Get 状態() As Variant
    状態 = var状態
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
        If var種類 = "本体" Or var種類 = "内部" Or var種類 = "周辺" Then
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
            'Property項目が設定されていればPropertyを更新
            recProperty.UpdateRecord objCon
            'HardWare項目が設定されていればHardWareを更新
            recHardWare.UpdateRecord objCon
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

