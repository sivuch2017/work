VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'サーバー名⇒オブジェクトID
Private hashServerName As Object

'種類リスト
Private strTypeList As String

'Code
Private colCode As Collection

'起動時にエントリーシートリボンに移動
Private Sub Workbook_Open()
    Application.OnTime [Now() + "0:00:00.1"], "Ribbon.CallMe"
End Sub

Private Sub Workbook_Activate()
    CustomUI.ActivateTab
End Sub

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal TARGET As Hyperlink)
    Application.GoTo Selection, True
End Sub

Public Sub clearSheets()
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = "ソフトウェア" Or sht.Name = "ハードウェア" Or sht.Name = "破棄済一覧" Or sht.Name = "GX環境一覧" Or sht.Name = IPList.strSheetName Or sht.Name = EditData.strSheetName Or sht.Name = StopStartList.strSheetName Or sht.Name = ServerList.strSheetName Or sht.Name = "外部ディスク" Or sht.Name = "仮想マシン" Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    If TObjectMaster.Visible = xlSheetVisible Then
        TObjectMaster.ClearList
        TObjectMaster.Visible = xlSheetHidden
    End If
    If THardWare.Visible = xlSheetVisible Then
        THardWare.ClearList
        THardWare.Visible = xlSheetHidden
    End If
    If TServer.Visible = xlSheetVisible Then
        TServer.ClearList
        TServer.Visible = xlSheetHidden
    End If
    If TCode.Visible = xlSheetVisible Then
        TCode.ClearList
        TCode.Visible = xlSheetHidden
    End If
    If TProperty.Visible = xlSheetVisible Then
        TProperty.ClearList
        TProperty.Visible = xlSheetHidden
    End If
    If TAddress.Visible = xlSheetVisible Then
        TAddress.ClearList
        TAddress.Visible = xlSheetHidden
    End If
End Sub

Public Sub openMaintenance()
    TObjectMaster.Visible = xlSheetVisible
    TObjectMaster.InitSheet
    THardWare.Visible = xlSheetVisible
    THardWare.InitSheet
    TServer.Visible = xlSheetVisible
    TServer.InitSheet
    TCode.Visible = xlSheetVisible
    TCode.InitSheet
    TProperty.InitSheet
    TProperty.Visible = xlSheetVisible
    TAddress.InitSheet
    TAddress.Visible = xlSheetVisible
    TObjectMaster.Activate
End Sub

Public Function getObjcetId(ByVal Name As String) As String
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    
    If hashServerName Is Nothing Then
        conQCSDB.ConnectionString = getConInfo
        conQCSDB.Open
        
        rsBase.Open "SELECT host_name, object_id FROM Address WHERE host_name IS NOT NULL AND object_id IS NOT NULL", conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
        If rsBase.EOF Then
        Else
            Set hashServerName = CreateObject("Scripting.Dictionary")
            Do Until rsBase.EOF
                hashServerName.Add Trim(rsBase.Fields("host_name").Value), Trim(rsBase.Fields("object_id").Value)
                rsBase.MoveNext
            Loop
        End If
        rsBase.Close
    
        conQCSDB.Close
        Set rsBase = Nothing
        Set conQCSDB = Nothing
    End If
    
    getObjcetId = hashServerName(Name)

End Function

'CSV:type_id
Public Function getTypeList() As String
    If strTypeList = "" Then
        loadCode
    End If
    getTypeList = strTypeList
End Function

'ToTypeId
Public Function toTypeId(ByVal Value As String) As String
    Dim Item As Variant
    For Each Item In colCode
        
    Next
End Function

'Code
Private Sub loadCode()
    
    Dim conQCSDB As New ADODB.Connection
    Dim rsBase As New ADODB.Recordset
    Dim colWork As Collection
    
    conQCSDB.ConnectionString = getConInfo
    conQCSDB.Open
    
    rsBase.Open "SELECT column_id, code, value FROM Code", conQCSDB, adOpenStatic, adLockReadOnly, adCmdText
    If rsBase.EOF Then
    Else
        Set colCode = New Collection
        Do Until rsBase.EOF
            'Collection
            If isExists(colCode, rsBase.Fields("column_id").Value) Then
                Set colWork = colCode.Item(rsBase.Fields("column_id").Value)
                If isExists(colWork, rsBase.Fields("code").Value) Then
                    colWork.Item(rsBase.Fields("code").Value) = rsBase.Fields("value").Value
                Else
                    colWork.Add rsBase.Fields("value").Value, rsBase.Fields("code").Value
                End If
            Else
                Set colWork = New Collection
                colWork.Add rsBase.Fields("value").Value, rsBase.Fields("code").Value
                colCode.Add colWork, rsBase.Fields("column_id").Value
            End If
            
            'type_id
            If rsBase.Fields("column_id") = "type_id" Then
                If strTypeList <> "" Then
                    strTypeList = strTypeList & ","
                End If
                strTypeList = strTypeList & rsBase.Fields("value").Value
            End If
            
            rsBase.MoveNext
        Loop
    End If
    rsBase.Close

    conQCSDB.Close
    Set rsBase = Nothing
    Set conQCSDB = Nothing
    
End Sub

'Collection.isExists
Public Function isExists(ByVal col As Collection, ByVal Key As String) As Boolean
    On Error GoTo ErrHandle
    col.Item Key
    isExists = True
    Exit Function
ErrHandle:
    isExists = False
End Function

'PING
Public Function pingIp(ip As String) As Boolean

    Dim PingSet As SWbemObjectSet
    Dim ping As SWbemObject
    Dim Locator As SWbemLocator
    Dim Service As SWbemServices
    Dim MesStr As String
    
    pingIp = False
    
    Set Locator = New WbemScripting.SWbemLocator
    Set Service = Locator.ConnectServer
    Set PingSet = Service.ExecQuery("Select * From Win32_PingStatus Where Address = '" & ip & "'")
    
    For Each ping In PingSet
        Debug.Print ip & ":" & ping.StatusCode
        If ping.StatusCode = 0 Then
            pingIp = True
        End If
    Next
    
    Set PingSet = Nothing
    Set ping = Nothing
    Set Service = Nothing
    Set Locator = Nothing

End Function

