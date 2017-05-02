Attribute VB_Name = "CustomUI"
Option Explicit

'シート名：ビューオブジェクトのハッシュ
Private ArrayWorkSheet As Object

'リボン
Private objRibbon As IRibbonUI

'シート名ID
Private Enum Names
    仮想マシン = 8
    ソフトウェア = 9
End Enum

'シート名
Private Property Get NameById(ByVal ID As Long) As String
    Select Case ID
        Case Names.仮想マシン: NameById = "仮想マシン"
        Case Names.ソフトウェア: NameById = "ソフトウェア"
    End Select
End Property

Sub Ribbon_onLoad(ribbon As IRibbonUI)
  Set objRibbon = ribbon
  Call objRibbon.ActivateTab("CustomTab")
End Sub

Sub ActivateTab()
    If objRibbon Is Nothing Then
        '将来的にはIRibbonUIを再設定するロジックを追加
        'https://social.msdn.microsoft.com/Forums/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba?forum=exceldev
    Else
        Call objRibbon.ActivateTab("CustomTab")
    End If
End Sub

'シート存在チェック
Private Function ExistsSheet(ByVal Name As String) As Boolean
    Dim objSheet As Worksheet
    ExistsSheet = False
    For Each objSheet In Worksheets
        If objSheet.Name = Name Then
            ExistsSheet = True
        End If
    Next objSheet
End Function

'汎用更新処理
Private Sub 更新(ByVal ID As Long)
    On Error GoTo ErrHandle
    Dim strName As String
    strName = NameById(ID)
    If ArrayWorkSheet Is Nothing Then
        MsgBox "内部保持データが破棄されました。" + vbCrLf + "修正個所を退避し再読み込みしてください。", vbOKOnly, "警告"
    ElseIf ArrayWorkSheet.Exists(strName) And ExistsSheet(strName) Then
        ArrayWorkSheet.Item(strName).SheetObject.Activate
        If MsgBox("「" + strName + "」シートの内容で更新します。" + vbCrLf + "よろしいですか。", vbOKCancel) = vbOK Then
            ArrayWorkSheet.Item(strName).UpdateDatabase
        End If
    Else
        MsgBox "「" + strName + "」データが読み込まれていません。" + vbCrLf + "修正個所を退避し再読み込みしてください。", vbOKOnly, "警告"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'汎用削除処理
Private Sub 削除(ByVal ID As Long)
    On Error GoTo ErrHandle
    Dim strName As String
    strName = NameById(ID)
    If ArrayWorkSheet Is Nothing Then
        MsgBox "内部保持データが破棄されました。" + vbCrLf + "修正個所を退避し再読み込みしてください。", vbOKOnly, "警告"
    ElseIf ArrayWorkSheet.Exists(strName) And ExistsSheet(strName) Then
        ArrayWorkSheet.Item(strName).SheetObject.Activate
        If MsgBox("「" + strName + "」シート " & ActiveCell.row & "行目のデータを削除します。" + vbCrLf + "よろしいですか。", vbOKCancel) = vbOK Then
            ArrayWorkSheet.Item(strName).DeleteRecord ActiveCell.row
        End If
    Else
        MsgBox "「" + strName + "」データが読み込まれていません。" + vbCrLf + "修正個所を退避し再読み込みしてください。", vbOKOnly, "警告"
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
    '宣言
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New Summary
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM Summary WHERE 状況<>N'破棄済' OR 状況 IS NULL"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "P,VM名,依存先,貸出依頼者,担当者,Type,ホスト名,IPアドレス,内容,状況,予定,ラックNO,保守契約,シリアル番号,製品番号,備考,メーカー,マシン,資産番号,現物設置フロア,管理外"
    objList.GroupColNames = "ID,VM名,依存先,貸出依頼者,担当者,保守,保守窓口,保守契約,シリアル番号,製品番号,備考,メーカー,ユニットコード,ユニット名,資産名１(資産名称),資産名２(メーカー名・型番),数量,取得日,ビル名,現物照合用情報(付属機器・設置場所等),現物管理担当者（使用者）名,現物設置フロア,区分,ラベル枚数,IPソート用"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "開発環境"
    
    'オブジェクト保持
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
    If ActiveSheet.Name = "開発環境" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "内部メモリがクリアされました。" & vbCrLf & "変更箇所を退避し、一覧を読み直してください。"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "「開発環境」シート専用の機能です"
    End If
    Exit Sub
ErrHandle:
    Debug.Print Err.Source
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'GXサーバー一覧
Sub Button21(ByVal control As IRibbonControl)
    Dim objDB As DatabaseInterface
    Dim objList As DataListInterface
    Dim objSheet As ViewControllerInterface
    Dim strViewName As String
    Dim strSql As String

    'SQL作成
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "    GXServer.ID, "
    strSql = strSql & "    GXServer.ホスト名, "
    strSql = strSql & "    GXServer.IPアドレス, "
    strSql = strSql & "    GXServer.種類, "
    strSql = strSql & "    GXServer.VM名, "
    strSql = strSql & "    GXServer.VMサーバー名, "
    strSql = strSql & "    GXServer.貸出依頼者, "
    strSql = strSql & "    GXServer.担当者, "
    strSql = strSql & "    GXServer.内容, "
    strSql = strSql & "    GXServer.状況, "
    strSql = strSql & "    GXServer.予定, "
    strSql = strSql & "    GXServer.マシン, "
    strSql = strSql & "    GXServer.割り当てCPU, "
    strSql = strSql & "    GXServer.割り当てメモリ, "
    strSql = strSql & "    GXServer.ディスク容量, "
    strSql = strSql & "    SoftWare.software_name AS OS, "
    strSql = strSql & "    STUFF((SELECT CAST(',' AS VARCHAR(max)) + software_name FROM SoftWare WHERE dependence_place = GXServer.ID ORDER BY SoftWare.software_name FOR XML PATH('')),1,1,'') AS ソフト, "
    strSql = strSql & "    GXServer.備考, "
    strSql = strSql & "    GXServer.対象フラグ, "
    strSql = strSql & "    GXServer.IPソート用 "
    strSql = strSql & "FROM "
    strSql = strSql & "    GXServer "
    strSql = strSql & "    LEFT JOIN "
    strSql = strSql & "    SoftWare ON "
    strSql = strSql & "    GXServer.ID = SoftWare.object_id "
    strSql = strSql & "WHERE "
    strSql = strSql & "    状況<>N'破棄済' OR "
    strSql = strSql & "    状況 IS NULL "
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New GXServer
    objList.StartCol = "A1"
    objList.QuerySQL = strSql
    objList.KeyColNames = "ID"
    objList.UpdColNames = "ホスト名,IPアドレス,種類,VM名,VMサーバー名,貸出依頼者,担当者,内容,状況,予定,マシン,割り当てCPU,割り当てメモリ,ディスク容量,備考,対象フラグ"
    objList.GroupColNames = "割り当てCPU,割り当てメモリ,ディスク容量,対象フラグ,IPソート用"
    objList.GroupRow = "対象フラグ,EQ,0"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "GX環境一覧"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button22(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "GX環境一覧" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "内部メモリがクリアされました。" & vbCrLf & "変更箇所を退避し、一覧を読み直してください。"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "「GX環境一覧」シート専用の機能です"
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
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New HardWareList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM HWList WHERE 状態<>N'破棄済' OR 状態 IS NULL ORDER BY 管理外,ラックNo,[GROUP],接続先,種類"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "接続先,種類,メーカー,マシン,製品番号,CPU,メモリ,ディスク容量,シリアル番号,フロア,ラックNo,位置,内容,資産コード,保守契約,ラベル番号,ラベル枚数,管理外,プロダクト,状態,予定,貸出依頼者,担当者,備考"
    objList.HiddenColNames = "GROUP"
    objList.GroupRow = "接続先,NE,"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "ハードウェア"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button32(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "ハードウェア" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "内部メモリがクリアされました。" & vbCrLf & "変更箇所を退避し、一覧を読み直してください。"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "「ハードウェア」シート専用の機能です"
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
    '宣言
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New Summary
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM Summary WHERE 状況=N'破棄済'"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "P,VM名,依存先,貸出依頼者,担当者,Type,ホスト名,IPアドレス,内容,状況,予定,ラックNO,保守契約,シリアル番号,製品番号,備考,メーカー,マシン,資産番号,現物設置フロア,管理外"
    objList.GroupColNames = "ID,VM名,依存先,貸出依頼者,担当者,保守,保守窓口,保守契約,シリアル番号,製品番号,備考,メーカー,ユニットコード,ユニット名,資産名１(資産名称),資産名２(メーカー名・型番),数量,取得日,ビル名,現物照合用情報(付属機器・設置場所等),現物管理担当者（使用者）名,現物設置フロア,区分,ラベル枚数,IPソート用"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "破棄済一覧"
    
    'オブジェクト保持
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button62(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "破棄済一覧" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "内部メモリがクリアされました。" & vbCrLf & "変更箇所を退避し、一覧を読み直してください。"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).UpdateDatabase
                End If
            Next
        End If
    Else
        MsgBox "「破棄済一覧」シート専用の機能です"
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
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New ExDiskList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM ExDiskList"
    objList.GroupColNames = "object_id,dependence_place,IPソート用"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet "外部ディスク"
    
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button81(ByVal control As IRibbonControl)
    '宣言
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New VMList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM VMList"
    objList.GroupColNames = "object_id,dependence_place,管理外"
    objList.KeyColNames = "object_id"
    objList.UpdColNames = "VMサーバー名,VM名,ホスト名,IPアドレス,貸出依頼者,担当者,内容,状況,予定,管理外,備考"
    objList.ColToTable = "Server,Server,Address,Server,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.仮想マシン)
    
    'データ保持
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button82(ByVal control As IRibbonControl)
    更新 Names.仮想マシン
End Sub

Sub Button83(ByVal control As IRibbonControl)
    On Error GoTo ErrHandle
    Dim strKey As Variant
    If ActiveSheet.Name = "仮想マシン" Then
        If ArrayWorkSheet Is Nothing Then
            MsgBox "内部保持データが破棄されました。" + vbCrLf + "修正個所を退避し再読み込みしてください。"
        Else
            For Each strKey In ArrayWorkSheet.Keys
                If ArrayWorkSheet.Item(strKey).SheetObject.Name = ActiveSheet.Name Then
                    ArrayWorkSheet.Item(strKey).DeleteRecord ActiveCell.row
                End If
            Next
        End If
    Else
        MsgBox "「仮想マシン」シート専用の機能です"
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
        MsgBox "内部保持データが破棄されました。" + vbCrLf + "修正個所を退避し再読み込みしてください。", vbOKOnly, "警告"
        Exit Sub
    ElseIf ArrayWorkSheet.Exists("仮想マシン") Then
    Else
        MsgBox "対象シートが読み込まれていません。", vbOKOnly, "警告"
        Exit Sub
    End If
    
    ArrayWorkSheet.Item("仮想マシン").SheetObject.Activate
    
    For intRow = 2 To Range("A1").End(xlDown).row
        If ThisWorkbook.pingIp(Cells(intRow, 6).Value) Then
            Cells(intRow, 6).Font.Color = RGB(0, 0, 0)
        Else
            Cells(intRow, 6).Font.Color = RGB(255, 0, 0)
        End If
    Next
End Sub

Sub Button85()
    '宣言
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New VMList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM VMHostList UNION ALL SELECT * FROM VMGuestList ORDER BY dependence_place, is_guest"
    objList.GroupColNames = "object_id,dependence_place,is_guest,管理外"
'    objList.KeyColNames = "object_id"
'    objList.UpdColNames = "VMサーバー名,VM名,ホスト名,IPアドレス,貸出依頼者,担当者,内容,状況,予定,管理外,備考"
'    objList.ColToTable = "Server,Server,Address,Server,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster,ObjectMaster"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.仮想マシン)
    
    'データ保持
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

'ソフトウェア一覧
Sub Button91(ByVal control As IRibbonControl)
    '宣言
    Dim objDB As DatabaseInterface, objList As DataListInterface, objSheet As ViewControllerInterface
    
    '接続
    Set objDB = getConObj
    
    'リスト
    Set objList = New SoftwareList
    objList.StartCol = "A1"
    objList.QuerySQL = "SELECT * FROM SWList ORDER BY 導入先ID,依存先"
    objList.GroupColNames = "ID,依存先,シリアル番号,導入先ID"
    objList.KeyColNames = "ID"
    objList.UpdColNames = "依存先,ソフトウェア名,メーカー,シリアル番号"
    objList.ColToTable = "SoftWare,SoftWare,SoftWare,SoftWare"
    objList.TitleThemeColor = xlThemeColorAccent3
    Set objList.Connector = objDB
    
    'シート
    Set objSheet = New ViewControllerInterfaceImpl
    objSheet.BoolAddAfterSheet = False
    objSheet.AddDataList objList
    objSheet.CreateSheet NameById(Names.ソフトウェア)
    
    'データ保持
    If ArrayWorkSheet Is Nothing Then
        Set ArrayWorkSheet = CreateObject("Scripting.Dictionary")
    End If
    Set ArrayWorkSheet.Item(objSheet.SheetObject.Name) = objSheet
End Sub

Sub Button92(ByVal control As IRibbonControl)
    更新 Names.ソフトウェア
End Sub

Sub Button93(ByVal control As IRibbonControl)
    削除 Names.ソフトウェア
End Sub

