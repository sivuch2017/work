VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseInterfaceImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DatabaseInterface

Dim ThisProvider As String
Dim ThisDriver As String
Dim ThisDataSource As String
Dim ThisServer As String
Dim ThisPersistSecurityInfo As String
Dim ThisInitialCatalog As String
Dim ThisDatabase As String
Dim ThisUserID As String
Dim ThisPassword As String
Dim ThisCashArray As Object

'接続済みコネクション
Dim ThisCon As ADODB.Connection

'Provider
Property Let DatabaseInterface_Provider(ByVal Value As String)
    ThisProvider = Value
End Property
Private Property Get Provider() As String
    If ThisProvider <> "" Then
        Provider = "Provider=" & ThisProvider & ";"
    Else
        Provider = ""
    End If
End Property

'Driver
Property Let DatabaseInterface_Driver(ByVal Value As String)
    ThisDriver = Value
End Property
Private Property Get Driver() As String
    If ThisDriver <> "" Then
        Driver = "Driver=" & ThisDriver & ";"
    Else
        Driver = ""
    End If
End Property

'Data Source
Property Let DatabaseInterface_DataSource(ByVal Value As String)
    ThisDataSource = Value
End Property
Private Property Get DataSource() As String
    If ThisDataSource <> "" Then
        DataSource = "Data Source=" & ThisDataSource & ";"
    Else
        DataSource = ""
    End If
End Property

'Server
Property Let DatabaseInterface_Server(ByVal Value As String)
    ThisServer = Value
End Property
Private Property Get Server() As String
    If ThisServer <> "" Then
        Server = "Server=" & ThisServer & ";"
    Else
        Server = ""
    End If
End Property

'Persist Security Info
Property Let DatabaseInterface_PersistSecurityInfo(ByVal Value As String)
    ThisPersistSecurityInfo = Value
End Property
Private Property Get PersistSecurityInfo() As String
    If ThisPersistSecurityInfo <> "" Then
        PersistSecurityInfo = "Persist Security Info=" & ThisPersistSecurityInfo & ";"
    Else
        PersistSecurityInfo = ""
    End If
End Property

'Initial Catalog
Property Let DatabaseInterface_InitialCatalog(ByVal Value As String)
    ThisInitialCatalog = Value
End Property
Private Property Get InitialCatalog() As String
    If ThisInitialCatalog <> "" Then
        InitialCatalog = "Initial Catalog=" & ThisInitialCatalog & ";"
    Else
        InitialCatalog = ""
    End If
End Property

'Database
Property Let DatabaseInterface_Database(ByVal Value As String)
    ThisDatabase = Value
End Property
Private Property Get Database() As String
    If ThisDatabase <> "" Then
        Database = "Database=" & ThisDatabase & ";"
    Else
        Database = ""
    End If
End Property

'User ID
Property Let DatabaseInterface_UserID(ByVal Value As String)
    ThisUserID = Value
End Property
Private Property Get UserID() As String
    If ThisUserID <> "" Then
        UserID = "User ID=" & ThisUserID & ";"
    Else
        UserID = ""
    End If
End Property

'Password
Property Let DatabaseInterface_Password(ByVal Value As String)
    ThisPassword = Value
End Property
Private Property Get Password() As String
    If ThisPassword <> "" Then
        Password = "Password=" & ThisPassword & ";"
    Else
        Password = ""
    End If
End Property

'接続文字列
Public Function DatabaseInterface_GetConnectString() As String
    DatabaseInterface_GetConnectString = Provider & Driver & DataSource & Server & PersistSecurityInfo & InitialCatalog & Database & UserID & Password
End Function

'DB接続
Public Function DatabaseInterface_ConnectDatabase() As ADODB.Connection
    'On Error Resume Next
    Dim objCon As ADODB.Connection
    
    If ThisCon Is Nothing Then
    Else
        If ThisCon.State = adStateOpen Then
             Set DatabaseInterface_ConnectDatabase = ThisCon
             Exit Function
        End If
    End If
    
    Set objCon = New ADODB.Connection
    objCon.ConnectionString = DatabaseInterface_GetConnectString
    objCon.Open
    Set ThisCon = objCon
    Set DatabaseInterface_ConnectDatabase = objCon
    
End Function

'Accessリンク型か？
Public Function DatabaseInterface_IsLinkType(ByVal TableName As String, ColumnName As String) As Boolean
    DatabaseInterface_IsLinkType = False
End Function


