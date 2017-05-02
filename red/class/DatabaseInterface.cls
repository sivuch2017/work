VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Provider
Property Let Provider(ByVal Value As String)
End Property

'Driver
Property Let Driver(ByVal Value As String)
End Property

'Data Source
Property Let DataSource(ByVal Value As String)
End Property

'Server
Property Let Server(ByVal Value As String)
End Property

'Persist Security Info
Property Let PersistSecurityInfo(ByVal Value As String)
End Property

'Initial Catalog
Property Let InitialCatalog(ByVal Value As String)
End Property

'Database
Property Let Database(ByVal Value As String)
End Property

'User ID
Property Let UserID(ByVal Value As String)
End Property

'Password
Property Let Password(ByVal Value As String)
End Property

'ê⁄ë±ï∂éöóÒ
Public Function GetConnectString() As String
End Function

'DBê⁄ë±
Public Function ConnectDatabase() As ADODB.Connection
End Function

'AccessÉäÉìÉNå^Ç©ÅH
Public Function IsLinkType(ByVal TableName As String, ColumnName As String) As Boolean
End Function

