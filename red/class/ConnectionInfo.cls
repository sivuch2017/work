VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'接続情報セル開始位置
Private Const ThisConPos As String = "D2"

'取得情報セル開始位置
Private Const ThisGetPos As String = "E13"

Public Function ConPos() As String
    ConPos = ThisConPos
End Function

Public Function GetPos() As String
    GetPos = ThisGetPos
End Function

