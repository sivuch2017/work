VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EXCELStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intMaxCount As Integer
Private intCount As Integer

'�R���X�g���N�^
Private Sub Class_Initialize()
    intMaxCount = 0
    intCount = 0
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Application.StatusBar = False
End Sub

Public Sub Init(Optional ByVal MaxCount As Integer = 0)
    intMaxCount = MaxCount
    Application.StatusBar = "������..." & String(intMaxCount, "��")
End Sub

Public Sub SetMax(ByVal MaxCount As Integer)
    intMaxCount = MaxCount
    If intCount >= intMaxCount Then
        Application.StatusBar = "������..." & String(intCount, "��")
    Else
        Application.StatusBar = "������..." & String(intCount, "��") & String(intMaxCount - intCount, "��")
    End If
End Sub

Public Sub CountUp()
    intCount = intCount + 1
    If intCount >= intMaxCount Then
        Application.StatusBar = "������..." & String(intCount, "��")
    Else
        Application.StatusBar = "������..." & String(intCount, "��") & String(intMaxCount - intCount, "��")
    End If
End Sub