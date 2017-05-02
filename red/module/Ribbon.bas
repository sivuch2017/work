Attribute VB_Name = "Ribbon"
'���{������p�Ɉȉ��̃T�C�g����R�s�y
'http://www.ka-net.org/ribbon/ri32.html

Option Explicit

Private Declare Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long

Private Const CHILDID_SELF = 0&
Private Const ROLE_SYSTEM_PAGETABLIST = &H3C
Private Const ROLE_SYSTEM_PAGETAB = &H25

Sub CallMe()
  '�����̓J�X�^���^�u(tab�v�f)��label�����̒l,��������"�A�h�C��"
  'Excel2013��API���ύX���ꂽ���ߎ��s�s��
  'Call SelRibbonTAB("�J����")
End Sub

Sub SelRibbonTAB(myTabName As String)
  Dim myAcc As Office.IAccessible
  Dim TimeLimit As Date
  
  TimeLimit = DateAdd("s", 2, Now())  '���[�v�̐�������:2�b
  Set myAcc = Application.CommandBars("Ribbon")
  Set myAcc = GetAcc(myAcc, "���{�� �^�u", ROLE_SYSTEM_PAGETABLIST)
  
  On Error Resume Next
  Do
    Set myAcc = GetAcc(myAcc, myTabName, ROLE_SYSTEM_PAGETAB)
    DoEvents
    If Now() > TimeLimit Then Exit Do  '�������Ԃ��߂����烋�[�v�𔲂���
  Loop While myAcc Is Nothing
  On Error GoTo 0
  
  If Not myAcc Is Nothing Then
    myAcc.accDoDefaultAction (CHILDID_SELF)
    Set myAcc = Nothing
  End If
End Sub

Private Function GetAcc(myAcc As Office.IAccessible, myAccName As String, myAccRole As Long) As Office.IAccessible
  Dim ReturnAcc As Office.IAccessible
  Dim ChildAcc As Office.IAccessible
  Dim List() As Variant
  Dim Count As Long
  Dim i As Long
  
  If (myAcc.accState(CHILDID_SELF) <> 32769) And _
     (myAcc.accName(CHILDID_SELF) = myAccName) And _
     (myAcc.accRole(CHILDID_SELF) = myAccRole) Then
    Set ReturnAcc = myAcc
  Else
    Count = myAcc.accChildCount
    
    If Count > 0& Then
      ReDim List(Count - 1&)
      If AccessibleChildren(myAcc, 0&, ByVal Count, List(0), Count) = 0& Then
        For i = LBound(List) To UBound(List)
          If TypeOf List(i) Is Office.IAccessible Then
            Set ChildAcc = List(i)
            Set ReturnAcc = GetAcc(ChildAcc, myAccName, myAccRole)
            If Not ReturnAcc Is Nothing Then Exit For
          End If
        Next
      End If
    End If
    
  End If
  
  Set GetAcc = ReturnAcc
End Function


