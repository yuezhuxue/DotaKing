VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegHotKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Private Type POINTAPI
          X   As Long
          Y   As Long
  End Type
    
  Private Type Msg
          hwnd   As Long
          Message   As Long
          wParam   As Long
          lParam   As Long
          time   As Long
          pt   As POINTAPI
  End Type
    
  Private Type KeyMsg
          ID   As Long       '   ����ע���ȼ�ʱ��ID
          Key   As String     '����ע���ȼ�ʱ�Ĺؼ���
  End Type
    
  Private Const PM_REMOVE = &H1
  Private Const WM_HOTKEY = &H312
    
  Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
  'id   ֵ��Χ   ��0X0000-0XBFFF
  Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
  Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
  Private Declare Function WaitMessage Lib "user32" () As Long
  '************************************************************
    
  Enum ShiftKeys
          AltKey = &H1
          CtrlKey = &H2
          ShiftKey = &H4
  End Enum
    
  '�ֲ�����
  Private bCancel     As Boolean
  Private clsHwnd     As Long
  Private KeyGroup     As Integer
  Private KeyID     As Long
  Private Keys()     As KeyMsg
    
  '�����¼�
  Public Event HotKeysDown(Key As String)
  'ע���ȼ�,����ע������ȼ�
  Sub RegHotKeys(ByVal hwnd As Long, ByVal ShiftKey As ShiftKeys, ByVal ComKey As KeyCodeConstants, ByVal Key As String)
          On Error Resume Next
          clsHwnd = hwnd
          KeyID = KeyID + 1
          KeyGroup = KeyGroup + 1
          ReDim Preserve Keys(KeyGroup)
          RegisterHotKey hwnd, KeyID, ShiftKey, ComKey                 'ע���ȼ�
          Keys(KeyGroup).ID = KeyID
          Keys(KeyGroup).Key = Trim(Key)
  End Sub
  'ȡ���ȼ�ע��
  Sub UnRegHotKeys(ByVal Key As String)
          On Error Resume Next
          If KeyGroup = 0 Then Exit Sub
          Dim i     As Integer
          For i = 0 To KeyGroup
                  If Trim(Key) = Trim(Keys(i).Key) Then
                          UnregisterHotKey clsHwnd, Keys(i).ID
                  End If
          Next
  End Sub
    
  'ȡ��ȫ���ȼ�ע��
  Sub UnRegAllHotKeys()
          On Error Resume Next
          If KeyGroup = 0 Then Exit Sub
          Dim i     As Integer
          For i = 0 To KeyGroup
                  UnregisterHotKey clsHwnd, Keys(i).ID
          Next
  End Sub
    
  '�Ⱥ򰴼���Ϣ
  Sub WaitMsg()
          On Error Resume Next
          bCancel = False
          Dim Message     As Msg, i       As Integer
          Do While Not bCancel
                  WaitMessage   '�Ⱥ򰴼���Ϣ
                  '�ж���Ϣ
                  If PeekMessage(Message, clsHwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
                          For i = 0 To KeyGroup
                                  If Keys(i).ID = Message.wParam Then           '�жϰ��������ȼ�
                                          RaiseEvent HotKeysDown(Keys(i).Key)     '�����¼�
                                  End If
                          Next
                  End If
                  DoEvents
          Loop
  End Sub
    
  'ȡ���Ⱥ���Ϣ
  Sub UnWaitMsg()
          bCancel = True
  End Sub
    
  Private Sub Class_Initialize()
          KeyID = &H1000&       '��ʼID
          KeyGroup = -1
          ReDim Keys(0)
  End Sub
    
  Private Sub Class_Terminate()
          On Error Resume Next
          bCancel = True
          UnRegAllHotKeys
  End Sub