VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm_DotaKing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�й������� - ģʽ����"
   ClientHeight    =   9765
   ClientLeft      =   1470
   ClientTop       =   3240
   ClientWidth     =   13575
   Icon            =   "frm_DotaKing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   13575
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   9240
   End
   Begin VB.CommandButton cmdhero 
      Caption         =   "���һ��Ӣ�۰�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   28
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "ģʽ�Ƽ�"
      Height          =   615
      Left            =   600
      TabIndex        =   23
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option3 
         Caption         =   "�������"
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��·��ͬӢ��5v5"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�Զ���"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CheckBox dm 
      Caption         =   "dm������ģʽ"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      ToolTipText     =   "������Ӣ��"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CheckBox ar 
      Caption         =   "ar�����Ӣ��"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   21
      ToolTipText     =   "���Ӣ��"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "����汾v1.1"
      Height          =   1695
      Left            =   9960
      TabIndex        =   17
      Top             =   7920
      Width           =   3495
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�������Դ��Ŀ��ַ"
         Height          =   180
         Left            =   600
         TabIndex        =   20
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�й���������QQȺ532519139"
         Height          =   180
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   2250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�й���������ètvֱ����"
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   1980
      End
   End
   Begin VB.CheckBox id 
      Caption         =   "id����������Ʒ����Ǯ"
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      ToolTipText     =   "���������Ǹ���������Ʒ�Ļ�"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������ŵ�����ս��"
      Height          =   615
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "����mp3�ļ�����Ϊ1.mp3��������Ŀ¼"
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   12
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox mi 
      Caption         =   "mi��Ӣ��ģ����Сһ��"
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      ToolTipText     =   "Ӣ��ģ����Сһ��"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   6495
   End
   Begin VB.CheckBox sp 
      Caption         =   "sp���������λ��"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "�������λ��"
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox om 
      Caption         =   "om��ֻ����·����"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "ֻ����·����"
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox em 
      Caption         =   "em�������Ǯ����"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "�����Ǯ����"
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox sh 
      Caption         =   "sh����ͬӢ��"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "��ͬӢ��"
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox np 
      Caption         =   "np����ֹ�������"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "��ֹ�������"
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox pm 
      Caption         =   "pm����Ʒ����������Ч"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "��Ʒ����������Ч"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox sc 
      Caption         =   "sc�������ӢС��"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "�����ӢС��"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox fr 
      Caption         =   "fr������ʱ�����"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "����ʱ�����"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   9240
   End
   Begin VB.CheckBox ap 
      Caption         =   "ap��ȫ��Ӫѡ��"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "ȫ��Ӫѡ��"
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Label lblhero 
      AutoSize        =   -1  'True
      Caption         =   "��·���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   720
      TabIndex        =   27
      Top             =   6720
      Width           =   3360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ģʽ�����2-15��������-noneutrals��ֹˢ��Ұ��"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   600
      TabIndex        =   15
      Top             =   4320
      Width           =   4050
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   1095
      Left            =   6000
      TabIndex        =   13
      Top             =   8640
      Visible         =   0   'False
      Width           =   3495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6165
      _cy             =   1931
   End
   Begin VB.Label Label1 
      Caption         =   "��Ctrl+C������������,��ɫ�仯���Ƴɹ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "frm_DotaKing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��������ҳ����
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private WithEvents hk As clsRegHotKeys
Attribute hk.VB_VarHelpID = -1
Dim FrTxt(1 To 13) As String
Dim a, b, c, i As Integer
Public Grtxt As String
Public I_second As Integer
Dim StrHero() As String
Public StrCount As Integer





Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Grtxt
a = Int(Rnd * 256)
b = Int(Rnd * 256)
c = Int(Rnd * 256)
Text1.ForeColor = RGB(a, b, c)
End Sub
'��������
Private Sub Command2_Click()
If Dir(App.Path & "\1.mp3") <> "" Then
WMP.URL = App.Path & "\1.mp3" ' ������·��\
'Print App.Path & "1.mp3"
WMP.settings.volume = 100
Else
MsgBox "���󣺳���Ŀ¼������1.mp3"
Exit Sub
End If
End Sub



Private Sub Form_Click()
'Print Grtxt
End Sub

Private Sub Form_Load()
FrTxt(1) = "ap"
FrTxt(2) = "sp"
FrTxt(3) = "om"
FrTxt(4) = "em"
FrTxt(5) = "sh"
FrTxt(6) = "np"
FrTxt(7) = ""
FrTxt(8) = ""
FrTxt(9) = ""
FrTxt(10) = ""
FrTxt(11) = ""
FrTxt(12) = ""
FrTxt(13) = ""
Grtxt = "-"
For i = 1 To 13
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
'Timer1.Enabled = True

          Set hk = New clsRegHotKeys
          hk.RegHotKeys Me.hwnd, ctrlKey, vbKeyC, "C"
          hk.RegHotKeys Me.hwnd, ctrlKey, vbKeyD, "D"
          Me.Show   '�������ʡ�ԣ��������޷���ʾ������
            
          hk.WaitMsg
    
  End Sub
    


Private Sub Form_Unload(Cancel As Integer)
End

End Sub

  Private Sub hk_HotKeysDown(Key As String)
          Select Case Key
                  'Case "C"
                  '        MsgBox "�㰴��Ctrl+C   !"
                  Case "C"
                          
                          
                          '����Ҫ��Clear����Ȼ�޷�����
                          Clipboard.Clear
                          Clipboard.SetText Grtxt
a = Int(Rnd * 256)
b = Int(Rnd * 256)
c = Int(Rnd * 256)
Text1.ForeColor = RGB(a, b, c)

                        'MsgBox "�ı�����ɹ�"
          End Select
  End Sub
  



Private Sub Option2_Click()
ap.Value = 1
sp.Value = 1
om.Value = 1
em.Value = 1
sh.Value = 1
np.Value = 1
pm.Value = 0
sc.Value = 1
fr.Value = 0
mi.Value = 0
ID.Value = 0
ar.Value = 0
dm.Value = 0
FrTxt(1) = "ap"
FrTxt(2) = "sp"
FrTxt(3) = "om"
FrTxt(4) = "em"
FrTxt(5) = "sh"
FrTxt(6) = "np"
FrTxt(7) = ""
FrTxt(8) = "sc"
FrTxt(9) = ""
FrTxt(10) = ""
FrTxt(11) = ""
FrTxt(12) = ""
FrTxt(13) = ""
Grtxt = "-"
For i = 1 To 13
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
End Sub

Private Sub Option3_Click()
ap.Value = 0
sp.Value = 1
om.Value = 1
em.Value = 1
sh.Value = 0
np.Value = 1
pm.Value = 0
sc.Value = 1
fr.Value = 0
mi.Value = 0
ID.Value = 0
ar.Value = 1
dm.Value = 1
FrTxt(1) = ""
FrTxt(2) = "sp"
FrTxt(3) = "om"
FrTxt(4) = "em"
FrTxt(5) = ""
FrTxt(6) = "np"
FrTxt(7) = ""
FrTxt(8) = "sc"
FrTxt(9) = ""
FrTxt(10) = ""
FrTxt(11) = ""
FrTxt(12) = "ar"
FrTxt(13) = "dm"
Grtxt = "-"
For i = 1 To 13
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
End Sub
Private Sub ap_Click()
If ap.Value = 1 Then
    FrTxt(1) = "ap"
    Call CheckString
    ar.Enabled = False
    ar.Value = 0
    Label2.Caption = "ģʽ�����2-15��������-noneutrals��ֹˢ��Ұ��"
Else
    FrTxt(1) = ""
    Call CheckString
    ar.Enabled = True
    ar.Value = 1
    Label2.Caption = "ģʽ�����2-15��������-ndȡ������ȴ�ʱ�䣬����-lives #����������"
End If
End Sub
Private Sub sp_Click()
If sp.Value = 1 Then
    FrTxt(2) = "sp"
    Call CheckString
Else
    FrTxt(2) = ""
    Call CheckString
End If
End Sub
Private Sub om_Click()
If om.Value = 1 Then
    FrTxt(3) = "om"
    Call CheckString
Else
    FrTxt(3) = ""
    Call CheckString
End If
End Sub
Private Sub em_Click()
If em.Value = 1 Then
    FrTxt(4) = "em"
    Call CheckString
Else
    FrTxt(4) = ""
    Call CheckString
End If
End Sub
Private Sub sh_Click()
If sh.Value = 1 Then
    FrTxt(5) = "sh"
    Call CheckString
Else
    FrTxt(5) = ""
    Call CheckString
End If
End Sub
Private Sub np_Click()
If np.Value = 1 Then
    FrTxt(6) = "np"
    Call CheckString
Else
    FrTxt(6) = ""
    Call CheckString
End If
End Sub
Private Sub pm_Click()
If pm.Value = 1 Then
    FrTxt(7) = "pm"
    Call CheckString
Else
    FrTxt(7) = ""
    Call CheckString
End If
End Sub
Private Sub sc_Click()
If sc.Value = 1 Then
    FrTxt(8) = "sc"
    Call CheckString
Else
    FrTxt(8) = ""
    Call CheckString
End If
End Sub
Private Sub fr_Click()
If fr.Value = 1 Then
    FrTxt(9) = "fr"
    Call CheckString
Else
    FrTxt(9) = ""
    Call CheckString
End If
End Sub
Private Sub mi_Click()
If mi.Value = 1 Then
    FrTxt(10) = "mi"
    Call CheckString
Else
    FrTxt(10) = ""
    Call CheckString
End If
End Sub
Private Sub id_Click()
If ID.Value = 1 Then
    FrTxt(11) = "id"
    Call CheckString
Else
    FrTxt(11) = ""
    Call CheckString
End If
End Sub
Private Sub ar_Click()
If ar.Value = 1 Then
    FrTxt(12) = "ar"
    Call CheckString
    ap.Enabled = False
    ap.Value = 0
    sh.Enabled = False
    sh.Value = 0
    Label2.Caption = "ģʽ�����2-15��������-ndȡ������ȴ�ʱ�䣬����-lives #����������"
Else
    FrTxt(12) = ""
    Call CheckString
    ap.Enabled = True
    ap.Value = 1
    sh.Enabled = True
    sh.Value = 1
    Label2.Caption = "ģʽ�����2-15��������-noneutrals��ֹˢ��Ұ��"
End If
End Sub
Private Sub dm_Click()
If dm.Value = 1 Then
    FrTxt(13) = "dm"
    Call CheckString
    sh.Value = 0
    sh.Enabled = False
Else
    FrTxt(13) = ""
    Call CheckString
    sh.Value = 1
    sh.Enabled = True
End If
End Sub
Private Sub CheckString()
Grtxt = "-"
For i = 1 To 13
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
If Len(Grtxt) >= 23 Then
    MsgBox "DotA���֧��10��ģʽ����ȡ��1����ѡ"
    ID.Value = False
    FrTxt(11) = ""
    pm.Value = False
    FrTxt(7) = ""
    mi.Value = False
    FrTxt(10) = ""
End If
End Sub


Private Sub cmdhero_Click()
cmdhero.Enabled = False
cmdhero.Caption = "5��"
Timer1.Interval = "100"
Dim i As Integer
i = 0
Timer1.Enabled = True
Open App.Path & "\omhero.ini" For Input As #123
Do While Not EOF(123)
i = i + 1
ReDim Preserve StrHero(i) '��ȷ�Ͻ죬��ֹ�±�Խ��
Line Input #123, StrHero(i)
Loop
StrCount = i
I_second = 5
Timer2.Enabled = True
Close #123
'Print StrHero(110)
End Sub

Private Sub Timer1_Timer()
Randomize

lblhero.Caption = StrHero(Int(Rnd * (StrCount + 1)))
End Sub
Private Sub Timer2_Timer()
I_second = I_second - 1
cmdhero.Caption = I_second & "��"
'If I_second = 4 Then Timer1.Interval = 100
'If I_second = 3 Then Timer1.Interval = 150
'If I_second = 2 Then Timer1.Interval = 200
'If I_second = 1 Then Timer1.Interval = 250




If I_second = 0 Then
Timer2.Enabled = False
cmdhero.Caption = "���һ��Ӣ�۰�"
Timer1.Enabled = False
cmdhero.Enabled = True
End If
End Sub

'������
Private Sub Label3_Click()
ShellExecute Me.hwnd, "open", "http://www.panda.tv/33861", "", "", 1
End Sub
Private Sub Label4_Click()
ShellExecute Me.hwnd, "open", "http://qun.qzone.qq.com/group#!/532519139/home", "", "", 1
End Sub
Private Sub Label5_Click()
ShellExecute Me.hwnd, "open", "https://github.com/yuezhuxue/DotaKing/releases", "", "", 1
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = True
Label3.ForeColor = vbRed
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontUnderline = True
Label4.ForeColor = vbRed
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontUnderline = True
Label5.ForeColor = vbRed
End Sub

'��������ɫ�ָ�
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Label3.ForeColor = vbBlue
Label4.FontUnderline = False
Label4.ForeColor = vbBlue
Label5.FontUnderline = False
Label5.ForeColor = vbBlue
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Label3.ForeColor = vbBlue
Label4.FontUnderline = False
Label4.ForeColor = vbBlue
Label5.FontUnderline = False
Label5.ForeColor = vbBlue
End Sub



