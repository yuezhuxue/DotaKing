VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm_DotaKing 
   Caption         =   "中国刀塔王 - 模式辅助"
   ClientHeight    =   7455
   ClientLeft      =   7530
   ClientTop       =   4710
   ClientWidth     =   10395
   Icon            =   "frm_DotaKing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10395
   Begin VB.Frame Frame1 
      Caption         =   "程序版本v1.0"
      Height          =   1695
      Left            =   6720
      TabIndex        =   17
      Top             =   5640
      Width           =   3495
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "本软件开源项目地址"
         Height          =   225
         Left            =   600
         TabIndex        =   20
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "中国刀塔王的QQ群532519139"
         Height          =   180
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   2250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "中国刀塔王熊猫tv直播间"
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   1980
      End
   End
   Begin VB.CheckBox id 
      Caption         =   "id：死亡掉物品不掉钱"
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      ToolTipText     =   "如果随机的那个格子有物品的话"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "点击播放刀塔王战歌"
      Height          =   615
      Left            =   840
      TabIndex        =   14
      ToolTipText     =   "任意mp3文件改名为1.mp3丢到程序目录"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点击复制左边文字"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   12
      Top             =   480
      Width           =   2055
   End
   Begin VB.CheckBox mi 
      Caption         =   "mi：英雄模型缩小一半"
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      ToolTipText     =   "英雄模型缩小一半"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   6495
   End
   Begin VB.CheckBox sp 
      Caption         =   "sp：随机交换位置"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "随机交换位置"
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox om 
      Caption         =   "om：只有中路出兵"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "只有中路出兵"
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox em 
      Caption         =   "em：经验金钱更多"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "经验金钱更多"
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox sh 
      Caption         =   "sh：相同英雄"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "相同英雄"
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox np 
      Caption         =   "np：禁止神符生成"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "禁止神符生成"
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox pm 
      Caption         =   "pm：物品他人拿了有效"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "物品他人拿了有效"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox sc 
      Caption         =   "sc：会出精英小兵"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "会出精英小兵"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox fr 
      Caption         =   "fr：复活时间减半"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "复活时间减半"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6960
   End
   Begin VB.CheckBox ap 
      Caption         =   "ap：全阵营选择"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "全阵营选择"
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "模式输入后2-15秒内输入-noneutrals禁止刷新野怪"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3960
      Width           =   4575
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   1095
      Left            =   840
      TabIndex        =   13
      Top             =   5520
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
      Caption         =   "按Ctrl+C复制以下文字,颜色变化则复制成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_DotaKing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'单击打开网页声明
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private WithEvents hk As clsRegHotKeys
Attribute hk.VB_VarHelpID = -1
Dim FrTxt(1 To 11) As String
Dim a, b, c, i As Integer
Public Grtxt As String



Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Grtxt
a = Int(Rnd * 256)
b = Int(Rnd * 256)
c = Int(Rnd * 256)
Text1.ForeColor = RGB(a, b, c)
End Sub
'播放音乐
Private Sub Command2_Click()
WMP.URL = App.Path & "\1.mp3" ' 歌曲的路径\
'Print App.Path & "1.mp3"
WMP.settings.volume = 100
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
Grtxt = "-"
For i = 1 To 11
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
Timer1.Enabled = True

          Set hk = New clsRegHotKeys
          hk.RegHotKeys Me.hwnd, ctrlKey, vbKeyC, "C"
          hk.RegHotKeys Me.hwnd, ctrlKey, vbKeyD, "D"
          Me.Show   '这个不能省略，否则窗体无法显示出来！
            
          hk.WaitMsg
    
  End Sub
    


Private Sub Form_Unload(Cancel As Integer)
End

End Sub

  Private Sub hk_HotKeysDown(Key As String)
          Select Case Key
                  'Case "C"
                  '        MsgBox "你按了Ctrl+C   !"
                  Case "C"
                          
                          
                          '必须要有Clear，不然无法复制
                          Clipboard.Clear
                          Clipboard.SetText Grtxt
a = Int(Rnd * 256)
b = Int(Rnd * 256)
c = Int(Rnd * 256)
Text1.ForeColor = RGB(a, b, c)

                        'MsgBox "文本输入成功"
          End Select
  End Sub
  
Private Sub ap_Click()
If ap.Value = 1 Then
    FrTxt(1) = "ap"
    Call CheckString
Else
    FrTxt(1) = ""
    Call CheckString
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
If id.Value = 1 Then
    FrTxt(11) = "id"
    Call CheckString
Else
    FrTxt(11) = ""
    Call CheckString
End If
End Sub


Private Sub CheckString()
Grtxt = "-"
For i = 1 To 11
Grtxt = Grtxt + FrTxt(i)
Next
Text1.Text = Grtxt
If Grtxt = "-apspomemshnppmscfrmiid" Then
    MsgBox "DotA最多支持10个模式，请取消1个勾选"
    id.Value = False
    FrTxt(11) = ""
End If
End Sub






'超链接
Private Sub Label3_Click()
ShellExecute Me.hwnd, "open", "http://www.panda.tv/33861", "", "", 1
End Sub
Private Sub Label4_Click()
ShellExecute Me.hwnd, "open", "http://qun.qzone.qq.com/group#!/532519139/home", "", "", 1
End Sub
Private Sub Label5_Click()
ShellExecute Me.hwnd, "open", "https://github.com/yuezhuxue/DotaKing", "", "", 1
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

'超链接颜色恢复
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

