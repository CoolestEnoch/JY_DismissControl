VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleMode       =   0  'User
   ScaleWidth      =   6675
   Begin VB.CommandButton win_dsk 
      Caption         =   "小窗桌面"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton loc_help 
      Caption         =   "不知道怎么看StudentMain.exe的文件位置?"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Frame location_guide 
      Caption         =   "如果上面的重启与启动按钮不奏效，请在此填入极域StudentMain.exe的目录..."
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   6495
      Begin VB.TextBox jiyu_loc 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton update_log 
      Caption         =   "更新日志"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer bald 
      Interval        =   300
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer gank_top 
      Interval        =   1
      Left            =   5880
      Top             =   1200
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关于作者"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton open 
      Caption         =   "启动控制器"
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton rest 
      Caption         =   "重启控制器"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton no_ctr 
      Caption         =   "解除控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label virus 
      Caption         =   "点按钮解除控制。软件使用vb制作，报毒正常。"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label title 
      Caption         =   "       极域电子教室控制解除工具"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ver As String '版本号变量定义,版本号修改在Form1_Load处修改
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '窗口置顶API
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long  '小窗申明
Private Declare Function GetForegroundWindow Lib "user32" () As Long  '小窗声明
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long  '小窗声明

Private Sub loc_help_Click()
MsgBox "右键桌面上的极域电子教室的快捷方式，点属性，找到“目标”栏里的文本，把它复制到输入框即，再点上面的重启或者启动控制器按钮即可。", okobly, "如何查看文件位置"
End Sub

Private Sub no_ctr_Click()
Shell "cmd /c taskkill /f /im Studentmain.exe", 0
MsgBox "解除成功!--by帥の智叟L", okonly, "提示"
End Sub

Private Sub rest_Click()
Dim stum As New FileSystemObject
lloc = jiyu_loc.Text
llong = Len(lloc)
If lloc = "" Then
 If stum.FileExists("C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe") Then
  Shell "cmd /c taskkill /f /im Studentmain.exe", 0
  Shell "C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe", 1
  MsgBox "启动成功!--by帥の智叟L", okonly, "提示"
 Else
  MsgBox "在预置目录没找到极域主程序！请确保这是学校信息教室电脑或此电脑已安装极域电子教室学生机版！", okonly, "提示"
 End If
Else
 If llong < 5 Then
  MsgBox "目录输入有误!请检查。", okonly, "提示"
  jiyu_loc.SetFocus
 ElseIf llong >= 5 Then
  sstr = Mid(lloc, 1, 1)
  eend = Mid(lloc, llong, 1)
   If Asc(sstr) = 34 And Asc(eend) = 34 Then
     flloc = Mid(lloc, 2, llong - 2)
     llong = llong - 2
    ElseIf Asc(sstr) = 34 And Asc(eend) <> 34 Then
     flloc = Mid(lloc, 2, llong - 1)
     llong = llong - 1
    ElseIf Asc(sstr) <> 34 And Asc(eend) = 34 Then
     flloc = Mid(lloc, 1, llong - 1)
     llong = llong - 1
    ElseIf Asc(sstr) <> 34 And Asc(eend) <> 34 Then
     flloc = Mid(lloc, 1, llong)
   End If
  nname = Mid(flloc, llong - 3, 4)
   If nname = ".exe" Then
     If stum.FileExists(flloc) Then
      Shell "cmd /c taskkill /f /im Studentmain.exe", 0
      Shell flloc, 0
      MsgBox "启动成功!--by帥の智叟L", okonly, "提示"
     Else
      MsgBox "目录输入有误!请检查。", okonly, "提示"
     End If
   ElseIf nname <> ".exe" Then
      MsgBox "目录输入有误!请检查。", okonly, "提示"
     End If
   End If
 End If
End Sub

Private Sub open_Click()
Dim stum As New FileSystemObject
lloc = jiyu_loc.Text
llong = Len(lloc)
If lloc = "" Then
 If stum.FileExists("C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe") Then
  Shell "C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe", 1
  MsgBox "启动成功!--by帥の智叟L", okonly, "提示"
 Else
  MsgBox "在预置目录没找到极域主程序！请确保这是学校信息教室电脑或此电脑已安装极域电子教室学生机版！", okonly, "提示"
 End If
Else
 If llong < 5 Then
  MsgBox "目录输入有误!请检查。", okonly, "提示"
  jiyu_loc.SetFocus
 ElseIf llong >= 5 Then
  sstr = Mid(lloc, 1, 1)
  eend = Mid(lloc, llong, 1)
   If Asc(sstr) = 34 And Asc(eend) = 34 Then
     flloc = Mid(lloc, 2, llong - 2)
     llong = llong - 2
    ElseIf Asc(sstr) = 34 And Asc(eend) <> 34 Then
     flloc = Mid(lloc, 2, llong - 1)
     llong = llong - 1
    ElseIf Asc(sstr) <> 34 And Asc(eend) = 34 Then
     flloc = Mid(lloc, 1, llong - 1)
     llong = llong - 1
    ElseIf Asc(sstr) <> 34 And Asc(eend) <> 34 Then
     flloc = Mid(lloc, 1, llong)
   End If
  nname = Mid(flloc, llong - 3, 4)
   If nname = ".exe" Then
     If stum.FileExists(flloc) Then
      Shell flloc, 0
      MsgBox "启动成功!--by帥の智叟L", okonly, "提示"
     Else
      MsgBox "目录输入有误!请检查。", okonly, "提示"
     End If
   ElseIf nname <> ".exe" Then
      MsgBox "目录输入有误!请检查。", okonly, "提示"
     End If
   End If
 End If
End Sub

Private Sub Command4_Click()
MsgBox "作者:b站 帥の智叟L.各位给一个关注就行!!!!", okonly, "关于作者"
End Sub

Private Sub Form_click()
MsgBox "啊啊啊痒", okonly, "干蛤啊你"
End Sub

Private Sub Form_Load()
ver = "2.5" '版本号
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1 '窗口置顶
Form1.Caption = "老师控制解除器" & ver & "-by b站、帥の智叟L     "
End Sub

Private Sub update_log_Click()
MsgBox "1.0：第一个解控器诞生。" + vbCrLf + "2.0：大更新！界面优化，增加“重启”、“启动”按钮。" + vbCrLf + "2.1：增加窗口强行置顶的功能，使屏幕广播时也能显示解控器。" + vbCrLf + "2.2：紧急修复强行置在屏幕广播时失效的bug。" + vbCrLf + "2.3：固定了窗口的大小，同时禁用了窗口的最大化，防止误触导致界面布局变化。顺便优化了下软件界面(Duang！加特效！很光很亮很油很..咳咳)" + vbCrLf + "2.4：增加自定义极域的源文件目录，并且智能判断目录语法是否有错误，增强了兼容性。" + vbCrLf + "2.5：修复了一些无故闪退问题，增加了稳定性，修复了自定义极域目录的找不到文件就闪退的bug。禁用了软件的最小化。改正了更新日志里一个错别字。修复了标题变色导致的几率性闪退。" + vbCrLf + "3.0：正在开发，但愿不要出些什么新bug就好。计划加入把极域的全屏屏幕广播强行变成小窗模式(欸？谁养的鸽子？“咕咕咕.....”)。", okonly, "更新日志"
End Sub

Private Sub virus_Click()
MsgBox "作者是不是很帅", okonly, "大帅比"
End Sub

Private Sub title_Click()
MsgBox "作者是不是很帅", okonly, "大帅比"
End Sub

Private Sub gank_top_Timer()
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1
Form1.Height = 4935
Form1.Width = 6915
End Sub

Private Sub bald_Timer()
title = title.Caption
frmt = Form1.Caption
ctr = no_ctr.Caption
 n = Len(frmt)
 m = Len(virus)
 Form1.Caption = Mid(frmt, 2, n - 1) + Mid(frmt, 1, 1)
 virus.Caption = Mid(virus, 2, n - 1) + Mid(virus, 1, 1) '跑马灯
'开始变色
Randomize
aa = Int(255 * Rnd + 1)
aaa = Int(250 * Rnd - 5)
aaaa = Int(260 * Rnd + 5)
bb = Int(255 * Rnd + 1)
bbb = Int(250 * Rnd - 5)
bbbb = Int(260 * Rnd + 5)
cc = Int(255 * Rnd + 1)
ccc = Int(255 * Rnd - 5)
cccc = Int(260 * Rnd + 5)
title.ForeColor = RGB(aa, bb, cc)
virus.ForeColor = RGB(aaa + 5, bbb + 5, ccc + 5)
End Sub

Private Sub win_dsk_Click() '小窗按钮
 Dim Hwnd As Long
 Shell "explorer", vbNormalFocus  '极域目录C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe
 Hwnd = GetForegroundWindow
 SetParent Hwnd, Me.Hwnd
 MoveWindow Hwnd, 0, 0, 200, 200, 1 ' 0，0 是位置，200，200是宽度和高度
End Sub
