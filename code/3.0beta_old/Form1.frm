VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "老师控制解除器3.0-by b站、帥の智叟L"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6675
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton win_jiyu 
      Caption         =   "小窗模式の极域"
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关于作者"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "启动控制器"
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重启控制器"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解除控制"
      Height          =   1335
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "点按钮解除控制。软件使用vb制作，报毒正常。"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
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
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '窗口置顶

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long  '置顶声明
Private Declare Function GetForegroundWindow Lib "user32" () As Long  '置顶声明
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long  '置顶声明

Private Sub Command1_Click()
Shell "cmd /c taskkill /f /im Studentmain.exe", 0
MsgBox "解除成功!--by帥の智叟L", okonly, "提示"
End Sub

Private Sub Command2_Click()
Shell "C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe", 1
MsgBox "重启成功!--by帥の智叟L", okonly, "提示"
End Sub

Private Sub Command3_Click()
Shell "C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe", 1
MsgBox "启动成功!--by帥の智叟L", okonly, "提示"
End Sub

Private Sub Command4_Click()
MsgBox "作者:b站 帥の智叟L.各位给一个关注就行!!!!", okonly, "关于作者"
End Sub

Private Sub Form_click()
MsgBox "啊啊啊痒", okonly, "干蛤啊你"
End Sub

Private Sub Form_Load()
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1
End Sub

Private Sub Label1_Click()
MsgBox "作者是不是很帅", okonly, "大帅比"
End Sub

Private Sub Label2_Click()
MsgBox "作者是不是很帅", okonly, "大帅比"
End Sub

Private Sub Timer1_Timer()
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1
End Sub

Private Sub win_jiyu_Click()
 Dim Hwnd As Long
 Shell "C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe", vbNormalFocus  '极域目录C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe
 Hwnd = GetForegroundWindow
 SetParent Hwnd, Me.Hwnd
 MoveWindow Hwnd, 0, 0, 200, 200, 1 ' 0，0 是位置，200，200是宽度和高度

End Sub
