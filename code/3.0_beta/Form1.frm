VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ʦ���ƽ����3.0-by bվ����������L"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6675
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton win_jiyu 
      Caption         =   "С��ģʽ�μ���"
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
      Caption         =   "��������"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����������"
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����������"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Height          =   1335
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "�㰴ť������ơ����ʹ��vb����������������"
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
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '�����ö�

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long  '�ö�����
Private Declare Function GetForegroundWindow Lib "user32" () As Long  '�ö�����
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long  '�ö�����

Private Sub Command1_Click()
Shell "cmd /c taskkill /f /im Studentmain.exe", 0
MsgBox "����ɹ�!--by��������L", okonly, "��ʾ"
End Sub

Private Sub Command2_Click()
Shell "C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe", 1
MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
End Sub

Private Sub Command3_Click()
Shell "C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe", 1
MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
End Sub

Private Sub Command4_Click()
MsgBox "����:bվ ��������L.��λ��һ����ע����!!!!", okonly, "��������"
End Sub

Private Sub Form_click()
MsgBox "��������", okonly, "�ɸ���"
End Sub

Private Sub Form_Load()
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1
End Sub

Private Sub Label1_Click()
MsgBox "�����ǲ��Ǻ�˧", okonly, "��˧��"
End Sub

Private Sub Label2_Click()
MsgBox "�����ǲ��Ǻ�˧", okonly, "��˧��"
End Sub

Private Sub Timer1_Timer()
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1
End Sub

Private Sub win_jiyu_Click()
 Dim Hwnd As Long
 Shell "C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe", vbNormalFocus  '����Ŀ¼C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe
 Hwnd = GetForegroundWindow
 SetParent Hwnd, Me.Hwnd
 MoveWindow Hwnd, 0, 0, 200, 200, 1 ' 0��0 ��λ�ã�200��200�ǿ�Ⱥ͸߶�

End Sub
