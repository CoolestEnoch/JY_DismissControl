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
      Caption         =   "С������"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton loc_help 
      Caption         =   "��֪����ô��StudentMain.exe���ļ�λ��?"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Frame location_guide 
      Caption         =   "��������������������ť����Ч�����ڴ����뼫��StudentMain.exe��Ŀ¼..."
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
      Caption         =   "������־"
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
      Caption         =   "��������"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton open 
      Caption         =   "����������"
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton rest 
      Caption         =   "����������"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton no_ctr 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�㰴ť������ơ����ʹ��vb����������������"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label title 
      Caption         =   "       ������ӽ��ҿ��ƽ������"
      BeginProperty Font 
         Name            =   "����"
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
Dim ver As String '�汾�ű�������,�汾���޸���Form1_Load���޸�
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '�����ö�API
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long  'С������
Private Declare Function GetForegroundWindow Lib "user32" () As Long  'С������
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long  'С������

Private Sub loc_help_Click()
MsgBox "�Ҽ������ϵļ�����ӽ��ҵĿ�ݷ�ʽ�������ԣ��ҵ���Ŀ�ꡱ������ı����������Ƶ�����򼴣��ٵ��������������������������ť���ɡ�", okobly, "��β鿴�ļ�λ��"
End Sub

Private Sub no_ctr_Click()
Shell "cmd /c taskkill /f /im Studentmain.exe", 0
MsgBox "����ɹ�!--by��������L", okonly, "��ʾ"
End Sub

Private Sub rest_Click()
Dim stum As New FileSystemObject
lloc = jiyu_loc.Text
llong = Len(lloc)
If lloc = "" Then
 If stum.FileExists("C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe") Then
  Shell "cmd /c taskkill /f /im Studentmain.exe", 0
  Shell "C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe", 1
  MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
 Else
  MsgBox "��Ԥ��Ŀ¼û�ҵ�������������ȷ������ѧУ��Ϣ���ҵ��Ի�˵����Ѱ�װ������ӽ���ѧ�����棡", okonly, "��ʾ"
 End If
Else
 If llong < 5 Then
  MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
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
      MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
     Else
      MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
     End If
   ElseIf nname <> ".exe" Then
      MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
     End If
   End If
 End If
End Sub

Private Sub open_Click()
Dim stum As New FileSystemObject
lloc = jiyu_loc.Text
llong = Len(lloc)
If lloc = "" Then
 If stum.FileExists("C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe") Then
  Shell "C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe", 1
  MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
 Else
  MsgBox "��Ԥ��Ŀ¼û�ҵ�������������ȷ������ѧУ��Ϣ���ҵ��Ի�˵����Ѱ�װ������ӽ���ѧ�����棡", okonly, "��ʾ"
 End If
Else
 If llong < 5 Then
  MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
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
      MsgBox "�����ɹ�!--by��������L", okonly, "��ʾ"
     Else
      MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
     End If
   ElseIf nname <> ".exe" Then
      MsgBox "Ŀ¼��������!���顣", okonly, "��ʾ"
     End If
   End If
 End If
End Sub

Private Sub Command4_Click()
MsgBox "����:bվ ��������L.��λ��һ����ע����!!!!", okonly, "��������"
End Sub

Private Sub Form_click()
MsgBox "��������", okonly, "�ɸ���"
End Sub

Private Sub Form_Load()
ver = "2.5" '�汾��
SetWindowPos Me.Hwnd, -1, 0, 0, 0, 0, 2 Or 1 '�����ö�
Form1.Caption = "��ʦ���ƽ����" & ver & "-by bվ����������L     "
End Sub

Private Sub update_log_Click()
MsgBox "1.0����һ�������������" + vbCrLf + "2.0������£������Ż������ӡ�������������������ť��" + vbCrLf + "2.1�����Ӵ���ǿ���ö��Ĺ��ܣ�ʹ��Ļ�㲥ʱҲ����ʾ�������" + vbCrLf + "2.2�������޸�ǿ��������Ļ�㲥ʱʧЧ��bug��" + vbCrLf + "2.3���̶��˴��ڵĴ�С��ͬʱ�����˴��ڵ���󻯣���ֹ�󴥵��½��沼�ֱ仯��˳���Ż������������(Duang������Ч���ܹ�������ͺ�..�ȿ�)" + vbCrLf + "2.4�������Զ��弫���Դ�ļ�Ŀ¼�����������ж�Ŀ¼�﷨�Ƿ��д�����ǿ�˼����ԡ�" + vbCrLf + "2.5���޸���һЩ�޹��������⣬�������ȶ��ԣ��޸����Զ��弫��Ŀ¼���Ҳ����ļ������˵�bug���������������С���������˸�����־��һ������֡��޸��˱����ɫ���µļ��������ˡ�" + vbCrLf + "3.0�����ڿ�������Ը��Ҫ��Щʲô��bug�ͺá��ƻ�����Ѽ����ȫ����Ļ�㲥ǿ�б��С��ģʽ(�G��˭���ĸ��ӣ���������.....��)��", okonly, "������־"
End Sub

Private Sub virus_Click()
MsgBox "�����ǲ��Ǻ�˧", okonly, "��˧��"
End Sub

Private Sub title_Click()
MsgBox "�����ǲ��Ǻ�˧", okonly, "��˧��"
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
 virus.Caption = Mid(virus, 2, n - 1) + Mid(virus, 1, 1) '�����
'��ʼ��ɫ
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

Private Sub win_dsk_Click() 'С����ť
 Dim Hwnd As Long
 Shell "explorer", vbNormalFocus  '����Ŀ¼C:\Program Files (x86)\Mythware\������ӽ������ v4.0 2015 ������\StudentMain.exe
 Hwnd = GetForegroundWindow
 SetParent Hwnd, Me.Hwnd
 MoveWindow Hwnd, 0, 0, 200, 200, 1 ' 0��0 ��λ�ã�200��200�ǿ�Ⱥ͸߶�
End Sub
