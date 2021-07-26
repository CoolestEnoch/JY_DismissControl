VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private m_Hwnd As Long

Private Sub Command1_Click()

    Dim dblPid As Long

    Call LockWindowUpdate(GetDesktopWindow)

    dblPid = Shell("c:\windows\explorer.exe", vbNormalFocus)
'极域目录C:\Program Files (x86)\Mythware\极域电子教室软件 v4.0 2015 豪华版\StudentMain.exe
'记事本目录c:\windows\notepad.exe
'CMDc:\windows\system32\cmd.exe
'资源管理器c:\windows\explorer.exe
    m_Hwnd = InstanceToWnd(dblPid) '根据进程PID找窗口句柄

    SetParent m_Hwnd, Me.hwnd

    Putfocus m_Hwnd                 '记事本设置焦点

    Call LockWindowUpdate(0)

End Sub

Function InstanceToWnd(ByVal target_pid As Long) As Long

    Dim i As Long, lHwnd As Long, lPid As Long, lThreadId As Long

    lHwnd = FindWindow(ByVal 0&, ByVal 0&)   '查找第一个窗口

    Do While lHwnd <> 0

        i = i + 1

        If i Mod 20 = 0 Then DoEvents

        '判断窗口是否没父窗口
        If GetParent(lHwnd) = 0 Then

            '获取该窗口的线程ID
            lThreadId = GetWindowThreadProcessId(lHwnd, lPid)

            If lPid = target_pid Then '找到PID所在窗口句柄

                InstanceToWnd = lHwnd
                Exit Do

            End If

        End If

        '继续查找下一个兄弟窗口
        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)

        Debug.Print Hex$(lHwnd)

    Loop

End Function

Private Sub Form_Unload(Cancel As Integer)

    Call DestroyWindow(m_Hwnd)
    'TerminateProcess GetCurrentProcess, 0    '野蛮了些
    Set Form1 = Nothing

End Sub

