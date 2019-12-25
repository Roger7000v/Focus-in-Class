VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Class"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12850
   Icon            =   "Class.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   12850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1810
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5880
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   4090
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Class.frx":238A
      Top             =   3840
      Width           =   6970
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "笔记"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1810
   End
   Begin VB.Label Label2 
      Caption         =   "时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   7800
      TabIndex        =   3
      Top             =   480
      Width           =   2050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 上课中"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2530
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   6250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim time As Long

Private Sub Command1_Click()
Form2.Show
End Sub

Private Function Fun_DisplayTaskBar(ByVal bShow As Boolean) As Integer
    Dim lTaskBarHWND     As Long
    Dim lRet     As Long
    Dim lFlags     As Long
    On Error GoTo vbErrorHandler
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    lTaskBarHWND = FindWindow("Shell_TrayWnd", "")
    lRet = ShowWindow(lTaskBarHWND, lFlags)
    If lRet < 0 Then
          Exit Function
    End If
vbErrorHandler:
  End Function
Private Sub Form_Load()
    Fun_DisplayTaskBar False
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Width
    
     ' 安装钩子
    lHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallKeyHookProc, App.hInstance, 0)
    
    time = 2400
End Sub

Private Sub Timer1_timer()
time = time - 1
If time = 0 Then
Fun_DisplayTaskBar True

' 卸载钩子
    UnhookWindowsHookEx lHook

Unload Me
End If
End Sub

Private Sub Timer2_Timer()
Label2.Caption = Format(Now, "HH:mm")
End Sub

Private Sub Command3_Click()
Shell "C:\WINDOWS\system32\calc.exe"
End Sub
