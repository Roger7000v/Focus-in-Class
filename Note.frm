VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Note"
   ClientHeight    =   4540
   ClientLeft      =   110
   ClientTop       =   750
   ClientWidth     =   7900
   Icon            =   "Note.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4540
   ScaleWidth      =   7900
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4450
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7690
      _ExtentX        =   13564
      _ExtentY        =   7849
      _Version        =   393217
      TextRTF         =   $"Note.frx":25CA
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件"
      Begin VB.Menu mnuNew 
         Caption         =   "新建"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "打开"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑"
      Begin VB.Menu mnuCopy 
         Caption         =   "复制"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴"
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "全选"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "搜索"
      Begin VB.Menu mnuFind 
         Caption         =   "查找"
      End
      Begin VB.Menu mnuFindOn 
         Caption         =   "查找下一个"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助"
      Begin VB.Menu mnuUsage 
         Caption         =   "使用说明"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'声明查找变量
Dim sFind As String
'声明文件类型
Dim FileType, FiType As String

'初始化程序
Private Sub Form_Load()
'设置程序启动时的大小
Me.Height = 6000
Me.Width = 9000
End Sub

'设置编辑框的位置和大小
Private Sub Form_Resize()
On Error Resume Next '出错处理
RichTextBox1.Top = 20
RichTextBox1.Left = 20
RichTextBox1.Height = ScaleHeight - 40
RichTextBox1.Width = ScaleWidth - 40
End Sub

'新建文件
Private Sub mnuNew_Click()
RichTextBox1.Text = "" '清空文本框
FileName = "未命名"
Me.Caption = FileName
End Sub


'打开文件
Private Sub mnuOpen_Click()
CommonDialog1.Filter = "文本文档(*.txt)|*.txt|RTF文档(*.rtf)|*.rtf|所有文件(*.*)|*.*"
CommonDialog1.ShowOpen
RichTextBox1.Text = "" '清空文本框
FileName = CommonDialog1.FileName
RichTextBox1.LoadFile FileName
Me.Caption = "笔记：" & FileName
End Sub

'保存文件
Private Sub mnuSave_Click()
CommonDialog1.Filter = "文本文档(*.txt)|*.txt|RTF文档(*.rtf)|*.rtf|所有文件(*.*)|*.*"
CommonDialog1.ShowSave
FileType = CommonDialog1.FileTitle
FiType = LCase(Right(FileType, 3))
FileName = CommonDialog1.FileName
Select Case FiType
Case "txt"
RichTextBox1.SaveFile FileName, rtfText
Case "rtf"
RichTextBox1.SaveFile FileName, rtfRTF
Case "*.*"
RichTextBox1.SaveFile FileName
End Select
Me.Caption = "笔记本：" & FileName
End Sub

'退出
Private Sub mnuExit_Click()
Unload Me
End Sub

'复制
Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End Sub

'剪切
Private Sub mnuCut_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End Sub

'全选
Private Sub mnuSelectAll_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub

'粘贴
Private Sub mnuPaste_Click()
RichTextBox1.SelText = Clipboard.GetText
End Sub

'查找
Private Sub mnuFind_Click()
sFind = InputBox("请输入要查找的字、词：", "查找内容", sFind)
RichTextBox1.Find sFind
End Sub

'继续查找
Private Sub mnuFindOn_Click()
RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength + 1
RichTextBox1.Find sFind, , Len(RichTextBox1)
End Sub

'使用说明
Private Sub mnuReadme_Click()
On Error GoTo handler
RichTextBox1.LoadFile "Readme.txt", rtfText '请写好Readme.txt文件并存入程序所在文件夹中
Me.Caption = "笔记本：" & "使用说明"
Exit Sub
handler:
MsgBox "使用说明文档可能已经被移除，请与作者联系。", vbOKOnly, " 错误信息"
End Sub

'关于
Private Sub mnuAbout_Click()
MsgBox "笔记本 Ver1.0 版权所有(C) 2017 Rooooger", vbOKOnly, "关于"
End Sub

'设置弹出式菜单（即在编辑框中单击鼠标右键时弹出的动态菜单）
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuEdit, vbPopupMenuLeftAlign
Else
Exit Sub
End If
End Sub

'防止在切换输入法时字体自变
Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
RichTextBox1.SelFontName = CommonDialog1.FontName
End If
End Sub
