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
      Caption         =   "�ļ�"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "��"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭"
      Begin VB.Menu mnuCopy 
         Caption         =   "����"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��"
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "ȫѡ"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "����"
      Begin VB.Menu mnuFind 
         Caption         =   "����"
      End
      Begin VB.Menu mnuFindOn 
         Caption         =   "������һ��"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����"
      Begin VB.Menu mnuUsage 
         Caption         =   "ʹ��˵��"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�������ұ���
Dim sFind As String
'�����ļ�����
Dim FileType, FiType As String

'��ʼ������
Private Sub Form_Load()
'���ó�������ʱ�Ĵ�С
Me.Height = 6000
Me.Width = 9000
End Sub

'���ñ༭���λ�úʹ�С
Private Sub Form_Resize()
On Error Resume Next '������
RichTextBox1.Top = 20
RichTextBox1.Left = 20
RichTextBox1.Height = ScaleHeight - 40
RichTextBox1.Width = ScaleWidth - 40
End Sub

'�½��ļ�
Private Sub mnuNew_Click()
RichTextBox1.Text = "" '����ı���
FileName = "δ����"
Me.Caption = FileName
End Sub


'���ļ�
Private Sub mnuOpen_Click()
CommonDialog1.Filter = "�ı��ĵ�(*.txt)|*.txt|RTF�ĵ�(*.rtf)|*.rtf|�����ļ�(*.*)|*.*"
CommonDialog1.ShowOpen
RichTextBox1.Text = "" '����ı���
FileName = CommonDialog1.FileName
RichTextBox1.LoadFile FileName
Me.Caption = "�ʼǣ�" & FileName
End Sub

'�����ļ�
Private Sub mnuSave_Click()
CommonDialog1.Filter = "�ı��ĵ�(*.txt)|*.txt|RTF�ĵ�(*.rtf)|*.rtf|�����ļ�(*.*)|*.*"
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
Me.Caption = "�ʼǱ���" & FileName
End Sub

'�˳�
Private Sub mnuExit_Click()
Unload Me
End Sub

'����
Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End Sub

'����
Private Sub mnuCut_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End Sub

'ȫѡ
Private Sub mnuSelectAll_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub

'ճ��
Private Sub mnuPaste_Click()
RichTextBox1.SelText = Clipboard.GetText
End Sub

'����
Private Sub mnuFind_Click()
sFind = InputBox("������Ҫ���ҵ��֡��ʣ�", "��������", sFind)
RichTextBox1.Find sFind
End Sub

'��������
Private Sub mnuFindOn_Click()
RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength + 1
RichTextBox1.Find sFind, , Len(RichTextBox1)
End Sub

'ʹ��˵��
Private Sub mnuReadme_Click()
On Error GoTo handler
RichTextBox1.LoadFile "Readme.txt", rtfText '��д��Readme.txt�ļ���������������ļ�����
Me.Caption = "�ʼǱ���" & "ʹ��˵��"
Exit Sub
handler:
MsgBox "ʹ��˵���ĵ������Ѿ����Ƴ�������������ϵ��", vbOKOnly, " ������Ϣ"
End Sub

'����
Private Sub mnuAbout_Click()
MsgBox "�ʼǱ� Ver1.0 ��Ȩ����(C) 2017 Rooooger", vbOKOnly, "����"
End Sub

'���õ���ʽ�˵������ڱ༭���е�������Ҽ�ʱ�����Ķ�̬�˵���
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuEdit, vbPopupMenuLeftAlign
Else
Exit Sub
End If
End Sub

'��ֹ���л����뷨ʱ�����Ա�
Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
RichTextBox1.SelFontName = CommonDialog1.FontName
End If
End Sub
