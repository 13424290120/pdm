VERSION 5.00
Begin VB.Form FrmPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ����Ҫ������ݿ��ļ� ��·��"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4980
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdSetNew 
      Caption         =   "��  ��"
      Height          =   435
      Left            =   3420
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdNo 
      Caption         =   "ȡ  ��"
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdYes 
      Caption         =   "ȷ  ��"
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   4635
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4635
   End
End
Attribute VB_Name = "FrmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdNo_Click()
    Unload Me
    FrmOption.Show 1
End Sub

Private Sub CmdSetNew_Click()   '�½��ļ���
Dim Message As String, Title As String, Default As String, MyValue As String
Dim StrPath   As String

On Error GoTo ErrHandle:

HandlePoint:
    Message = "�������½����ļ�������:"   ' ������ʾ��Ϣ��
    Title = "�½����ļ�������"   ' ���ñ��⡣
    Default = "�½��ļ���"   ' ����ȱʡֵ��
    ' ��ʾ��Ϣ�����⼰ȱʡֵ��
    MyValue = InputBox(Message, Title, Default)
    If Len(Trim(MyValue)) = 0 Then GoTo ExitPoint
    
    StrPath = Dir1.Path & "\" & Trim(MyValue)  'ѡ����·��������������ļ�����
    If Dir(StrPath, 16) = "" Then               '16 ָ���������ļ�����·�����ļ���
        MkDir (StrPath)                     '����һ���µ�Ŀ¼���ļ��С�
    Else
        MsgBox "���ļ����Ѵ���,��������������!", vbInformation, "��ʾ"
        GoTo HandlePoint:
    End If
    Dir1.Refresh
ExitPoint:
    Exit Sub
    
ErrHandle:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ExitPoint
End Sub

Private Sub CmdYes_Click()
    FrmOption.TxtDataPath.Text = Dir1.Path
    Unload Me
    FrmOption.Show 1
End Sub

Private Sub Drive1_Change()

On Error GoTo ErrHandle:
     Dir1.Path = Drive1.Drive
     
ExitPoint:
    Exit Sub
    
ErrHandle:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ExitPoint
End Sub
