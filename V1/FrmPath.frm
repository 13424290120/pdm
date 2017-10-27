VERSION 5.00
Begin VB.Form FrmPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择你要存放数据库文件 的路径"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdSetNew 
      Caption         =   "新  建"
      Height          =   435
      Left            =   3420
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdNo 
      Caption         =   "取  消"
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdYes 
      Caption         =   "确  定"
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

Private Sub CmdSetNew_Click()   '新建文件夹
Dim Message As String, Title As String, Default As String, MyValue As String
Dim StrPath   As String

On Error GoTo ErrHandle:

HandlePoint:
    Message = "请输入新建的文件夹名称:"   ' 设置提示信息。
    Title = "新建的文件夹名称"   ' 设置标题。
    Default = "新建文件夹"   ' 设置缺省值。
    ' 显示信息、标题及缺省值。
    MyValue = InputBox(Message, Title, Default)
    If Len(Trim(MyValue)) = 0 Then GoTo ExitPoint
    
    StrPath = Dir1.Path & "\" & Trim(MyValue)  '选定的路径加上输入的新文件夹名
    If Dir(StrPath, 16) = "" Then               '16 指定无属性文件及其路径和文件夹
        MkDir (StrPath)                     '创建一个新的目录或文件夹。
    Else
        MsgBox "此文件夹已存在,请输入重新命名!", vbInformation, "提示"
        GoTo HandlePoint:
    End If
    Dir1.Refresh
ExitPoint:
    Exit Sub
    
ErrHandle:
    MsgBox Err.Description, vbExclamation, "错误"
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
    MsgBox Err.Description, vbExclamation, "出错"
    GoTo ExitPoint
End Sub
