VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recovery 恢复选项"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "恢复选项"
      Height          =   3795
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      Begin VB.CommandButton CmdFile 
         Caption         =   "浏览"
         Height          =   375
         Left            =   6180
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtDataFile 
         Height          =   315
         Left            =   2820
         TabIndex        =   1
         Top             =   240
         Width           =   3195
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "确  定"
         Height          =   435
         Left            =   3660
         TabIndex        =   11
         Top             =   3240
         Width           =   1395
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "退  出"
         Height          =   435
         Left            =   5400
         TabIndex        =   10
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "恢复后的数据库文件存放目录"
         Height          =   1635
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   6555
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3180
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "默认安装目录"
            Height          =   375
            Index           =   0
            Left            =   300
            TabIndex        =   9
            Top             =   420
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "选定目录"
            Height          =   375
            Index           =   1
            Left            =   300
            TabIndex        =   8
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton CmdFind 
            Caption         =   "浏览"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5760
            TabIndex        =   7
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox TxtDataPath 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   1020
            Width           =   3915
         End
      End
      Begin VB.CheckBox ChkOver 
         Caption         =   "是否覆盖已经存在的数据"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   3180
         Width           =   2595
      End
      Begin VB.TextBox TxtDataBase 
         Height          =   375
         Left            =   2820
         TabIndex        =   2
         Top             =   780
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "请选择恢复数据库的文件:"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "请输入要恢复的数据库名称:(若保留数据库名称为空,则按备份的数据库命名)"
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdFile_Click()
On Error GoTo ErrHandle:

    CommonDialog1.CancelError = True    '判断是否取消操作
    CommonDialog1.DialogTitle = "打开要恢复的数据库文件"
    CommonDialog1.Filter = "数据库文件(*.MsDat)|*.MsDat|所有文件(*.*)|*.*"
    CommonDialog1.ShowOpen
    TxtDataFile.Text = CommonDialog1.filename
    
ErrHandle:
    Exit Sub
    
End Sub

Private Sub CmdFind_Click()
    Me.Hide
    FrmPath.Show 1
End Sub

Private Sub CmdYes_Click()
Dim flag As Boolean
Dim OptionFlag As Boolean
Dim filename As String
    MousePointer = vbHourglass
    
    If Len(Trim(TxtDataFile.Text)) = 0 Then MsgBox "请输入存放数据库文件的路径", vbInformation, "提示": GoTo ExitPoint
    
    If OptPath(1).Value = True And Len(Trim(TxtDataPath.Text)) = 0 Then MsgBox "请输入恢复的数据库文件存放的路径!", vbInformation, "提示": GoTo ExitPoint
       If Dir(TxtDataPath.Text, 16) = "" Then
           MsgBox "系统没有该文件夹,系统自动创建该文件夹!", vbInformation, "提示"
           MkDir (TxtDataPath.Text)
       End If
    If ChkOver.Value = 1 Then OptionFlag = True        'ChkOver是checkbox控件:是否覆盖已经存在的数据
    
    filename = (Trim(TxtDataFile.Text))
    'FrmServerBkup.Tag = "0"
    
    Call CheckServer(filename, flag)
      If flag = False Then MsgBox "数据库没有恢复!", vbExclamation, "错误": GoTo ExitPoint
    
    Call RestoreDatabase(filename, Trim(TxtDataBase), TxtDataPath.Text, 1, OptionFlag)
    Call HandleFile(optflag, flag)

ExitPoint:
    MousePointer = vbDefault
    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation, "错误"
    GoTo ExitPoint
End Sub


Private Sub OptPath_Click(Index As Integer)        'OptPath是一个单选项,恢复后的数据库文件存放目录默认或指定
    Select Case Index
        Case 0
            TxtDataPath.Enabled = False
            CmdFind.Enabled = False
        Case 1
            TxtDataPath.Enabled = True
            CmdFind.Enabled = True
    End Select
End Sub
