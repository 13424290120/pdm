VERSION 5.00
Begin VB.Form FrmSaveOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "请输入登录信息"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6705
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   $"FrmSaveOption.frx":0000
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton CmdNo 
         Caption         =   "取  消"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   2040
         Width           =   1155
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "确  定"
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox TxtPwd 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2040
         Width           =   2835
      End
      Begin VB.TextBox TxtName 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "Administrator"
         Top             =   1620
         Width           =   2835
      End
      Begin VB.TextBox TxtGroup 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Text            =   "工作组 或者 域名"
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   540
         TabIndex        =   9
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "密码:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "用户名:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "工作组:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmSaveOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdNo_Click()
    optflag = False
    Unload Me
End Sub

Private Sub CmdYes_Click()
    Me.Hide
    GroupName = Trim(TxtGroup.Text)
    UserName = Trim(TxtName.Text)
    Pwd = Trim(TxtPwd.Text)
    optflag = True
    Unload Me
End Sub

Private Sub TxtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then                            '判断是否为回车键
        SendKeys "{TAB}"                            '转换为Tab键
    End If
End Sub
