VERSION 5.00
Begin VB.Form FrmSystemAdmin 
   Caption         =   "PDM-System Admin 系统管理子系统"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmManManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   StartUpPosition =   2  '屏幕中心
   Begin VB.Image Image4 
      Height          =   300
      Left            =   6660
      Picture         =   "FrmManManage.frx":08CA
      Top             =   4260
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2040
      Picture         =   "FrmManManage.frx":0CE6
      Top             =   4260
      Width           =   300
   End
   Begin VB.Label LblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit 退出系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7140
      MouseIcon       =   "FrmManManage.frx":1102
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6420
      Width           =   2355
   End
   Begin VB.Shape Shape3 
      Height          =   555
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   2475
   End
   Begin VB.Label LblDatabaseManage 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmManManage.frx":140C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7500
      MouseIcon       =   "FrmManManage.frx":1436
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4260
      Width           =   3135
   End
   Begin VB.Label LblDatabaseSet 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmManManage.frx":1740
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7500
      MouseIcon       =   "FrmManManage.frx":1765
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   2955
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   6660
      Picture         =   "FrmManManage.frx":1A6F
      Top             =   3060
      Width           =   300
   End
   Begin VB.Label LblSysIni 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmManManage.frx":1E8B
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2700
      MouseIcon       =   "FrmManManage.frx":1EB0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2040
      Picture         =   "FrmManManage.frx":21BA
      Top             =   3060
      Width           =   300
   End
   Begin VB.Label LblMan 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"FrmManManage.frx":25D6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   1860
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label LblUserManage 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmManManage.frx":25FD
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      MouseIcon       =   "FrmManManage.frx":2620
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4260
      Width           =   2715
   End
End
Attribute VB_Name = "FrmSystemAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub LblDatabaseManage_Click()
  '数据库表管理
  FrmDatabaseManage.Show
End Sub

Private Sub LblDatabaseSet_Click()
  '数据库配置检查
  FrmDatabaseSet.Show
End Sub

Private Sub LblExit_Click()
  '卸载窗体
  Unload Me
FrmEngineeringSys.Show 0
End Sub

Private Sub LblSysIni_Click()
  '显示系统初始化设置窗体
  FrmSysIni.Show
End Sub

Private Sub lblUserManage_Click()
  '显示系统用户管理窗体
  FrmUserManage.Show
End Sub
