VERSION 5.00
Begin VB.Form FrmPurchasingSys 
   Caption         =   "PDM-Purchasing"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   10065
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   9975
      Begin VB.Image Image1 
         Height          =   300
         Left            =   360
         Picture         =   "FrmPurchasingSys.frx":0000
         Top             =   2430
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Part Library 标准件库管理"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   660
         Left            =   690
         MouseIcon       =   "FrmPurchasingSys.frx":041C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2400
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "New Part OverView 新物料使用查询 <F7>"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   660
         Left            =   5700
         MouseIcon       =   "FrmPurchasingSys.frx":056E
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1530
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Concession Number 工程让步编号管理"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   690
         Left            =   660
         MouseIcon       =   "FrmPurchasingSys.frx":06C0
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   720
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SER Number 单件物料承认编号管理"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   660
         Left            =   660
         MouseIcon       =   "FrmPurchasingSys.frx":0812
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1590
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CP/CN Number 工程变更编号管理"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   690
         Left            =   5700
         MouseIcon       =   "FrmPurchasingSys.frx":0964
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   690
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFinsGd 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   660
         Left            =   5730
         MouseIcon       =   "FrmPurchasingSys.frx":0AB6
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1560
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image18 
         Height          =   300
         Left            =   330
         Picture         =   "FrmPurchasingSys.frx":0C08
         Top             =   720
         Width           =   300
      End
      Begin VB.Image Image17 
         Height          =   300
         Left            =   330
         Picture         =   "FrmPurchasingSys.frx":1024
         Top             =   1620
         Width           =   300
      End
      Begin VB.Image Image7 
         Height          =   300
         Left            =   5340
         Picture         =   "FrmPurchasingSys.frx":1440
         Top             =   690
         Width           =   300
      End
      Begin VB.Image Image8 
         Height          =   300
         Left            =   5340
         Picture         =   "FrmPurchasingSys.frx":185C
         Top             =   1650
         Width           =   300
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   9975
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit 退出系统 <ESC>"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   660
         Left            =   5700
         MouseIcon       =   "FrmPurchasingSys.frx":1C78
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   480
         Width           =   2460
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image15 
         Height          =   300
         Left            =   5400
         Picture         =   "FrmPurchasingSys.frx":1DCA
         Top             =   510
         Width           =   300
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Purchasing Database V2.0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   675
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "FrmPurchasingSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LblCNCSN_Click()
'显示CONCESSION管理窗体
FrmCNCSN.Show
End Sub

Private Sub LblCPCN_Click()
'显示CP/CN管理窗体
FrmCPCN.Show
End Sub

Private Sub LblNPO_Click()
FrmBOMNPO.Show
End Sub

Private Sub LblSER_Click()
'显示SER管理窗体
FrmSER.Show
End Sub

Private Sub Label1_Click()
    Set FromForm2 = FrmPurchasingSys
    FrmPurchasingSys.Hide
    FrmStdPrtLibStructr.Show 0
End Sub

Private Sub Label3_Click()
FrmEngineeringSys.Hide
QuerytableName = "CPCN"                                  '##########告诉通用查询窗口是对哪个表进行操作

'@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
If SystemAdmin <> "Y" Then
'    MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
'    FrmQuery.CmdModify.Enabled = False
'    FrmQuery.CmdDel.Enabled = False

    FrmQuery.DataGrid1.AllowDelete = False
    FrmQuery.DataGrid1.AllowAddNew = False
    FrmQuery.DataGrid1.AllowUpdate = False
End If
'@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
Set FromForm2 = FrmPurchasingSys
FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
FrmQuery.Caption = "PDM-CPCN Number Admin 工程管理子系统"
End Sub

Private Sub Label4_Click()
    FrmPurchasingSys.Hide
    QuerytableName = "SER"                                  '##########告诉通用查询窗口是对哪个表进行操作
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
    '    MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
    '    FrmQuery.CmdModify.Enabled = False
    '    FrmQuery.CmdDel.Enabled = False
    
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    Set FromForm = FrmPurchasingSys
    FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
    FrmQuery.Caption = "PDM-SER Number Admin 工程管理子系统"
End Sub

Private Sub Label5_Click()
    FrmPurchasingSys.Hide
    QuerytableName = "CNCSN"                                  '##########告诉通用查询窗口是对哪个表进行操作
    
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
    '    MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
    '    FrmQuery.CmdModify.Enabled = False
    '    FrmQuery.CmdDel.Enabled = False
    
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    Set FromForm = FrmPurchasingSys
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
    FrmQuery.Caption = "PDM-CONCESSION Number Admin 工程管理子系统"
End Sub

Private Sub Label7_Click()
    Set FromForm2 = FrmPurchasingSys
    FrmPurchasingSys.Hide
    FrmBOMNPO.Show 0
End Sub

Private Sub LblExit_Click()
 '卸载窗体
Unload Me
End Sub
