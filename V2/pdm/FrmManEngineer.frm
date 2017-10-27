VERSION 5.00
Begin VB.Form FrmEngineeringSys 
   Appearance      =   0  'Flat
   BackColor       =   &H0053442A&
   BorderStyle     =   0  'None
   Caption         =   "Engineer Database"
   ClientHeight    =   7092
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10596
   ControlBox      =   0   'False
   FillColor       =   &H80000000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmManEngineer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7092
   ScaleWidth      =   10596
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit System"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   11820
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmManEngineer.frx":08CA
      TabIndex        =   14
      Top             =   10410
      Width           =   3765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V3.5.9.13-RC2"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7830
      TabIndex        =   15
      Top             =   450
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   8076
      MouseIcon       =   "FrmManEngineer.frx":3FA2
      MousePointer    =   99  'Custom
      Picture         =   "FrmManEngineer.frx":40F4
      Top             =   5856
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   9396
      MouseIcon       =   "FrmManEngineer.frx":67FA
      MousePointer    =   99  'Custom
      Picture         =   "FrmManEngineer.frx":694C
      Top             =   5856
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   684
      Index           =   1
      Left            =   9336
      Shape           =   1  'Square
      Top             =   5820
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time Writing Record"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   705
      Index           =   3
      Left            =   240
      MouseIcon       =   "FrmManEngineer.frx":7698
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5970
      Width           =   3195
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   825
      Index           =   3
      Left            =   240
      Top             =   5820
      Width           =   3195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Approval"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   705
      Index           =   2
      Left            =   7140
      MouseIcon       =   "FrmManEngineer.frx":77EA
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4920
      Width           =   3195
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   825
      Index           =   2
      Left            =   7140
      Top             =   4770
      Width           =   3195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Part Overview"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   705
      Index           =   1
      Left            =   3660
      MouseIcon       =   "FrmManEngineer.frx":793C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4920
      Width           =   3195
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   825
      Index           =   1
      Left            =   3690
      Top             =   4770
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   8
      Left            =   7140
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Part Lib."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   8
      Left            =   7140
      MouseIcon       =   "FrmManEngineer.frx":7A8E
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3510
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   7
      Left            =   3690
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   7
      Left            =   3780
      MouseIcon       =   "FrmManEngineer.frx":7BE0
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3510
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   6
      Left            =   240
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Goods Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   6
      Left            =   240
      MouseIcon       =   "FrmManEngineer.frx":7D32
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3510
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   5
      Left            =   240
      Top             =   2280
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   5
      Left            =   300
      MouseIcon       =   "FrmManEngineer.frx":7E84
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2430
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   4
      Left            =   7140
      Top             =   2280
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SER Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   4
      Left            =   7140
      MouseIcon       =   "FrmManEngineer.frx":7FD6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2430
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   3
      Left            =   3690
      Top             =   2280
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CP/CN Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   3
      Left            =   3660
      MouseIcon       =   "FrmManEngineer.frx":8128
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2430
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   2
      Left            =   7140
      Top             =   1200
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Concession Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   2
      Left            =   7140
      MouseIcon       =   "FrmManEngineer.frx":827A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1380
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   1
      Left            =   3690
      Top             =   1200
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   1
      Left            =   3720
      MouseIcon       =   "FrmManEngineer.frx":83CC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1380
      Width           =   3195
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   825
      Index           =   0
      Left            =   240
      Top             =   4770
      Width           =   3195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   705
      Index           =   0
      Left            =   240
      MouseIcon       =   "FrmManEngineer.frx":851E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4920
      Width           =   3195
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   684
      Index           =   0
      Left            =   8040
      Shape           =   1  'Square
      Top             =   5820
      Width           =   684
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ/RFS Admin."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   0
      Left            =   240
      MouseIcon       =   "FrmManEngineer.frx":8670
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1380
      Width           =   3195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   825
      Index           =   0
      Left            =   240
      Top             =   1200
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Engineer Data Base "
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   10095
   End
End
Attribute VB_Name = "FrmEngineeringSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'###########移动无边框无标题栏窗口######
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
'Private Const GWL_STYLE = (-16)
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim ReturnVal As Long
        x = ReleaseCapture()
        ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub
'###########移动无边框无标题栏窗口 END ############
Private Sub Image1_Click()
     '卸载窗体
    End
    Unload Me
End Sub

Private Sub Image2_Click()
    FrmEngineeringSys.Hide
    FrmUserManage.Show 0
End Sub

Private Sub Label2_Click(Index As Integer)
    Select Case Index
     Case 0
        QueryTableName = "RFSRFQ"                                  '##########告诉通用查询窗口是对哪个表进行操作
        FrmQuery.Caption = "PDM-RFS/RFQ Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 1
        QueryTableName = "PJNO"                                   '##########告诉通用查询窗口是对哪个表进行操作
        FrmQuery.Caption = "PDM-Project Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 2
        QueryTableName = "CNCSN"
        FrmQuery.Caption = "PDM-CONCESSION Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 3
        QueryTableName = "CPCN"
        FrmQuery.Caption = "PDM-CPCN Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 4
        QueryTableName = "SER"
        FrmQuery.Caption = "PDM-SER Number Admin 工程管理子系统"
        FrmQuery.Show 0
    Case 5
        FrmGlueSupplier.Show 0
     Case 6
        QueryTableName = "FinsGd"
        FrmQuery.Caption = "PDM-Finish Goods Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 7
        QueryTableName = "SglPrt"                                   '##########告诉通用查询窗口是对哪个表进行操作
        
        ''和料号申请无关的栏目隐藏
        FrmQuery.DataGrid1.Columns(15).Visible = False
        FrmQuery.DataGrid1.Columns(16).Visible = False
        FrmQuery.DataGrid1.Columns(17).Visible = False
        FrmQuery.DataGrid1.Columns(18).Visible = False
        FrmQuery.DataGrid1.Columns(19).Visible = False
        FrmQuery.DataGrid1.Columns(20).Visible = False
        FrmQuery.DataGrid1.Columns(21).Visible = False
        FrmQuery.DataGrid1.Columns(22).Visible = False
        FrmQuery.DataGrid1.Columns(23).Visible = False
        FrmQuery.DataGrid1.Columns(24).Visible = False
        FrmQuery.Caption = "PDM-Single Part Number Admin 工程管理子系统"
        FrmQuery.Show 0
     Case 8
        QueryTableName = "StdPrtLibStructr"
        FrmStdPrtLibStructr.Show 0
    End Select
    FrmEngineeringSys.Hide
    Set FromForm = FrmEngineeringSys
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyF8
                Call Label2_Click(7)
            Case vbKeyF5
                Call Label4_Click(0)
            Case vbKeyF6
                Call Label4_Click(2)
            Case vbKeyF7
                Call Label4_Click(1)
            Case vbKeyF9
                Call Label2_Click(8)
            Case vbKeyEscape
                Call Image1_Click
    End Select
End Sub


Private Sub Form_Load()
'Load Skin & Format Control
''LoadSkin Me
Me.Picture = LoadPicture("")

If SystemAdmin <> "Y" Then
    Shape2(0).Visible = False
    Image2.Visible = False
End If
End Sub


Private Sub LblExit_Click()
     '卸载窗体
    End
    Unload Me
End Sub

Private Sub cmdCPCNImport_Click()
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Access Right is denied ", vbInformation, "System Info."
        Exit Sub
    End If
    
    FrmCPCNImport.Show 1
End Sub

Private Sub cmdSERImport_Click()
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Access Right is denied ", vbInformation, "System Info."
        Exit Sub
    End If
    
    FrmSERImport.Show 1
End Sub

Private Sub cmdSglPrtImport_Click()
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Access Right is denied ", vbInformation, "System Info."
        Exit Sub
    End If
    
    FrmSglPrtImport.Show 1
End Sub

Private Sub cmdFinsGdImport_Click()
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Access Right is denied ", vbInformation, "System Info."
        Exit Sub
    End If
    
    FrmFinsGdImport.Show 1
End Sub

Private Sub Label4_Click(Index As Integer)
    FrmEngineeringSys.Hide
    Set FromForm = FrmEngineeringSys
    Select Case Index
    Case 0
        FrmBOMAdmin.Show 0
    Case 1
        FrmBOMNPO.Show 0
    Case 2
        FrmBOMApproval.Show 0
    Case 3
        FrmTimeWriting.Show 0
    End Select
End Sub
