VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSEREdit 
   Caption         =   "SER Number Edit. SER 号码编辑"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSEREdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   12360
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdToQueryFinsGd 
      Caption         =   "Search 查询"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8415
      TabIndex        =   43
      Top             =   8160
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seek PjtNO. from PjtName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   8415
      TabIndex        =   39
      Top             =   6300
      Width           =   3345
      Begin VB.ComboBox ComboPJNOIndex 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   41
         Text            =   "ComboPJNOIndex"
         Top             =   300
         Width           =   3135
      End
      Begin VB.ComboBox ComboPjtName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   40
         Text            =   "ComboPjtName"
         Top             =   1065
         Width           =   3135
      End
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5430
      TabIndex        =   14
      Top             =   3225
      Width           =   2775
   End
   Begin VB.TextBox TxtCAorA 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   13
      Top             =   1650
      Width           =   1380
   End
   Begin VB.TextBox TxtSERIndex 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   210
      Width           =   2775
   End
   Begin VB.TextBox TxtIDSO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   11
      Top             =   4005
      Width           =   1380
   End
   Begin VB.TextBox TxtOpnDate 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   10
      Top             =   4845
      Width           =   1350
   End
   Begin VB.TextBox TxtClosDate 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   9
      Top             =   5700
      Width           =   1350
   End
   Begin VB.ComboBox CombIDSO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4005
      Width           =   1410
   End
   Begin VB.TextBox TxtApplicant 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   930
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   6
      Top             =   6540
      Width           =   2775
   End
   Begin VB.TextBox TxtPjtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   5
      Top             =   7335
      Width           =   2775
   End
   Begin VB.TextBox TxtFinsGdNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   4
      Top             =   8160
      Width           =   2775
   End
   Begin VB.TextBox TxtSglPrtNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   3
      Top             =   2430
      Width           =   2775
   End
   Begin VB.TextBox TxtCommtNote 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   2
      Top             =   8985
      Width           =   2775
   End
   Begin VB.CommandButton CmdSysDistrb 
      Caption         =   "System Distribute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8295
      TabIndex        =   1
      Top             =   210
      Width           =   1740
   End
   Begin VB.ComboBox CombCAorA 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   1410
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8025
      Top             =   9765
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   6810
      TabIndex        =   15
      Top             =   5685
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   6810
      TabIndex        =   16
      Top             =   4845
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmSEREdit.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      Left            =   8865
      TabIndex        =   42
      Top             =   3150
      Width           =   2715
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OK 确 定"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3645
      MouseIcon       =   "FrmSEREdit.frx":090F
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3045
      Picture         =   "FrmSEREdit.frx":0C19
      Top             =   9810
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel 取 消"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5925
      MouseIcon       =   "FrmSEREdit.frx":1035
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5325
      Picture         =   "FrmSEREdit.frx":133F
      Top             =   9810
      Width           =   300
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description 物品描述"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2325
      MouseIcon       =   "FrmSEREdit.frx":175B
      TabIndex        =   36
      Top             =   3225
      Width           =   3000
   End
   Begin VB.Label LblCAorA 
      BackStyle       =   0  'Transparent
      Caption         =   "CA or FA or RJ是哪个状态"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1590
      MouseIcon       =   "FrmSEREdit.frx":1A65
      TabIndex        =   35
      Top             =   1650
      Width           =   3705
   End
   Begin VB.Label LblSERIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "SER NO. SER 编号"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2625
      MouseIcon       =   "FrmSEREdit.frx":1D6F
      TabIndex        =   34
      Top             =   210
      Width           =   2670
   End
   Begin VB.Label LblIDSO 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Open/Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MouseIcon       =   "FrmSEREdit.frx":2079
      TabIndex        =   33
      Top             =   4020
      Width           =   2685
   End
   Begin VB.Label LblOpnDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date 开始日期"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2475
      MouseIcon       =   "FrmSEREdit.frx":2383
      TabIndex        =   32
      Top             =   4860
      Width           =   2865
   End
   Begin VB.Label LblClosDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Close Date 结束日期"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2445
      MouseIcon       =   "FrmSEREdit.frx":268D
      TabIndex        =   31
      Top             =   5715
      Width           =   2895
   End
   Begin VB.Label LblOld1 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5940
      TabIndex        =   30
      Top             =   3750
      Width           =   285
   End
   Begin VB.Label LblNew1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7305
      TabIndex        =   29
      Top             =   3750
      Width           =   390
   End
   Begin VB.Label LblOld2 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5925
      TabIndex        =   28
      Top             =   4590
      Width           =   285
   End
   Begin VB.Label LblNew2 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7305
      TabIndex        =   27
      Top             =   4590
      Width           =   390
   End
   Begin VB.Label LblOld3 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5940
      TabIndex        =   26
      Top             =   5415
      Width           =   285
   End
   Begin VB.Label LblNew3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7305
      TabIndex        =   25
      Top             =   5415
      Width           =   390
   End
   Begin VB.Label LblApplicant 
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant 申请人"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      MouseIcon       =   "FrmSEREdit.frx":2997
      TabIndex        =   24
      Top             =   960
      Width           =   2355
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   2850
      Shape           =   4  'Rounded Rectangle
      Top             =   9675
      Width           =   4740
   End
   Begin VB.Label LblPJNOIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Number 所属项目编号"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1110
      MouseIcon       =   "FrmSEREdit.frx":2CA1
      TabIndex        =   23
      Top             =   6540
      Width           =   4200
   End
   Begin VB.Label LblPjtName 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name 项目名称描述"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1395
      MouseIcon       =   "FrmSEREdit.frx":2FAB
      TabIndex        =   22
      Top             =   7350
      Width           =   3915
   End
   Begin VB.Label LblFinsGdNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good NO. 相关的成品编号"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      MouseIcon       =   "FrmSEREdit.frx":32B5
      TabIndex        =   21
      Top             =   8160
      Width           =   4650
   End
   Begin VB.Label LblSglPrtNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part NO. SER单件物品编号"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   555
      MouseIcon       =   "FrmSEREdit.frx":35BF
      TabIndex        =   20
      Top             =   2430
      Width           =   4770
   End
   Begin VB.Label LblCommtNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment  Note. 注释和备注"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1335
      MouseIcon       =   "FrmSEREdit.frx":38C9
      TabIndex        =   19
      Top             =   8985
      Width           =   3975
   End
   Begin VB.Label LblNew0 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7260
      TabIndex        =   18
      Top             =   1410
      Width           =   390
   End
   Begin VB.Label LblOld0 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5910
      TabIndex        =   17
      Top             =   1410
      Width           =   285
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   2535
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7590
   End
End
Attribute VB_Name = "FrmSEREdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriSERIndex As String                       '############变量改成对应的表字段名字

Private Sub CmdSysDistrb_Click()
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim BitNum As Integer   '定义取出的去头数SER和零后的位数
Dim i As Integer
Dim j As Integer
Dim OriString, StrSql As String

MousePointer = vbHourglass   '搜索时间较长，需要定义鼠标状态
Conn.ConnectionString = connString
Conn.Open
Set rcds.ActiveConnection = Conn
On Error Resume Next           '############以下相关改成对应的表字段名字
StrSql = "select TOP 1 Right(Max(SERIndex),9)+1 from SER group by SERIndex order by SERIndex DESC"

rcds.Open StrSql, Conn, adOpenKeyset, adOpenStatic  'SERIndex+1表示每1位申请一个号，也就是从末位开始递增1
BitNum = Len(rcds.Fields(0).Value)  '判断实际查询去头数SER和零后的数字是几位

OriString = "SER"
For i = 0 To (12 - 3 - BitNum - 1)  '判断SER和实际数字之间有几个0,有几个就增加几个
    OriString = OriString + "0"
Next i
If Modify = False Then
    TxtSERIndex.Text = OriString + Trim(Str(rcds.Fields(0).Value))
    MousePointer = vbDefault               '恢复鼠标状态
    'MsgBox "Succeed to Add" + vbCrLf + "增加成功"   这句可以不用，用了还要关窗口，麻烦
End If
If rcds.EOF Or rcds.BOF Then
    MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
    MousePointer = vbDefault                  '恢复鼠标状态
    Exit Sub
End If

'如果不能查到记录
If rcds.RecordCount = 0 Then
  '系统提示信息，没有推荐号，请自行选择
MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
MousePointer = vbDefault                  '恢复鼠标状态
End If
Conn.Close
End Sub

Private Sub CmdToQueryFinsGd_Click()
QuerytableName = "FinsGd"                                  '##########告诉通用查询窗口是对哪个表进行操作

    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False

        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    
FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
End Sub

Private Sub ComboPJNOIndex_Click()
TxtPJNOIndex.Text = ComboPJNOIndex.Text
TxtPjtName.Text = ComboPjtName.List(ComboPJNOIndex.ListIndex)
End Sub

Private Sub ComboPjtName_Click()
TxtPjtName.Text = ComboPjtName.Text
TxtPJNOIndex.Text = ComboPJNOIndex.List(ComboPjtName.ListIndex)
End Sub

Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub

Private Sub TxtPJNOIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1         'sqlSDBC1为用户自定义控件

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNO为要查询的表名
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始

    Do While sqlUsrCtrl.IfBOForEOF = False
        sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '其中1UseEquel代表= 2UseLike是代表Like

       ComboPJNOIndex.AddItem (FormatNumber6(CStr(UsrCtlFind(0))))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind括号中的3()是对应Description的字段序号
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
    
    Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub TxtPjtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNO为要查询的表名
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始

     Do While sqlUsrCtrl.IfBOForEOF = False
       sqlUsrCtrl.FindRecord "Description", UseLike, Trim(TxtPjtName.Text)  '其中1UseEquel代表= 2UseLike是代表Like

       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind括号中的3()是对应Description的字段序号
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
  
     Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub Form_Load()               '############以下相关改成对应的控件,表的字段名字
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me
TxtApplicant.Text = PDMUserName
CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0
CombCAorA.AddItem ("CA")
CombCAorA.AddItem ("FA")
CombCAorA.AddItem ("RJ")
CombCAorA.ListIndex = 0

DTPickerOpnDate.Value = Date
DTPickerClosDate.Value = Date
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
  Function Isnum(Str As String) As Boolean    '判断一个字符串中是否含有数字  用IsNumeric判断0000d031为真(当成double型数字)
  Isnum = True
  Dim i  As Integer
  For i = 1 To Len(Str)
      Select Case Mid(Str, i, 1)
          Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          ' Isnum = True  这里写Isnum = True就出错,因为如果中间是字母false了后面有数字的话又成为true了
          Case Else
            Isnum = False
      End Select
  Next
  End Function

Private Function Check() As Boolean                        '############以下相关改成对应的控件,表的字段名字
If Trim(TxtSERIndex) = "" Or (Len(TxtSERIndex) <> 12) Then
    MsgBox "Please input SER Number" + vbCrLf + "请输入SER号", vbInformation, "System Info."
    TxtSERIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (left(TxtSERIndex, 3) = "SER" And Isnum(right(TxtSERIndex, 9))) Then  '其中Left() Right()是从左边和右边截取字符串
    MsgBox "SER Series Number is 3 Letter SER + 9 Number" + vbCrLf + "SER是3位字符SER + 9位数字的编号", vbInformation, "System Info."
    TxtSERIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtApplicant) = "" Then
    MsgBox "Please input Applicant Name" + vbCrLf + "请输入申请人名", vbInformation, "System Info."
    TxtApplicant.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtDescription) = "" Then
    MsgBox "Please input Description" + vbCrLf + "请输入物品描述", vbInformation, "System Info."
    TxtDescription.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPJNOIndex) = "" Or (Not Isnum(TxtPJNOIndex)) Or (Len(TxtPJNOIndex) <> 6) Then
    MsgBox "Please input Project Number, 6 number" + vbCrLf + "请输入涉及项目编号, 6位的数字", vbInformation, "System Info."
    TxtPJNOIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPjtName) = "" Then
    MsgBox "Please input Project Name" + vbCrLf + "请输入涉及项目名称", vbInformation, "System Info."
    TxtPjtName.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtFinsGdNO) = "" Then
    MsgBox "Please input relevant finish goods 12NC" + vbCrLf + "请输入涉及成品的12NC", vbInformation, "System Info."
    TxtFinsGdNO.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtSglPrtNO) = "" Or (Not Isnum(TxtSglPrtNO)) Or (Len(TxtSglPrtNO) <> 12) Then
    MsgBox "Please input relevant single part 12NC, Must be 12 Number" + vbCrLf + "请输入涉及零件的12NC,必须是12位数字", vbInformation, "System Info."
    TxtSglPrtNO.SetFocus
    Check = False
    Exit Function
  End If
  
  
   Check = True
End Function


Private Sub lblOk_Click()
    
   '判断要编辑信息是否完整
   If Check = False Then
    Exit Sub
   End If
     
   With MySER              '已经定义Public MySER As New ClsSER, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字
    .SERIndex = TxtSERIndex.Text
    .Applicant = TxtApplicant.Text
    .CAorA = CombCAorA.Text
    .Description = TxtDescription.Text
    .IDSO = CombIDSO.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
    .PJNOIndex = TxtPJNOIndex.Text
    .PjtName = TxtPjtName.Text
    .FinsGdNO = TxtFinsGdNO.Text
    .SglPrtNO = TxtSglPrtNO.Text
    .CommtNote = TxtCommtNote.Text
    
   
            '判断操作是添加还是修改
       If Modify = False Then         '判断为添加操作
     
           '判断SERIndex序号是否已经存在
                If .In_DB(TxtSERIndex.Text) = True Then
                   MsgBox "SER number exists, Please re-input" + vbCrLf + "SER号重复，请重新设置", vbInformation, "System Info."
                   TxtSERIndex.SetFocus
                   TxtSERIndex.SelStart = 0
                   TxtSERIndex.SelLength = Len(TxtSERIndex)
                   Exit Sub
                Else
                   .Insert                   '添加
                    MsgBox "Succeed to Add" + vbCrLf + "添加成功", vbInformation, "System Info."
                End If
       Else  '判断为修改操作
        .Update (OriSERIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "修改成功", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '关闭自身窗口
End Sub


