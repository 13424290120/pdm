VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCNCSNEdit 
   Caption         =   "CONCESSION Number Edit. CONCESSION 号码编辑"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCNCSNEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   12300
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdToQuerySglPrt 
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
      Left            =   8700
      TabIndex        =   44
      Top             =   8115
      Width           =   1620
   End
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
      Left            =   8700
      TabIndex        =   43
      Top             =   7275
      Width           =   1620
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
      Left            =   8670
      TabIndex        =   39
      Top             =   5415
      Width           =   3345
      Begin VB.ComboBox ComboPjtName 
         Height          =   345
         Left            =   105
         TabIndex        =   41
         Text            =   "ComboPjtName"
         Top             =   1065
         Width           =   3135
      End
      Begin VB.ComboBox ComboPJNOIndex 
         Height          =   345
         Left            =   105
         TabIndex        =   40
         Text            =   "ComboPJNOIndex"
         Top             =   270
         Width           =   3135
      End
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   5790
      TabIndex        =   14
      Top             =   2340
      Width           =   2775
   End
   Begin VB.TextBox TxtCPCNMP 
      Height          =   375
      Left            =   5790
      TabIndex        =   13
      Top             =   1620
      Width           =   1305
   End
   Begin VB.TextBox TxtCNCSNIndex 
      Height          =   375
      Left            =   5790
      TabIndex        =   12
      Top             =   180
      Width           =   2775
   End
   Begin VB.TextBox TxtIDSO 
      Height          =   375
      Left            =   5805
      TabIndex        =   11
      Top             =   3120
      Width           =   1380
   End
   Begin VB.TextBox TxtOpnDate 
      Height          =   375
      Left            =   5805
      TabIndex        =   10
      Top             =   3960
      Width           =   1350
   End
   Begin VB.TextBox TxtClosDate 
      Height          =   375
      Left            =   5805
      TabIndex        =   9
      Top             =   4815
      Width           =   1350
   End
   Begin VB.ComboBox CombIDSO 
      Height          =   345
      Left            =   7185
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3120
      Width           =   1410
   End
   Begin VB.TextBox TxtApplicant 
      Height          =   375
      Left            =   5790
      TabIndex        =   7
      Top             =   900
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   375
      Left            =   5775
      TabIndex        =   6
      Top             =   5655
      Width           =   2775
   End
   Begin VB.TextBox TxtPjtName 
      Height          =   375
      Left            =   5775
      TabIndex        =   5
      Top             =   6450
      Width           =   2775
   End
   Begin VB.TextBox TxtFinsGdNO 
      Height          =   375
      Left            =   5775
      TabIndex        =   4
      Top             =   7275
      Width           =   2775
   End
   Begin VB.TextBox TxtSglPrtNO 
      Height          =   375
      Left            =   5775
      TabIndex        =   3
      Top             =   8115
      Width           =   2775
   End
   Begin VB.TextBox TxtCommtNote 
      Height          =   375
      Left            =   5775
      TabIndex        =   2
      Top             =   8955
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
      Left            =   8685
      TabIndex        =   1
      Top             =   180
      Width           =   1740
   End
   Begin VB.ComboBox CombCPCNMP 
      Height          =   345
      Left            =   7170
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1635
      Width           =   1410
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8415
      Top             =   9795
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   7170
      TabIndex        =   15
      Top             =   4800
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   7170
      TabIndex        =   16
      Top             =   3960
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmCNCSNEdit.frx":08CA
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
      Left            =   9090
      TabIndex        =   42
      Top             =   2265
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
      Left            =   4035
      MouseIcon       =   "FrmCNCSNEdit.frx":090F
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   9870
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3435
      Picture         =   "FrmCNCSNEdit.frx":0C19
      Top             =   9840
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
      Left            =   6315
      MouseIcon       =   "FrmCNCSNEdit.frx":1035
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   9870
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5715
      Picture         =   "FrmCNCSNEdit.frx":133F
      Top             =   9840
      Width           =   300
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description 让步描述"
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
      Left            =   2685
      MouseIcon       =   "FrmCNCSNEdit.frx":175B
      TabIndex        =   36
      Top             =   2340
      Width           =   3000
   End
   Begin VB.Label LblCPCNMP 
      BackStyle       =   0  'Transparent
      Caption         =   "CD/DR/CR/MP 是哪个状态"
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
      Left            =   2010
      MouseIcon       =   "FrmCNCSNEdit.frx":1A65
      TabIndex        =   35
      Top             =   1620
      Width           =   3675
   End
   Begin VB.Label LblCNCSNIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "CONCESSION NO. CONCESSION 编号"
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
      Left            =   195
      MouseIcon       =   "FrmCNCSNEdit.frx":1D6F
      TabIndex        =   34
      Top             =   180
      Width           =   5490
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
      Left            =   3000
      MouseIcon       =   "FrmCNCSNEdit.frx":2079
      TabIndex        =   33
      Top             =   3135
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
      Left            =   2835
      MouseIcon       =   "FrmCNCSNEdit.frx":2383
      TabIndex        =   32
      Top             =   3975
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
      Left            =   2805
      MouseIcon       =   "FrmCNCSNEdit.frx":268D
      TabIndex        =   31
      Top             =   4830
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
      Left            =   6300
      TabIndex        =   30
      Top             =   2865
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
      Left            =   7665
      TabIndex        =   29
      Top             =   2865
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
      Left            =   6285
      TabIndex        =   28
      Top             =   3705
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
      Left            =   7665
      TabIndex        =   27
      Top             =   3705
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
      Left            =   6300
      TabIndex        =   26
      Top             =   4545
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
      Left            =   7665
      TabIndex        =   25
      Top             =   4545
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
      Left            =   3330
      MouseIcon       =   "FrmCNCSNEdit.frx":2997
      TabIndex        =   24
      Top             =   930
      Width           =   2355
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   9705
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
      Left            =   1470
      MouseIcon       =   "FrmCNCSNEdit.frx":2CA1
      TabIndex        =   23
      Top             =   5655
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
      Left            =   1755
      MouseIcon       =   "FrmCNCSNEdit.frx":2FAB
      TabIndex        =   22
      Top             =   6465
      Width           =   3915
   End
   Begin VB.Label LblFinsGdNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good NO. 更改涉及成品编号"
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
      Left            =   690
      MouseIcon       =   "FrmCNCSNEdit.frx":32B5
      TabIndex        =   21
      Top             =   7275
      Width           =   4980
   End
   Begin VB.Label LblSglPrtNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part NO. 更改涉及单件编号"
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
      Left            =   870
      MouseIcon       =   "FrmCNCSNEdit.frx":35BF
      TabIndex        =   20
      Top             =   8115
      Width           =   4800
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
      Left            =   1695
      MouseIcon       =   "FrmCNCSNEdit.frx":38C9
      TabIndex        =   19
      Top             =   8955
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
      Left            =   7650
      TabIndex        =   18
      Top             =   1380
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
      Left            =   6300
      TabIndex        =   17
      Top             =   1380
      Width           =   285
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   10395
   End
End
Attribute VB_Name = "FrmCNCSNEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriCNCSNIndex As String                       '############变量改成对应的表字段名字

Private Sub CmdSysDistrb_Click()
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim BitNum As Integer   '定义取出的去头数HK1FS和零后的位数
Dim i As Integer
Dim j As Integer
Dim OriString As String
Conn.ConnectionString = connString
Conn.Open
Set rcds.ActiveConnection = Conn
On Error Resume Next           '############以下相关改成对应的表字段名字
rcds.Open "select top 1 Right(CNCSNIndex,7)+1 from CNCSN WHERE ((Right(CNCSNIndex,7)+1) Not In (select Right(CNCSNIndex,7) from CNCSN))  order by Right(CNCSNIndex,7)", Conn, adOpenKeyset, adOpenStatic  'CNCSNIndex+1表示每1位申请一个号，也就是从末位开始递增1
BitNum = Len(rcds.Fields(0).Value)  '判断实际查询去头数HK1FS和零后的数字是几位

OriString = "CNCSN"
For i = 0 To (12 - 5 - BitNum - 1)  '判断HK1FS和实际数字之间有几个0,有几个就增加几个
OriString = OriString + "0"
Next i
        If Modify = False Then
            TxtCNCSNIndex.Text = OriString + Trim(Str(rcds.Fields(0).Value))
            'MsgBox "Succeed to Add" + vbCrLf + "增加成功"   这句可以不用，用了还要关窗口，麻烦
        End If
    If rcds.EOF Or rcds.BOF Then
        MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
        Exit Sub
    End If

  '如果不能查到记录
    If rcds.RecordCount = 0 Then
      '系统提示信息，没有推荐号，请自行选择
    MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
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

Private Sub CmdToQuerySglPrt_Click()
QuerytableName = "SglPrt"                                  '##########告诉通用查询窗口是对哪个表进行操作

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
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNO为要查询的表名
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始

    Do While sqlUsrCtrl.IfBOForEOF = False
        sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '其中1UseEquel代表= 2UseLike是代表Like

       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
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
TxtApplicant.Text = PDMUserName
CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0
CombCPCNMP.AddItem ("CD")
CombCPCNMP.AddItem ("DR")
CombCPCNMP.AddItem ("CR")
CombCPCNMP.AddItem ("MP")
CombCPCNMP.ListIndex = 2

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
If Trim(TxtCNCSNIndex) = "" Or (Len(TxtCNCSNIndex) <> 12) Then
    MsgBox "Please input CONCESSION Number" + vbCrLf + "请输入CONCESSION号", vbInformation, "System Info."
    TxtCNCSNIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (left(TxtCNCSNIndex, 5) = "CNCSN" And Isnum(right(TxtCNCSNIndex, 7))) Then  '其中Left() Right()是从左边和右边截取字符串
    MsgBox "CONCESSION Series Number is 5 Letter CNCSN + 7 Number" + vbCrLf + "CONCESSION是7位字符CNCSN + 7位数字的编号", vbInformation, "System Info."
    TxtCNCSNIndex.SetFocus
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
    MsgBox "Please input Description" + vbCrLf + "请输入让步描述", vbInformation, "System Info."
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
If Trim(TxtSglPrtNO) = "" Then
    MsgBox "Please input relevant single part 12NC" + vbCrLf + "请输入涉及零件的12NC", vbInformation, "System Info."
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
     
   With MyCNCSN              '已经定义Public MyCNCSN As New ClsCNCSN, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字
    .CNCSNIndex = TxtCNCSNIndex.Text
    .Applicant = TxtApplicant.Text
    .CPCNMP = CombCPCNMP.Text
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
     
           '判断CNCSNIndex序号是否已经存在
                If .In_DB(TxtCNCSNIndex.Text) = True Then
                   MsgBox "CONCESSION number exists, Please re-input" + vbCrLf + "CONCESSION号重复，请重新设置", vbInformation, "System Info."
                   TxtCNCSNIndex.SetFocus
                   TxtCNCSNIndex.SelStart = 0
                   TxtCNCSNIndex.SelLength = Len(TxtCNCSNIndex)
                   Exit Sub
                Else
                   .Insert                   '添加
                    MsgBox "Succeed to Add" + vbCrLf + "添加成功", vbInformation, "System Info."
                End If
       Else  '判断为修改操作
        .Update (OriCNCSNIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "修改成功", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '关闭自身窗口
End Sub


