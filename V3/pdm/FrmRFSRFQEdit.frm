VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRFSRFQEdit 
   Caption         =   "RFS/RFQ Number Edit. RFS/RFQ 号码编辑"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRFSRFQEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10635
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtApplicant 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5565
      TabIndex        =   29
      Top             =   2835
      Width           =   2775
   End
   Begin VB.CommandButton CmdSysAdd3 
      Caption         =   "Add"
      Height          =   495
      Left            =   8490
      TabIndex        =   28
      Top             =   1515
      Width           =   930
   End
   Begin VB.CommandButton CmdSysAdd2 
      Caption         =   "Add"
      Height          =   495
      Left            =   8475
      TabIndex        =   26
      Top             =   915
      Width           =   930
   End
   Begin VB.CommandButton CmdSysAdd1 
      Caption         =   "Add"
      Height          =   495
      Left            =   8460
      TabIndex        =   24
      Top             =   315
      Width           =   930
   End
   Begin VB.ComboBox CombIDSQ 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5055
      Width           =   1410
   End
   Begin VB.TextBox TxtClosDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   15
      Top             =   6750
      Width           =   1350
   End
   Begin VB.TextBox TxtOpnDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   14
      Top             =   5895
      Width           =   1350
   End
   Begin VB.TextBox TxtIDSQ 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   13
      Top             =   5055
      Width           =   1380
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   6945
      TabIndex        =   10
      Top             =   6750
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   6945
      TabIndex        =   9
      Top             =   5895
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.TextBox TxtRFSRFQIndex 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5565
      TabIndex        =   2
      Top             =   2115
      Width           =   2775
   End
   Begin VB.TextBox TxtLeader 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5565
      TabIndex        =   1
      Top             =   3555
      Width           =   2775
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5565
      TabIndex        =   0
      Top             =   4275
      Width           =   2775
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmRFSRFQEdit.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2010
      Left            =   8805
      TabIndex        =   31
      Top             =   5085
      Width           =   1755
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   2790
      Shape           =   4  'Rounded Rectangle
      Top             =   7395
      Width           =   4740
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
      Left            =   3105
      MouseIcon       =   "FrmRFSRFQEdit.frx":090F
      TabIndex        =   30
      Top             =   2865
      Width           =   2355
   End
   Begin VB.Label LblNote3 
      Caption         =   "116000 To 119999  : LoB Automotive roadmap & application studies"
      Height          =   330
      Left            =   1005
      TabIndex        =   27
      Top             =   1560
      Width           =   7305
   End
   Begin VB.Label LblNote2 
      Caption         =   "113000 To 115999  : LoB AV-MM roadmap & application studies"
      Height          =   330
      Left            =   990
      TabIndex        =   25
      Top             =   960
      Width           =   7305
   End
   Begin VB.Label LblNote1 
      Caption         =   "110000 To 112999  : General "
      Height          =   330
      Left            =   990
      TabIndex        =   23
      Top             =   405
      Width           =   7305
   End
   Begin VB.Shape Shape1 
      Height          =   2460
      Left            =   675
      Shape           =   4  'Rounded Rectangle
      Top             =   195
      Width           =   9150
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
      Left            =   7440
      TabIndex        =   21
      Top             =   6465
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
      Left            =   6075
      TabIndex        =   20
      Top             =   6465
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
      Left            =   7440
      TabIndex        =   19
      Top             =   5640
      Width           =   390
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
      Left            =   6075
      TabIndex        =   18
      Top             =   5640
      Width           =   285
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
      Left            =   7440
      TabIndex        =   17
      Top             =   4800
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
      Left            =   6075
      TabIndex        =   16
      Top             =   4800
      Width           =   285
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
      Left            =   2580
      MouseIcon       =   "FrmRFSRFQEdit.frx":0C19
      TabIndex        =   12
      Top             =   6765
      Width           =   2895
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
      Left            =   2610
      MouseIcon       =   "FrmRFSRFQEdit.frx":0F23
      TabIndex        =   11
      Top             =   5910
      Width           =   2865
   End
   Begin VB.Label LblIDSQ 
      BackStyle       =   0  'Transparent
      Caption         =   "RFS or RFQ"
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
      Left            =   3735
      MouseIcon       =   "FrmRFSRFQEdit.frx":122D
      TabIndex        =   8
      Top             =   5070
      Width           =   1740
   End
   Begin VB.Label LblRFQRFSIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "RFS/RFQ NO. RFS/RFQ编号"
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
      Left            =   1500
      MouseIcon       =   "FrmRFSRFQEdit.frx":1537
      TabIndex        =   7
      Top             =   2115
      Width           =   3960
   End
   Begin VB.Label LblLeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Leader 主导人"
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
      Left            =   3435
      MouseIcon       =   "FrmRFSRFQEdit.frx":1841
      TabIndex        =   6
      Top             =   3555
      Width           =   2025
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description 项目描述"
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
      Left            =   2460
      MouseIcon       =   "FrmRFSRFQEdit.frx":1B4B
      TabIndex        =   5
      Top             =   4275
      Width           =   3000
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5265
      Picture         =   "FrmRFSRFQEdit.frx":1E55
      Top             =   7530
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel 取 消"
      Height          =   315
      Left            =   5865
      MouseIcon       =   "FrmRFSRFQEdit.frx":2271
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2985
      Picture         =   "FrmRFSRFQEdit.frx":257B
      Top             =   7530
      Width           =   300
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OK 确 定"
      Height          =   315
      Left            =   3585
      MouseIcon       =   "FrmRFSRFQEdit.frx":2997
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7560
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRFSRFQEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriRFSRFQIndex As String                       '############变量改成对应的表字段名字

Private Sub CmdSysAdd1_Click()     '添加号码段110000 To 112999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next           '############以下相关改成对应的表字段名字
rcds.Open "select top 1 RFSRFQIndex+10 from RFSRFQ WHERE (((RFSRFQIndex+10) Not In (select RFSRFQIndex from RFSRFQ))and (RFSRFQIndex+10) between 110000 and 112999) order by RFSRFQIndex+10", Conn, adOpenKeyset, adOpenStatic  'RFSRFQIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtRFSRFQIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub CmdSysAdd2_Click()     '添加号码段113000 To 115999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 RFSRFQIndex+10 from RFSRFQ WHERE (((RFSRFQIndex+10) Not In (select RFSRFQIndex from RFSRFQ))and (RFSRFQIndex+10) between 113000 and 115999) order by RFSRFQIndex+10", Conn, adOpenKeyset, adOpenStatic  'RFSRFQIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtRFSRFQIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub CmdSysAdd3_Click()     '添加号码段116000 To 119999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 RFSRFQIndex+10 from RFSRFQ WHERE (((RFSRFQIndex+10) Not In (select RFSRFQIndex from RFSRFQ))and (RFSRFQIndex+10) between 116000 and 119999) order by RFSRFQIndex+10", Conn, adOpenKeyset, adOpenStatic  'RFSRFQIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtRFSRFQIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub Form_Load()               '############以下相关改成对应的控件,表的字段名字
'Load Skin & Format Control
'LoadSkin Me
'Call ResizeInit(Me)
TxtApplicant.Text = PDMUserName
CombIDSQ.AddItem ("RFS")
CombIDSQ.AddItem ("RFQ")
CombIDSQ.ListIndex = 0
DTPickerOpnDate.Value = Date
DTPickerClosDate.Value = Date
End Sub

Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
Private Function Check() As Boolean                        '############以下相关改成对应的控件,表的字段名字
If Trim(TxtRFSRFQIndex) = "" Then
    MsgBox "Please input RFS/RFQ Number" + vbCrLf + "请输入RFS/RFQ号", vbInformation, "System Info."
    TxtRFSRFQIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (Len(TxtRFSRFQIndex) = 6 And IsNumeric(TxtRFSRFQIndex)) Then
    MsgBox "RFSRFQIndex is 6 Number, No letter " + vbCrLf + "请输入6位的数字,无字母", vbInformation, "System Info."
    TxtRFSRFQIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtApplicant) = "" Then
    MsgBox "Please input Applicant Name" + vbCrLf + "请输入申请人名", vbInformation, "System Info."
    TxtLeader.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtLeader) = "" Then
    MsgBox "Please input Leader Name" + vbCrLf + "请输入主导人名", vbInformation, "System Info."
    TxtLeader.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtDescription) = "" Then
    MsgBox "Please input Description" + vbCrLf + "请输入项目描述", vbInformation, "System Info."
    TxtDescription.SetFocus
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
     
   With MyRFSRFQ              '已经定义Public MyRFSRFQ As New ClsRFSRFQ, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字
    .RFSRFQIndex = TxtRFSRFQIndex.Text
    .Applicant = TxtApplicant.Text
    .Leader = TxtLeader.Text
    .Description = TxtDescription.Text
    .IDSQ = CombIDSQ.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
   
            '判断操作是添加还是修改
       If Modify = False Then         '判断为添加操作
     
           '判断RFQRFSIndex序号是否已经存在
                If .In_DB(TxtRFSRFQIndex.Text) = True Then
                   MsgBox "RFSRFQIndex exists, Please re-input" + vbCrLf + "RFSRFQIndex号重复，请重新设置", vbInformation, "System Info."
                   TxtRFSRFQIndex.SetFocus
                   TxtRFSRFQIndex.SelStart = 0
                   TxtRFSRFQIndex.SelLength = Len(TxtRFSRFQIndex)
                   Exit Sub
                Else
                   .Insert                   '添加
                    MsgBox "Succeed to Add" + vbCrLf + "添加成功", vbInformation, "System Info."
                End If
       Else  '判断为修改操作
        .Update (OriRFSRFQIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "修改成功", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '关闭自身窗口
End Sub

