VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPJNOEdit 
   Caption         =   "Project Number Edit. Project 号码编辑"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPJNOEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10650
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdSysAdd5 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8430
      TabIndex        =   34
      Top             =   1710
      Width           =   930
   End
   Begin VB.CommandButton CmdSysAdd4 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8430
      TabIndex        =   33
      Top             =   1350
      Width           =   930
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   5535
      TabIndex        =   12
      Top             =   4275
      Width           =   2775
   End
   Begin VB.TextBox TxtLeader 
      Height          =   375
      Left            =   5535
      TabIndex        =   11
      Top             =   3555
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   375
      Left            =   5535
      TabIndex        =   10
      Top             =   2115
      Width           =   2775
   End
   Begin VB.TextBox TxtIDSQ 
      Height          =   375
      Left            =   5550
      TabIndex        =   7
      Top             =   5055
      Width           =   1380
   End
   Begin VB.TextBox TxtOpnDate 
      Height          =   375
      Left            =   5550
      TabIndex        =   6
      Top             =   5895
      Width           =   1350
   End
   Begin VB.TextBox TxtClosDate 
      Height          =   375
      Left            =   5550
      TabIndex        =   5
      Top             =   6750
      Width           =   1350
   End
   Begin VB.ComboBox CombIDSQ 
      Height          =   345
      Left            =   6930
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5055
      Width           =   1410
   End
   Begin VB.CommandButton CmdSysAdd1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8430
      TabIndex        =   3
      Top             =   255
      Width           =   930
   End
   Begin VB.CommandButton CmdSysAdd2 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8430
      TabIndex        =   2
      Top             =   615
      Width           =   930
   End
   Begin VB.CommandButton CmdSysAdd3 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8430
      TabIndex        =   1
      Top             =   990
      Width           =   930
   End
   Begin VB.TextBox TxtApplicant 
      Height          =   375
      Left            =   5535
      TabIndex        =   0
      Top             =   2835
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   6915
      TabIndex        =   8
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
      Left            =   6915
      TabIndex        =   9
      Top             =   5895
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmPJNOEdit.frx":08CA
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
      Left            =   8775
      TabIndex        =   35
      Top             =   5070
      Width           =   1755
   End
   Begin VB.Label LblNote5 
      Caption         =   "160000 To 169999  : Automotive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   32
      Top             =   1740
      Width           =   7305
   End
   Begin VB.Label LblNote4 
      Caption         =   "150000 To 159999  : Multi Media External"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   31
      Top             =   1395
      Width           =   7305
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
      Left            =   3555
      MouseIcon       =   "FrmPJNOEdit.frx":090F
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2955
      Picture         =   "FrmPJNOEdit.frx":0C19
      Top             =   7530
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
      Left            =   5835
      MouseIcon       =   "FrmPJNOEdit.frx":1035
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5235
      Picture         =   "FrmPJNOEdit.frx":133F
      Top             =   7530
      Width           =   300
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
      Left            =   2430
      MouseIcon       =   "FrmPJNOEdit.frx":175B
      TabIndex        =   28
      Top             =   4275
      Width           =   3000
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
      Left            =   3405
      MouseIcon       =   "FrmPJNOEdit.frx":1A65
      TabIndex        =   27
      Top             =   3555
      Width           =   2025
   End
   Begin VB.Label LblPJNOIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Project NO. Project 编号"
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
      Left            =   1965
      MouseIcon       =   "FrmPJNOEdit.frx":1D6F
      TabIndex        =   26
      Top             =   2115
      Width           =   3465
   End
   Begin VB.Label LblIDSQ 
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
      Left            =   2745
      MouseIcon       =   "FrmPJNOEdit.frx":2079
      TabIndex        =   25
      Top             =   5070
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
      Left            =   2580
      MouseIcon       =   "FrmPJNOEdit.frx":2383
      TabIndex        =   24
      Top             =   5910
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
      Left            =   2550
      MouseIcon       =   "FrmPJNOEdit.frx":268D
      TabIndex        =   23
      Top             =   6765
      Width           =   2895
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
      Left            =   6045
      TabIndex        =   22
      Top             =   4800
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
      Left            =   7410
      TabIndex        =   21
      Top             =   4800
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
      Left            =   6045
      TabIndex        =   20
      Top             =   5640
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
      Left            =   7410
      TabIndex        =   19
      Top             =   5640
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
      Left            =   6045
      TabIndex        =   18
      Top             =   6465
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
      Left            =   7410
      TabIndex        =   17
      Top             =   6465
      Width           =   390
   End
   Begin VB.Shape Shape1 
      Height          =   2460
      Left            =   645
      Shape           =   4  'Rounded Rectangle
      Top             =   195
      Width           =   9150
   End
   Begin VB.Label LblNote1 
      Caption         =   "120000 To 129999  : Audio Video PSS Internal OEM/ODM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   16
      Top             =   315
      Width           =   7305
   End
   Begin VB.Label LblNote2 
      Caption         =   "130000 To 139999  : Audio Video External OEM/ODM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   15
      Top             =   690
      Width           =   7305
   End
   Begin VB.Label LblNote3 
      Caption         =   "140000 To 149999  : Multi Media PSS Internal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   14
      Top             =   1050
      Width           =   7305
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
      Left            =   3075
      MouseIcon       =   "FrmPJNOEdit.frx":2997
      TabIndex        =   13
      Top             =   2865
      Width           =   2355
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   7395
      Width           =   4740
   End
End
Attribute VB_Name = "FrmPJNOEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriPJNOIndex As String                       '############变量改成对应的表字段名字

Private Sub CmdSysAdd1_Click()     '添加号码段120000 To 129999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next           '############以下相关改成对应的表字段名字
rcds.Open "select top 1 PJNOIndex+10 from PJNO WHERE (((PJNOIndex+10) Not In (select PJNOIndex from PJNO))and (PJNOIndex+10) between 120000 and 129999) order by PJNOIndex+10", Conn, adOpenKeyset, adOpenStatic  'PJNOIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtPJNOIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub CmdSysAdd2_Click()     '添加号码段130000 To 139999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 PJNOIndex+10 from PJNO WHERE (((PJNOIndex+10) Not In (select PJNOIndex from PJNO))and (PJNOIndex+10) between 130000 and 139999) order by PJNOIndex+10", Conn, adOpenKeyset, adOpenStatic  'PJNOIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtPJNOIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub CmdSysAdd3_Click()     '添加号码段140000 To 149999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 PJNOIndex+10 from PJNO WHERE (((PJNOIndex+10) Not In (select PJNOIndex from PJNO))and (PJNOIndex+10) between 140000 and 149999) order by PJNOIndex+10", Conn, adOpenKeyset, adOpenStatic  'PJNOIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtPJNOIndex.Text = Trim(rcds.Fields(0).Value)
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

Private Sub CmdSysAdd4_Click()     '添加号码段150000 To 159999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 PJNOIndex+10 from PJNO WHERE (((PJNOIndex+10) Not In (select PJNOIndex from PJNO))and (PJNOIndex+10) between 150000 and 159999) order by PJNOIndex+10", Conn, adOpenKeyset, adOpenStatic  'PJNOIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtPJNOIndex.Text = Trim(rcds.Fields(0).Value)
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
Private Sub CmdSysAdd5_Click()     '添加号码段160000 To 169999######
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next              '############以下相关改成对应的表字段名字
rcds.Open "select top 1 PJNOIndex+10 from PJNO WHERE (((PJNOIndex+10) Not In (select PJNOIndex from PJNO))and (PJNOIndex+10) between 160000 and 169999) order by PJNOIndex+10", Conn, adOpenKeyset, adOpenStatic  'PJNOIndex+10表示每10位申请一个号，也就是从第2位开始递增1

        If Modify = False Then
            TxtPJNOIndex.Text = Trim(rcds.Fields(0).Value)
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
'ResizeInit Me
TxtApplicant.Text = PDMUserName
CombIDSQ.AddItem ("Open")
CombIDSQ.AddItem ("Close")
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
If Trim(TxtPJNOIndex) = "" Then
    MsgBox "Please input Project Number" + vbCrLf + "请输入Project Number号", vbInformation, "System Info."
    TxtPJNOIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (Len(TxtPJNOIndex) = 6 And IsNumeric(TxtPJNOIndex)) Then
    MsgBox "Project Number is 6 Number, No letter " + vbCrLf + "请输入6位的数字,无字母", vbInformation, "System Info."
    TxtPJNOIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtApplicant) = "" Then
    MsgBox "Please input Leader Name" + vbCrLf + "请输入申请人名", vbInformation, "System Info."
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
     
   With MyPJNO              '已经定义Public MyPJNO As New ClsPJNO, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字
    .PJNOIndex = TxtPJNOIndex.Text
    .Applicant = TxtApplicant.Text
    .Leader = TxtLeader.Text
    .Description = TxtDescription.Text
    .IDSQ = CombIDSQ.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
   
            '判断操作是添加还是修改
       If Modify = False Then         '判断为添加操作
     
           '判断PJNOIndex序号是否已经存在
                If .In_DB(TxtPJNOIndex.Text) = True Then
                   MsgBox "Project number exists, Please re-input" + vbCrLf + "Project号重复，请重新设置", vbInformation, "System Info."
                   TxtPJNOIndex.SetFocus
                   TxtPJNOIndex.SelStart = 0
                   TxtPJNOIndex.SelLength = Len(TxtPJNOIndex)
                   Exit Sub
                Else
                   .Insert                   '添加
                    MsgBox "Succeed to Add" + vbCrLf + "添加成功", vbInformation, "System Info."
                End If
       Else  '判断为修改操作
        .Update (OriPJNOIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "修改成功", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '关闭自身窗口
End Sub


