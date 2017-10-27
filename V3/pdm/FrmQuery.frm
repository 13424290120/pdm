VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmQuery 
   Caption         =   "General Search Window 通用查询窗口"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   876
   StartUpPosition =   2  '屏幕中心
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8745
      Left            =   60
      TabIndex        =   21
      Top             =   1950
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   15425
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdExportExcel 
      Caption         =   "Export / Print Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton CmdAdd 
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
      Height          =   495
      Left            =   10140
      TabIndex        =   18
      Top             =   1380
      Width           =   885
   End
   Begin VB.TextBox TxtQry2 
      Height          =   345
      Left            =   7890
      TabIndex        =   17
      Top             =   480
      Width           =   2000
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12090
      TabIndex        =   14
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11070
      TabIndex        =   13
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   11
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton CmdExecQ 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   10
      Top             =   180
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   60
      TabIndex        =   4
      Top             =   1050
      Width           =   9915
      Begin VB.TextBox txtReason 
         Height          =   405
         Left            =   7350
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ComboBox CmboDate 
         Height          =   345
         ItemData        =   "FrmQuery.frx":08CA
         Left            =   450
         List            =   "FrmQuery.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   285
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   3670017
         CurrentDate     =   39974
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2445
         TabIndex        =   6
         Top             =   285
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         Format          =   3670017
         CurrentDate     =   39974
      End
      Begin VB.CheckBox ChkBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   345
         Width           =   255
      End
      Begin VB.Label lblReason 
         Caption         =   "Reason:"
         Height          =   375
         Left            =   6540
         TabIndex        =   22
         Top             =   390
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4050
         TabIndex        =   9
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1935
         TabIndex        =   8
         Top             =   345
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Condition      设置查询的项目和条件"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9915
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   450
         Width           =   255
      End
      Begin VB.ComboBox CmboEqut2 
         Height          =   345
         Left            =   6810
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   975
      End
      Begin VB.CheckBox ChkBox3 
         Caption         =   "And"
         Height          =   375
         Left            =   6090
         TabIndex        =   15
         Top             =   420
         Width           =   765
      End
      Begin VB.TextBox TxtQry1 
         Height          =   345
         Left            =   4020
         TabIndex        =   3
         Top             =   420
         Width           =   2000
      End
      Begin VB.ComboBox CmboEqut1 
         Height          =   345
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
      Begin VB.ComboBox CmboItem 
         Height          =   345
         Left            =   450
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   2595
      End
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'调用此通用查询窗口需要参数
' 1- QueryTableName通用查询窗口查询运作的表名字 见Form Load段

Option Explicit
Dim QryItem As Long             '定义当前字段数量的循环查询变量
Dim QrysqlStr As String         '定义SQL查询语句字符变量
Dim DtGrdLen As Long            '定义Grid某一格所在的栏号0，1，2...或者对应记录集字段的字段序号
Dim response As Integer         'msgbox返回值判断
Dim Qcnn As New ADODB.Connection   '定义一个ADO连接
Dim QRS As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录

Private Sub Check1_Click()
    If Check1.Value = 0 Then ChkBox2.Value = 0
End Sub

Private Sub ChkBox3_Click()
    If ChkBox3.Value = 1 Then
        CmboEqut1.ListIndex = 2    '如果区段查找标志ChkBox3.Value设置为1这里就只能用>号
        Check1.Value = 1
        CmboEqut1.Enabled = False
    Else
        CmboEqut1.Enabled = True
    End If
End Sub

Private Sub CmboEqut1_Click()
    If ChkBox3.Value = 1 Then
        CmboEqut1.ListIndex = 2    '设置为>符号,只是显示用
    End If
End Sub

Private Sub cmdAdd_Click()
    If QueryTableName = "SglPrt" Then
        FrmSglPrtEdit.Caption = "Add One Single Part Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmSglPrtEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        
        '如果是添加模式要隐藏一些控件
        
        FrmSglPrtEdit.TxtSglPrtVer.Visible = False
        FrmSglPrtEdit.TxtPrtUnit.Visible = False
        FrmSglPrtEdit.TxtIDSO.Visible = False
        FrmSglPrtEdit.TxtNewOldStatus.Visible = False
        FrmSglPrtEdit.TxtOpnDate.Visible = False
        FrmSglPrtEdit.TxtClosDate.Visible = False
        FrmSglPrtEdit.TxtProductLine.Visible = False
        FrmSglPrtEdit.TxtItemType.Visible = False
        FrmSglPrtEdit.TxtLocation.Visible = False
        
        FrmSglPrtEdit.LblOld0.Visible = False
        FrmSglPrtEdit.LblOld1.Visible = False
        FrmSglPrtEdit.LblOld2.Visible = False
        FrmSglPrtEdit.LblOld3.Visible = False
        FrmSglPrtEdit.LblOld4.Visible = False
        FrmSglPrtEdit.LblOld5.Visible = False
        FrmSglPrtEdit.LblOld6.Visible = False
        FrmSglPrtEdit.LblOld7.Visible = False
        FrmSglPrtEdit.LblOld8.Visible = False
        FrmSglPrtEdit.LblReminder.Visible = False
                
        FrmSglPrtEdit.Show 0                                     '##########对应编辑窗口打开
    ElseIf QueryTableName = "RFSRFQ" Then
        FrmRFSRFQEdit.Caption = "Add One RFQ or RFS Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmRFSRFQEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        '如果是添加模式要隐藏一些控件
        FrmRFSRFQEdit.TxtIDSQ.Visible = False
        FrmRFSRFQEdit.TxtOpnDate.Visible = False
        FrmRFSRFQEdit.TxtClosDate.Visible = False
        FrmRFSRFQEdit.LblOld0.Visible = False
        FrmRFSRFQEdit.LblOld1.Visible = False
        FrmRFSRFQEdit.LblOld2.Visible = False
        FrmRFSRFQEdit.LblReminder.Visible = False
        FrmRFSRFQEdit.Show 0                                     '##########对应编辑窗口打开
    ElseIf QueryTableName = "CNCSN" Then
        FrmCNCSNEdit.Caption = "Add One CONCESSION Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmCNCSNEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        
        '如果是添加模式要隐藏一些控件
        FrmCNCSNEdit.TxtCPCNMP.Visible = False
        FrmCNCSNEdit.TxtIDSO.Visible = False
        FrmCNCSNEdit.TxtOpnDate.Visible = False
        FrmCNCSNEdit.TxtClosDate.Visible = False
        
        FrmCNCSNEdit.LblOld0.Visible = False
        FrmCNCSNEdit.LblOld1.Visible = False
        FrmCNCSNEdit.LblOld2.Visible = False
        FrmCNCSNEdit.LblOld3.Visible = False
        FrmCNCSNEdit.LblReminder.Visible = False
                
        FrmCNCSNEdit.Show 0                                     '##########对应编辑窗口打开
    ElseIf QueryTableName = "SER" Then
        FrmSEREdit.Caption = "Add One SER Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmSEREdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        
        '如果是添加模式要隐藏一些控件
        FrmSEREdit.TxtCAorA.Visible = False
        FrmSEREdit.TxtIDSO.Visible = False
        FrmSEREdit.TxtOpnDate.Visible = False
        FrmSEREdit.TxtClosDate.Visible = False
        FrmSEREdit.LblOld0.Visible = False
        FrmSEREdit.LblOld1.Visible = False
        FrmSEREdit.LblOld2.Visible = False
        FrmSEREdit.LblOld3.Visible = False
        FrmSEREdit.LblReminder.Visible = False
        FrmSEREdit.Show 0                                     '##########对应编辑窗口打开
    ElseIf QueryTableName = "PJNO" Then
        FrmPJNOEdit.Caption = "Add One Project Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmPJNOEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        
        '如果是添加模式要隐藏一些控件
        FrmPJNOEdit.TxtIDSQ.Visible = False
        FrmPJNOEdit.TxtOpnDate.Visible = False
        FrmPJNOEdit.TxtClosDate.Visible = False
        FrmPJNOEdit.LblOld0.Visible = False
        FrmPJNOEdit.LblOld1.Visible = False
        FrmPJNOEdit.LblOld2.Visible = False
        FrmPJNOEdit.LblReminder.Visible = False
        FrmPJNOEdit.Show 0                                    '##########对应编辑窗口打开
    ElseIf QueryTableName = "CPCN" Then
        FrmCPCNEdit.Caption = "Add One CP/CN Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmCPCNEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        '如果是添加模式要隐藏一些控件
        FrmCPCNEdit.TxtCPCNMP.Visible = False
        FrmCPCNEdit.TxtIDSO.Visible = False
        FrmCPCNEdit.TxtOpnDate.Visible = False
        FrmCPCNEdit.TxtClosDate.Visible = False
        FrmCPCNEdit.LblOld0.Visible = False
        FrmCPCNEdit.LblOld1.Visible = False
        FrmCPCNEdit.LblOld2.Visible = False
        FrmCPCNEdit.LblOld3.Visible = False
        FrmCPCNEdit.LblReminder.Visible = False
        FrmCPCNEdit.Show 0                                     '##########对应编辑窗口打开
    ElseIf QueryTableName = "FinsGd" Then
        FrmFinsGdEdit.Caption = "Add One Finish Goods Number."     '##########在对应打开窗口中标题要赋值
        '标明操作为添加而非修改
        FrmFinsGdEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值
        
        '如果是添加模式要隐藏一些控件
        FrmFinsGdEdit.TxtIDSO.Visible = False
        FrmFinsGdEdit.TxtOpnDate.Visible = False
        FrmFinsGdEdit.TxtClosDate.Visible = False
        FrmFinsGdEdit.LblReminder.Visible = False
        FrmFinsGdEdit.TxtProductLine.Visible = False
        FrmFinsGdEdit.TxtItemType.Visible = False
        FrmFinsGdEdit.TxtLocation.Visible = False
        
        FrmFinsGdEdit.LblOld0.Visible = False
        FrmFinsGdEdit.LblOld1.Visible = False
        FrmFinsGdEdit.LblOld2.Visible = False
        FrmFinsGdEdit.LblOld3.Visible = False
        FrmFinsGdEdit.LblOld4.Visible = False
        FrmFinsGdEdit.LblOld5.Visible = False
        FrmFinsGdEdit.CmdSysDistrb.Enabled = True
        FrmFinsGdEdit.Show 0                                     '##########对应编辑窗口打开
    End If
    CmdExecQ_Click
End Sub

'删除单元行(记录)
Private Sub CmdDel_Click()
    If QRS.EOF Then
        MsgBox "No Chosed Record!"
        Exit Sub
    Else
        Dim TempID As String

        '@@@@@@@@@@判断是否是管理员用户，否则不能删除
        If SystemAdmin <> "Y" Then
          MsgBox "you are not administrator, No right to delete", vbInformation, "System Info."
          Exit Sub
        End If
        '@@@@@@@@@@判断是否是管理员用户，否则不能删除
        
        TempID = QRS.Fields(0)                      '##########TempSglPrtID更换成对应表中Key字段名
        '弹出删除确认对话框 Str是数字变字符串的函数,这里如果不用Str会出错
        If MsgBox("Confirm to delete" + CStr(QRS.Fields(0)) + "?" + vbCrLf + "是否删除" + CStr(QRS.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete 确认删除") = vbYes Then
            If QueryTableName = "SglPrt" Then
                '调用类中Delete函数删除SglPrt信息
                MySglPrt.Delete (TempID)
            ElseIf QueryTableName = "RFSRFQ" Then
                  '调用类中Delete函数删除RFQ/RFS信息
                MyRFSRFQ.Delete (TempID)
            ElseIf QueryTableName = "CNCSN" Then
                MyCNCSN.Delete (TempID)
            ElseIf QueryTableName = "PJNO" Then
                MyPJNO.Delete (TempID)
            ElseIf QueryTableName = "CPCN" Then
                MyCPCN.Delete (TempID)
            ElseIf QueryTableName = "FinsGd" Then
                MyFinsGd.Delete (TempID)
            ElseIf QueryTableName = "SER" Then
                MySER.Delete (TempID)
            End If
            MsgBox "Succeed to delete, 删除成功", vbInformation, "System Info."
        End If
    End If
    CmdExecQ_Click
End Sub



'修改单元格(记录)
Private Sub cmdModify_Click()
    If QRS.EOF Then
        MsgBox "No Chosed Record!"
    Else
        If PDMUserName <> Trim(QRS.Fields("Applicant")) And SystemAdmin <> "Y" Then MsgBox "No right to modify it.", vbInformation: Exit Sub
        If QueryTableName = "SglPrt" Then
            '保存待修改记录的原始ID
            FrmSglPrtEdit.OriSglPrtIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))       '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmSglPrtEdit.TxtSglPrtIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))       '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtSglPrtVer = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))       '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtPrtUnit = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))       '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtApplicant = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))             '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtProductLine = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtIDSO = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))                   '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtNewOldStatus = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))                   '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))                '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtClosDate = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))              '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(12)), "", FormatNumber6(Trim(QRS.Fields(11))))               '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtPjtName = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))              '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtItemType = IIf(QRS.Fields(13) = Null, "", Trim(QRS.Fields(13)))              '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtLocation = IIf(QRS.Fields(14) = Null, "", Trim(QRS.Fields(14)))             '##########对应编辑窗口控件赋值
            FrmSglPrtEdit.TxtCommtNote = IIf(QRS.Fields(15) = Null, "", Trim(QRS.Fields(15)))              '##########对应编辑窗口控件赋值
            
            FrmSglPrtEdit.TxtSglPrtIndex.Locked = True   '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmSglPrtEdit.TxtApplicant.Locked = True       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmSglPrtEdit.Caption = "Modify One Single Part Number."                                  '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmSglPrtEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            FrmSglPrtEdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "RFSRFQ" Then
            '保存待修改记录的原始ID
            FrmRFSRFQEdit.OriRFSRFQIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmRFSRFQEdit.TxtRFSRFQIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))             '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtLeader = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                 '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtIDSQ = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                   '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########对应编辑窗口控件赋值
            FrmRFSRFQEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########对应编辑窗口控件赋值
        
            FrmRFSRFQEdit.TxtRFSRFQIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmRFSRFQEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmRFSRFQEdit.Caption = "Modify one RFQ/RFS Number."                                           '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmRFSRFQEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            FrmRFSRFQEdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "CNCSN" Then
            '保存待修改记录的原始ID
            FrmCNCSNEdit.OriCNCSNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmCNCSNEdit.TxtCNCSNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))             '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtCPCNMP = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))         '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                  '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))                '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))             '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))                '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))               '##########对应编辑窗口控件赋值
            FrmCNCSNEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))            '##########对应编辑窗口控件赋值
            
            FrmCNCSNEdit.TxtCNCSNIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmCNCSNEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmCNCSNEdit.Caption = "Modify One CONCESSION Number."                                  '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmCNCSNEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            
            FrmCNCSNEdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "PJNO" Then
            '保存待修改记录的原始ID
            FrmPJNOEdit.OriPJNOIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))        '##########对应编辑窗口变量赋值
            '把待修改信息添加到编辑窗口
            FrmPJNOEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))            '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtLeader = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtIDSQ = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                 '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########对应编辑窗口控件赋值
            FrmPJNOEdit.TxtPJNOIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmPJNOEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmPJNOEdit.Caption = "Modify One Project Number."                                           '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmPJNOEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            FrmPJNOEdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "SER" Then
            '保存待修改记录的原始ID
            FrmSEREdit.OriSERIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmSEREdit.TxtSERIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))        '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))           '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtCAorA = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))          '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                  '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))            '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))             '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))              '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))              '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))             '##########对应编辑窗口控件赋值
            FrmSEREdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))              '##########对应编辑窗口控件赋值
            
            FrmSEREdit.TxtSERIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmSEREdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmSEREdit.Caption = "Modify One SER Number."                                  '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmSEREdit.Modify = True                                     '##########对应编辑窗口变量赋值
            
            FrmSEREdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "CPCN" Then
            '保存待修改记录的原始ID
            FrmCPCNEdit.OriCPCNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))         '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmCPCNEdit.TxtCPCNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))           '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtCPCNMP = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))               '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))           '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))              '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))             '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))             '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))             '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))              '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))              '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))              '##########对应编辑窗口控件赋值
            FrmCPCNEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))             '##########对应编辑窗口控件赋值
            FrmCPCNEdit.txtReason = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))             '##########对应编辑窗口控件赋值
            
            FrmCPCNEdit.TxtCPCNIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmCPCNEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmCPCNEdit.Caption = "Modify One CP/CN Number."                                  '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmCPCNEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            FrmCPCNEdit.Show 0                                            '##########对应编辑窗口打开
        ElseIf QueryTableName = "FinsGd" Then
            '保存待修改记录的原始ID
            FrmFinsGdEdit.OriFinsGdIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))         '##########对应编辑窗口变量赋值
            
            '把待修改信息添加到编辑窗口
            FrmFinsGdEdit.TxtFinsGdIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))            '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtProductLine = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))        '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtIDSO = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                  '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))               '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtClosDate = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))               '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))             '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtPjtName = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))               '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtItemType = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))               '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtLocation = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))               '##########对应编辑窗口控件赋值
            FrmFinsGdEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))                '##########对应编辑窗口控件赋值
            
            FrmFinsGdEdit.TxtFinsGdIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
            FrmFinsGdEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
            FrmFinsGdEdit.Caption = "Modify One Finish Goods Number."                                  '##########对应编辑窗口标题
            '标明操作为修改操作
            FrmFinsGdEdit.Modify = True                                     '##########对应编辑窗口变量赋值
            FrmFinsGdEdit.CmdSysDistrb.Enabled = False
            FrmFinsGdEdit.Show 0                                            '##########对应编辑窗口打开

        Else
        End If
    End If
    CmdExecQ_Click
End Sub
'本通用查询窗口退出
Private Sub CmdExit_Click()
    On Error Resume Next
    If QRS.State = adStateOpen Then QRS.Close
    Set QRS = Nothing
    If Qcnn.State = adStateOpen Then Qcnn.Close
    Set Qcnn = Nothing
    Unload Me
    FromForm.Show 0
End Sub

Private Sub CmdExportExcel_Click()
    On Error Resume Next

    Dim i As Integer
    Dim sHeader As String
    Set xlApp = CreateObject("Excel.Application")   '创建Excel文件
    Set xlApp = New excel.Application
    
'        '解决出现部件挂起提示
'    xlApp.OleRequestPendingTimeout = 10000   '10000毫秒后出现忙对话框
'    xlApp.OleServerBusyTimeout = 1000     '请求超时1秒
'    xlApp.OleServerBusyRaiseError = True '不显示忙对话框
    
    
    xlApp.SheetsInNewWorkbook = 1                   '将新建的工作薄数量设为1
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)              '第1张工作表
    If QueryTableName = "SglPrt" Then
        sHeader = "Single Part Number"
    ElseIf QueryTableName = "RFSRFQ" Then
        sHeader = "RFS/RFQ Number"
    ElseIf QueryTableName = "CNCSN" Then
        sHeader = "CONCESSION Number"
    ElseIf QueryTableName = "PJNO" Then
        sHeader = "Project Number"
    ElseIf QueryTableName = "SER" Then
        sHeader = "SER Number"
    ElseIf QueryTableName = "CPCN" Then
        sHeader = "CPCN Number"
    ElseIf QueryTableName = "FinsGd" Then
        sHeader = "Finish Goods Number"
    End If
    xlSheet.Cells(1, 1) = sHeader
    For i = 0 To DataGrid1.Columns.count - 1
        xlSheet.Cells(3, i + 1) = DataGrid1.Columns(i).Caption
    Next i
    
    xlSheet.Cells(2, i - 3) = "Table Maker:": xlSheet.Cells(2, i - 2) = PDMUserName
    xlSheet.Cells(2, i - 1) = "Print Date:": xlSheet.Cells(2, i) = Now()
        
    xlSheet.Cells(4, 1).CopyFromRecordset Qcnn.Execute(QrysqlStr)       '此行是粘贴数据
    xlSheet.Columns("K").NumberFormat = "############"

    xlApp.ActiveWorkbook.Close True     '关闭工作簿并保存
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If QRS.State = adStateOpen Then QRS.Close
    Set QRS = Nothing
    If Qcnn.State = adStateOpen Then Qcnn.Close
    Set Qcnn = Nothing
    Unload Me
    FromForm.Show 0
End Sub


Private Sub DataGrid1_Error(ByVal dataerror As Integer, response As Integer)
response = 0
MsgBox "check the Input Data type or Length", vbInformation, "Error Info!"
End Sub
Private Sub DataGrid1_colEdit(ByVal colindex As Integer)

On Error Resume Next
DataGrid1.SetFocus
DataGrid1.SelStart = 0
DataGrid1.SelLength = Len(QRS.Fields(DtGrdLen))   'DataGrid单元格对应记录集字段的字段序号的长度
 
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)  '注意括号内参数不能少,少了编译出错
    
    Select Case DataGrid1.Col
        Case 0                             '0列设置为不可更新  如果加1, 2, 3 不允许更新的列还可以设置1, 2, 3...
            DataGrid1.AllowUpdate = False
        Case Else
            DataGrid1.AllowUpdate = True
            
                 If SystemAdmin <> "Y" Then             '多一次判断，如果不是sa用户的话还是不能修改
                 DataGrid1.AllowUpdate = False
                 End If

            DtGrdLen = DataGrid1.Col      'DataGrid1.Col是点击DataGrid单元格返回该单元格所在的栏号0，1，2...
    ' TxtTest.Text = DataGrid1.Col  测试用语句调试使用时在窗口中设置为可见visible = true
    End Select
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me

If QueryTableName = "SglPrt" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "RFSRFQ" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "CNCSN" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "PJNO" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "CPCN" Then
    CmdExportExcel.Visible = True
    lblReason.Visible = True
    txtReason.Visible = True
ElseIf QueryTableName = "SER" Then
    CmdExportExcel.Visible = True
ElseIf QueryTableName = "FinsGd" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "BOMSubmitApprove" Then
    CmdExportExcel.Visible = False
    cmdAdd.Visible = False
    cmdModify.Visible = False
    CmdDel.Visible = False
End If

Qcnn.ConnectionString = connString
Qcnn.Open


'SER要求加载已经申请的数据
If QueryTableName = "SER" Then
    QRS.Open "select * from " + QueryTableName + " where applicant<>''", Qcnn, adOpenKeyset, adOpenStatic  ' 传递表的名字 Public QueryTableName As String
Else
    QRS.Open "select * from " + QueryTableName + " where 1>2", Qcnn, adOpenKeyset, adOpenStatic  ' 传递表的名字 Public QueryTableName As String
End If
Set DataGrid1.DataSource = QRS
DataGrid1.Columns(0).Width = 150  '第一列在显示是有宽度不够的问题，所以单独设置列宽
'Call AutoFitWidth(DataGrid1)

          For QryItem = 0 To QRS.Fields.count - 1          '查询项目值初始化
          
          Check1.Value = 1                                '第一个查询框默认可以用
          
             If InStr(QRS.Fields(QryItem).Name, "Date") <> 0 Then  'instr指定一字符串在另一字符串中最先出现的位置
               ChkBox2.Value = 0                                '第二个查询框也可以用
               CmboDate.AddItem (QRS.Fields(QryItem).Name)
               CmboDate.ListIndex = 0                           '第一项目默认选中
           GoTo NextLine                                    '如果有Date字段值则加入到第2个查询框
             End If
        
          CmboItem.AddItem (QRS.Fields(QryItem).Name)
NextLine:
          Next
          CmboItem.ListIndex = 0                           '第一项目默认选中
          
          CmboEqut1.AddItem ("=")                           '查询等式值初始化
          CmboEqut1.AddItem ("like")
          CmboEqut1.AddItem (">")
          CmboEqut1.AddItem ("<")
          
          CmboEqut1.ListIndex = 0                           '第一项目默认选中
          
          CmboEqut2.AddItem ("<")
          CmboEqut2.ListIndex = 0                           '第一项目默认选中
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub


Private Sub CmdExecQ_Click()
On Error Resume Next
'SQL查询语句字符串格式  SELECT   Field_1   FROM   Table_A   WHERE Field_1 = 'DDD'  ORDER  BY  Field_1
'例子:  SqlStmt = "SELECT Glue12NC FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"

QrysqlStr = "Select * from" + " "                                     '需要输出结果的查询字段
QrysqlStr = QrysqlStr + QueryTableName + " Where 1=1"                          '输入查询的表名 传递表的名字 Public QueryTableName As String

If Check1.Value = 1 Then
'    For QryItem = 0 To Qrs.Fields.Count - 1
'        '如果第0字段的类型是数字(12NC)并且第0字段的字段名在Combo中选中要查询
'        If Qrs.Fields(QryItem).Name = CmboItem.Text Then    'adBigInt(值为20)8字节带符号整形,adUnsignedBigInt(值为21)8字节不带符号整形
'           If (Qrs.Fields(QryItem).Type = 20) And (Not IsNumeric(TxtQry1.Text)) Then   '申请待查询的数据段恰好不是数字(12NC)
'              MsgBox " Input query type is not matching, check if it should be Number", vbInformation, "Error Info!"
'              Exit Sub
'           Else
                If CmboEqut1.Text = "like" Then                   '按照SQL语句规则，Like匹配查询的字符串要加%号，否则和=符号一样
                    QrysqlStr = QrysqlStr + " And " & CmboItem.Text & " like '%" + Trim(TxtQry1.Text) + "%'"
                Else
                    QrysqlStr = QrysqlStr + " And " & CmboItem.Text & CmboEqut1.Text & "'" & Trim(TxtQry1.Text) & "'"
                End If
'           End If
'        End If
'    Next
    If ChkBox3.Value = 1 Then
       '字符串定义格式：and XXDate Between #2009-01-20# and #2009-05-15#  注意#是用于Access数据库的
       'Select * from GlueSupplier Where (Glue12NC Between XXXXXXXA and XXXXXXXB)  And (RdDate Between '2007-04-10' and '2008-11-10')Order By Glue12NC
        QrysqlStr = QrysqlStr + " And " & CmboItem.Text & CmboEqut2.Text & "'" & Trim(TxtQry2.Text) & "'"
       'TxtTest.Text = QrysqlStr  '测试用语句调试使用时在窗口中设置为可见visible = true
    End If
End If


' Format(DTPicker1.Value, "YYYY/MM/DD hh:mm:ss") 时间数格式化成字符方法
If ChkBox2.Value = 1 Then                                               '第二个查询框关于时间默认有待查询项目
        QrysqlStr = QrysqlStr + " And (" + CmboDate.Text + " Between " + "'" + Format(DTPicker1.Value, "YYYY/MM/DD") + "'" + " and " + "'" + Format(DTPicker2.Value, "YYYY/MM/DD") + "'" + ")" '加入日期的查询字符串
       '上面的Format(DTPicker1.Value, "YYYY/MM/DD")左右需要 '符号隔开，否则出错。 注意 '是用于SQL数据库的
End If

If QueryTableName = "CPCN" Then QrysqlStr = QrysqlStr + " And Reason like '%" & txtReason.Text & "%'"

If Check1.Value = 1 Then QrysqlStr = QrysqlStr + " Order By" + " " + CmboItem.Text
Debug.Print QrysqlStr
Set QRS = Nothing  '原记录中的内容需要先清空才能写
QRS.Open QrysqlStr, Qcnn, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = QRS
DataGrid1.Columns(1).Width = 150  '第一列在显示是有宽度不够的问题，所以单独设置列宽

End Sub

Private Sub TxtQry1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then CmdExecQ_Click
End Sub

Public Sub AutoFitWidth(ByRef dg As DataGrid)

Dim tr As ADODB.Recordset
Dim r As ADODB.Recordset
Set tr = dg.DataSource
If tr Is Nothing Then Exit Sub
If tr.State = 0 Then Exit Sub
If tr.RecordCount = 0 Then Exit Sub

Set r = tr.Clone

Dim m
Dim Width

For m = 0 To dg.Columns.count - 1
    Width = Len(dg.Columns(m).Caption)
    r.MoveFirst
   While Not r.EOF
      If Len(r(m)) > Width Then
          Width = Len(r(m))
      End If
      r.MoveNext
   Wend
   dg.Columns(m).Width = Width * 229
Next m

End Sub
