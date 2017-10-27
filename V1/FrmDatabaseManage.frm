VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDatabaseManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBase Form Administrate 数据库表管理"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDatabaseManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "实用数据库表直接SQL语句查询"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   75
      TabIndex        =   1
      Top             =   1860
      Width           =   11895
      Begin VB.TextBox TxtSQL 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1770
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   450
         Width           =   9855
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   2205
         Top             =   4530
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2160
         Left            =   120
         TabIndex        =   2
         Top             =   2250
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   3810
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   165
         Top             =   4530
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label LblSQL 
         Caption         =   $"FrmDatabaseManage.frx":08CA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   270
         TabIndex        =   6
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label LblOK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search 查 询"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6045
         MouseIcon       =   "FrmDatabaseManage.frx":08E1
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1860
         Width           =   1440
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   5445
         Picture         =   "FrmDatabaseManage.frx":0BEB
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label LblCancel 
         BackStyle       =   0  'Transparent
         Caption         =   "Return 返 回"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8775
         MouseIcon       =   "FrmDatabaseManage.frx":1007
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   4770
         Width           =   1530
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   8235
         Picture         =   "FrmDatabaseManage.frx":1311
         Top             =   4755
         Width           =   300
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " 点击这里开始备份/恢复数据库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3000
      TabIndex        =   0
      Top             =   615
      Width           =   6495
   End
End
Attribute VB_Name = "FrmDatabaseManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
Dim I As Integer
Dim sqlApp As New SQLDMO.Application   '定义一个SQL应用对象sqlApp
Dim ServerName As SQLDMO.NameList

    Label1.Caption = "开始检测网络上数据库服务器"
    MousePointer = vbHourglass                        '鼠标设置成沙漏
    '遍历网内所有可用的SQL SERVER 服务器
    Set ServerName = sqlApp.ListAvailableSQLServers        '前面Dim ServerName 没定义类型,这里自动分配成集合
    
    For I = 1 To ServerName.Count
        FrmServerBkup.CobServer.AddItem (ServerName.Item(I))  'CobServer是下一个打开窗FrmServerBkup中的一个组合框控件
    Next
    
    Call FrmServerBkup.LocalInfo     ' LocalInfo是下一个打开窗FrmServerBkup中的子过程。取得本机名称,和返回给定机器名的Ip地址
    MousePointer = vbDefault
    
    'FrmServerBkup.Show 1          '打开备份数据库主窗口
    
    If StrComp("(local)", Trim(FrmServerBkup.CobServer.Text), 1) = 0 Then    'StrComp 为字符串比较的结果。string1 小于 string2返回-1,等于返回0,大于返回1.   括号中最后参数1 执行一个按照原文的比较。
        FrmServerBkup.LabServerName.Caption = FrmServerBkup.LabComputer.Caption
        FrmServerBkup.LabServerIP.Caption = FrmServerBkup.LabIp.Caption
    End If
     MousePointer = vbDefault
End Sub

Private Sub lblOk_Click()
If TxtSQL.Text <> "" Then  '如果SQL语句框内不为空
    If LblOK.Caption = "Search 查 询" Then  '如果LblOK的Caption是Search查询
    '用Adodc1连接数据库进行查询
    Adodc1.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(DBUser) + ";pwd=" + Trim(Password) + ";database=" + Trim(DataBase) + ""
    Adodc1.RecordSource = TxtSQL.Text
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    '文本输入框不可用
    TxtSQL.Enabled = False
    '如果LblOK的Caption变为重新查询
    LblOK.Caption = "Re-Search 重新查询"
    Else     'LblOK的Caption是重新查询
    '把DataGrid与Adodc2相连
     Set DataGrid1.DataSource = Adodc2 '将DataGrid1和没数据库的Adodc2相连等于是清空
     TxtSQL.Text = ""
     LblOK.Caption = "查 询"
     TxtSQL.Enabled = True
    End If
Else
MsgBox "Please Input SQL query" + vbCrLf + "请输入查询语句" + vbCrLf + "Example: select * from SglPrt", vbInformation, "System Info."
End If
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
