VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUserManage 
   Caption         =   "User Administrate 系统用户管理"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUserManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11880
   StartUpPosition =   2  '屏幕中心
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmUserManage.frx":08CA
      Height          =   5685
      Left            =   180
      TabIndex        =   5
      Top             =   900
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   10028
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "UserName"
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
         DataField       =   "UserTitle"
         Caption         =   "Title"
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
      BeginProperty Column02 
         DataField       =   "GrantGroup"
         Caption         =   "Grant"
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
      BeginProperty Column03 
         DataField       =   "UserGroup"
         Caption         =   "Group"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   6720
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.Image Image4 
      Height          =   300
      Left            =   8820
      Picture         =   "FrmUserManage.frx":08DF
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label LblBack 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUserManage.frx":0CFB
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9420
      MouseIcon       =   "FrmUserManage.frx":0D0B
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6840
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   7140
      Picture         =   "FrmUserManage.frx":1015
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUserManage.frx":1431
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7740
      MouseIcon       =   "FrmUserManage.frx":1441
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6840
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   5160
      Picture         =   "FrmUserManage.frx":174B
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label LblModify 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUserManage.frx":1B67
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5760
      MouseIcon       =   "FrmUserManage.frx":1B77
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3600
      Picture         =   "FrmUserManage.frx":1E81
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label LblAdd 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUserManage.frx":229D
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4200
      MouseIcon       =   "FrmUserManage.frx":22AA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Administrate 系统用户管理"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   0
      Top             =   240
      Width           =   5475
   End
End
Attribute VB_Name = "FrmUserManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim QConn As New ADODB.Connection   '定义一个ADO连接
Dim QRS As New ADODB.Recordset

Private Sub Form_Load()
    'Load Skin & Format Control
    'LoadSkin Me
    
    If QConn.State = adStateOpen Then QConn.Close
    QConn.Open connString
    
    '查询Users表中全部信息
    QRS.Open "select [name],[password],[usergroup],[usertitle],[grantgroup] from [Users]", QConn, adOpenKeyset, adOpenStatic
    '用DataGrid显示查询信息
    Set DataGrid1.DataSource = QRS
End Sub

Private Sub Form_Resize()
 '确保窗体改变时控件随之改变
    Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    FrmEngineeringSys.Show 0
End Sub

Private Sub LblAdd_Click()
'标明操作为添加而非修改，FrmUsersEdit窗体中有Public Modify As Boolean
FrmUsersEdit.Modify = False '设置窗体FrmUsersEdit一个常量Modify为False表明要进行的操作不是修改
'显示用户信息编辑窗口
FrmUsersEdit.Show 1
'刷新系统用户管理界面
Form_Load '参见下面的Refresh_Users子过程
End Sub

Private Sub LblBack_Click()
Unload Me
FrmEngineeringSys.Show 0
End Sub


Private Sub LblDelete_Click()
Dim TempName As String
'保存待删除的用户名
  TempName = Trim(QRS.Fields(0))
'弹出删除确认对话框
  If MsgBox("Delete A User " + Trim(QRS.Fields(0)) + "?" + vbCrLf + "是否删除用户" + Trim(QRS.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirmation 确认") = vbYes Then
'调用Users类中Delete函数删除用户
    MyUsers.Delete (TempName) '使用类模块ClsUsers对象MyUsers中的Delete函数删除一个用户
    MsgBox "Successfully Delete" + vbCrLf + "删除成功"
  End If
Form_Load
End Sub

Private Sub LblModify_Click()
'标明操作为修改操作

FrmUsersEdit.Modify = True
'保存修改前的用户名
FrmUsersEdit.OriName = Trim(QRS.Fields(0)) 'FrmUsersEdit窗体中有Public OriName As String

'填充用户信息编辑窗口的各项信息
FrmUsersEdit.TxtName = Trim(QRS.Fields(0))
FrmUsersEdit.TxtName.Enabled = False
FrmUsersEdit.TxtPassword = Trim(QRS.Fields(1))
FrmUsersEdit.TxtPassword2 = Trim(QRS.Fields(1))
FrmUsersEdit.cmbGroup = Trim(QRS.Fields(4))
FrmUsersEdit.txtTitle = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))
FrmUsersEdit.txtUserGroup = Trim(QRS.Fields(2))
'显示用户信息编辑窗口
FrmUsersEdit.Show 1
'刷新系统用户管理界面
Form_Load
End Sub
