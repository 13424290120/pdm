VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.OCX"
Begin VB.Form FrmSysIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统初始化设置"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5145
   Icon            =   "FrmSysIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5145
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   720
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.TextBox TxtServer 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "SQL Server服务器名"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox TxtUser 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "服务器用户名"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "服务器密码"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3480
      Picture         =   "FrmSysIn.frx":0CCA
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label LblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "返 回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      MouseIcon       =   "FrmSysIn.frx":10E6
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LblRead 
      BackStyle       =   0  'Transparent
      Caption         =   "读 取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      MouseIcon       =   "FrmSysIn.frx":13F0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   360
      Picture         =   "FrmSysIn.frx":16FA
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label LblSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设 定"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      MouseIcon       =   "FrmSysIn.frx":1B16
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2520
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   1800
      Picture         =   "FrmSysIn.frx":1E20
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "服务器名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "用 户 名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "密    码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSysIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'判断文件是否存在

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
    FileExist = (Dir(Fname) <> "")
End Function


Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
    Dim File As Long
    '分配文件句柄
    File = FreeFile
    '如果文件不存在则创建一个默认的Setup.ini文件
    If FileExist(Tmp_File) = False Then
        GetKey = ""
        Call WritePrivateProfileString("Setup Information", "Server Name ", "", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "UserName ", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "PasswordName ", " ", App.Path + "\Setup.ini")
        Exit Function
    End If
    '读取数据项值
    Open Tmp_File For Input As File
    Do While Not EOF(1)
        Line Input #File, buffer
        If Left(buffer, Len(Tmp_Key)) = Tmp_Key Then
            pos = InStr(buffer, "=")
            GetKey = Trim(Mid(buffer, pos + 1))
        End If
    Loop
    Close File
End Function





Private Sub LblBack_Click()
    Unload Me
End Sub

Private Sub LblRead_Click()
    '从Setup.ini中读取服务器的名字
    ServerName = GetKey(App.Path + "\Setup.ini", "Server")
    UserName = GetKey(App.Path + "\Setup.ini", "UserName")
    PasswordName = GetKey(App.Path + "\Setup.ini", "Password")
    
    If ServerName = "" Then
        MsgBox "初始信息没有设置，请填入初始信息"
        LblRead.Enabled = False
        LblSet.Enabled = True
    Else
        TxtServer.Text = ServerName
        TxtUser.Text = UserName
        TxtPassword.Text = PasswordName
        
        
        LblSet.Enabled = True
        LblRead.Enabled = False
    End If
End Sub


Private Sub Test()
    Adodc1.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(User) + ";pwd=" + Trim(Password) + ";database=ERP"
    Adodc1.RecordSource = "select * from Users"
    Set Text1.DataSource = Adodc1
    Text1.DataField = "Name"
End Sub


Private Sub LblSet_Click()
    If TxtServer.Text = "" Then
        MsgBox "请填入服务器名称"
        TxtServer.SetFocus
        Exit Sub
    End If
    
    Call WritePrivateProfileString("Setup Information", "Server Name ", TxtServer.Text, App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "UserName ", TxtUser.Text, App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "PasswordName ", TxtPassword.Text, App.Path + "\Setup.ini")
    
    Server = TxtServer.Text
    User = TxtUser.Text
    Password = TxtPassword.Text
    
    DataEnvironmentItem.Item.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(User) + ";pwd=" + Trim(Password) + ";database=ERP"
    
    Test
    
    If Text1.Text <> "" Then
        
        MsgBox "设置成功"
        FrmMan.LogIn.Enabled = True
        FrmMan.Toolbar1.Buttons.Item("Ini").Enabled = False
    Else
        MsgBox "设置失败，设置参数与SQL服务器参数不符，请重新设置"
        TxtServer.Text = ""
        TxtUser.Text = ""
        TxtPassword.Text = ""
        TxtServer.SetFocus
        Exit Sub
    End If
    Unload Me
    
End Sub
