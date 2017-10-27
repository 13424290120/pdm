VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDM Database - Login Form"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6540
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   6285
      Begin VB.CheckBox LblChangePwdIn 
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4020
         TabIndex        =   9
         Top             =   270
         Width           =   2205
      End
      Begin VB.OptionButton optLive 
         Caption         =   "Live DB"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optTest 
         Caption         =   "Test DB"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1950
         TabIndex        =   7
         Top             =   270
         Width           =   2055
      End
   End
   Begin VB.CommandButton LblCancel 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5010
      TabIndex        =   3
      Top             =   990
      Width           =   1275
   End
   Begin VB.CommandButton LblOK 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   2
      Top             =   360
      Width           =   1275
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2460
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "输入用户名，不超过20个字符"
      Top             =   390
      Width           =   2400
   End
   Begin VB.TextBox TxtPwd 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2460
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "输入密码，不超过20个字符"
      Top             =   990
      Width           =   2400
   End
   Begin VB.Label LblPassWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   825
      MouseIcon       =   "FrmLogin.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   945
      Width           =   1590
   End
   Begin VB.Label LblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      MouseIcon       =   "FrmLogin.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   390
      Width           =   1665
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PasswordKey As String
Public NameKey As String
Private Conn As New ADODB.Connection



Private Sub Form_Load()
    'Load Skin & Format Control
    ''LoadSkin Me
    SkinH_Attach
    If optLive.Value Then
        SetupFile = "Setup.ini"
    Else
        SetupFile = "Test.ini"
    End If
    Server = GetKey(App.Path + "\" + SetupFile, "Server")
    DataBase = GetKey(App.Path + "\" + SetupFile, "DataBase")
    DBUser = GetKey(App.Path + "\" + SetupFile, "UserName")
    Password = GetKey(App.Path + "\" + SetupFile, "Password")
    
    connString = "driver={SQL Server};server=" + Server + ";uid=" + DBUser + ";pwd=" + Password + ";database=" & DataBase  'ERP是数据库名
End Sub

Private Sub LblChangePwdIn_Click()
    '数据有效性检查
    If TxtUser = "" Then
        MsgBox "Please input User Name" + vbCrLf + "请输入用户名", vbInformation, "System Info."
        TxtUser.SetFocus
        Exit Sub
    End If
    
    If TxtPwd = "" Then
        MsgBox "Please input PassWord" + vbCrLf + "请输入密码", vbInformation, "System Info."
        TxtPwd.SetFocus
        Exit Sub
    End If
    
    '判断用户是否存在
    If MyUsers.In_DB(Trim(TxtUser.Text)) = False Then
        MsgBox "User Name Not Exist" + vbCrLf + "用户名不存在", vbInformation, "System Info."
        Try_times = Try_times + 1
        If Try_times >= 3 Then
            MsgBox "You tried three times, failed to logon. System will Exit" + vbCrLf + "您已经三次尝试进入本系统，均不成功，系统将关闭", vbInformation, "System Info."
           End
        Else
            Exit Sub
        End If
    
     End If

'判断密码是否正确
If MyUsers.GetInfo(TxtUser.Text, TxtPwd.Text) = False Then
    MsgBox "Wrong PassWord" + vbCrLf + "密码错误", vbInformation, "System Info."
    Try_times = Try_times + 1
    If Try_times >= 3 Then
           MsgBox "You tried three times, failed to logon. System will Exit" + vbCrLf + "您已经三次尝试进入本系统，均不成功，系统将关闭", vbInformation, "System Info."
         End
    Else
    Exit Sub
    End If
Else
FrmChangePwd.TxtName.Text = Trim(TxtUser.Text)
FrmChangePwd.Show 0
Unload Me
End If
End Sub


Private Sub LblCancel_Click()
    Unload Me
End Sub

Private Sub lblOk_Click()
    On Error Resume Next
    
    Dim LoginRS As New ADODB.Recordset
    Dim SetupFile As String
    '数据有效性检查
    If TxtUser = "" Then
        MsgBox "Please input User Name" + vbCrLf + "请输入用户名", vbInformation, "System Info."
        TxtUser.SetFocus
        Exit Sub
    End If
    
    If TxtPwd = "" Then
        MsgBox "Please input PassWord" + vbCrLf + "请输入密码", vbInformation, "System Info."
        TxtPwd.SetFocus
        Exit Sub
    End If
    

    

    Conn.Open connString
    
    If Err Then
        MsgBox "Connect Database Error, please contact System Admin" & vbCrLf & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf, vbCritical
        Exit Sub
    Else
            
        '判断用户是否存在
        StrSql = "SELECT [Name] FROM Users WHERE [Name]='" + Trim(TxtUser.Text) + "'"
        LoginRS.Open StrSql, Conn, adOpenKeyset, adOpenStatic
        If LoginRS.RecordCount <= 0 Then
            MsgBox "User Name Not Exist" + vbCrLf + "用户名不存在", vbInformation, "System Info."
            Try_times = Try_times + 1
            If Try_times >= 3 Then
                MsgBox "You tried three times, failed to logon. System will Exit" + vbCrLf + "您已经三次尝试进入本系统，均不成功，系统将关闭", vbInformation, "System Info."
               End
            Else
                LoginRS.Close
                Set LoginRS = Nothing
                Conn.Close
                Set Conn = Nothing
                Exit Sub
            End If
        End If
        LoginRS.Close
    
        '判断密码是否正确
        StrSql = "SELECT * FROM Users WHERE [Name]='" + Trim(TxtUser.Text) + "'"
        LoginRS.Open StrSql, Conn, adOpenKeyset, adOpenStatic
        
        If Trim(LoginRS.Fields("Password")) <> TxtPwd.Text Then
            MsgBox "Wrong PassWord" + vbCrLf + "密码错误", vbInformation, "System Info."
            Try_times = Try_times + 1
            If Try_times >= 3 Then
                   MsgBox "You tried three times, failed to logon. System will Exit" + vbCrLf + "您已经三次尝试进入本系统，均不成功，系统将关闭", vbInformation, "System Info."
                 End
            Else
                LoginRS.Close
                Set LoginRS = Nothing
                Conn.Close
                Set Conn = Nothing
                Exit Sub
            End If
        Else
            
            PDMUserName = Trim(TxtUser.Text)   '保存登陆PDM的用户名
            PDMUserGroup = Trim(LoginRS.Fields("UserGroup"))
            '获取用户权限
            Select Case Trim(LoginRS.Fields("GrantGroup"))
                Case "管理组"
                SystemAdmin = "Y"
                Case "开发组"
                EngineeringSys = "Y"
                Case "采购组"
                PurchasingSys = "Y"
                Case "Development" ''for guilao
                EngineeringSys = "Y"
            End Select
        
            '断开与数据库的连接
        
            LoginRS.Close
            Set LoginRS = Nothing
            Conn.Close
            Set Conn = Nothing
            If SystemAdmin = "Y" Or EngineeringSys = "Y" Then
                FrmEngineeringSys.Show 0
            ElseIf PurchasingSys = "Y" Then
                FrmPurchasingSys.Show 0
            End If
        End If
    End If
    Unload Me
End Sub


Private Sub optLive_Click()
    optLive.Value = True
    optTest.Value = False
    
    If optLive.Value Then
        SetupFile = "Setup.ini"
    Else
        SetupFile = "Test.ini"
    End If
    Server = GetKey(App.Path + "\" + SetupFile, "Server")
    DataBase = GetKey(App.Path + "\" + SetupFile, "DataBase")
    DBUser = GetKey(App.Path + "\" + SetupFile, "UserName")
    Password = GetKey(App.Path + "\" + SetupFile, "Password")
    
    connString = "driver={SQL Server};server=" + Server + ";uid=" + DBUser + ";pwd=" + Password + ";database=" & DataBase  'ERP是数据库名
End Sub

Private Sub optTest_Click()
    optLive.Value = False
    optTest.Value = True
    
    If optLive.Value Then
        SetupFile = "Setup.ini"
    Else
        SetupFile = "Test.ini"
    End If
    Server = GetKey(App.Path + "\" + SetupFile, "Server")
    DataBase = GetKey(App.Path + "\" + SetupFile, "DataBase")
    DBUser = GetKey(App.Path + "\" + SetupFile, "UserName")
    Password = GetKey(App.Path + "\" + SetupFile, "Password")
    
    connString = "driver={SQL Server};server=" + Server + ";uid=" + DBUser + ";pwd=" + Password + ";database=" & DataBase  'ERP是数据库名
End Sub

Private Sub TxtPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lblOk_Click
End Sub

Private Sub txtUser_Change()
    TxtPwd.Enabled = True
End Sub

Private Sub TxtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TxtPwd.SetFocus
End Sub
