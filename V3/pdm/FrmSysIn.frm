VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.OCX"
Begin VB.Form FrmSysIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ��ʼ������"
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
   StartUpPosition =   2  '��Ļ����
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
      ToolTipText     =   "SQL Server��������"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox TxtUser 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "�������û���"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "����������"
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
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�� ȡ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�� ��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�� �� ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
'�ж��ļ��Ƿ����

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
    FileExist = (Dir(Fname) <> "")
End Function


Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
    Dim File As Long
    '�����ļ����
    File = FreeFile
    '����ļ��������򴴽�һ��Ĭ�ϵ�Setup.ini�ļ�
    If FileExist(Tmp_File) = False Then
        GetKey = ""
        Call WritePrivateProfileString("Setup Information", "Server Name ", "", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "UserName ", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "PasswordName ", " ", App.Path + "\Setup.ini")
        Exit Function
    End If
    '��ȡ������ֵ
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
    '��Setup.ini�ж�ȡ������������
    ServerName = GetKey(App.Path + "\Setup.ini", "Server")
    UserName = GetKey(App.Path + "\Setup.ini", "UserName")
    PasswordName = GetKey(App.Path + "\Setup.ini", "Password")
    
    If ServerName = "" Then
        MsgBox "��ʼ��Ϣû�����ã��������ʼ��Ϣ"
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
        MsgBox "���������������"
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
        
        MsgBox "���óɹ�"
        FrmMan.LogIn.Enabled = True
        FrmMan.Toolbar1.Buttons.Item("Ini").Enabled = False
    Else
        MsgBox "����ʧ�ܣ����ò�����SQL��������������������������"
        TxtServer.Text = ""
        TxtUser.Text = ""
        TxtPassword.Text = ""
        TxtServer.SetFocus
        Exit Sub
    End If
    Unload Me
    
End Sub
