VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmSysIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Initialize ϵͳ��ʼ������"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSysIni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   780
      Top             =   300
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "SQL Server��������"
      Top             =   1020
      Width           =   2775
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "�������û���"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "����������"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4020
      Picture         =   "FrmSysIni.frx":08CA
      Top             =   3660
      Width           =   300
   End
   Begin VB.Label LblBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmSysIni.frx":0CE6
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4560
      MouseIcon       =   "FrmSysIni.frx":0CF7
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3660
      Width           =   855
   End
   Begin VB.Label LblRead 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmSysIni.frx":1001
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1140
      MouseIcon       =   "FrmSysIni.frx":1013
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3660
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   600
      Picture         =   "FrmSysIni.frx":131D
      Top             =   3660
      Width           =   300
   End
   Begin VB.Label LblSet 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   $"FrmSysIni.frx":1739
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
      Height          =   480
      Left            =   2880
      MouseIcon       =   "FrmSysIni.frx":174C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3660
      Width           =   945
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2400
      Picture         =   "FrmSysIni.frx":1A56
      Top             =   3660
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmSysIni.frx":1E72
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmSysIni.frx":1E8A
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmSysIni.frx":1EA0
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "FrmSysIni"
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
    Call WritePrivateProfileString("Setup Information", "Server", "", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "DataBase", "", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "UserName", " ", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "Password", " ", App.Path + "\Setup.ini")
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
  Server = GetKey(App.Path + "\Setup.ini", "Server")
  DataBase = GetKey(App.Path + "\Setup.ini", "DataBase")
  UserName = GetKey(App.Path + "\Setup.ini", "UserName")
  Password = GetKey(App.Path + "\Setup.ini", "Password")
  
  If Server = "" Then     '������ܶ�����������
    MsgBox "No Server please input" + vbCrLf + "��ʼ��Ϣû�����ã��������ʼ��Ϣ", vbInformation, "System Info."
  Else
  '�ܶ�ȡ����������
    TxtServer.Text = Server
    TxtUser.Text = UserName
    TxtPassword.Text = Password
  End If
    '�趨��ť����
    LblSet.Enabled = True
    '��ȡ��ť������
    LblRead.Enabled = False
End Sub
Private Sub Test()
'��Adodc1�����õĲ����������ݿ�
Adodc1.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(User) + ";pwd=" + Trim(Password) + ";database=" + Trim(DataBase) + ""
'ѡ�����ݿ��е�Users���ݱ�
Adodc1.RecordSource = "select * from Users"
'��Text1��ʾ��¼
Set Text1.DataSource = Adodc1
    Text1.DataField = "Name"
End Sub
Private Sub LblSet_Click()
Screen.MousePointer = vbHourglass   '����ʱ��ϳ�����Ҫ�������״̬
If TxtServer.Text = "" Then
 MsgBox "please input Server Name " + vbCrLf + "���������������", vbInformation, "System Info."
 TxtServer.SetFocus
 Exit Sub
End If

'����ini��д���趨��Ϣ
Call WritePrivateProfileString("Setup Information", "Server", TxtServer.Text, App.Path + "\Setup.ini")
Call WritePrivateProfileString("Setup Information", "UserName ", TxtUser.Text, App.Path + "\Setup.ini")
Call WritePrivateProfileString("Setup Information", "Password ", TxtPassword.Text, App.Path + "\Setup.ini")

'Ϊϵͳ��������������ֵ
Server = TxtServer.Text
User = TxtUser.Text
Password = TxtPassword.Text

Test
Screen.MousePointer = vbDefault     '�ָ����״̬

'�жϲ�����ȷ��
    If Text1.Text <> "" Then  '���óɹ�
        MsgBox "Ok to Save into INI" + vbCrLf + "���óɹ�,OK����", vbInformation, "System Info."
        '�������ϵ��������ܿ���
        FrmMan.LogIn.Enabled = True
        FrmMan.Toolbar1.Buttons.Item("Ini").Enabled = False
    Else                      '����ʧ��
        MsgBox "Fail to save into INI, need to reset" + vbCrLf + "����ʧ�ܣ����ò�����SQL��������������������������", vbInformation, "System Info."
        TxtServer.Text = ""
        TxtUser.Text = ""
        TxtPassword.Text = ""
        TxtServer.SetFocus
        Exit Sub
    End If
    Unload Me

End Sub
