VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDatabaseManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBase Form Administrate ���ݿ�����"
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
      Caption         =   "ʵ�����ݿ��ֱ��SQL����ѯ"
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
            Name            =   "����"
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
         Caption         =   "Search �� ѯ"
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
         Caption         =   "Return �� ��"
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
      Caption         =   " ������￪ʼ����/�ָ����ݿ�"
      BeginProperty Font 
         Name            =   "����"
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
Dim sqlApp As New SQLDMO.Application   '����һ��SQLӦ�ö���sqlApp
Dim ServerName As SQLDMO.NameList

    Label1.Caption = "��ʼ������������ݿ������"
    MousePointer = vbHourglass                        '������ó�ɳ©
    '�����������п��õ�SQL SERVER ������
    Set ServerName = sqlApp.ListAvailableSQLServers        'ǰ��Dim ServerName û��������,�����Զ�����ɼ���
    
    For I = 1 To ServerName.Count
        FrmServerBkup.CobServer.AddItem (ServerName.Item(I))  'CobServer����һ���򿪴�FrmServerBkup�е�һ����Ͽ�ؼ�
    Next
    
    Call FrmServerBkup.LocalInfo     ' LocalInfo����һ���򿪴�FrmServerBkup�е��ӹ��̡�ȡ�ñ�������,�ͷ��ظ�����������Ip��ַ
    MousePointer = vbDefault
    
    'FrmServerBkup.Show 1          '�򿪱������ݿ�������
    
    If StrComp("(local)", Trim(FrmServerBkup.CobServer.Text), 1) = 0 Then    'StrComp Ϊ�ַ����ȽϵĽ����string1 С�� string2����-1,���ڷ���0,���ڷ���1.   ������������1 ִ��һ������ԭ�ĵıȽϡ�
        FrmServerBkup.LabServerName.Caption = FrmServerBkup.LabComputer.Caption
        FrmServerBkup.LabServerIP.Caption = FrmServerBkup.LabIp.Caption
    End If
     MousePointer = vbDefault
End Sub

Private Sub lblOk_Click()
If TxtSQL.Text <> "" Then  '���SQL�����ڲ�Ϊ��
    If LblOK.Caption = "Search �� ѯ" Then  '���LblOK��Caption��Search��ѯ
    '��Adodc1�������ݿ���в�ѯ
    Adodc1.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(DBUser) + ";pwd=" + Trim(Password) + ";database=" + Trim(DataBase) + ""
    Adodc1.RecordSource = TxtSQL.Text
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    '�ı�����򲻿���
    TxtSQL.Enabled = False
    '���LblOK��Caption��Ϊ���²�ѯ
    LblOK.Caption = "Re-Search ���²�ѯ"
    Else     'LblOK��Caption�����²�ѯ
    '��DataGrid��Adodc2����
     Set DataGrid1.DataSource = Adodc2 '��DataGrid1��û���ݿ��Adodc2�������������
     TxtSQL.Text = ""
     LblOK.Caption = "�� ѯ"
     TxtSQL.Enabled = True
    End If
Else
MsgBox "Please Input SQL query" + vbCrLf + "�������ѯ���" + vbCrLf + "Example: select * from SglPrt", vbInformation, "System Info."
End If
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
