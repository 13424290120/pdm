VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSEREdit 
   Caption         =   "SER Number Edit. SER ����༭"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSEREdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   12360
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdToQueryFinsGd 
      Caption         =   "Search ��ѯ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8415
      TabIndex        =   43
      Top             =   8160
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seek PjtNO. from PjtName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   8415
      TabIndex        =   39
      Top             =   6300
      Width           =   3345
      Begin VB.ComboBox ComboPJNOIndex 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   41
         Text            =   "ComboPJNOIndex"
         Top             =   300
         Width           =   3135
      End
      Begin VB.ComboBox ComboPjtName 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   40
         Text            =   "ComboPjtName"
         Top             =   1065
         Width           =   3135
      End
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5430
      TabIndex        =   14
      Top             =   3225
      Width           =   2775
   End
   Begin VB.TextBox TxtCAorA 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   13
      Top             =   1650
      Width           =   1380
   End
   Begin VB.TextBox TxtSERIndex 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   210
      Width           =   2775
   End
   Begin VB.TextBox TxtIDSO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   11
      Top             =   4005
      Width           =   1380
   End
   Begin VB.TextBox TxtOpnDate 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   10
      Top             =   4845
      Width           =   1350
   End
   Begin VB.TextBox TxtClosDate 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   9
      Top             =   5700
      Width           =   1350
   End
   Begin VB.ComboBox CombIDSO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4005
      Width           =   1410
   End
   Begin VB.TextBox TxtApplicant 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   930
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   6
      Top             =   6540
      Width           =   2775
   End
   Begin VB.TextBox TxtPjtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   5
      Top             =   7335
      Width           =   2775
   End
   Begin VB.TextBox TxtFinsGdNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   4
      Top             =   8160
      Width           =   2775
   End
   Begin VB.TextBox TxtSglPrtNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   3
      Top             =   2430
      Width           =   2775
   End
   Begin VB.TextBox TxtCommtNote 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5415
      TabIndex        =   2
      Top             =   8985
      Width           =   2775
   End
   Begin VB.CommandButton CmdSysDistrb 
      Caption         =   "System Distribute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8295
      TabIndex        =   1
      Top             =   210
      Width           =   1740
   End
   Begin VB.ComboBox CombCAorA 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   1410
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8025
      Top             =   9765
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   6810
      TabIndex        =   15
      Top             =   5685
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   6810
      TabIndex        =   16
      Top             =   4845
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmSEREdit.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      Left            =   8865
      TabIndex        =   42
      Top             =   3150
      Width           =   2715
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OK ȷ ��"
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
      Left            =   3645
      MouseIcon       =   "FrmSEREdit.frx":090F
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3045
      Picture         =   "FrmSEREdit.frx":0C19
      Top             =   9810
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel ȡ ��"
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
      Left            =   5925
      MouseIcon       =   "FrmSEREdit.frx":1035
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5325
      Picture         =   "FrmSEREdit.frx":133F
      Top             =   9810
      Width           =   300
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description ��Ʒ����"
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
      Left            =   2325
      MouseIcon       =   "FrmSEREdit.frx":175B
      TabIndex        =   36
      Top             =   3225
      Width           =   3000
   End
   Begin VB.Label LblCAorA 
      BackStyle       =   0  'Transparent
      Caption         =   "CA or FA or RJ���ĸ�״̬"
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
      Left            =   1590
      MouseIcon       =   "FrmSEREdit.frx":1A65
      TabIndex        =   35
      Top             =   1650
      Width           =   3705
   End
   Begin VB.Label LblSERIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "SER NO. SER ���"
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
      Left            =   2625
      MouseIcon       =   "FrmSEREdit.frx":1D6F
      TabIndex        =   34
      Top             =   210
      Width           =   2670
   End
   Begin VB.Label LblIDSO 
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
      Left            =   2640
      MouseIcon       =   "FrmSEREdit.frx":2079
      TabIndex        =   33
      Top             =   4020
      Width           =   2685
   End
   Begin VB.Label LblOpnDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date ��ʼ����"
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
      Left            =   2475
      MouseIcon       =   "FrmSEREdit.frx":2383
      TabIndex        =   32
      Top             =   4860
      Width           =   2865
   End
   Begin VB.Label LblClosDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Close Date ��������"
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
      Left            =   2445
      MouseIcon       =   "FrmSEREdit.frx":268D
      TabIndex        =   31
      Top             =   5715
      Width           =   2895
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
      Left            =   5940
      TabIndex        =   30
      Top             =   3750
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
      Left            =   7305
      TabIndex        =   29
      Top             =   3750
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
      Left            =   5925
      TabIndex        =   28
      Top             =   4590
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
      Left            =   7305
      TabIndex        =   27
      Top             =   4590
      Width           =   390
   End
   Begin VB.Label LblOld3 
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
      Left            =   5940
      TabIndex        =   26
      Top             =   5415
      Width           =   285
   End
   Begin VB.Label LblNew3 
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
      Left            =   7305
      TabIndex        =   25
      Top             =   5415
      Width           =   390
   End
   Begin VB.Label LblApplicant 
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant ������"
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
      Left            =   2940
      MouseIcon       =   "FrmSEREdit.frx":2997
      TabIndex        =   24
      Top             =   960
      Width           =   2355
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   2850
      Shape           =   4  'Rounded Rectangle
      Top             =   9675
      Width           =   4740
   End
   Begin VB.Label LblPJNOIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Number ������Ŀ���"
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
      Left            =   1110
      MouseIcon       =   "FrmSEREdit.frx":2CA1
      TabIndex        =   23
      Top             =   6540
      Width           =   4200
   End
   Begin VB.Label LblPjtName 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name ��Ŀ��������"
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
      Left            =   1395
      MouseIcon       =   "FrmSEREdit.frx":2FAB
      TabIndex        =   22
      Top             =   7350
      Width           =   3915
   End
   Begin VB.Label LblFinsGdNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good NO. ��صĳ�Ʒ���"
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
      Left            =   660
      MouseIcon       =   "FrmSEREdit.frx":32B5
      TabIndex        =   21
      Top             =   8160
      Width           =   4650
   End
   Begin VB.Label LblSglPrtNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part NO. SER������Ʒ���"
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
      Left            =   555
      MouseIcon       =   "FrmSEREdit.frx":35BF
      TabIndex        =   20
      Top             =   2430
      Width           =   4770
   End
   Begin VB.Label LblCommtNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment  Note. ע�ͺͱ�ע"
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
      Left            =   1335
      MouseIcon       =   "FrmSEREdit.frx":38C9
      TabIndex        =   19
      Top             =   8985
      Width           =   3975
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
      Left            =   7260
      TabIndex        =   18
      Top             =   1410
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
      Left            =   5910
      TabIndex        =   17
      Top             =   1410
      Width           =   285
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   2535
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7590
   End
End
Attribute VB_Name = "FrmSEREdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriSERIndex As String                       '############�����ĳɶ�Ӧ�ı��ֶ�����

Private Sub CmdSysDistrb_Click()
Dim Conn As New ADODB.Connection   '����һ��ADO����
Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
Dim BitNum As Integer   '����ȡ����ȥͷ��SER������λ��
Dim i As Integer
Dim j As Integer
Dim OriString, StrSql As String

MousePointer = vbHourglass   '����ʱ��ϳ�����Ҫ�������״̬
Conn.ConnectionString = connString
Conn.Open
Set rcds.ActiveConnection = Conn
On Error Resume Next           '############������ظĳɶ�Ӧ�ı��ֶ�����
StrSql = "select TOP 1 Right(Max(SERIndex),9)+1 from SER group by SERIndex order by SERIndex DESC"

rcds.Open StrSql, Conn, adOpenKeyset, adOpenStatic  'SERIndex+1��ʾÿ1λ����һ���ţ�Ҳ���Ǵ�ĩλ��ʼ����1
BitNum = Len(rcds.Fields(0).Value)  '�ж�ʵ�ʲ�ѯȥͷ��SER�����������Ǽ�λ

OriString = "SER"
For i = 0 To (12 - 3 - BitNum - 1)  '�ж�SER��ʵ������֮���м���0,�м��������Ӽ���
    OriString = OriString + "0"
Next i
If Modify = False Then
    TxtSERIndex.Text = OriString + Trim(Str(rcds.Fields(0).Value))
    MousePointer = vbDefault               '�ָ����״̬
    'MsgBox "Succeed to Add" + vbCrLf + "���ӳɹ�"   �����Բ��ã����˻�Ҫ�ش��ڣ��鷳
End If
If rcds.EOF Or rcds.BOF Then
    MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
    MousePointer = vbDefault                  '�ָ����״̬
    Exit Sub
End If

'������ܲ鵽��¼
If rcds.RecordCount = 0 Then
  'ϵͳ��ʾ��Ϣ��û���Ƽ��ţ�������ѡ��
MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
MousePointer = vbDefault                  '�ָ����״̬
End If
Conn.Close
End Sub

Private Sub CmdToQueryFinsGd_Click()
QuerytableName = "FinsGd"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���

    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False

        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    
FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub ComboPJNOIndex_Click()
TxtPJNOIndex.Text = ComboPJNOIndex.Text
TxtPjtName.Text = ComboPjtName.List(ComboPJNOIndex.ListIndex)
End Sub

Private Sub ComboPjtName_Click()
TxtPjtName.Text = ComboPjtName.Text
TxtPJNOIndex.Text = ComboPJNOIndex.List(ComboPjtName.ListIndex)
End Sub

Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Resize_ALL Me
End Sub

Private Sub TxtPJNOIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1         'sqlSDBC1Ϊ�û��Զ���ؼ�

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNOΪҪ��ѯ�ı���
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord�Ѿ�ȡ���ӵ�һ��ʼ�ң���������Ҫ���õ���ʼ

    Do While sqlUsrCtrl.IfBOForEOF = False
        sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '����1UseEquel����= 2UseLike�Ǵ���Like

       ComboPJNOIndex.AddItem (FormatNumber6(CStr(UsrCtlFind(0))))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
    
    Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub TxtPjtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNOΪҪ��ѯ�ı���
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord�Ѿ�ȡ���ӵ�һ��ʼ�ң���������Ҫ���õ���ʼ

     Do While sqlUsrCtrl.IfBOForEOF = False
       sqlUsrCtrl.FindRecord "Description", UseLike, Trim(TxtPjtName.Text)  '����1UseEquel����= 2UseLike�Ǵ���Like

       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
  
     Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub Form_Load()               '############������ظĳɶ�Ӧ�Ŀؼ�,����ֶ�����
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me
TxtApplicant.Text = PDMUserName
CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0
CombCAorA.AddItem ("CA")
CombCAorA.AddItem ("FA")
CombCAorA.AddItem ("RJ")
CombCAorA.ListIndex = 0

DTPickerOpnDate.Value = Date
DTPickerClosDate.Value = Date
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
  Function Isnum(Str As String) As Boolean    '�ж�һ���ַ������Ƿ�������  ��IsNumeric�ж�0000d031Ϊ��(����double������)
  Isnum = True
  Dim i  As Integer
  For i = 1 To Len(Str)
      Select Case Mid(Str, i, 1)
          Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          ' Isnum = True  ����дIsnum = True�ͳ���,��Ϊ����м�����ĸfalse�˺��������ֵĻ��ֳ�Ϊtrue��
          Case Else
            Isnum = False
      End Select
  Next
  End Function

Private Function Check() As Boolean                        '############������ظĳɶ�Ӧ�Ŀؼ�,����ֶ�����
If Trim(TxtSERIndex) = "" Or (Len(TxtSERIndex) <> 12) Then
    MsgBox "Please input SER Number" + vbCrLf + "������SER��", vbInformation, "System Info."
    TxtSERIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (left(TxtSERIndex, 3) = "SER" And Isnum(right(TxtSERIndex, 9))) Then  '����Left() Right()�Ǵ���ߺ��ұ߽�ȡ�ַ���
    MsgBox "SER Series Number is 3 Letter SER + 9 Number" + vbCrLf + "SER��3λ�ַ�SER + 9λ���ֵı��", vbInformation, "System Info."
    TxtSERIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtApplicant) = "" Then
    MsgBox "Please input Applicant Name" + vbCrLf + "��������������", vbInformation, "System Info."
    TxtApplicant.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtDescription) = "" Then
    MsgBox "Please input Description" + vbCrLf + "��������Ʒ����", vbInformation, "System Info."
    TxtDescription.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPJNOIndex) = "" Or (Not Isnum(TxtPJNOIndex)) Or (Len(TxtPJNOIndex) <> 6) Then
    MsgBox "Please input Project Number, 6 number" + vbCrLf + "�������漰��Ŀ���, 6λ������", vbInformation, "System Info."
    TxtPJNOIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPjtName) = "" Then
    MsgBox "Please input Project Name" + vbCrLf + "�������漰��Ŀ����", vbInformation, "System Info."
    TxtPjtName.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtFinsGdNO) = "" Then
    MsgBox "Please input relevant finish goods 12NC" + vbCrLf + "�������漰��Ʒ��12NC", vbInformation, "System Info."
    TxtFinsGdNO.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtSglPrtNO) = "" Or (Not Isnum(TxtSglPrtNO)) Or (Len(TxtSglPrtNO) <> 12) Then
    MsgBox "Please input relevant single part 12NC, Must be 12 Number" + vbCrLf + "�������漰�����12NC,������12λ����", vbInformation, "System Info."
    TxtSglPrtNO.SetFocus
    Check = False
    Exit Function
  End If
  
  
   Check = True
End Function


Private Sub lblOk_Click()
    
   '�ж�Ҫ�༭��Ϣ�Ƿ�����
   If Check = False Then
    Exit Sub
   End If
     
   With MySER              '�Ѿ�����Public MySER As New ClsSER, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .SERIndex = TxtSERIndex.Text
    .Applicant = TxtApplicant.Text
    .CAorA = CombCAorA.Text
    .Description = TxtDescription.Text
    .IDSO = CombIDSO.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
    .PJNOIndex = TxtPJNOIndex.Text
    .PjtName = TxtPjtName.Text
    .FinsGdNO = TxtFinsGdNO.Text
    .SglPrtNO = TxtSglPrtNO.Text
    .CommtNote = TxtCommtNote.Text
    
   
            '�жϲ�������ӻ����޸�
       If Modify = False Then         '�ж�Ϊ��Ӳ���
     
           '�ж�SERIndex����Ƿ��Ѿ�����
                If .In_DB(TxtSERIndex.Text) = True Then
                   MsgBox "SER number exists, Please re-input" + vbCrLf + "SER���ظ�������������", vbInformation, "System Info."
                   TxtSERIndex.SetFocus
                   TxtSERIndex.SelStart = 0
                   TxtSERIndex.SelLength = Len(TxtSERIndex)
                   Exit Sub
                Else
                   .Insert                   '���
                    MsgBox "Succeed to Add" + vbCrLf + "��ӳɹ�", vbInformation, "System Info."
                End If
       Else  '�ж�Ϊ�޸Ĳ���
        .Update (OriSERIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "�޸ĳɹ�", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '�ر�������
End Sub


