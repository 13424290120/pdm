VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCPCNEdit 
   Caption         =   "CP/CN Number Edit. CP/CN ����༭"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCPCNEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   11955
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox txtReason 
      Height          =   345
      ItemData        =   "FrmCPCNEdit.frx":08CA
      Left            =   5385
      List            =   "FrmCPCNEdit.frx":08DD
      TabIndex        =   44
      Top             =   8970
      Width           =   2805
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
      Left            =   8340
      TabIndex        =   39
      Top             =   5370
      Width           =   3345
      Begin VB.ComboBox ComboPJNOIndex 
         Height          =   345
         Left            =   105
         TabIndex        =   41
         Text            =   "ComboPJNOIndex"
         Top             =   270
         Width           =   3135
      End
      Begin VB.ComboBox ComboPjtName 
         Height          =   345
         Left            =   105
         TabIndex        =   40
         Text            =   "ComboPjtName"
         Top             =   1050
         Width           =   3135
      End
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8025
      Top             =   9615
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ComboBox CombCPCNMP 
      Height          =   345
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   1530
      Width           =   1410
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
      TabIndex        =   35
      Top             =   135
      Width           =   1740
   End
   Begin VB.TextBox TxtCommtNote 
      Height          =   375
      Left            =   5385
      TabIndex        =   33
      Top             =   8370
      Width           =   2775
   End
   Begin VB.TextBox TxtSglPrtNO 
      Height          =   375
      Left            =   5385
      TabIndex        =   31
      Top             =   7710
      Width           =   2775
   End
   Begin VB.TextBox TxtFinsGdNO 
      Height          =   375
      Left            =   5385
      TabIndex        =   29
      Top             =   6990
      Width           =   2775
   End
   Begin VB.TextBox TxtPjtName 
      Height          =   375
      Left            =   5385
      TabIndex        =   27
      Top             =   6285
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   375
      Left            =   5385
      TabIndex        =   25
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox TxtApplicant 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   780
      Width           =   2775
   End
   Begin VB.ComboBox CombIDSO 
      Height          =   345
      Left            =   6795
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3015
      Width           =   1410
   End
   Begin VB.TextBox TxtClosDate 
      Height          =   375
      Left            =   5415
      TabIndex        =   5
      Top             =   4710
      Width           =   1350
   End
   Begin VB.TextBox TxtOpnDate 
      Height          =   375
      Left            =   5415
      TabIndex        =   4
      Top             =   3855
      Width           =   1350
   End
   Begin VB.TextBox TxtIDSO 
      Height          =   375
      Left            =   5415
      TabIndex        =   3
      Top             =   3015
      Width           =   1380
   End
   Begin VB.TextBox TxtCPCNIndex 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   135
      Width           =   2775
   End
   Begin VB.TextBox TxtCPCNMP 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1515
      Width           =   1305
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   2235
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   6780
      TabIndex        =   8
      Top             =   4695
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   96993281
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   6780
      TabIndex        =   9
      Top             =   3855
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   96993281
      CurrentDate     =   39979
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason ԭ�������"
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
      Left            =   1320
      MouseIcon       =   "FrmCPCNEdit.frx":0936
      TabIndex        =   43
      Top             =   8970
      Width           =   3975
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmCPCNEdit.frx":0C40
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
      Left            =   8760
      TabIndex        =   42
      Top             =   2220
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   1935
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   8190
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
      TabIndex        =   38
      Top             =   1275
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
      Left            =   7260
      TabIndex        =   37
      Top             =   1275
      Width           =   390
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
      Left            =   1305
      MouseIcon       =   "FrmCPCNEdit.frx":0C85
      TabIndex        =   34
      Top             =   8370
      Width           =   3975
   End
   Begin VB.Label LblSglPrtNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part NO. �����漰�������"
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
      Left            =   480
      MouseIcon       =   "FrmCPCNEdit.frx":0F8F
      TabIndex        =   32
      Top             =   7710
      Width           =   4800
   End
   Begin VB.Label LblFinsGdNO 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good NO. �����漰��Ʒ���"
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
      Left            =   300
      MouseIcon       =   "FrmCPCNEdit.frx":1299
      TabIndex        =   30
      Top             =   6990
      Width           =   4980
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
      Left            =   1365
      MouseIcon       =   "FrmCPCNEdit.frx":15A3
      TabIndex        =   28
      Top             =   6300
      Width           =   3915
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
      Left            =   1080
      MouseIcon       =   "FrmCPCNEdit.frx":18AD
      TabIndex        =   26
      Top             =   5520
      Width           =   4200
   End
   Begin VB.Shape Shape2 
      Height          =   720
      Left            =   2790
      Shape           =   4  'Rounded Rectangle
      Top             =   9735
      Width           =   4740
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
      MouseIcon       =   "FrmCPCNEdit.frx":1BB7
      TabIndex        =   24
      Top             =   825
      Width           =   2355
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
      Left            =   7275
      TabIndex        =   23
      Top             =   4425
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
      Left            =   5910
      TabIndex        =   22
      Top             =   4425
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
      Left            =   7275
      TabIndex        =   21
      Top             =   3600
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
      Left            =   5895
      TabIndex        =   20
      Top             =   3600
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
      Left            =   7275
      TabIndex        =   19
      Top             =   2760
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
      Left            =   5910
      TabIndex        =   18
      Top             =   2760
      Width           =   285
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
      Left            =   2415
      MouseIcon       =   "FrmCPCNEdit.frx":1EC1
      TabIndex        =   17
      Top             =   4725
      Width           =   2895
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
      Left            =   2445
      MouseIcon       =   "FrmCPCNEdit.frx":21CB
      TabIndex        =   16
      Top             =   3870
      Width           =   2865
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
      Left            =   2610
      MouseIcon       =   "FrmCPCNEdit.frx":24D5
      TabIndex        =   15
      Top             =   3030
      Width           =   2685
   End
   Begin VB.Label LblCPCNIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "CP/CN NO. CP/CN ���"
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
      Left            =   2025
      MouseIcon       =   "FrmCPCNEdit.frx":27DF
      TabIndex        =   14
      Top             =   135
      Width           =   3270
   End
   Begin VB.Label LblCPCNMP 
      BackStyle       =   0  'Transparent
      Caption         =   "CP/CN/MP ���ĸ�״̬"
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
      Left            =   2145
      MouseIcon       =   "FrmCPCNEdit.frx":2AE9
      TabIndex        =   13
      Top             =   1515
      Width           =   3150
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description �������"
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
      Left            =   2295
      MouseIcon       =   "FrmCPCNEdit.frx":2DF3
      TabIndex        =   12
      Top             =   2235
      Width           =   3000
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5265
      Picture         =   "FrmCPCNEdit.frx":30FD
      Top             =   9870
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
      Height          =   435
      Left            =   5865
      MouseIcon       =   "FrmCPCNEdit.frx":3519
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9900
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2985
      Picture         =   "FrmCPCNEdit.frx":3823
      Top             =   9870
      Width           =   300
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
      Height          =   435
      Left            =   3585
      MouseIcon       =   "FrmCPCNEdit.frx":3C3F
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   9900
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCPCNEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriCPCNIndex As String                       '############�����ĳɶ�Ӧ�ı��ֶ�����

Private Sub CmdSysDistrb_Click()
Dim Conn As New ADODB.Connection   '����һ��ADO����
Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
Dim BitNum As Integer   '����ȡ����ȥͷ��HK1FS������λ��
Dim i As Integer
Dim J As Integer
Dim OriString As String
Conn.ConnectionString = connString
Conn.Open
Set rcds.ActiveConnection = Conn
On Error Resume Next           '############������ظĳɶ�Ӧ�ı��ֶ�����
rcds.Open "select top 1 Right(CPCNIndex,7)+1 from CPCN WHERE ((Right(CPCNIndex,7)+1) Not In (select Right(CPCNIndex,7) from CPCN))  order by Right(CPCNIndex,7)", Conn, adOpenKeyset, adOpenStatic  'CPCNIndex+1��ʾÿ1λ����һ���ţ�Ҳ���Ǵ�ĩλ��ʼ����1
BitNum = Len(rcds.Fields(0).Value)  '�ж�ʵ�ʲ�ѯȥͷ��HK1FS�����������Ǽ�λ

OriString = "HK1FS"
For i = 0 To (12 - 5 - BitNum - 1)  '�ж�HK1FS��ʵ������֮���м���0,�м��������Ӽ���
OriString = OriString + "0"
Next i
        If Modify = False Then
            TxtCPCNIndex.Text = OriString + Trim(Str(rcds.Fields(0).Value))
            'MsgBox "Succeed to Add" + vbCrLf + "���ӳɹ�"   �����Բ��ã����˻�Ҫ�ش��ڣ��鷳
        End If
    If rcds.EOF Or rcds.BOF Then
        MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
        Exit Sub
    End If

  '������ܲ鵽��¼
    If rcds.RecordCount = 0 Then
      'ϵͳ��ʾ��Ϣ��û���Ƽ��ţ�������ѡ��
    MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
    End If
Conn.Close
End Sub

Private Sub CmdToQueryFinsGd_Click()
QueryTableName = "FinsGd"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���

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
    
FrmQuery.Show 1 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub CmdToQuerySglPrt_Click()
QueryTableName = "SglPrt"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���

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
    
FrmQuery.Show 1 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
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
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNOΪҪ��ѯ�ı���
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord�Ѿ�ȡ���ӵ�һ��ʼ�ң���������Ҫ���õ���ʼ

    Do While sqlUsrCtrl.IfBOForEOF = False
        sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '����1UseEquel����= 2UseLike�Ǵ���Like

       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
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
''ResizeInit Me
CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0
CombCPCNMP.AddItem ("CN")
CombCPCNMP.AddItem ("CP")
CombCPCNMP.AddItem ("CP/CN")
CombCPCNMP.AddItem ("MP")
CombCPCNMP.ListIndex = 2
TxtApplicant.Text = PDMUserName

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
If Trim(TxtCPCNIndex) = "" Or (Len(TxtCPCNIndex) <> 12) Then
    MsgBox "Please input CP/CN Number" + vbCrLf + "������CP/CN��", vbInformation, "System Info."
    TxtCPCNIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (left(TxtCPCNIndex, 5) = "HK1FS" And Isnum(right(TxtCPCNIndex, 7))) Then  '����Left() Right()�Ǵ���ߺ��ұ߽�ȡ�ַ���
    MsgBox "CP/CN Series Number is 5 Letter HK1FS + 7 Number" + vbCrLf + "CP/CN��7λ�ַ�HK1FS + 7λ���ֵı��", vbInformation, "System Info."
    TxtCPCNIndex.SetFocus
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
    MsgBox "Please input Description" + vbCrLf + "������������", vbInformation, "System Info."
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
If Trim(TxtSglPrtNO) = "" Then
    MsgBox "Please input relevant single part 12NC" + vbCrLf + "�������漰�����12NC", vbInformation, "System Info."
    TxtSglPrtNO.SetFocus
    Check = False
    Exit Function
  End If
  
 If Trim(txtReason.Text) = "" Then
    MsgBox "Please choose the reason.", vbInformation, "System Info."
    txtReason.SetFocus
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
     
   With MyCPCN              '�Ѿ�����Public MyCPCN As New ClsCPCN, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .CPCNIndex = TxtCPCNIndex.Text
    .Applicant = TxtApplicant.Text
    .CPCNMP = CombCPCNMP.Text
    .Description = TxtDescription.Text
    .IDSO = CombIDSO.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
    .PJNOIndex = TxtPJNOIndex.Text
    .PjtName = TxtPjtName.Text
    .FinsGdNO = TxtFinsGdNO.Text
    .SglPrtNO = TxtSglPrtNO.Text
    .CommtNote = TxtCommtNote.Text
    .Reason = txtReason.Text
    
    
   
            '�жϲ�������ӻ����޸�
       If Modify = False Then         '�ж�Ϊ��Ӳ���
     
           '�ж�CPCNIndex����Ƿ��Ѿ�����
                If .In_DB(TxtCPCNIndex.Text) = True Then
                   MsgBox "CP/CN number exists, Please re-input" + vbCrLf + "CP/CN���ظ�������������", vbInformation, "System Info."
                   TxtCPCNIndex.SetFocus
                   TxtCPCNIndex.SelStart = 0
                   TxtCPCNIndex.SelLength = Len(TxtCPCNIndex)
                   Exit Sub
                Else
                   .Insert                   '���
                    MsgBox "Succeed to Add" + vbCrLf + "��ӳɹ�", vbInformation, "System Info."
                End If
       Else  '�ж�Ϊ�޸Ĳ���
        .Update (OriCPCNIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "�޸ĳɹ�", vbInformation, "System Info."
       End If
       
    End With
    Unload Me    '�ر�������
End Sub

