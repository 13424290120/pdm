VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinsGdEdit 
   Caption         =   "Finish Goods Number Edit.   Finish Goods����༭"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFinsGdEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   12345
   StartUpPosition =   2  '��Ļ����
   Visible         =   0   'False
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
      Left            =   8880
      TabIndex        =   49
      Top             =   180
      Width           =   1740
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
      Left            =   8850
      TabIndex        =   45
      Top             =   4290
      Width           =   3345
      Begin VB.ComboBox ComboPJNOIndex 
         Height          =   345
         Left            =   105
         TabIndex        =   47
         Text            =   "ComboPJNOIndex"
         Top             =   285
         Width           =   3135
      End
      Begin VB.ComboBox ComboPjtName 
         Height          =   345
         Left            =   105
         TabIndex        =   46
         Text            =   "ComboPjtName"
         Top             =   1065
         Width           =   3135
      End
   End
   Begin VB.ComboBox CombLocation 
      Height          =   345
      Left            =   7335
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   7725
      Width           =   1410
   End
   Begin VB.ComboBox CombItemType 
      Height          =   345
      Left            =   7335
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   6870
      Width           =   1410
   End
   Begin VB.ComboBox CombProductLine 
      Height          =   345
      ItemData        =   "FrmFinsGdEdit.frx":08CA
      Left            =   7335
      List            =   "FrmFinsGdEdit.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6105
      Width           =   1410
   End
   Begin VB.TextBox TxtCommtNote 
      Height          =   375
      Left            =   5940
      TabIndex        =   12
      Top             =   8490
      Width           =   2775
   End
   Begin VB.TextBox TxtLocation 
      Height          =   375
      Left            =   5940
      TabIndex        =   11
      Top             =   7695
      Width           =   1305
   End
   Begin VB.TextBox TxtItemType 
      Height          =   375
      Left            =   5940
      TabIndex        =   10
      Top             =   6855
      Width           =   1305
   End
   Begin VB.TextBox TxtPjtName 
      Height          =   375
      Left            =   5925
      TabIndex        =   9
      Top             =   5340
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   375
      Left            =   5925
      TabIndex        =   8
      Top             =   4545
      Width           =   2775
   End
   Begin VB.TextBox TxtApplicant 
      Height          =   375
      Left            =   5940
      TabIndex        =   7
      Top             =   765
      Width           =   2775
   End
   Begin VB.ComboBox CombIDSO 
      Height          =   345
      Left            =   7335
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2130
      Width           =   1410
   End
   Begin VB.TextBox TxtClosDate 
      Height          =   375
      Left            =   5955
      TabIndex        =   5
      Top             =   3825
      Width           =   1350
   End
   Begin VB.TextBox TxtOpnDate 
      Height          =   375
      Left            =   5955
      TabIndex        =   4
      Top             =   2970
      Width           =   1350
   End
   Begin VB.TextBox TxtIDSO 
      Height          =   375
      Left            =   5955
      TabIndex        =   3
      Top             =   2130
      Width           =   1380
   End
   Begin VB.TextBox TxtFinsGdIndex 
      Height          =   375
      Left            =   5940
      TabIndex        =   2
      Top             =   150
      Width           =   2775
   End
   Begin VB.TextBox TxtProductLine 
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   6090
      Width           =   1305
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   1455
      Width           =   2775
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8250
      Top             =   9345
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   7320
      TabIndex        =   14
      Top             =   3810
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   7320
      TabIndex        =   15
      Top             =   2970
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmFinsGdEdit.frx":08CE
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
      Left            =   9255
      TabIndex        =   48
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label LblReminder1 
      Caption         =   $"FrmFinsGdEdit.frx":0913
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   8865
      TabIndex        =   44
      Top             =   5820
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label LblNew5 
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
      Left            =   7830
      TabIndex        =   43
      Top             =   7410
      Width           =   390
   End
   Begin VB.Label LblOld5 
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
      Left            =   6465
      TabIndex        =   42
      Top             =   7410
      Width           =   285
   End
   Begin VB.Label LblNew4 
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
      Left            =   7830
      TabIndex        =   41
      Top             =   6600
      Width           =   390
   End
   Begin VB.Label LblOld4 
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
      Left            =   6465
      TabIndex        =   40
      Top             =   6600
      Width           =   285
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
      Left            =   6450
      TabIndex        =   37
      Top             =   5850
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
      Left            =   7800
      TabIndex        =   36
      Top             =   5850
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
      Left            =   1890
      MouseIcon       =   "FrmFinsGdEdit.frx":0A9D
      TabIndex        =   35
      Top             =   8490
      Width           =   3960
   End
   Begin VB.Label LblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Number. ��Ʒ�����"
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
      Left            =   1350
      MouseIcon       =   "FrmFinsGdEdit.frx":0DA7
      TabIndex        =   34
      Top             =   7695
      Width           =   4500
   End
   Begin VB.Label LblItemType 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type Number. ��Ʒ�����"
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
      Left            =   1170
      MouseIcon       =   "FrmFinsGdEdit.frx":10B1
      TabIndex        =   33
      Top             =   6855
      Width           =   4680
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
      Left            =   1905
      MouseIcon       =   "FrmFinsGdEdit.frx":13BB
      TabIndex        =   32
      Top             =   5355
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
      Left            =   1620
      MouseIcon       =   "FrmFinsGdEdit.frx":16C5
      TabIndex        =   31
      Top             =   4545
      Width           =   4200
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   3390
      Shape           =   4  'Rounded Rectangle
      Top             =   9270
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
      Left            =   3480
      MouseIcon       =   "FrmFinsGdEdit.frx":19CF
      TabIndex        =   30
      Top             =   795
      Width           =   2355
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
      Left            =   7815
      TabIndex        =   29
      Top             =   3540
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
      Left            =   6450
      TabIndex        =   28
      Top             =   3540
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
      Left            =   7815
      TabIndex        =   27
      Top             =   2715
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
      Left            =   6435
      TabIndex        =   26
      Top             =   2715
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
      Left            =   7815
      TabIndex        =   25
      Top             =   1875
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
      Left            =   6450
      TabIndex        =   24
      Top             =   1875
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
      Left            =   2955
      MouseIcon       =   "FrmFinsGdEdit.frx":1CD9
      TabIndex        =   23
      Top             =   3840
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
      Left            =   2985
      MouseIcon       =   "FrmFinsGdEdit.frx":1FE3
      TabIndex        =   22
      Top             =   2985
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
      Left            =   3150
      MouseIcon       =   "FrmFinsGdEdit.frx":22ED
      TabIndex        =   21
      Top             =   2145
      Width           =   2685
   End
   Begin VB.Label LblFinsGdIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Goods NO.    Finish Goods ���"
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
      Left            =   345
      MouseIcon       =   "FrmFinsGdEdit.frx":25F7
      TabIndex        =   20
      Top             =   150
      Width           =   5490
   End
   Begin VB.Label LblProductLine 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Product Line ��Ʒ�߱��"
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
      Left            =   2340
      MouseIcon       =   "FrmFinsGdEdit.frx":2901
      TabIndex        =   19
      Top             =   6090
      Width           =   3495
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
      Left            =   2850
      MouseIcon       =   "FrmFinsGdEdit.frx":2C0B
      TabIndex        =   18
      Top             =   1455
      Width           =   2985
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5865
      Picture         =   "FrmFinsGdEdit.frx":2F15
      Top             =   9405
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
      Left            =   6465
      MouseIcon       =   "FrmFinsGdEdit.frx":3331
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   9435
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3585
      Picture         =   "FrmFinsGdEdit.frx":363B
      Top             =   9405
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
      Height          =   315
      Left            =   4185
      MouseIcon       =   "FrmFinsGdEdit.frx":3A57
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   9435
      Width           =   1095
   End
End
Attribute VB_Name = "FrmFinsGdEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriFinsGdIndex As String                       '############�����ĳɶ�Ӧ�ı��ֶ�����

Private Sub CmdSysDistrb_Click()
FrmFinsGDNOSection.ModifyFm = Modify          '�ѵ�ǰ���ڵ�״̬�̳и�����һ������
FrmFinsGDNOSection.Show 1
End Sub



Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modify Then Modify = False
    Unload Me
End Sub

Private Sub LblProductLine_Click()
LblReminder1.Visible = True
End Sub

Private Sub LblReminder1_Click()
LblReminder1.Visible = False
End Sub
Private Sub ComboPJNOIndex_Click()
TxtPJNOIndex.Text = ComboPJNOIndex.Text
TxtPjtName.Text = ComboPjtName.List(ComboPJNOIndex.ListIndex)
End Sub

Private Sub ComboPjtName_Click()
TxtPjtName.Text = ComboPjtName.Text
TxtPJNOIndex.Text = ComboPJNOIndex.List(ComboPjtName.ListIndex)
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
LblReminder1.Visible = False

CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0

CombProductLine.AddItem ("0000")
CombProductLine.AddItem ("5000")
CombProductLine.AddItem ("5010")
CombProductLine.AddItem ("6000")
CombProductLine.AddItem ("6400")
CombProductLine.AddItem ("7100")
CombProductLine.AddItem ("7200")
CombProductLine.AddItem ("7300")
CombProductLine.AddItem ("7400")
CombProductLine.AddItem ("7500")
CombProductLine.AddItem ("7600")
CombProductLine.AddItem ("7900")
CombProductLine.AddItem ("8100")
CombProductLine.AddItem ("8200")
CombProductLine.AddItem ("8300")
CombProductLine.AddItem ("9100")
CombProductLine.AddItem ("9200")
CombProductLine.AddItem ("9300")
CombProductLine.AddItem ("9400")
CombProductLine.AddItem ("9900")
CombProductLine.ListIndex = 5

CombItemType.AddItem ("400")
CombItemType.ListIndex = 0

CombLocation.AddItem ("AV")
CombLocation.AddItem ("MM")
CombLocation.AddItem ("CAR")
CombLocation.AddItem ("AV-KIT")
CombLocation.ListIndex = 0

DTPickerOpnDate.Value = Date
DTPickerClosDate.Value = Date
If Modify Then CmdSysDistrb.Enabled = False
End Sub

Private Sub LblCancel_Click()
    If Modify Then Modify = False
    Unload Me
End Sub
Private Function Isnum(Str As String) As Boolean     '�ж�һ���ַ������Ƿ�������  ��IsNumeric�ж�0000d031Ϊ��(����double������)
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
If Trim(TxtFinsGdIndex) = "" Then
    MsgBox "Please input Finish Goods Number" + vbCrLf + "������Finish Goods��", vbInformation, "System Info."
    TxtFinsGdIndex.SetFocus
    Check = False
    Exit Function
  End If
 If Not (Len(TxtFinsGdIndex) = 12 And Isnum(TxtFinsGdIndex)) Then  '����Left() Right()�Ǵ���ߺ��ұ߽�ȡ�ַ���
    MsgBox "Finish Goods Series Number is 12 Number, no Letter" + vbCrLf + "Finish Goods��12λ���ֵı��,����ĸ", vbInformation, "System Info."
    TxtFinsGdIndex.SetFocus
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
    MsgBox "Please input Description" + vbCrLf + "�������Ʒ����", vbInformation, "System Info."
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
  
'If Trim(TxtItemType) = "" Then
'    MsgBox "Please input relevant Item Type" + vbCrLf + "�������Ʒ�����", vbInformation, "System Info."
'    TxtItemType.SetFocus
'    Check = False
'    Exit Function
'  End If
'If Trim(TxtLocation) = "" Then
'    MsgBox "Please input relevant Location Number" + vbCrLf + "�������Ʒ�����", vbInformation, "System Info."
'    TxtLocation.SetFocus
'    Check = False
'    Exit Function
'  End If
  
  
   Check = True
End Function


Private Sub lblOk_Click()
    
   '�ж�Ҫ�༭��Ϣ�Ƿ�����
   If Check = False Then
    Exit Sub
   End If
     
   With MyFinsGd              '�Ѿ�����Public MyFinsGd As New ClsFinsGd, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .FinsGdIndex = TxtFinsGdIndex.Text
    .Applicant = TxtApplicant.Text
    .Description = TxtDescription.Text
    .IDSO = CombIDSO.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
    .PJNOIndex = TxtPJNOIndex.Text
    .PjtName = TxtPjtName.Text
    .ProductLine = CombProductLine.Text
    .ItemType = CombItemType.Text
    .Location = CombLocation.Text
    .CommtNote = TxtCommtNote.Text
    
   
            '�жϲ�������ӻ����޸�
       If Modify = False Then         '�ж�Ϊ��Ӳ���
     
           '�ж�FinsGdIndex����Ƿ��Ѿ�����
                If .In_DB(TxtFinsGdIndex.Text) = True Then
                   MsgBox "Finish Goods number exists, Please re-input" + vbCrLf + "Finish Goods���ظ�������������", vbInformation, "System Info."
                   TxtFinsGdIndex.SetFocus
                   TxtFinsGdIndex.SelStart = 0
                   TxtFinsGdIndex.SelLength = Len(TxtFinsGdIndex)
                   Exit Sub
                Else
                   .Insert                   '���
                    MsgBox "Succeed to Add" + vbCrLf + "��ӳɹ�", vbInformation, "System Info."
                End If
       Else  '�ж�Ϊ�޸Ĳ���
        .Update (OriFinsGdIndex)
         MsgBox "Succeed to Modify" + vbCrLf + "�޸ĳɹ�", vbInformation, "System Info."
         Modify = False
       End If
       
    End With
    Unload Me    '�ر�������
End Sub



