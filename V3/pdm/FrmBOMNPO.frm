VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmBOMNPO 
   Caption         =   "PDM-BOM NPO(New Part Overview) ���̹�����ϵͳ"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13752
   Icon            =   "FrmBOMNPO.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10800
   ScaleWidth      =   13752
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdExportExcel 
      Caption         =   "Export Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9330
      TabIndex        =   41
      Top             =   10260
      Width           =   1635
   End
   Begin VB.ComboBox cmbAuthor 
      Height          =   300
      Left            =   1770
      TabIndex        =   40
      Top             =   1710
      Width           =   2000
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print NPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11040
      TabIndex        =   38
      Top             =   10260
      Width           =   1155
   End
   Begin VB.CommandButton CmdNewtoOld 
      Caption         =   "All New to Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4980
      TabIndex        =   37
      Top             =   10260
      Width           =   1455
   End
   Begin VB.CommandButton CmdNPOinPjt 
      Caption         =   "NPO in Project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      TabIndex        =   34
      Top             =   10260
      Width           =   1410
   End
   Begin VB.CommandButton CmdNPOinBOM 
      Caption         =   "New Part OverView"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      TabIndex        =   30
      Top             =   10245
      Width           =   1830
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1995
      Top             =   10335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSearchSglPrt 
      Caption         =   "Search SglPrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7920
      TabIndex        =   32
      Top             =   10260
      Width           =   1320
   End
   Begin VB.CommandButton CmdExportBOM 
      Caption         =   "Export BOM / NPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9270
      TabIndex        =   12
      Top             =   1650
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton CmdSearchFinsGd 
      Caption         =   "Search FinsGd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      TabIndex        =   33
      Top             =   10260
      Width           =   1395
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   11055
      TabIndex        =   31
      Top             =   10620
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   106037249
      CurrentDate     =   40037
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12270
      TabIndex        =   1
      Top             =   10260
      Width           =   1350
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1770
      TabIndex        =   27
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton CmdRunBOM 
      Caption         =   "Refresh BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   60
      TabIndex        =   2
      Top             =   10230
      Width           =   1425
   End
   Begin VB.TextBox txtSERNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10185
      TabIndex        =   22
      Top             =   960
      Width           =   2265
   End
   Begin VB.TextBox txtCPCNNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5445
      TabIndex        =   21
      Top             =   960
      Width           =   2160
   End
   Begin VB.TextBox txtSERlocate 
      Height          =   360
      Left            =   9390
      TabIndex        =   20
      Top             =   1350
      Width           =   3060
   End
   Begin VB.TextBox txtCPCNlocate 
      Height          =   360
      Left            =   4530
      TabIndex        =   19
      Top             =   1350
      Width           =   3060
   End
   Begin VB.CommandButton CmdSERView 
      Caption         =   $"FrmBOMNPO.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12510
      TabIndex        =   16
      Top             =   975
      Width           =   1125
   End
   Begin VB.CommandButton CmdCPCNView 
      Caption         =   $"FrmBOMNPO.frx":08D7
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7665
      TabIndex        =   15
      Top             =   960
      Width           =   1125
   End
   Begin VB.TextBox MSFlexGrid1EditText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12585
      TabIndex        =   13
      Text            =   "MsFleGrdTxt"
      Top             =   10605
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton CmdDrwView 
      Caption         =   "See Drw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12510
      TabIndex        =   7
      Top             =   195
      Width           =   1125
   End
   Begin VB.TextBox txtNodeDrwlocate 
      Height          =   360
      Left            =   8535
      TabIndex        =   6
      Top             =   540
      Width           =   3930
   End
   Begin VB.TextBox txtNodePrtUnit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7785
      TabIndex        =   5
      Top             =   540
      Width           =   705
   End
   Begin VB.TextBox txtNodeSglPrt12NC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3975
      TabIndex        =   4
      Top             =   540
      Width           =   1410
   End
   Begin VB.TextBox txtNodeDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5445
      TabIndex        =   3
      Top             =   540
      Width           =   2310
   End
   Begin VB.TextBox TxtFinsGdIndex 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1770
      TabIndex        =   0
      Top             =   210
      Width           =   2000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8040
      Left            =   30
      TabIndex        =   29
      Top             =   2160
      Width           =   13650
      _ExtentX        =   24067
      _ExtentY        =   14182
      _Version        =   393216
      Rows            =   38
      Cols            =   15
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   5475
      Top             =   10365
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ComboBox ComboPJNOIndex 
      Height          =   300
      Left            =   1770
      TabIndex        =   36
      Top             =   960
      Width           =   2000
   End
   Begin VB.ComboBox ComboPjtName 
      Height          =   300
      Left            =   1770
      TabIndex        =   35
      Top             =   1350
      Width           =   2000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      MouseIcon       =   "FrmBOMNPO.frx":08E5
      TabIndex        =   39
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label LblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   630
      MouseIcon       =   "FrmBOMNPO.frx":0BEF
      TabIndex        =   28
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label LblCPCN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "CP/CN Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   26
      Top             =   1005
      Width           =   1425
   End
   Begin VB.Label LblFinsGdNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Goods NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      MouseIcon       =   "FrmBOMNPO.frx":0EF9
      TabIndex        =   25
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label LblPJNOIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      MouseIcon       =   "FrmBOMNPO.frx":1203
      TabIndex        =   24
      Top             =   990
      Width           =   1485
   End
   Begin VB.Label LblPjtName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      MouseIcon       =   "FrmBOMNPO.frx":150D
      TabIndex        =   23
      Top             =   1380
      Width           =   1290
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8865
      TabIndex        =   18
      Top             =   1395
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   17
      Top             =   1395
      Width           =   480
   End
   Begin VB.Label LblSER 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SER Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8865
      TabIndex        =   14
      Top             =   1005
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      Height          =   1995
      Left            =   45
      Top             =   120
      Width           =   3795
   End
   Begin VB.Shape Shape1 
      Height          =   1995
      Left            =   3915
      Top             =   135
      Width           =   9795
   End
   Begin VB.Label LblDrw 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Drawing location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8535
      TabIndex        =   11
      Top             =   225
      Width           =   3870
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7755
      TabIndex        =   10
      Top             =   225
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5445
      TabIndex        =   9
      Top             =   225
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "  Selected 12NC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3990
      TabIndex        =   8
      Top             =   225
      Width           =   1440
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCode 
         Caption         =   "&Add New Code Here"
      End
      Begin VB.Menu mnuDeleteCode 
         Caption         =   "&Delete Selected Code"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename Code Here"
      End
   End
End
Attribute VB_Name = "FrmBOMNPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RowNum As Integer
Private FinishGoodsNO As String
Private ApprovalStatus As Boolean    '����BOM�Ƿ���׼�ı��
Private OpennerSubmiter As Boolean   '����BOM�����Ƿ�����(�ύ��)�ı��
Private NPOWorking As Boolean        '����ÿһ�е�NewOld״̬�Ƿ�ΪNew�ı��
Private BOMExist As Boolean          '����Refresh_FlexGridģ������BOM�Ƿ���ڵı��
Private StrSql As String
'Private scr As Object         '�����ר�õı��ʽ�����Ļ����������ò���

'Mid(myNode, 2, Len(myNode))    'ȥ�������һ���ַ�
'Mid(myNode, 1, Len(myNode)-1)  'ȥ�����ұ�һ���ַ�

Private Sub MSFlexGridColumnColorChange(MSFlexGridName As Object, ByVal ColNo As Integer, ByVal RowSum As Integer, Optional ByVal ColColor As Long = &HC0E0FF)      '&HC0E0FFΪǳ�ۺ�ɫ
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Col = ColNo                    '�ӵ�ColNo�е�0�п�ʼ
    MSFlexGridName.Row = 0                        '�ӵ�ColNo�е�0�п�ʼ
    MSFlexGridName.RowSel = RowSum - 1            '����ѡ��ֱ�����һ��RowSum
    MSFlexGridName.CellBackColor = ColColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub
Private Sub MSFlexGridRowColorChange(MSFlexGridName As Object, ByVal RowNo As Integer, ByVal ColSum As Integer, Optional ByVal RowColor As Long = &HFFFF&)          '&HFFFFΪ��ɫ
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Row = RowNo                    '�ӵ�RowNo�е�0�п�ʼ
    MSFlexGridName.Col = 1                        '�ӵ�RowNo�е�0�п�ʼ
    MSFlexGridName.ColSel = ColSum - 1            '����ѡ��ֱ�����һ��ColSum
    MSFlexGridName.CellBackColor = RowColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub

'�ж�һ���ַ������Ƿ�������  ��IsNumeric�ж�0000d031Ϊ��(����double������)
Private Function Isnum(Str As String) As Boolean
    Isnum = True
    Dim i  As Integer
    For i = 1 To Len(Str)
        Select Case Mid(Str, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            'Isnum = True  ����дIsnum = True�ͳ���,��Ϊ����м�����ĸfalse�˺��������ֵĻ��ֳ�Ϊtrue��
        Case Else
            Isnum = False
        End Select
    Next i
End Function

Private Sub GeneralDocView(ByVal InputPathName As String)
    Dim OpnDocPathName As String
    OpnDocPathName = Trim(InputPathName)
    If OpnDocPathName = "" Then
        MsgBox "The Drawing(Document) Path/name is Null", vbInformation, "System Info."
        Exit Sub
    End If
    OpnShllExcFile (OpnDocPathName)
End Sub


Private Sub cmbAuthor_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Dim Conn As New ADODB.Connection
    
        Conn.Open connString
        
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        
        MSFlexGrid1EditText.Visible = False
        DTPicker1.Visible = False
        
        If Len(Trim(TxtFinsGdIndex.Text)) = 0 Then
            MsgBox "You must enter a new 12NC for the Finish Goods", vbInformation, "System Info."
            Exit Sub
        ElseIf Not (Len(Trim(TxtFinsGdIndex.Text)) = 12 And Isnum(Trim(TxtFinsGdIndex.Text))) Then
            MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
            Exit Sub
        End If
        
        '�ж�BOM��¼�Ƿ�Ǽǲ����Ѿ���׼
        rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            If Trim(rs("Approver")) <> "" Then
                ApprovalStatus = True
            Else
                ApprovalStatus = False
            End If
            'ͬʱ�ж�BOM�����Ƿ�BOM����(�ύ��)
            If InStr(Trim(rs("Submiter")), PDMUserName) Then
                OpennerSubmiter = True
            Else
                OpennerSubmiter = False
            End If
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        
        
        StrSql = "Select Top 1 isNull(CPCNNmbr,''),isNull(CPCNLocate,'') From BOMCPCN Where BOMID =" & TxtFinsGdIndex.Text & " Order by BOMVersion Desc"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            txtCPCNNO.Text = rs(0)
            txtCPCNlocate.Text = rs(1)
        End If
        rs.Close
        

        
        Dim i As Integer
        With MSFlexGrid1
            For i = MSFlexGrid1.Rows - 1 To 1 Step -1
                StrSql = "SELECT * FROM SglPrt Where Applicant='" & Trim(cmbAuthor.Text) & "' And SglPrtIndex='" & left(.TextMatrix(i, 3), 11) & "0'"
                rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                    If i <> 1 Then .RemoveItem i: .Refresh
                End If
                rs.Close
            Next
        End With
        Conn.Close
    End If
    
End Sub

Private Sub CmdCPCNView_Click()
    GeneralDocView (txtCPCNlocate.Text)
End Sub

Private Sub CmdExportExcel_Click()
    If Not ApprovalStatus Then
        MsgBox "The BOM/NPO is NOT Approved, Please do NOT use it Formally(Offically)", vbInformation, "System Info."
    End If

    
    If MsgBox("You are going to Export BOM/NPO Data to an Excel File, Continue��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then

        Dim i, J As Integer

        Set xlApp = CreateObject("Excel.Application")   '����Excel�ļ�
        Set xlApp = New excel.Application
        xlApp.SheetsInNewWorkbook = 1                   '���½��Ĺ�����������Ϊ1
        
        '������ֲ���������ʾ
        'xlApp.OleRequestPendingTimeout = 10000   '10000��������æ�Ի���
        'xlApp.OleServerBusyTimeout = 1000     '����ʱ1��
        'xlApp.OleServerBusyRaiseError = True '����ʾæ�Ի���
    
    
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)              '��1�Ź�����
        xlSheet.Cells(1, 1) = "New Part Overview"
        If ComboPJNOIndex.Text <> "" Then xlSheet.Cells(2, 1) = "PJNOIndex:": xlSheet.Cells(2, 2) = ComboPJNOIndex.Text
        If ComboPjtName.Text <> "" Then xlSheet.Cells(2, 3) = "PjtName:": xlSheet.Cells(2, 4) = ComboPjtName.Text
        If cmbAuthor.Text <> "" Then xlSheet.Cells(2, 5) = "Applicant:": xlSheet.Cells(2, 6) = cmbAuthor.Text
        For i = 0 To MSFlexGrid1.Cols - 1
            xlSheet.Cells(3, i + 1) = MSFlexGrid1.TextMatrix(0, i)
        Next i
        
        xlSheet.Cells(2, i - 3) = "Table Maker:": xlSheet.Cells(2, i - 2) = PDMUserName
        xlSheet.Cells(2, i - 1) = "Print Date:": xlSheet.Cells(2, i) = Now()
        
        For J = 1 To MSFlexGrid1.Rows - 1
                For i = 0 To MSFlexGrid1.Cols - 1
                    xlSheet.Cells(J + 3, i + 1) = "'" + MSFlexGrid1.TextMatrix(J, i)
                Next i
        Next J
        'xlSheet.Cells(4, 1).CopyFromRecordset Conn.Execute(strSql)       '������ճ������
    
        xlApp.ActiveWorkbook.Close True     '�رչ�����������
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim i As Long, J As Long
    Dim rtMargin As RECT, rtCell As RECT, rtText As RECT

    If MsgBox("Are you sure that the default printer has set up Horizontal printing?", vbYesNo, "ERP") = vbNo Then Exit Sub
    '���ô�ӡ��Ϣ
    Printer.PaperSize = vbPRPSA4
    Printer.DrawMode = vbPixels
    SetRect rtMargin, 100, 100, 100, 100 'ҳ�߾�
    '��ʼ��ӡ
    Printer.CurrentX = rtMargin.left
    Printer.CurrentY = rtMargin.top
    Printer.Print "" '��ֽ
    SetRect rtCell, rtMargin.left, rtMargin.top, 0, 0
    With MSFlexGrid1
        For i = 0 To .Rows - 1
            .Row = i
            'ȷ���Ƿ�Ҫ��ҳ
            If Printer.ScaleHeight - .RowHeight(i) <= rtMargin.bottom Then
                Printer.NewPage
                rtCell.top = rtMargin.top
            End If
            For J = 0 To .Cols - 1
                .Col = J
                '��ӡ��Ԫ��߿�
                rtCell.right = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
                rtCell.bottom = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
                Rectangle Printer.hDC, rtCell.left, rtCell.top, rtCell.right + 1, rtCell.bottom + 1
                '���õ�Ԫ������
                Printer.FontName = .CellFontName
                Printer.FontSize = .CellFontSize
                Printer.FontBold = .CellFontBold
                Printer.FontItalic = .CellFontItalic
                Printer.FontUnderline = .CellFontUnderline
                '��ӡ��Ԫ�����֣������ڱ߾�Ϊ4��
                SetRect rtText, rtCell.left + 4, rtCell.top + 4, rtCell.right - 4, rtCell.bottom - 4
                DrawText Printer.hDC, .TextMatrix(i, J), LenB(StrConv(.TextMatrix(i, J), vbFromUnicode)), rtText, _
                DT_SINGLELINE Or GetAlign(.CellAlignment)
                rtCell.left = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
            Next
            rtCell.left = rtMargin.left
            rtCell.top = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
        Next
    End With
    '��ӡ���
    Printer.EndDoc
End Sub


'Private Sub cmdPrint_Click()
'Dim strLine As String
'Dim i, x As Integer
'
'    Screen.MousePointer = vbHourglass
'
'    Open App.Path & "\NPO_" & Year(Now) & Month(Now) & Day(Now) & Minute(Now) & Second(Now) & ".txt" For Output As #1
'        With MSFlexGrid1
'            For x = 0 To .Rows - 1
'                strLine = ""
'                For i = 0 To .Cols - 1
'                    strLine = strLine & "" & .TextMatrix(x, i)
'                    If i < .Cols - 1 Then
'                        strLine = strLine & vbTab
'                    End If
'                Next i
'
'                Print #1, strLine
'            Next x
'        End With
'    Close #1
'
'    Screen.MousePointer = vbDefault
'    MsgBox "Export Successfully!", vbInformation
'End Sub

Private Sub CmdSERView_Click()
    GeneralDocView (txtSERlocate.Text)
End Sub

Private Sub CmdDrwView_Click()
    GeneralDocView (txtNodeDrwlocate.Text)
End Sub

Private Sub CmdNPOinBOM_Click()
    'On Error Resume Next
    NPOWorking = True
    
    MSFlexGrid1.ColWidth(2) = 0
    MSFlexGrid1.ColWidth(1) = 0
    MSFlexGrid1.ColWidth(10) = 0
    'MSFlexGrid1.ColWidth(11) = 0
    
    Dim RowSumVar As Integer
    For RowSumVar = 2 To RowNum
        If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) <> "NEW" And UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) <> "EXISTING" Then   '���ز���New/Existing Part����
            MSFlexGrid1.RowHeight(RowSumVar) = 0
        End If
    Next RowSumVar
    
End Sub
Private Sub CmdNPOinPjt_Click()
    On Error GoTo vbErrorHandler
    Dim J As Integer
    Dim Conn As New ADODB.Connection
    
    Dim PJNOIndex, PjtName, AuthorName As String
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    '��Project��NPO��,�����޷����BOM��Approve״̬ �� BOM�Ĵ����Ƿ�����(�ύ��)��״̬,  ������ر�Ƕ�Ҫ��Ϊfalse
    ApprovalStatus = False
    OpennerSubmiter = False
    
    MSFlexGrid1EditText.Visible = False
    DTPicker1.Visible = False
    NPOWorking = True
    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    If Len(Trim(ComboPJNOIndex.Text)) = 0 And Len(Trim(cmbAuthor.Text)) = 0 Then
        MsgBox "You must enter a new 6NC for the Project Number, or choose an anthor", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(ComboPJNOIndex.Text)) = 6 And Isnum(Trim(ComboPJNOIndex.Text))) And Len(Trim(cmbAuthor.Text)) = 0 Then
        MsgBox "Project Number is 6 Number, no Letter" + vbCrLf + "������6λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    Else
        PJNOIndex = CLng(ComboPJNOIndex.Text) 'ȥ��
        PjtName = CStr(ComboPjtName.Text)
        AuthorName = CStr(cmbAuthor.Text)
    End If
    '�ж�SinglePart��¼�еķ���Project Number��New Part
    StrSql = "Select * from SglPrt where NewOldStatus ='New'"
    If PJNOIndex <> "" Then StrSql = StrSql & " And PJNOIndex ='" & CStr(PJNOIndex) & "'"
    If PjtName <> "" Then StrSql = StrSql & " And PJTName like '%" & Trim(PjtName) & "%'"
    If AuthorName <> "" Then StrSql = StrSql & " And Applicant = '" & Trim(AuthorName) & "'"
    
    rs.Open StrSql, Conn, adOpenStatic, adLockOptimistic
    J = 2      'MSFlexGrid1�ӵ�2�п�ʼд,��0����Title,��1�����ո�BOM�е�Finish Goods
    MSFlexGrid1.Rows = 3   '����������
    MSFlexGrid1.RowHeight(1) = 225     '������б����ع��Ļ���Ҫ�ָ��и�Ĭ��ֵ
    MSFlexGrid1.RowHeight(2) = 225     '������б����ع��Ļ���Ҫ�ָ��и�Ĭ��ֵ
    Do While rs.EOF = False
        MSFlexGrid1.TextMatrix(J, 0) = J - 1                               '����ÿ�е��к�,���������1�пո�BOM�е�Finish Goods,��������Ҫ -1����ŲŴ�1��ʼ
        MSFlexGrid1.TextMatrix(J, 3) = rs.Fields("SglPrtIndex")            '�������ϵ�12NC
        MSFlexGrid1.TextMatrix(J, 5) = Trim(rs.Fields("PrtUnit"))          '�������ϵĵ�λ
        MSFlexGrid1.TextMatrix(J, 6) = Trim(rs.Fields("Description"))      '�������ϵ�����
        MSFlexGrid1.TextMatrix(J, 7) = Trim(rs.Fields("ItemType"))         '�������ϵ�ItemType
        MSFlexGrid1.TextMatrix(J, 8) = Trim(rs.Fields("ProductLine"))      '�������ϵ�ProductionLine
        MSFlexGrid1.TextMatrix(J, 9) = Trim(rs.Fields("NewOldStatus"))     '�������ϵ�New/old״̬
        
        
        If IsNull(rs.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ rstCX3.Fields("SERNmbr") = Null
            MSFlexGrid1.TextMatrix(J, 12) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 12) = rs.Fields("SERNmbr")
        End If
        
        If IsNull(rs.Fields("CommtNote")) Then
            MSFlexGrid1.TextMatrix(J, 13) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 13) = Trim(rs.Fields("CommtNote"))
        End If
        
        If IsNull(rs.Fields("ETA")) Then
            MSFlexGrid1.TextMatrix(J, 15) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 15) = Format(rs.Fields("ETA"), "YYYY/MM/DD")
        End If
        
        If IsNull(rs.Fields("SampleQuantity")) Then
            MSFlexGrid1.TextMatrix(J, 16) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 16) = Trim(rs.Fields("SampleQuantity"))
        End If
        
        If IsNull(rs.Fields("Supplier")) Then
            MSFlexGrid1.TextMatrix(J, 17) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 17) = Trim(rs.Fields("Supplier"))
        End If
        
        If IsNull(rs.Fields("ArrivalDate")) Then
            MSFlexGrid1.TextMatrix(J, 18) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 18) = Format(rs.Fields("ArrivalDate"), "YYYY/MM/DD")
        End If
        
        If IsNull(rs.Fields("RoHsReport")) Then
            MSFlexGrid1.TextMatrix(J, 19) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 19) = Trim(rs.Fields("RoHsReport"))
        End If
        
        If IsNull(rs.Fields("LLTI")) Then
            MSFlexGrid1.TextMatrix(J, 20) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 20) = Trim(rs.Fields("LLTI"))
        End If
        
        If IsNull(rs.Fields("RiskOrderQ")) Then
            MSFlexGrid1.TextMatrix(J, 21) = ""
        Else
            MSFlexGrid1.TextMatrix(J, 21) = Trim(rs.Fields("RiskOrderQ"))
        End If
        
        rs.MoveNext
        J = J + 1
        MSFlexGrid1.Rows = J + 1
    Loop
    RowNum = J
    If rs.State = adStateOpen Then rs.Close
    Refresh_FlexGridColumnSupplierPN
    
    MSFlexGrid_ChgStatus_HightlightRow (10)                   '��ChangeStatus��11���������ݵ�������Ϊ��ɫ
    MSFlexGrid_NewOld_HightlightRow (9)                       '��New/Old��9����������ΪNew��������Ϊ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 9, J             '����New/Old��(��9��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 12, J            '����Comments��(��13��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 21, J            '����Supplier PN��(��22��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 13, J, &H404040  '���÷ָ���(��14��)Ϊ��ɫ
    MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 0                  '���õ�Ԫ��λ��ȡ������ı亯���е�ĳ�и�����ʾ
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdNPOinPjt__Click"
End Sub

Private Sub CmdNewtoOld_Click()
    On Error GoTo vbErrorHandler
    
    Dim Conn As New ADODB.Connection
    

    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    Dim RowSumVar As Integer
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    
    If Not OpennerSubmiter Then
        MsgBox "You are not the BOM/NPO Author, No Right to Update", vbInformation, "System Info."
        Exit Sub
    End If
    
    If MsgBox("You are going to Change All New Parts to Old Parts, Continue��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
        '�ȱ����ʼ��Col��Row�ŵ�ֵ
        ColNoTemp = MSFlexGrid1.Col
        RowNoTemp = MSFlexGrid1.Row
        For RowSumVar = 2 To RowNum
            If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) = "NEW" Then   '�ж��ǲ���New Part����
                '�ҵ�SinglePart�еĶ�Ӧ12NC,���Ҹ���NewOld�ֶγ�Old
                If rs.State = adStateOpen Then rs.Close
                rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                    MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
                Else
                    MSFlexGrid1.Col = 9
                    MSFlexGrid1.Row = RowSumVar
                    MSFlexGrid1.Text = "Old"
                    rs("NewOldStatus") = "Old"
                    rs.Update
                End If
                If rs.State = adStateOpen Then rs.Close
            End If
        Next RowSumVar
        MSFlexGrid1.Row = RowNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
        MSFlexGrid1.Col = ColNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
    End If
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdNewtoOld_Click"
End Sub

Private Sub CmdSearchFinsGd_Click()
    QueryTableName = "FinsGd"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        If PurchasingSys = "Y" Then FrmQuery.cmdAdd.Enabled = False
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    Set FromForm = FrmBOMNPO
    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub CmdSearchSglPrt_Click()

    QueryTableName = "SglPrt"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        If PurchasingSys = "Y" Then FrmQuery.cmdAdd.Enabled = False
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    Set FromForm = FrmBOMNPO
    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
    MousePointer = vbDefault                  '�ָ����״̬
End Sub

Private Sub CmdExportBOM_Click()
    If Not ApprovalStatus Then
        MsgBox "The BOM/NPO is NOT Approved, Please do NOT use it Formally(Offically)", vbInformation, "System Info."
    End If
    
    If MsgBox("You are going to Export BOM/NPO Data to an Excel File, Continue��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
        ExportFlexDataToExcel FrmBOMNPO.MSFlexGrid1, FrmBOMNPO.CommonDialog1
    End If
End Sub

Private Sub CmdRunBOM_Click()
    On Error GoTo vbErrorHandler
    
    Dim Conn As New ADODB.Connection
    

    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    MSFlexGrid1EditText.Visible = False
    DTPicker1.Visible = False
    
    If Len(Trim(TxtFinsGdIndex.Text)) = 0 Then
        MsgBox "You must enter a new 12NC for the Finish Goods", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(TxtFinsGdIndex.Text)) = 12 And Isnum(Trim(TxtFinsGdIndex.Text))) Then
        MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    End If
    
    '�ж�BOM��¼�Ƿ�Ǽǲ����Ѿ���׼
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If Trim(rs("Approver")) <> "" Then
            ApprovalStatus = True
        Else
            ApprovalStatus = False
        End If
        'ͬʱ�ж�BOM�����Ƿ�BOM����(�ύ��)
        If InStr(Trim(rs("Submiter")), PDMUserName) Then
            OpennerSubmiter = True
        Else
            OpennerSubmiter = False
        End If
    End If
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    
    
    StrSql = "Select Top 1 isNull(CPCNNmbr,''),isNull(CPCNLocate,'') From BOMCPCN Where BOMID ='" & TxtFinsGdIndex.Text & "' Order by BOMVersion Desc"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        txtCPCNNO.Text = rs(0)
        txtCPCNlocate.Text = rs(1)
    End If
    rs.Close
    
    Conn.Close
    

    FinishGoodsNO = Trim(TxtFinsGdIndex.Text)
    MSFlexGrid1.Rows = 3   '����������
    MSFlexGrid1.RowHeight(1) = 225     '������б����ع��Ļ���Ҫ�ָ��и�Ĭ��ֵ
    MSFlexGrid1.RowHeight(2) = 225     '������б����ع��Ļ���Ҫ�ָ��и�Ĭ��ֵ
    Refresh_FlexGrid
    If BOMExist Then Refresh_FlexGridColumnSupplierPN
    NPOWorking = True
    Exit Sub

vbErrorHandler:
   MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdRunBOM_Click"
End Sub
Private Sub Refresh_FlexGridColumnSupplierPN()
    'On Error GoTo vbErrorHandler
    Dim RowSumVar As Integer
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    Dim Conn As New ADODB.Connection
    

    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    '�ȱ����ʼ��Col��Row�ŵ�ֵ
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    If RowNum > 1 Then
    For RowSumVar = 2 To RowNum
        'If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) = "NEW" Then   '�ж��ǲ���New Part����
        '�ҵ�GlueSupplier�еĶ�Ӧ12NC,������ȡ�ֶγ�Old
        If rs.State = adStateOpen Then rs.Close
        '������԰汾�ű��ʽΪ(Mid(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3))) - 1) & "0")
        rs.Open "Select * from GlueSupplier Where Glue12NC ='" & Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            If rs.State = adStateOpen Then rs.Close
            GoTo NextRowforSupplierPN
        Else
            MSFlexGrid1.Col = 21      'ָ����22��Supplier PN
            MSFlexGrid1.Row = RowSumVar
            MSFlexGrid1.Text = rs("SupplierPN")
        End If
        If rs.State = adStateOpen Then rs.Close
        'End If
NextRowforSupplierPN:
    Next RowSumVar
    End If
    MSFlexGrid1.Row = RowNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
    MSFlexGrid1.Col = ColNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
    
    'Exit Sub

'vbErrorHandler:
'    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:Refresh_FlexGridColumnSupplierPN"
End Sub
Private Sub Refresh_FlexGrid(Optional ByVal Applicant As String = "")
    Dim i As Integer                              '���������ʱ���е�ѭ������
    Dim J As Integer                              'MSFlexGrid�����������ı���
    Dim k As Integer                              'ÿ��������һ������ļ�������
    Dim m() As Integer                            'ÿ��������һ������ʱ,ÿ�㼶���е���������
    Dim MI As Integer                             'ÿ��������һ������ʱ,ÿ�㼶���е���������������ʱ����ɾ���仯����Ҫ�ı���
    Dim Z01 As Integer                            'һ��0��1�任�ı���.��K����������ڶ�λ����
    Dim TempParentName()  As String               'ÿ��������һ������ļ������Ӧ�ĸ���Name(ID)
    Dim TempForm As String                        '��ѯBOM����ʱʹ�õı������ַ�������
    
    Dim SERNmbr, TempSER As String
    
    
    Dim myCnn As New ADODB.Connection

    myCnn.Open connString
    
    '//����������¼��  rstCX��¼����Ӧ��ʱ��  rstCX2��¼����Ӧ��ʱ����һ����¼����Ķ�Ӧ�Ӽ�¼
    Dim rstCX As New ADODB.Recordset
    Dim rstCX2 As New ADODB.Recordset
    Dim rstCX3 As New ADODB.Recordset
    Set rstCX.ActiveConnection = myCnn
    Set rstCX2.ActiveConnection = myCnn
    Set rstCX3.ActiveConnection = myCnn
    
    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    
    '//�ж������ͼ���Ƿ�ײ���� �����û�и���ĵײ���������ʾ���˳�
    StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + FinishGoodsNO + "'"
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
    If Not rstCX.RecordCount > 0 Then
        MsgBox " This item is not assembly, has no Child", vbInformation, "System Info."
        rstCX.Close
        BOMExist = False   'BOMExist�����ж��Ƿ�ִ�� If BOMExist Then Refresh_FlexGridColumnSupplierPN
        Exit Sub
    Else
        BOMExist = True    'BOMExist�����ж��Ƿ�ִ�� If BOMExist Then Refresh_FlexGridColumnSupplierPN
    End If
    rstCX.Close
    
    '//������ʱ��ʾ��
    'myCnn.Execute "create table testTable (" & "colTime datetime NULL ," & _
    "colFlt float NULL ," & _
    "myImg image NULL  ," & _
    "myInt int NULL ," & _
    "myNText ntext COLLATE Chinese_PRC_CI_AS NULL )"
    
    '��������Ϊtb�ı�ʾ��
    'strSQL = "CREATE TABLE tb(" & "username varchar(20) not null primary key," & "pass char(10) not null)"
    
    
    
    '������ʱ��ǰ��Ҫ��һ���жϣ����ͬ�����������Ҫ����
    Dim temp12NC As String
    Dim temptbleID As Integer
    Dim TblExist As Boolean
    temptbleID = 0
    TempForm = "TempForm0"
    TblExist = False
    Do
        temptbleID = temptbleID + 1
        TempForm = Mid(TempForm, 1, Len(TempForm) - 1) & Trim(temptbleID)        'str=mid(str,1,len(str)-n) ɾ�����ұߵ�n���ַ�
        StrSql = "select * from sysobjects where name = '" & TempForm & "'"
        rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic   'rstCX��¼����ͼȡ��ʱ���м�¼
        If rstCX.EOF Then TblExist = rstCX.EOF
        rstCX.Close
    Loop Until TblExist
    
    StrSql = "CREATE TABLE " + TempForm + "(TempID INT IDENTITY(1,1) PRIMARY KEY CLUSTERED, Temp12NC varchar(12), TempLevel INT)"
    '����IDENTITY(1,1)���Զ����ӱ�Ŷ����1��ʼ��������1
    myCnn.Execute StrSql
    
    StrSql = "INSERT INTO " + TempForm + " VALUES (" + FinishGoodsNO + ",0)"         '0ΪTempLevel
    myCnn.Execute StrSql
    
    '//��ADODB�ؼ�������ʱ��TempForm��Ϣ
    StrSql = "SELECT * FROM " + TempForm + ""
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic   'rstCX��¼����Ӧ��ʱ��
    
    '//�ж���ʱ���¼�Ƿ�Ϊ��
    J = 1   'MSFlexGrid1���ӵ�1(ʵ���ǵ�2��)��ʼ
    k = 0   'ÿ��������һ������ļ�������, ��ʼֵΪ0
    Do While Not rstCX.EOF   'һֱѭ������ʱ���еļ�¼Ϊ0
        rstCX.MoveLast                  'ָ���ƶ�����ʱ���е�����¼
        Z01 = 0                         'һ��0��1�任�ı���.��K����������ڶ�λ����,û���ҵ�����Ϊ0
        
        '//�ж������Ƿ�������
        Set rstCX2 = Nothing
        If Applicant = "" Then
            StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + CStr(rstCX.Fields("Temp12NC")) + "' ORDER BY ChildID"
        Else
            StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + CStr(rstCX.Fields("Temp12NC")) + "' And ChildID in (SELECT LEFT(SglPrtIndex,11)+ '' + CAST(CAST(RIGHT(SglPrtIndex,1) AS INT)+SglPrtVer AS VARCHAR(12)) FROM SglPrt Where Applicant='" & Applicant & "') ORDER BY ChildID"
        End If

        rstCX2.Open StrSql, myCnn, adOpenStatic, adLockOptimistic 'rstCX2��¼����Ӧ��ʱ����һ����¼����Ķ�Ӧ�Ӽ�¼
        
        If rstCX2.RecordCount >= 1 Then
            Z01 = 1                         'һ��0��1�任�ı���.��K����������ڶ�λ����,�ҵ�����Ϊ1
            If k >= 1 Then m(k) = MI                      '��Kֵ(��������)����ǰ����ʣ���������
            k = k + 1                       '���������,��ôTempParentName������Ҫ����
            
            ReDim Preserve TempParentName(1 To k)
            TempParentName(k) = CStr(rstCX.Fields("Temp12NC"))
            
            ReDim Preserve m(1 To k)
            m(k) = rstCX2.RecordCount
            MI = m(k) + 1              '+1����Ϊ���ҵ����Ӽ�¼�⻹Ҫ��ɾ���ҳ����Ӽ�¼���ĸ���¼
            
            rstCX2.MoveFirst
            For i = 1 To rstCX2.RecordCount
                '//�����������ʱ����
                StrSql = "INSERT INTO " + TempForm + " VALUES ('" + CStr(rstCX2.Fields("ChildID")) + "'," + (CStr(rstCX.Fields("TempLevel") + 1)) + ")"
                myCnn.Execute (StrSql)
                rstCX2.MoveNext
            Next i
            rstCX2.MoveLast
        End If
        If rstCX.Fields("TempLevel") = 0 Then           'TempLevel= 0 ��Ϊ����Ŀ,���ݵ�����д
                '//��MSFlexGrid����������Ϣ
            MSFlexGrid1.TextMatrix(J, 0) = J        '����ÿ�е��к�
            MSFlexGrid1.TextMatrix(J, 1) = rstCX.Fields("TempLevel")
            MSFlexGrid1.TextMatrix(J, 2) = "Top"
            MSFlexGrid1.TextMatrix(J, 3) = CStr(rstCX.Fields("Temp12NC"))
            MSFlexGrid1.TextMatrix(J, 4) = 1            'Root���ڵ�(Finish Goods ��������Ϊ1)
            rstCX3.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(MSFlexGrid1.TextMatrix(J, 3)) & "'", myCnn, adOpenKeyset, adLockOptimistic
            If rstCX3.RecordCount > 0 Then
                MSFlexGrid1.TextMatrix(J, 5) = "Piece"     '���ڸ���Ŀ��Unit,����Ϊ"Piece"
                MSFlexGrid1.TextMatrix(J, 6) = Trim(rstCX3.Fields("Description"))
                
                TxtDescription.Text = Trim(rstCX3.Fields("Description"))
                ComboPJNOIndex.Text = FormatProjectCode(Trim(rstCX3.Fields("PJNOIndex")))
                ComboPjtName.Text = Trim(rstCX3.Fields("PjtName"))
                'cmbAuthor.Text = Trim(rstCX3.Fields("Applicant"))
                
                MSFlexGrid1.TextMatrix(J, 7) = Trim(rstCX3.Fields("ItemType"))
                MSFlexGrid1.TextMatrix(J, 8) = Trim(rstCX3.Fields("ProductLine"))
                MSFlexGrid1.TextMatrix(J, 9) = ""
                
                MSFlexGrid1.TextMatrix(J, 10) = ""         '���ڸ���Ŀ��ChangeStatus,����Ϊ��
                
                If Not IsNull(Trim(rstCX3.Fields("SERLocate"))) Then
                    TempSER = Mid(Replace(Trim(rstCX3.Fields("SERLocate")), "----", ""), 32, 5)
                    If TempSER = "EASE " Then
                        TempSER = "RELEASREPORT"
                    Else
                        If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "��" Then
                            TempSER = "SER00000" & left(TempSER, 4)
                        Else
                            TempSER = "SER0000" & TempSER
                        End If
                    End If
                Else
                    TempSER = ""
                End If
                
                If IsNull(rstCX3.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ rstCX3.Fields("SERNmbr") = Null
                    If TempSER <> "" Then
                        MSFlexGrid1.TextMatrix(J, 12) = TempSER
                    Else
                        MSFlexGrid1.TextMatrix(J, 12) = ""
                    End If
                    SERNmbr = ""
                Else
                    If TempSER <> "" Then
                        MSFlexGrid1.TextMatrix(J, 12) = TempSER
                    Else
                        MSFlexGrid1.TextMatrix(J, 12) = rstCX3.Fields("SERNmbr")
                    End If
                    SERNmbr = Trim(rstCX3.Fields("SERNmbr"))
                End If
                
                '����SER
                If SERNmbr <> TempSER Then Call UpdateSER(TempSER, Trim(MSFlexGrid1.TextMatrix(J, 3)), "FinsGd")
                
                If IsNull(rstCX3.Fields("CommtNote")) Then
                    MSFlexGrid1.TextMatrix(J, 13) = ""
                Else
                    MSFlexGrid1.TextMatrix(J, 13) = Trim(rstCX3.Fields("CommtNote"))
                End If
                
            End If
            If rstCX3.State = adStateOpen Then rstCX3.Close
        Else
            temp12NC = rstCX.Fields("Temp12NC")
            rstCX2.Close
            Set rstCX2 = Nothing
            StrSql = "SELECT * FROM BOMOrigData WHERE ChildID = '" + Trim(CStr(rstCX.Fields("Temp12NC"))) + "' and  ParentID = '" + Trim(TempParentName(k - Z01)) + "'"  '(K - Z01)û���ҵ��������0,�ҵ��������1
            rstCX2.Open StrSql, myCnn, adOpenStatic, adLockOptimistic 'rstCX2��¼����Ӧ��ʱ����һ����¼����Ķ�Ӧ�Ӽ�¼
            If rstCX2.RecordCount > 0 Then
                MSFlexGrid1.TextMatrix(J, 0) = J        '����ÿ�е��к�
                MSFlexGrid1.TextMatrix(J, 1) = rstCX.Fields("TempLevel")
                MSFlexGrid1.TextMatrix(J, 3) = temp12NC
            
                MSFlexGrid1.TextMatrix(J, 2) = rstCX2.Fields("ParentID")
                MSFlexGrid1.TextMatrix(J, 4) = rstCX2.Fields("Quantity")
                If IsNull(rstCX2.Fields("ChgStatus")) Then          '������IsNull�����ж�,������ rstCX2.Fields("ChgStatus") = Null
                    MSFlexGrid1.TextMatrix(J, 10) = ""
                Else
                    MSFlexGrid1.TextMatrix(J, 10) = rstCX2.Fields("ChgStatus")
                End If
                
                rstCX3.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(J, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(J, 3))) - 1) & "0") & "'", myCnn, adOpenKeyset, adLockOptimistic
                If rstCX3.RecordCount > 0 Then
                    MSFlexGrid1.TextMatrix(J, 5) = Trim(rstCX3.Fields("PrtUnit"))
                    MSFlexGrid1.TextMatrix(J, 6) = Trim(rstCX3.Fields("Description"))
                    MSFlexGrid1.TextMatrix(J, 7) = Trim(rstCX3.Fields("ItemType"))
                    MSFlexGrid1.TextMatrix(J, 8) = Trim(rstCX3.Fields("ProductLine"))
                    MSFlexGrid1.TextMatrix(J, 9) = get12NCStatusIsExisting(MSFlexGrid1.TextMatrix(J, 3), MSFlexGrid1.TextMatrix(J, 2), (rstCX3.Fields("NewOldStatus")))
                    
                    If Not IsNull(Trim(rstCX3.Fields("SERLocate"))) Then
                        TempSER = Mid(Replace(Trim(rstCX3.Fields("SERLocate")), "----", ""), 32, 5)
                        If TempSER = "EASE " Then
                            TempSER = "RELEASREPORT"
                        Else
                            If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "��" Then
                                TempSER = "SER00000" & left(TempSER, 4)
                            Else
                                TempSER = "SER0000" & TempSER
                            End If
                        End If
                    Else
                        TempSER = ""
                    End If
    
                    
                    If IsNull(rstCX3.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ rstCX3.Fields("SERNmbr") = Null
                        If TempSER <> "" Then
                            MSFlexGrid1.TextMatrix(J, 11) = TempSER
                        Else
                            MSFlexGrid1.TextMatrix(J, 11) = ""
                        End If
                        SERNmbr = ""
                    Else
                        If TempSER <> "" Then
                            MSFlexGrid1.TextMatrix(J, 11) = TempSER
                        Else
                            MSFlexGrid1.TextMatrix(J, 11) = rstCX3.Fields("SERNmbr")
                        End If
                        SERNmbr = Trim(rstCX3.Fields("SERNmbr"))
                    End If
                    
                    '����SER
                    If SERNmbr <> TempSER Then Call UpdateSER(TempSER, (Mid(Trim(MSFlexGrid1.TextMatrix(J, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(J, 3))) - 1) & "0"), "SglPrt")
                    
                    If IsNull(rstCX3.Fields("CommtNote")) Then
                        MSFlexGrid1.TextMatrix(J, 12) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 12) = Trim(rstCX3.Fields("CommtNote"))
                    End If
                    
                    If IsNull(rstCX3.Fields("ETA")) Then
                        MSFlexGrid1.TextMatrix(J, 14) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 14) = Format(rstCX3.Fields("ETA"), "YYYY/MM/DD")
                    End If
                    
                    If IsNull(rstCX3.Fields("SampleQuantity")) Then
                        MSFlexGrid1.TextMatrix(J, 15) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 15) = Trim(rstCX3.Fields("SampleQuantity"))
                    End If
                    
                    If IsNull(rstCX3.Fields("Supplier")) Then
                        MSFlexGrid1.TextMatrix(J, 16) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 16) = Trim(rstCX3.Fields("Supplier"))
                    End If
                    
                    If IsNull(rstCX3.Fields("ArrivalDate")) Then
                        MSFlexGrid1.TextMatrix(J, 17) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 17) = Format(rstCX3.Fields("ArrivalDate"), "YYYY/MM/DD")
                    End If
                    
                    If IsNull(rstCX3.Fields("RoHsReport")) Then
                        MSFlexGrid1.TextMatrix(J, 18) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 18) = Trim(rstCX3.Fields("RoHsReport"))
                    End If
                    
                    If IsNull(rstCX3.Fields("LLTI")) Then
                        MSFlexGrid1.TextMatrix(J, 19) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 19) = Trim(rstCX3.Fields("LLTI"))
                    End If
                    
                    If IsNull(rstCX3.Fields("RiskOrderQ")) Then
                        MSFlexGrid1.TextMatrix(J, 20) = ""
                    Else
                        MSFlexGrid1.TextMatrix(J, 20) = Trim(rstCX3.Fields("RiskOrderQ"))
                    End If
                End If
                If rstCX3.State = adStateOpen Then rstCX3.Close
            Else
                J = J - 1
            End If
        End If
        
        '// ɾ����ʱ���м�¼
        StrSql = "DELETE FROM " + TempForm + " WHERE Temp12NC = '" + CStr(rstCX.Fields("Temp12NC")) + "' and TempLevel = '" + CStr(k - Z01) + "'"    'һ��Ҫ����and TempLevel = '" + cStr(K - Z01) + "'" ��������level�Ĳ�ͬparent����ĿҲһͬ��ɾ����
        myCnn.Execute (StrSql)
        MI = MI - 1                  'ÿ�δ���ʱ����ɾ��һ���,���������ݼ�1
        
        If k - 1 >= 1 And Z01 = 1 Then         '��ɾ���ļ�¼��������һ���(��һ���Ǳ�����2����)
            m(k - 1) = m(k - 1) - 1                '������һ�����������Ҳ�ݼ�1
        End If
        
        If MI = 0 Then                          'MI = 0 ��ʾÿ�㼶���е���������ȫ�����ʱ (Ҫ������һ��ʱ)
KMinus2:
            k = k - 1                               'Ҫ������һ��ʱ,����������1
            If k >= 1 Then                       '��������С��1,�±�Խ��,����Ҫ�˳�
                If m(k) = 0 Then GoTo KMinus2       '�����һ��������ҲΪ0,������ݼ�������һ��
                MI = m(k)                           '�����������һ��(��������һ��)��ʣ�µ��Ӽ�¼����ֵ��MI
            Else
                k = k + 1
            End If
        End If
        
        Set rstCX = Nothing         '��¼��rstCX���
        StrSql = "SELECT * FROM " + TempForm + ""     '��¼��rstCX��պ����¼��������һ���Ӽ�¼����ʱ�����Ŀ
        rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
        J = J + 1
        MSFlexGrid1.Rows = J + 1
    Loop
    'ɾ����ʱ�ñ�
    StrSql = "Drop TABLE " + TempForm + ""
    myCnn.Execute StrSql
    
    rstCX.Close
    rstCX2.Close
    RowNum = J
    
    MSFlexGrid_ChgStatus_HightlightRow (10)                   '��ChangeStatus��11���������ݵ�������Ϊ��ɫ
    MSFlexGrid_NewOld_HightlightRow (9)                       '��New/Old��9����������ΪNew��������Ϊ�ۺ�ɫ
    'MSFlexGridColumnColorChange MSFlexGrid1, 9, j             '����New/Old��(��9��)Ϊǳ�ۺ�ɫ
    'MSFlexGridColumnColorChange MSFlexGrid1, 12, j            '����Comments��(��13��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 21, J            '����Supplier PN��(��22��)Ϊǳ�ۺ�ɫ
    MSFlexGrid_ApproveStatus_HightlightRow (ApprovalStatus)   '�Ե�1������Ϊ��ɫ������Ѿ���׼��BOM
    MSFlexGridColumnColorChange MSFlexGrid1, 13, J, &H404040  '���÷ָ���(��14��)Ϊ��ɫ
    MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 0                  '���õ�Ԫ��λ��ȡ������ı亯���е�ĳ�и�����ʾ
    
End Sub

Private Sub MSFlexGrid_ChgStatus_HightlightRow(ByVal CheckColNO As Integer)
    Dim RowSumVar As Integer
    For RowSumVar = 1 To RowNum
        If MSFlexGrid1.TextMatrix(RowSumVar, CheckColNO) <> "" Then
            MSFlexGridRowColorChange MSFlexGrid1, RowSumVar, MSFlexGrid1.Cols
        End If
    Next RowSumVar
End Sub

Private Sub MSFlexGrid_NewOld_HightlightRow(ByVal CheckColNO As Integer)
    Dim RowSumVar As Integer
    For RowSumVar = 1 To RowNum
        If UCase(left(Trim(MSFlexGrid1.TextMatrix(RowSumVar, CheckColNO)), 1)) = "N" Then
            MSFlexGridRowColorChange MSFlexGrid1, RowSumVar, MSFlexGrid1.Cols, &HFFC0FF      '&HFFC0FFΪ�ۺ�ɫ
        End If
    Next RowSumVar
End Sub

Private Sub MSFlexGrid_ApproveStatus_HightlightRow(ByVal ApproverOK As Boolean)
    
    If ApproverOK Then
        MSFlexGridRowColorChange MSFlexGrid1, 1, MSFlexGrid1.Cols, &H80FF80     '&H80FF80Ϊ��ɫ
    End If
    
End Sub

Private Sub ComboPJNOIndex_Click()
    ComboPjtName.ListIndex = ComboPJNOIndex.ListIndex
End Sub

Private Sub ComboPjtName_Click()
    ComboPJNOIndex.ListIndex = ComboPjtName.ListIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    FromForm.Show 0
End Sub

Private Sub MSFlexGrid1_Click()
    'On Error Resume Next
    Dim temp12NC As String
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    Dim Conn As New ADODB.Connection
    

    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    txtNodeSglPrt12NC = ""            '�����ԭ��������
    txtNodeDescription = ""
    txtNodePrtUnit = ""
    txtNodeDrwlocate = ""
    txtSERNO = ""
    txtSERlocate = ""
    
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    
    If MSFlexGrid1.Row = RowNum Then Exit Sub   '��������һ���������˳�
    
    If MSFlexGrid1.Row = 1 Then                 'MSFlexGrid1.Row = 1 ��ʾ��ȡ����Finish Goods
        MSFlexGrid1.Col = 3
        temp12NC = MSFlexGrid1.Text            '��3���е�FinishGoods��12NC��ֵ
        rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(temp12NC) & "'", Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            txtNodeSglPrt12NC = rs("FinsGdIndex")
            txtNodeDescription = rs("Description") & ""
            txtNodePrtUnit = "Piece"
            If IsNull(rs("Drwlocate")) Then
                txtNodeDrwlocate = ""
            Else
                txtNodeDrwlocate = rs("Drwlocate") & ""
            End If
            
            If IsNull(rs("SERNmbr")) Then
                txtSERNO = ""
            Else
                txtSERNO = rs("SERNmbr") & ""
            End If
            
            If IsNull(rs("SERlocate")) Then
                txtSERlocate = ""
            Else
                txtSERlocate = rs("SERlocate") & ""
            End If
            
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
        
    Else
        MSFlexGrid1.Col = 3
        temp12NC = MSFlexGrid1.Text
        temp12NC = Mid(temp12NC, 1, (Len(temp12NC) - 1)) & "0"
        If Not Isnum(temp12NC) Then Exit Sub      '�����ȡ���Ǳ�����,��temp12NC��������,�������������
        rs.Open "Select * from SglPrt where SglPrtIndex ='" & temp12NC & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            
            txtNodeSglPrt12NC = left(rs("SglPrtIndex"), 11) + CStr(CInt(right(rs("SglPrtIndex"), 1)) + rs("SglPrtVer"))
            txtNodeDescription = rs("Description") & ""
            txtNodePrtUnit = rs("PrtUnit") & ""
            
            If IsNull(rs("Drwlocate")) Then
                txtNodeDrwlocate = ""
            Else
                txtNodeDrwlocate = rs("Drwlocate") & ""
            End If
            
            If IsNull(rs("SERNmbr")) Then
                txtSERNO = ""
            Else
                txtSERNO = rs("SERNmbr") & ""
            End If
            
            If IsNull(rs("SERlocate")) Then
                txtSERlocate = ""
            Else
                txtSERlocate = rs("SERlocate") & ""
            End If
            
            
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
    End If
    
    MSFlexGrid1.Row = RowNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
    MSFlexGrid1.Col = ColNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
    
    Select Case MSFlexGrid1.Col
    Case 9, 12, 21, 22                          '��9��New/Old,��13��Comment/Note,��22��Supplier PartNumber
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        If SystemAdmin = "Y" Then
            'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
            GoTo AdminGoAhead1
        End If
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        
        If ApprovalStatus Then
            MsgBox "The BOM was already approved, Please Re-submit if you want to update", vbInformation, "System Info."
            Exit Sub
        End If
        
        If Not OpennerSubmiter Then
            MsgBox "You are not the BOM Author, No Right to update" & vbCrLf & " Or Check if NPO is Openned in Project Search", vbInformation, "System Info."
            Exit Sub
        End If
AdminGoAhead1:
        If MSFlexGrid1.Row = 1 Then
RowHeighZeroContinue1:
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            If MSFlexGrid1.Row = RowNum Then Exit Sub
            If MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0 Then GoTo RowHeighZeroContinue1
            
            MsgBox "The Root(Top) Item is not Editable" & vbCrLf & "Please Edit from next Row", vbInformation, "System Info."
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(TxtFinsGdIndex.Text)
        
    Case 15, 16, 18, 19, 20, 22                           '��15����20��Ϊ�ɹ���д
        
        If Not NPOWorking Then
            MsgBox "Current status is not New Part OverView, Can NOT Edit", vbInformation, "System Info."
            Exit Sub
        End If
        
        rs.Open "Select * from Users where Name ='" & PDMUserName & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
            If SystemAdmin <> "Y" Then
                'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
                GoTo AdminGoAhead2
            End If
            '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
            If Trim(rs("UserGroup")) <> "�ɹ���" Then
                MsgBox "You are not Purchasing Department Employee. Unable to Edit Purchasing Department Content", vbInformation, "System Info."
                Exit Sub
            End If
        End If
AdminGoAhead2:
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
        
        If MSFlexGrid1.Row = 1 Then
RowHeighZeroContinue2:
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            If MSFlexGrid1.Row = RowNum Then Exit Sub
            If MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0 Then GoTo RowHeighZeroContinue2
            
            MsgBox "The Root(Top) Item is not Editable" & vbCrLf & "Please Edit from next Row", vbInformation, "System Info."
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(TxtFinsGdIndex.Text)
        
    Case 14, 17                            '��14��,17��Ϊ�ɹ���д
        
        If Not NPOWorking Then
            MsgBox "Current status is not New Part OverView, Can NOT Edit", vbInformation, "System Info."
            Exit Sub
        End If
        
        rs.Open "Select * from Users where Name ='" & PDMUserName & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
            If SystemAdmin <> "Y" Then
                'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
                GoTo AdminGoAhead3
            End If
            '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
            If Trim(rs("usergroup")) <> "�ɹ���" Then
                MsgBox "You are not Purchasing Department Employee. Unable to Edit Purchasing Department Content", vbInformation, "System Info."
                Exit Sub
            End If
        End If
AdminGoAhead3:
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
        
        If MSFlexGrid1.Row = 1 Then
RowHeighZeroContinue3:
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            If MSFlexGrid1.Row = RowNum Then Exit Sub
            If MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0 Then GoTo RowHeighZeroContinue3
            
            MsgBox "The Root(Top) Item is not Editable" & vbCrLf & "Please Edit from next Row", vbInformation, "System Info."
        End If
        DTPicker1.Visible = True
        DTPicker1.Width = MSFlexGrid1.CellWidth
        DTPicker1.Height = MSFlexGrid1.CellHeight
        DTPicker1.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        DTPicker1.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        DTPicker1.SetFocus
        DTPicker1.Value = Date
        
    Case Else
        'MsgBox " This column is not editable", vbInformation, "System Info."
        Exit Sub
    End Select
    
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid1EditText_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error GoTo vbErrorHandler
    'Set scr = CreateObject("MSScriptControl.ScriptControl")
    'scr.Language = "vbscript"
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    Dim Conn As New ADODB.Connection
    

    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    If MSFlexGrid1.Col <> 9 And MSFlexGrid1.Col <> 12 And MSFlexGrid1.Col <> 15 And MSFlexGrid1.Col <> 16 And MSFlexGrid1.Col <> 18 And MSFlexGrid1.Col <> 19 And MSFlexGrid1.Col <> 20 And MSFlexGrid1.Col <> 21 And MSFlexGrid1.Col <> 22 Then Exit Sub
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    
    Select Case KeyCode
    Case vbKeyEscape
        MSFlexGrid1EditText.Visible = False
        'MSFlexGrid1.SetFocus
        Exit Sub
    Case vbKeyReturn
        'MSFlexGrid1.Text = scr.Eval(TxtFinsGdIndex.Text)                                 '��ScriptControl������������ʽ
        
        MSFlexGrid1.Text = Trim(MSFlexGrid1EditText.Text)
        
        Select Case ColNoTemp
        Case 9
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                If UCase(left(Trim(MSFlexGrid1.Text), 1)) = "N" Then
                    MSFlexGrid1.Text = "New"
                    rs("NewOldStatus") = "New"
                Else
                    If Trim(MSFlexGrid1.Text) = "" Then
                        MSFlexGrid1.Text = ""
                        rs("NewOldStatus") = ""
                    Else
                        MSFlexGrid1.Text = "Old"
                        rs("NewOldStatus") = "Old"
                    End If
                End If
                rs.Update
                If UCase(left(Trim(MSFlexGrid1.Text), 1)) = "N" Then
                    MSFlexGridRowColorChange MSFlexGrid1, MSFlexGrid1.Row, MSFlexGrid1.Cols, &HFF80FF    '&HFF80FFΪ�ۺ�ɫ
                End If
                MSFlexGrid1.Row = RowNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
                MSFlexGrid1.Col = ColNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 12
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("CommtNote") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 15
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("SampleQuantity") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 16
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("Supplier") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 18
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("RoHsReport") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 19
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                If UCase(left(Trim(MSFlexGrid1.Text), 1)) = "Y" Then
                    MSFlexGrid1.Text = "Yes"
                    rs("LLTI") = "Yes"
                Else
                    If Trim(MSFlexGrid1.Text) = "" Then
                        MSFlexGrid1.Text = ""
                        rs("LLTI") = ""
                    Else
                        MSFlexGrid1.Text = "No"
                        rs("LLTI") = "No"
                    End If
                End If
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 20
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("RiskOrderQ") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 21
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '������԰汾�ű��ʽΪ(Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0")
            rs.Open "Select * from GlueSupplier Where Glue12NC ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is still not existing in Glue/Electro Database" & vbCrLf & "Please Register the Part 12NC in Glue/Electro Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("SupplierPN") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            
        Case 22 '############����Remark##########
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '������԰汾�ű��ʽΪ(Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0")
            rs.Open "Select * from BOMOrigData Where ParentID='" & MSFlexGrid1.TextMatrix(RowNoTemp, 2) & "' AND ChildID ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is still not existing in Database" & vbCrLf, vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("Remark") = Trim(MSFlexGrid1.Text)
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case Else
        End Select
        
        If MSFlexGrid1.Row < RowNum Then                                         '��Ҫ��һ���ж��Ƿ񳬳����ֵ
RowHeighZeroGoAhead1:
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            If MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0 Then GoTo RowHeighZeroGoAhead1
            
        Else
            MSFlexGrid1EditText.Visible = False
            Exit Sub
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(TxtFinsGdIndex.Text)
    End Select
'    Exit Sub
'vbErrorHandler:
'    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMNPO:MSFlexGrid1EditText_KeyDown"
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo vbErrorHandler
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    If MSFlexGrid1.Col <> 14 And MSFlexGrid1.Col <> 17 Then Exit Sub
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    
    Select Case KeyCode
    Case vbKeyEscape
        DTPicker1.Visible = False
        Exit Sub
    Case vbKeyReturn
        MSFlexGrid1.Text = Format(DTPicker1.Value, "YYYY/MM/DD")
        
        Select Case ColNoTemp
        Case 14
            If MSFlexGrid1.Row = RowNum Then
                DTPicker1.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                If IsNull(DTPicker1.Value) Then
                    rs("ETA") = Null
                Else
                    rs("ETA") = MSFlexGrid1.Text
                End If
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 17
            If MSFlexGrid1.Row = RowNum Then
                DTPicker1.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                If IsNull(DTPicker1.Value) Then
                    rs("ArrivalDate") = Null
                Else
                    rs("ArrivalDate") = MSFlexGrid1.Text
                End If
                rs.Update
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            
        Case Else
        End Select
        
        If MSFlexGrid1.Row < RowNum Then                                         '��Ҫ��һ���ж��Ƿ񳬳����ֵ
RowHeighZeroGoAhead2:
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            If MSFlexGrid1.Row = RowNum Then
                DTPicker1.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            If MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0 Then GoTo RowHeighZeroGoAhead2
            
        Else
            DTPicker1.Visible = False
            Exit Sub
        End If
        DTPicker1.Visible = True
        DTPicker1.Width = MSFlexGrid1.CellWidth
        DTPicker1.Height = MSFlexGrid1.CellHeight
        DTPicker1.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        DTPicker1.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        DTPicker1.SetFocus
        DTPicker1.Value = Date
    End Select
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMNPO:DTPicker1_Change"
End Sub

Private Sub TxtFinsGdIndex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then CmdRunBOM_Click
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Private Sub ComboPJNOIndex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim PJNOIndex As String
    PJNOIndex = Trim(ComboPJNOIndex.Text)
    If KeyCode = vbKeyReturn Then
        ComboPJNOIndex.Clear
        ComboPjtName.Clear
        If Not Isnum(PJNOIndex) Then Exit Sub

        Dim Conn As New ADODB.Connection
        Dim StrSql As String

        Conn.Open connString
        
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        StrSql = "Select PJNOIndex, Description from PJNO Where PJNOIndex ='" & PJNOIndex & "' Order by PJNOIndex"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic

        Do While Not rs.EOF
            ComboPJNOIndex.AddItem (FormatProjectCode(Trim(CStr(rs(0))))) 'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
            ComboPjtName.AddItem (Trim(rs(1))) 'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
            rs.MoveNext
        Loop
        ComboPJNOIndex.ListIndex = 0
        ComboPjtName.ListIndex = 0
        rs.Close
        Set rs = Nothing
        Conn.Close
        Set Conn = Nothing
    End If
End Sub
Private Function FormatProjectCode(ByVal PJNOIndex As String) As String
    Dim i As Integer
    FormatProjectCode = PJNOIndex
    For i = 1 To 6 - Len(PJNOIndex)
        FormatProjectCode = "0" & FormatProjectCode
    Next
End Function

Private Sub ComboPjtName_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    Dim PjtName As String
    PjtName = Trim(ComboPjtName.Text)
    If KeyCode = vbKeyReturn Then
        ComboPJNOIndex.Clear
        ComboPjtName.Clear
        
        Dim Conn As New ADODB.Connection
        Dim StrSql As String

        Conn.Open connString
        
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        StrSql = "Select PJNOIndex, Description from PJNO Where Description like '" & PjtName & "%' Order by PJNOIndex"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        Do While Not rs.EOF
            ComboPJNOIndex.AddItem (Trim(rs(0))) 'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
            ComboPjtName.AddItem (Trim(rs(1))) 'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
            rs.MoveNext
        Loop
        ComboPJNOIndex.ListIndex = 0
        ComboPjtName.ListIndex = 0
        rs.Close
        Set rs = Nothing
        Conn.Close
        Set Conn = Nothing
    End If
End Sub
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Private Sub CmdExit_Click()
    Unload Me
    FromForm.Show 0
End Sub

Private Sub Form_Resize()
    'ȷ������ı�ʱ�ؼ���֮�ı�
    Resize_ALL Me
End Sub
Private Sub Form_Load()
    'Load Skin & Format Control
    'LoadSkin Me
    '''Call ResizeInit(Me)
    MSFlexGrid1.Rows = 3   '����������
    MSFlexGrid1.Cols = 23   '����������
    
    MSFlexGrid1.ColAlignment(0) = 3     '()��Ϊ�еı��
    MSFlexGrid1.ColAlignment(1) = 1
    MSFlexGrid1.ColAlignment(4) = 1
    MSFlexGrid1.ColAlignment(5) = 1
    MSFlexGrid1.ColAlignment(6) = 1
    MSFlexGrid1.ColAlignment(7) = 1
    MSFlexGrid1.ColAlignment(8) = 1
    MSFlexGrid1.ColAlignment(9) = 1
    MSFlexGrid1.ColAlignment(10) = 1
    MSFlexGrid1.ColAlignment(11) = 1
    MSFlexGrid1.ColAlignment(12) = 1
    MSFlexGrid1.ColAlignment(13) = 1
    MSFlexGrid1.ColAlignment(14) = 1
    MSFlexGrid1.ColAlignment(15) = 1
    MSFlexGrid1.ColAlignment(16) = 1
    MSFlexGrid1.ColAlignment(17) = 1
    MSFlexGrid1.ColAlignment(18) = 1
    MSFlexGrid1.ColAlignment(19) = 1
    MSFlexGrid1.ColAlignment(20) = 1
    MSFlexGrid1.ColAlignment(21) = 1
    MSFlexGrid1.ColAlignment(22) = 1
    'flexAlignLeftTop 0 ��Ԫ��������󡢶������롣
    'flexAlignLeftCenter 1 �ַ�����ȱʡ���뷽ʽ����Ԫ��������󡢾��ж��롣
    'flexAlignLeftBottom 2 ��Ԫ��������󡢵ײ����롣
    'flexAlignCenterTop 3 ��Ԫ������ݾ��С��������롣
    'flexAlignCenterCenter 4 ��Ԫ������ݾ��С����ж��롣
    'flexAlignCenterBottom 5 ��Ԫ������ݾ��С��ײ����롣
    'flexAlignRightTop 6 ��Ԫ��������ҡ��������롣
    'flexAlignRightCenter 7 ��ֵ��ȱʡ���뷽ʽ����Ԫ��������ҡ����ж��롣
    'flexAlignRightBottom 8 ��Ԫ��������ҡ��ײ����롣
    'flexAlignGeneral 9 ��Ԫ������ݰ�һ�㷽ʽ���ж��롣�ַ��������󡢾��С���ʾ�����ְ����ҡ����С���ʾ��
    
    
    'Load User information
    Dim Conn As New ADODB.Connection
    Dim StrSql As String
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "Select [Name] from Users Order by [Name]"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        cmbAuthor.AddItem (rs(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
    
    MSFlexGridTileInitialize
End Sub

Private Sub MSFlexGridTileInitialize()
    MSFlexGrid1.ColWidth(0) = 12 * 25 * 2
    MSFlexGrid1.ColWidth(1) = 12 * 25 * 1.8
    MSFlexGrid1.ColWidth(2) = 12 * 25 * 4
    MSFlexGrid1.ColWidth(3) = 12 * 25 * 4
    MSFlexGrid1.ColWidth(4) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(5) = 12 * 25 * 2.5
    MSFlexGrid1.ColWidth(6) = 12 * 25 * 6
    MSFlexGrid1.ColWidth(7) = 12 * 25 * 2.8
    MSFlexGrid1.ColWidth(8) = 12 * 25 * 3.6
    MSFlexGrid1.ColWidth(9) = 12 * 25 * 2.4
    MSFlexGrid1.ColWidth(10) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(11) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(12) = 12 * 25 * 4
    MSFlexGrid1.ColWidth(13) = 0
    MSFlexGrid1.ColWidth(14) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(15) = 12 * 25 * 3.6
    MSFlexGrid1.ColWidth(16) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(17) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(18) = 12 * 25 * 3.6
    MSFlexGrid1.ColWidth(19) = 12 * 25 * 3.2
    MSFlexGrid1.ColWidth(20) = 12 * 25 * 2
    MSFlexGrid1.ColWidth(21) = 12 * 25 * 4
    MSFlexGrid1.ColWidth(22) = 12 * 25 * 4
    
    MSFlexGrid1.TextMatrix(0, 0) = "Index"
    MSFlexGrid1.TextMatrix(0, 1) = "Level"
    MSFlexGrid1.TextMatrix(0, 2) = "Parent12NC"
    MSFlexGrid1.TextMatrix(0, 3) = "Child12NC"
    MSFlexGrid1.TextMatrix(0, 4) = "Quantity"
    MSFlexGrid1.TextMatrix(0, 5) = "PrtUnit"
    MSFlexGrid1.TextMatrix(0, 6) = "Description"
    MSFlexGrid1.TextMatrix(0, 7) = "ItemType"
    MSFlexGrid1.TextMatrix(0, 8) = "ProductLine"
    MSFlexGrid1.TextMatrix(0, 9) = "New/Old/Existing"
    MSFlexGrid1.TextMatrix(0, 10) = "ChgStatus"
    MSFlexGrid1.TextMatrix(0, 11) = "SER NO."
    MSFlexGrid1.TextMatrix(0, 12) = "Note"
    MSFlexGrid1.TextMatrix(0, 13) = ""
    MSFlexGrid1.TextMatrix(0, 14) = "ETA"
    MSFlexGrid1.TextMatrix(0, 15) = "SampleQuantity"
    MSFlexGrid1.TextMatrix(0, 16) = "Supplier"
    MSFlexGrid1.TextMatrix(0, 17) = "ArrivalDate"
    MSFlexGrid1.TextMatrix(0, 18) = "RoHsReport"
    MSFlexGrid1.TextMatrix(0, 19) = "LLTI"
    MSFlexGrid1.TextMatrix(0, 20) = "RiskOrder Q"
    MSFlexGrid1.TextMatrix(0, 21) = "Supplier PN"
    MSFlexGrid1.TextMatrix(0, 22) = "Remark"
End Sub

Private Function get12NCStatusIsExisting(s12NC As String, sParentId As String, sStatus As String)

    Dim Conn As New ADODB.Connection
    Dim sReturn As String
    Dim iIndex As Integer
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    StrSql = "select * from SER where SglPrtNo = '" & s12NC & "' and (CAorA='FA' or CAorA='A')"
    Set rs1.ActiveConnection = Conn
    rs1.Open StrSql, Conn, adOpenStatic, adLockOptimistic
    If rs1.RecordCount > 0 Then
        sReturn = "Old"
    Else
        'new part�ŷ�existing
        StrSql = "select * from BOMOrigData where childid = '" & s12NC & "' And ParentId<>'" & sParentId & "'"
        rs.Open StrSql, Conn, adOpenStatic, adLockOptimistic
        If rs.RecordCount > 0 Then
            StrSql = "select * from BOMOrigData where childid = '" & s12NC & "' and parentid = '" & sParentId & "' and [index]=(select top 1 [Index] from BOMOrigData where childid =  '" & s12NC & "'  order by [Index])"
            rs2.Open StrSql, Conn, adOpenStatic, adLockOptimistic
            '��һ��������BOM����Ȼ��NEW
            If rs2.RecordCount > 0 Then
                sReturn = "New"
            Else
                sReturn = "Existing"
            End If
            rs2.Close
        Else
            sReturn = sStatus
        End If
        rs.Close
    End If
    rs1.Close
    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set Conn = Nothing
    get12NCStatusIsExisting = sReturn
End Function

Private Sub UpdateSER(ByVal SERNO As String, PartNo As String, TableName As String)
    Dim Conn As New ADODB.Connection
    
    
    
    Conn.Open connString
    If TableName = "SglPrt" Then
        StrSql = "Update " & TableName & "  Set SERNmbr='" & SERNO & "' where SglPrtIndex='" & PartNo & "'"
    ElseIf TableName = "FinsGd" Then
        StrSql = "Update " & TableName & "  Set SERNmbr='" & SERNO & "' where FinsGdIndex='" & PartNo & "'"
    End If
    Conn.Execute StrSql
    Conn.Close
    Set Conn = Nothing
End Sub

