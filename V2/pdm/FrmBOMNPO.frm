VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmBOMNPO 
   Caption         =   "PDM-BOM NPO(New Part Overview) 工程管理子系统"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13752
   Icon            =   "FrmBOMNPO.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10800
   ScaleWidth      =   13752
   StartUpPosition =   2  '屏幕中心
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
         Name            =   "宋体"
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
Private ApprovalStatus As Boolean    '定义BOM是否批准的标记
Private OpennerSubmiter As Boolean   '定义BOM打开者是否作者(提交者)的标记
Private NPOWorking As Boolean        '定义每一行的NewOld状态是否为New的标记
Private BOMExist As Boolean          '定义Refresh_FlexGrid模块中找BOM是否存在的标记
Private StrSql As String
'Private scr As Object         '如果用专用的表达式函数的话这个定义就用不上

'Mid(myNode, 2, Len(myNode))    '去掉最左边一个字符
'Mid(myNode, 1, Len(myNode)-1)  '去掉最右边一个字符

Private Sub MSFlexGridColumnColorChange(MSFlexGridName As Object, ByVal ColNo As Integer, ByVal RowSum As Integer, Optional ByVal ColColor As Long = &HC0E0FF)      '&HC0E0FF为浅桔红色
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Col = ColNo                    '从第ColNo列第0行开始
    MSFlexGridName.Row = 0                        '从第ColNo列第0行开始
    MSFlexGridName.RowSel = RowSum - 1            '高亮选中直到最后一行RowSum
    MSFlexGridName.CellBackColor = ColColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub
Private Sub MSFlexGridRowColorChange(MSFlexGridName As Object, ByVal RowNo As Integer, ByVal ColSum As Integer, Optional ByVal RowColor As Long = &HFFFF&)          '&HFFFF为黄色
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Row = RowNo                    '从第RowNo行第0列开始
    MSFlexGridName.Col = 1                        '从第RowNo行第0列开始
    MSFlexGridName.ColSel = ColSum - 1            '高亮选中直到最后一列ColSum
    MSFlexGridName.CellBackColor = RowColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub

'判断一个字符串中是否含有数字  用IsNumeric判断0000d031为真(当成double型数字)
Private Function Isnum(Str As String) As Boolean
    Isnum = True
    Dim i  As Integer
    For i = 1 To Len(Str)
        Select Case Mid(Str, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            'Isnum = True  这里写Isnum = True就出错,因为如果中间是字母false了后面有数字的话又成为true了
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
            MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
            Exit Sub
        End If
        
        '判断BOM记录是否登记并且已经批准
        rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            If Trim(rs("Approver")) <> "" Then
                ApprovalStatus = True
            Else
                ApprovalStatus = False
            End If
            '同时判断BOM打开者是否BOM作者(提交者)
            If InStr(Trim(rs("Submiter")), PDMUserName) Then
                OpennerSubmiter = True
            Else
                OpennerSubmiter = False
            End If
        End If
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
        
        
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

    
    If MsgBox("You are going to Export BOM/NPO Data to an Excel File, Continue？", vbYesNo + vbInformation, "SystemInfo") = vbYes Then

        Dim i, J As Integer

        Set xlApp = CreateObject("Excel.Application")   '创建Excel文件
        Set xlApp = New excel.Application
        xlApp.SheetsInNewWorkbook = 1                   '将新建的工作薄数量设为1
        
        '解决出现部件挂起提示
        'xlApp.OleRequestPendingTimeout = 10000   '10000毫秒后出现忙对话框
        'xlApp.OleServerBusyTimeout = 1000     '请求超时1秒
        'xlApp.OleServerBusyRaiseError = True '不显示忙对话框
    
    
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)              '第1张工作表
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
        'xlSheet.Cells(4, 1).CopyFromRecordset Conn.Execute(strSql)       '此行是粘贴数据
    
        xlApp.ActiveWorkbook.Close True     '关闭工作簿并保存
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
    '设置打印信息
    Printer.PaperSize = vbPRPSA4
    Printer.DrawMode = vbPixels
    SetRect rtMargin, 100, 100, 100, 100 '页边距
    '开始打印
    Printer.CurrentX = rtMargin.left
    Printer.CurrentY = rtMargin.top
    Printer.Print "" '进纸
    SetRect rtCell, rtMargin.left, rtMargin.top, 0, 0
    With MSFlexGrid1
        For i = 0 To .Rows - 1
            .Row = i
            '确定是否要换页
            If Printer.ScaleHeight - .RowHeight(i) <= rtMargin.bottom Then
                Printer.NewPage
                rtCell.top = rtMargin.top
            End If
            For J = 0 To .Cols - 1
                .Col = J
                '打印单元格边框
                rtCell.right = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
                rtCell.bottom = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
                Rectangle Printer.hDC, rtCell.left, rtCell.top, rtCell.right + 1, rtCell.bottom + 1
                '设置单元格字体
                Printer.FontName = .CellFontName
                Printer.FontSize = .CellFontSize
                Printer.FontBold = .CellFontBold
                Printer.FontItalic = .CellFontItalic
                Printer.FontUnderline = .CellFontUnderline
                '打印单元格文字（假设内边距为4）
                SetRect rtText, rtCell.left + 4, rtCell.top + 4, rtCell.right - 4, rtCell.bottom - 4
                DrawText Printer.hDC, .TextMatrix(i, J), LenB(StrConv(.TextMatrix(i, J), vbFromUnicode)), rtText, _
                DT_SINGLELINE Or GetAlign(.CellAlignment)
                rtCell.left = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
            Next
            rtCell.left = rtMargin.left
            rtCell.top = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
        Next
    End With
    '打印完毕
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
        If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) <> "NEW" And UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) <> "EXISTING" Then   '隐藏不是New/Existing Part的行
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
    
    '在Project的NPO中,由于无法辨别BOM的Approve状态 和 BOM的打开者是否作者(提交者)的状态,  所以相关标记都要设为false
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
        MsgBox "Project Number is 6 Number, no Letter" + vbCrLf + "必须是6位数字的编号,无字母", vbInformation, "System Info."
        Exit Sub
    Else
        PJNOIndex = CLng(ComboPJNOIndex.Text) '去零
        PjtName = CStr(ComboPjtName.Text)
        AuthorName = CStr(cmbAuthor.Text)
    End If
    '判断SinglePart记录中的符合Project Number的New Part
    StrSql = "Select * from SglPrt where NewOldStatus ='New'"
    If PJNOIndex <> "" Then StrSql = StrSql & " And PJNOIndex ='" & CStr(PJNOIndex) & "'"
    If PjtName <> "" Then StrSql = StrSql & " And PJTName like '%" & Trim(PjtName) & "%'"
    If AuthorName <> "" Then StrSql = StrSql & " And Applicant = '" & Trim(AuthorName) & "'"
    
    rs.Open StrSql, Conn, adOpenStatic, adLockOptimistic
    J = 2      'MSFlexGrid1从第2行开始写,第0行是Title,第1行留空给BOM中的Finish Goods
    MSFlexGrid1.Rows = 3   '设置总行数
    MSFlexGrid1.RowHeight(1) = 225     '如果此行被隐藏过的话需要恢复行高默认值
    MSFlexGrid1.RowHeight(2) = 225     '如果此行被隐藏过的话需要恢复行高默认值
    Do While rs.EOF = False
        MSFlexGrid1.TextMatrix(J, 0) = J - 1                               '输入每行的行号,由于上面第1行空给BOM中的Finish Goods,所以这里要 -1，序号才从1开始
        MSFlexGrid1.TextMatrix(J, 3) = rs.Fields("SglPrtIndex")            '输入物料的12NC
        MSFlexGrid1.TextMatrix(J, 5) = Trim(rs.Fields("PrtUnit"))          '输入物料的单位
        MSFlexGrid1.TextMatrix(J, 6) = Trim(rs.Fields("Description"))      '输入物料的描述
        MSFlexGrid1.TextMatrix(J, 7) = Trim(rs.Fields("ItemType"))         '输入物料的ItemType
        MSFlexGrid1.TextMatrix(J, 8) = Trim(rs.Fields("ProductLine"))      '输入物料的ProductionLine
        MSFlexGrid1.TextMatrix(J, 9) = Trim(rs.Fields("NewOldStatus"))     '输入物料的New/old状态
        
        
        If IsNull(rs.Fields("SERNmbr")) Then    '必须用IsNull函数判断,不能用 rstCX3.Fields("SERNmbr") = Null
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
    
    MSFlexGrid_ChgStatus_HightlightRow (10)                   '对ChangeStatus第11列中有内容的行设置为黄色
    MSFlexGrid_NewOld_HightlightRow (9)                       '对New/Old第9列中有内容为New的行设置为粉红色
    MSFlexGridColumnColorChange MSFlexGrid1, 9, J             '设置New/Old列(第9列)为浅桔红色
    MSFlexGridColumnColorChange MSFlexGrid1, 12, J            '设置Comments列(第13列)为浅桔红色
    MSFlexGridColumnColorChange MSFlexGrid1, 21, J            '设置Supplier PN列(第22列)为浅桔红色
    MSFlexGridColumnColorChange MSFlexGrid1, 13, J, &H404040  '设置分隔列(第14列)为黑色
    MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 0                  '设置单元格位置取消上面改变函数中的某列高亮显示
    
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
    
    If MsgBox("You are going to Change All New Parts to Old Parts, Continue？", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
        '先保存初始的Col和Row号的值
        ColNoTemp = MSFlexGrid1.Col
        RowNoTemp = MSFlexGrid1.Row
        For RowSumVar = 2 To RowNum
            If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) = "NEW" Then   '判断是不是New Part的行
                '找到SinglePart中的对应12NC,并且更新NewOld字段成Old
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
        MSFlexGrid1.Row = RowNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的行数
        MSFlexGrid1.Col = ColNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的列数
    End If
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdNewtoOld_Click"
End Sub

Private Sub CmdSearchFinsGd_Click()
    QueryTableName = "FinsGd"                                  '##########告诉通用查询窗口是对哪个表进行操作
    
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        If PurchasingSys = "Y" Then FrmQuery.cmdAdd.Enabled = False
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    Set FromForm = FrmBOMNPO
    FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
End Sub

Private Sub CmdSearchSglPrt_Click()

    QueryTableName = "SglPrt"                                  '##########告诉通用查询窗口是对哪个表进行操作
    
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        If PurchasingSys = "Y" Then FrmQuery.cmdAdd.Enabled = False
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
    Set FromForm = FrmBOMNPO
    FrmQuery.Show 0 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
    MousePointer = vbDefault                  '恢复鼠标状态
End Sub

Private Sub CmdExportBOM_Click()
    If Not ApprovalStatus Then
        MsgBox "The BOM/NPO is NOT Approved, Please do NOT use it Formally(Offically)", vbInformation, "System Info."
    End If
    
    If MsgBox("You are going to Export BOM/NPO Data to an Excel File, Continue？", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
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
        MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
        Exit Sub
    End If
    
    '判断BOM记录是否登记并且已经批准
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If Trim(rs("Approver")) <> "" Then
            ApprovalStatus = True
        Else
            ApprovalStatus = False
        End If
        '同时判断BOM打开者是否BOM作者(提交者)
        If InStr(Trim(rs("Submiter")), PDMUserName) Then
            OpennerSubmiter = True
        Else
            OpennerSubmiter = False
        End If
    End If
    If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
    
    
    StrSql = "Select Top 1 isNull(CPCNNmbr,''),isNull(CPCNLocate,'') From BOMCPCN Where BOMID ='" & TxtFinsGdIndex.Text & "' Order by BOMVersion Desc"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        txtCPCNNO.Text = rs(0)
        txtCPCNlocate.Text = rs(1)
    End If
    rs.Close
    
    Conn.Close
    

    FinishGoodsNO = Trim(TxtFinsGdIndex.Text)
    MSFlexGrid1.Rows = 3   '设置总行数
    MSFlexGrid1.RowHeight(1) = 225     '如果此行被隐藏过的话需要恢复行高默认值
    MSFlexGrid1.RowHeight(2) = 225     '如果此行被隐藏过的话需要恢复行高默认值
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
    
    '先保存初始的Col和Row号的值
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    If RowNum > 1 Then
    For RowSumVar = 2 To RowNum
        'If UCase(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 9))) = "NEW" Then   '判断是不是New Part的行
        '找到GlueSupplier中的对应12NC,并且提取字段成Old
        If rs.State = adStateOpen Then rs.Close
        '如果忽略版本号表达式为(Mid(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3))) - 1) & "0")
        rs.Open "Select * from GlueSupplier Where Glue12NC ='" & Trim(MSFlexGrid1.TextMatrix(RowSumVar, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            If rs.State = adStateOpen Then rs.Close
            GoTo NextRowforSupplierPN
        Else
            MSFlexGrid1.Col = 21      '指定第22栏Supplier PN
            MSFlexGrid1.Row = RowSumVar
            MSFlexGrid1.Text = rs("SupplierPN")
        End If
        If rs.State = adStateOpen Then rs.Close
        'End If
NextRowforSupplierPN:
    Next RowSumVar
    End If
    MSFlexGrid1.Row = RowNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的行数
    MSFlexGrid1.Col = ColNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的列数
    
    'Exit Sub

'vbErrorHandler:
'    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:Refresh_FlexGridColumnSupplierPN"
End Sub
Private Sub Refresh_FlexGrid(Optional ByVal Applicant As String = "")
    Dim i As Integer                              '子项插入临时表中的循环变量
    Dim J As Integer                              'MSFlexGrid表行数递增的变量
    Dim k As Integer                              '每次增加下一级子项的级数变量
    Dim m() As Integer                            '每次增加下一级子项时,每层级数中的子项数量
    Dim MI As Integer                             '每次增加下一级子项时,每层级数中的子项数量随着临时表中删除变化而需要的变量
    Dim Z01 As Integer                            '一个0和1变换的变量.和K变量相关用于定位级数
    Dim TempParentName()  As String               '每次增加下一级子项的级数后对应的父项Name(ID)
    Dim TempForm As String                        '查询BOM中临时使用的表名的字符串变量
    
    Dim SERNmbr, TempSER As String
    
    
    Dim myCnn As New ADODB.Connection

    myCnn.Open connString
    
    '//定义两个记录集  rstCX记录集对应临时表  rstCX2记录集对应临时表中一个记录查出的对应子记录
    Dim rstCX As New ADODB.Recordset
    Dim rstCX2 As New ADODB.Recordset
    Dim rstCX3 As New ADODB.Recordset
    Set rstCX.ActiveConnection = myCnn
    Set rstCX2.ActiveConnection = myCnn
    Set rstCX3.ActiveConnection = myCnn
    
    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    
    '//判断输入的图号是否底层子项， 如果是没有父项的底层子项则提示后退出
    StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + FinishGoodsNO + "'"
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
    If Not rstCX.RecordCount > 0 Then
        MsgBox " This item is not assembly, has no Child", vbInformation, "System Info."
        rstCX.Close
        BOMExist = False   'BOMExist用于判断是否执行 If BOMExist Then Refresh_FlexGridColumnSupplierPN
        Exit Sub
    Else
        BOMExist = True    'BOMExist用于判断是否执行 If BOMExist Then Refresh_FlexGridColumnSupplierPN
    End If
    rstCX.Close
    
    '//创建临时表示例
    'myCnn.Execute "create table testTable (" & "colTime datetime NULL ," & _
    "colFlt float NULL ," & _
    "myImg image NULL  ," & _
    "myInt int NULL ," & _
    "myNText ntext COLLATE Chinese_PRC_CI_AS NULL )"
    
    '创建表名为tb的表示例
    'strSQL = "CREATE TABLE tb(" & "username varchar(20) not null primary key," & "pass char(10) not null)"
    
    
    
    '创建临时表前需要加一个判断，如果同名表存在则需要换名
    Dim temp12NC As String
    Dim temptbleID As Integer
    Dim TblExist As Boolean
    temptbleID = 0
    TempForm = "TempForm0"
    TblExist = False
    Do
        temptbleID = temptbleID + 1
        TempForm = Mid(TempForm, 1, Len(TempForm) - 1) & Trim(temptbleID)        'str=mid(str,1,len(str)-n) 删除最右边的n个字符
        StrSql = "select * from sysobjects where name = '" & TempForm & "'"
        rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic   'rstCX记录集试图取临时表中记录
        If rstCX.EOF Then TblExist = rstCX.EOF
        rstCX.Close
    Loop Until TblExist
    
    StrSql = "CREATE TABLE " + TempForm + "(TempID INT IDENTITY(1,1) PRIMARY KEY CLUSTERED, Temp12NC varchar(12), TempLevel INT)"
    '其中IDENTITY(1,1)是自动增加标号定义从1开始，增量是1
    myCnn.Execute StrSql
    
    StrSql = "INSERT INTO " + TempForm + " VALUES (" + FinishGoodsNO + ",0)"         '0为TempLevel
    myCnn.Execute StrSql
    
    '//用ADODB控件储存临时表TempForm信息
    StrSql = "SELECT * FROM " + TempForm + ""
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic   'rstCX记录集对应临时表
    
    '//判断临时表记录是否为空
    J = 1   'MSFlexGrid1表格从第1(实际是第2行)开始
    k = 0   '每次增加下一级子项的级数变量, 初始值为0
    Do While Not rstCX.EOF   '一直循环到临时表中的记录为0
        rstCX.MoveLast                  '指针移动到临时表中的最后记录
        Z01 = 0                         '一个0和1变换的变量.和K变量相关用于定位级数,没有找到子项为0
        
        '//判断物料是否有子项
        Set rstCX2 = Nothing
        If Applicant = "" Then
            StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + CStr(rstCX.Fields("Temp12NC")) + "' ORDER BY ChildID"
        Else
            StrSql = "SELECT * FROM BOMOrigData WHERE ParentID = '" + CStr(rstCX.Fields("Temp12NC")) + "' And ChildID in (SELECT LEFT(SglPrtIndex,11)+ '' + CAST(CAST(RIGHT(SglPrtIndex,1) AS INT)+SglPrtVer AS VARCHAR(12)) FROM SglPrt Where Applicant='" & Applicant & "') ORDER BY ChildID"
        End If

        rstCX2.Open StrSql, myCnn, adOpenStatic, adLockOptimistic 'rstCX2记录集对应临时表中一个记录查出的对应子记录
        
        If rstCX2.RecordCount >= 1 Then
            Z01 = 1                         '一个0和1变换的变量.和K变量相关用于定位级数,找到子项为1
            If k >= 1 Then m(k) = MI                      '在K值(级数变量)增加前保存剩余的子项数
            k = k + 1                       '如果有子项,那么TempParentName的内容要更新
            
            ReDim Preserve TempParentName(1 To k)
            TempParentName(k) = CStr(rstCX.Fields("Temp12NC"))
            
            ReDim Preserve m(1 To k)
            m(k) = rstCX2.RecordCount
            MI = m(k) + 1              '+1是因为除找到的子记录外还要先删除找出了子记录集的父记录
            
            rstCX2.MoveFirst
            For i = 1 To rstCX2.RecordCount
                '//把子项插入临时表中
                StrSql = "INSERT INTO " + TempForm + " VALUES ('" + CStr(rstCX2.Fields("ChildID")) + "'," + (CStr(rstCX.Fields("TempLevel") + 1)) + ")"
                myCnn.Execute (StrSql)
                rstCX2.MoveNext
            Next i
            rstCX2.MoveLast
        End If
        If rstCX.Fields("TempLevel") = 0 Then           'TempLevel= 0 则为根项目,内容单独填写
                '//在MSFlexGrid中输出相关信息
            MSFlexGrid1.TextMatrix(J, 0) = J        '输入每行的行号
            MSFlexGrid1.TextMatrix(J, 1) = rstCX.Fields("TempLevel")
            MSFlexGrid1.TextMatrix(J, 2) = "Top"
            MSFlexGrid1.TextMatrix(J, 3) = CStr(rstCX.Fields("Temp12NC"))
            MSFlexGrid1.TextMatrix(J, 4) = 1            'Root根节点(Finish Goods 数量总是为1)
            rstCX3.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(MSFlexGrid1.TextMatrix(J, 3)) & "'", myCnn, adOpenKeyset, adLockOptimistic
            If rstCX3.RecordCount > 0 Then
                MSFlexGrid1.TextMatrix(J, 5) = "Piece"     '对于根项目的Unit,总是为"Piece"
                MSFlexGrid1.TextMatrix(J, 6) = Trim(rstCX3.Fields("Description"))
                
                TxtDescription.Text = Trim(rstCX3.Fields("Description"))
                ComboPJNOIndex.Text = FormatProjectCode(Trim(rstCX3.Fields("PJNOIndex")))
                ComboPjtName.Text = Trim(rstCX3.Fields("PjtName"))
                'cmbAuthor.Text = Trim(rstCX3.Fields("Applicant"))
                
                MSFlexGrid1.TextMatrix(J, 7) = Trim(rstCX3.Fields("ItemType"))
                MSFlexGrid1.TextMatrix(J, 8) = Trim(rstCX3.Fields("ProductLine"))
                MSFlexGrid1.TextMatrix(J, 9) = ""
                
                MSFlexGrid1.TextMatrix(J, 10) = ""         '对于根项目的ChangeStatus,总是为空
                
                If Not IsNull(Trim(rstCX3.Fields("SERLocate"))) Then
                    TempSER = Mid(Replace(Trim(rstCX3.Fields("SERLocate")), "----", ""), 32, 5)
                    If TempSER = "EASE " Then
                        TempSER = "RELEASREPORT"
                    Else
                        If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "（" Then
                            TempSER = "SER00000" & left(TempSER, 4)
                        Else
                            TempSER = "SER0000" & TempSER
                        End If
                    End If
                Else
                    TempSER = ""
                End If
                
                If IsNull(rstCX3.Fields("SERNmbr")) Then    '必须用IsNull函数判断,不能用 rstCX3.Fields("SERNmbr") = Null
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
                
                '更新SER
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
            StrSql = "SELECT * FROM BOMOrigData WHERE ChildID = '" + Trim(CStr(rstCX.Fields("Temp12NC"))) + "' and  ParentID = '" + Trim(TempParentName(k - Z01)) + "'"  '(K - Z01)没有找到子项则减0,找到子项则减1
            rstCX2.Open StrSql, myCnn, adOpenStatic, adLockOptimistic 'rstCX2记录集对应临时表中一个记录查出的对应子记录
            If rstCX2.RecordCount > 0 Then
                MSFlexGrid1.TextMatrix(J, 0) = J        '输入每行的行号
                MSFlexGrid1.TextMatrix(J, 1) = rstCX.Fields("TempLevel")
                MSFlexGrid1.TextMatrix(J, 3) = temp12NC
            
                MSFlexGrid1.TextMatrix(J, 2) = rstCX2.Fields("ParentID")
                MSFlexGrid1.TextMatrix(J, 4) = rstCX2.Fields("Quantity")
                If IsNull(rstCX2.Fields("ChgStatus")) Then          '必须用IsNull函数判断,不能用 rstCX2.Fields("ChgStatus") = Null
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
                            If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "（" Then
                                TempSER = "SER00000" & left(TempSER, 4)
                            Else
                                TempSER = "SER0000" & TempSER
                            End If
                        End If
                    Else
                        TempSER = ""
                    End If
    
                    
                    If IsNull(rstCX3.Fields("SERNmbr")) Then    '必须用IsNull函数判断,不能用 rstCX3.Fields("SERNmbr") = Null
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
                    
                    '更新SER
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
        
        '// 删除临时表中记录
        StrSql = "DELETE FROM " + TempForm + " WHERE Temp12NC = '" + CStr(rstCX.Fields("Temp12NC")) + "' and TempLevel = '" + CStr(k - Z01) + "'"    '一定要加上and TempLevel = '" + cStr(K - Z01) + "'" 否则其它level的不同parent的项目也一同被删掉了
        myCnn.Execute (StrSql)
        MI = MI - 1                  '每次从临时表中删除一项后,子项数量递减1
        
        If k - 1 >= 1 And Z01 = 1 Then         '刚删除的记录是属于上一层的(上一层是必须是2以上)
            m(k - 1) = m(k - 1) - 1                '所以上一层的子项数量也递减1
        End If
        
        If MI = 0 Then                          'MI = 0 表示每层级数中的子项数量全部清空时 (要进到上一级时)
KMinus2:
            k = k - 1                               '要进到上一级时,级数变量减1
            If k >= 1 Then                       '级数变量小于1,下标越界,则需要退出
                If m(k) = 0 Then GoTo KMinus2       '如果上一级的子项也为0,则继续递减到上上一级
                MI = m(k)                           '把所到达的上一级(或者上上一级)还剩下的子记录数赋值给MI
            Else
                k = k + 1
            End If
        End If
        
        Set rstCX = Nothing         '记录集rstCX清空
        StrSql = "SELECT * FROM " + TempForm + ""     '记录集rstCX清空后重新加入包括下一层子记录的临时表的项目
        rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
        J = J + 1
        MSFlexGrid1.Rows = J + 1
    Loop
    '删除临时用表
    StrSql = "Drop TABLE " + TempForm + ""
    myCnn.Execute StrSql
    
    rstCX.Close
    rstCX2.Close
    RowNum = J
    
    MSFlexGrid_ChgStatus_HightlightRow (10)                   '对ChangeStatus第11列中有内容的行设置为黄色
    MSFlexGrid_NewOld_HightlightRow (9)                       '对New/Old第9列中有内容为New的行设置为粉红色
    'MSFlexGridColumnColorChange MSFlexGrid1, 9, j             '设置New/Old列(第9列)为浅桔红色
    'MSFlexGridColumnColorChange MSFlexGrid1, 12, j            '设置Comments列(第13列)为浅桔红色
    MSFlexGridColumnColorChange MSFlexGrid1, 21, J            '设置Supplier PN列(第22列)为浅桔红色
    MSFlexGrid_ApproveStatus_HightlightRow (ApprovalStatus)   '对第1行设置为绿色如果是已经批准的BOM
    MSFlexGridColumnColorChange MSFlexGrid1, 13, J, &H404040  '设置分隔列(第14列)为黑色
    MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 0                  '设置单元格位置取消上面改变函数中的某列高亮显示
    
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
            MSFlexGridRowColorChange MSFlexGrid1, RowSumVar, MSFlexGrid1.Cols, &HFFC0FF      '&HFFC0FF为粉红色
        End If
    Next RowSumVar
End Sub

Private Sub MSFlexGrid_ApproveStatus_HightlightRow(ByVal ApproverOK As Boolean)
    
    If ApproverOK Then
        MSFlexGridRowColorChange MSFlexGrid1, 1, MSFlexGrid1.Cols, &H80FF80     '&H80FF80为绿色
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
    
    txtNodeSglPrt12NC = ""            '先清除原来的内容
    txtNodeDescription = ""
    txtNodePrtUnit = ""
    txtNodeDrwlocate = ""
    txtSERNO = ""
    txtSERlocate = ""
    
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    
    If MSFlexGrid1.Row = RowNum Then Exit Sub   '如果是最后一个空行则退出
    
    If MSFlexGrid1.Row = 1 Then                 'MSFlexGrid1.Row = 1 表示点取的是Finish Goods
        MSFlexGrid1.Col = 3
        temp12NC = MSFlexGrid1.Text            '第3列中的FinishGoods的12NC赋值
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
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
        Set rs = Nothing
        
    Else
        MSFlexGrid1.Col = 3
        temp12NC = MSFlexGrid1.Text
        temp12NC = Mid(temp12NC, 1, (Len(temp12NC) - 1)) & "0"
        If Not Isnum(temp12NC) Then Exit Sub      '如果点取的是标题行,则temp12NC不是数字,则以下语句会出错
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
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
        Set rs = Nothing
    End If
    
    MSFlexGrid1.Row = RowNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的行数
    MSFlexGrid1.Col = ColNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的列数
    
    Select Case MSFlexGrid1.Col
    Case 9, 12, 21, 22                          '第9栏New/Old,第13栏Comment/Note,第22栏Supplier PartNumber
        '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
        If SystemAdmin = "Y" Then
            'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
            GoTo AdminGoAhead1
        End If
        '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
        
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
        
    Case 15, 16, 18, 19, 20, 22                           '第15栏到20栏为采购填写
        
        If Not NPOWorking Then
            MsgBox "Current status is not New Part OverView, Can NOT Edit", vbInformation, "System Info."
            Exit Sub
        End If
        
        rs.Open "Select * from Users where Name ='" & PDMUserName & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
            If SystemAdmin <> "Y" Then
                'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
                GoTo AdminGoAhead2
            End If
            '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
            If Trim(rs("UserGroup")) <> "采购组" Then
                MsgBox "You are not Purchasing Department Employee. Unable to Edit Purchasing Department Content", vbInformation, "System Info."
                Exit Sub
            End If
        End If
AdminGoAhead2:
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
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
        
    Case 14, 17                            '第14栏,17栏为采购填写
        
        If Not NPOWorking Then
            MsgBox "Current status is not New Part OverView, Can NOT Edit", vbInformation, "System Info."
            Exit Sub
        End If
        
        rs.Open "Select * from Users where Name ='" & PDMUserName & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
            If SystemAdmin <> "Y" Then
                'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
                GoTo AdminGoAhead3
            End If
            '@@@@@@@@@@判断是否是管理员用户，如果是直接跳转进行
            If Trim(rs("usergroup")) <> "采购组" Then
                MsgBox "You are not Purchasing Department Employee. Unable to Edit Purchasing Department Content", vbInformation, "System Info."
                Exit Sub
            End If
        End If
AdminGoAhead3:
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
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
        'MSFlexGrid1.Text = scr.Eval(TxtFinsGdIndex.Text)                                 '用ScriptControl对象来计算表达式
        
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
                    MSFlexGridRowColorChange MSFlexGrid1, MSFlexGrid1.Row, MSFlexGrid1.Cols, &HFF80FF    '&HFF80FF为粉红色
                End If
                MSFlexGrid1.Row = RowNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的行数
                MSFlexGrid1.Col = ColNoTemp         '因为上面有对MSFlexGrid1的单元格的操作所以需要恢复原来的列数
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
            '如果忽略版本号表达式为(Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0")
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
            
        Case 22 '############新增Remark##########
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            
            If rs.State = adStateOpen Then rs.Close
            '如果忽略版本号表达式为(Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0")
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
        
        If MSFlexGrid1.Row < RowNum Then                                         '需要加一个判断是否超出最大值
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
        
        If MSFlexGrid1.Row < RowNum Then                                         '需要加一个判断是否超出最大值
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
            ComboPJNOIndex.AddItem (FormatProjectCode(Trim(CStr(rs(0))))) 'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
            ComboPjtName.AddItem (Trim(rs(1))) 'UsrCtlFind括号中的3()是对应Description的字段序号
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
            ComboPJNOIndex.AddItem (Trim(rs(0))) 'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
            ComboPjtName.AddItem (Trim(rs(1))) 'UsrCtlFind括号中的3()是对应Description的字段序号
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
    '确保窗体改变时控件随之改变
    Resize_ALL Me
End Sub
Private Sub Form_Load()
    'Load Skin & Format Control
    'LoadSkin Me
    '''Call ResizeInit(Me)
    MSFlexGrid1.Rows = 3   '设置总行数
    MSFlexGrid1.Cols = 23   '设置总列数
    
    MSFlexGrid1.ColAlignment(0) = 3     '()中为列的编号
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
    'flexAlignLeftTop 0 单元格的内容左、顶部对齐。
    'flexAlignLeftCenter 1 字符串的缺省对齐方式。单元格的内容左、居中对齐。
    'flexAlignLeftBottom 2 单元格的内容左、底部对齐。
    'flexAlignCenterTop 3 单元格的内容居中、顶部对齐。
    'flexAlignCenterCenter 4 单元格的内容居中、居中对齐。
    'flexAlignCenterBottom 5 单元格的内容居中、底部对齐。
    'flexAlignRightTop 6 单元格的内容右、顶部对齐。
    'flexAlignRightCenter 7 数值的缺省对齐方式。单元格的内容右、居中对齐。
    'flexAlignRightBottom 8 单元格的内容右、底部对齐。
    'flexAlignGeneral 9 单元格的内容按一般方式进行对齐。字符串按“左、居中”显示，数字按“右、居中”显示。
    
    
    'Load User information
    Dim Conn As New ADODB.Connection
    Dim StrSql As String
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "Select [Name] from Users Order by [Name]"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        cmbAuthor.AddItem (rs(0))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
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
        'new part才分existing
        StrSql = "select * from BOMOrigData where childid = '" & s12NC & "' And ParentId<>'" & sParentId & "'"
        rs.Open StrSql, Conn, adOpenStatic, adLockOptimistic
        If rs.RecordCount > 0 Then
            StrSql = "select * from BOMOrigData where childid = '" & s12NC & "' and parentid = '" & sParentId & "' and [index]=(select top 1 [Index] from BOMOrigData where childid =  '" & s12NC & "'  order by [Index])"
            rs2.Open StrSql, Conn, adOpenStatic, adLockOptimistic
            '第一个创建的BOM，仍然是NEW
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

