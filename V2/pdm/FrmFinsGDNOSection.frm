VERSION 5.00
Begin VB.Form FrmFinsGDNOSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Goods Number Section Overview"
   ClientHeight    =   6444
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5940
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6444
   ScaleWidth      =   5940
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdSysDistrb15 
      Caption         =   "9999-00000/09999 SAP"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   5760
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb14 
      Caption         =   "9041754-35000/39999 Ford group    "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   13
      Top             =   4860
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb13 
      Caption         =   "9041754-26000/29999 Fiat group          "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb12 
      Caption         =   "9041754-24000/26999 Renault"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb11 
      Caption         =   "9041754-17000/19999 Daimler Chrysler Group"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2340
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb10 
      Caption         =   "9041754-14000/16999 BMW Group"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1980
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb1 
      Caption         =   "3141130-03000/09999 Electronic"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   5310
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb2 
      Caption         =   "9041754-10000/13999 VW Group"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1620
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb3 
      Caption         =   "9041754-20000/23999 PSA"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb4 
      Caption         =   "9041754-30000/34999 General Motors"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   5
      Top             =   4500
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb5 
      Caption         =   "9041754-40000/59999 Japanese,Korean Car group,Other"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   4140
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb6 
      BackColor       =   &H80000000&
      Caption         =   "2441257-30000/39999 Audio Video D&&M Internal OEM/ODM"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   30
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb7 
      BackColor       =   &H80000000&
      Caption         =   "2441257-40000/49999 Audio Video Customer External OEM/ODM        "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   420
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb8 
      BackColor       =   &H80000000&
      Caption         =   "2441257-50000/59999 Audio Video: CONSUMER- BACK UP"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   750
      Width           =   5900
   End
   Begin VB.CommandButton CmdSysDistrb9 
      BackColor       =   &H80000000&
      Caption         =   "2441257-60000/69999 AUDIO VIDEO:SERVICE PARTS            "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   1140
      Width           =   5900
   End
End
Attribute VB_Name = "FrmFinsGDNOSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModifyFm As Boolean
Private Sub CmdSysDistrb1_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("314113003000", "314113009999")
Unload Me
End Sub

Private Sub CmdSysDistrb10_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175414000", "904175416999")
Unload Me
End Sub

Private Sub CmdSysDistrb11_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175417000", "904175419999")
Unload Me
End Sub

Private Sub CmdSysDistrb12_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175424000", "904175426999")
Unload Me
End Sub

Private Sub CmdSysDistrb13_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175427000", "904175429999")
Unload Me
End Sub

Private Sub CmdSysDistrb14_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175435000", "904175439999")
Unload Me
End Sub

Private Sub CmdSysDistrb15_Click(Index As Integer)
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("999910000101", "999999999999")
Unload Me
End Sub

Private Sub CmdSysDistrb2_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175410000", "904175413999")
Unload Me
End Sub

Private Sub CmdSysDistrb3_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175420000", "904175423999")
Unload Me
End Sub

Private Sub CmdSysDistrb4_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175430000", "904175434999")
Unload Me
End Sub

Private Sub CmdSysDistrb5_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("904175440000", "904175459999")
Unload Me
End Sub

Private Sub CmdSysDistrb6_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("244125730000", "244125739999")
Unload Me
End Sub

Private Sub CmdSysDistrb7_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("244125740000", "244125749999")
Unload Me
End Sub

Private Sub CmdSysDistrb8_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("244125750000", "244125759999")
Unload Me
End Sub

Private Sub CmdSysDistrb9_Click()
FrmFinsGdEdit.TxtFinsGdIndex.Text = DataSection_DB("244125760000", "244125769999")
Unload Me
End Sub


Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me

End Sub
Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub


Private Function DataSection_DB(ByVal StartString As String, EndString As String) As String    '###########变量改成对应的表字段名
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim Rscntstring As String

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next           '############以下相关改成对应的表字段名字
'Rscntstring = "select top 1 FinsGdIndex+10 from FinsGd WHERE (((FinsGdIndex+10) Not In (select FinsGdIndex from FinsGd))and (FinsGdIndex+10) between " + StartString + " and " + EndString + ") order by FinsGdIndex+10"
'Rscntstring = "select Max(FinsGdIndex) from FinsGd WHERE FinsGdIndex between " + StartString + " and " + EndString + ""
Rscntstring = "select top 1 CAST(LEFT(FinsGdIndex,11)+'0' AS BIGINT)+10 from FinsGd WHERE (((CAST(LEFT(FinsGdIndex,11)+'0' AS BIGINT)+10) Not In (select FinsGdIndex from FinsGd))and (CAST(LEFT(FinsGdIndex,11)+'0' AS BIGINT)+10) between " + StartString + " and " + EndString + ") order by CAST(LEFT(FinsGdIndex,11)+'0' AS BIGINT)+10"

'Debug.Print Rscntstring
rcds.Open Rscntstring, Conn, adOpenKeyset, adOpenStatic     'PJNOIndex+10表示每加1位申请一个号

  '如果不能查到记录
If rcds.RecordCount = 0 Then
  '系统提示信息，没有推荐号，请自行选择
    MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
    DataSection_DB = StartString
Else
    If Modify = False Then
        DataSection_DB = Trim(rcds.Fields(0).Value)
    Else
       DataSection_DB = Trim(FrmFinsGdEdit.TxtFinsGdIndex.Text)
    End If
End If
Conn.Close
  
End Function

