VERSION 5.00
Begin VB.Form FrmSERImport 
   Caption         =   "Import SER data from Excel to SQL DataBase table"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10530
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtEndRow 
      Height          =   495
      Left            =   6150
      TabIndex        =   5
      Top             =   3360
      Width           =   2220
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit / Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6480
      TabIndex        =   4
      Top             =   5220
      Width           =   2505
   End
   Begin VB.CommandButton CmdWrite 
      Caption         =   "Start to write into SQL Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1755
      TabIndex        =   3
      Top             =   5220
      Width           =   3480
   End
   Begin VB.TextBox TxtStartRow 
      Height          =   495
      Left            =   6150
      TabIndex        =   2
      Top             =   2535
      Width           =   2220
   End
   Begin VB.CommandButton CmdCloseExcel 
      Caption         =   "Close Excel"
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
      Left            =   6420
      TabIndex        =   1
      Top             =   1245
      Width           =   2535
   End
   Begin VB.CommandButton CmdOpenExcel 
      Caption         =   "Open Excel"
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
      Left            =   2145
      TabIndex        =   0
      Top             =   1245
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Please input end row number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   7
      Top             =   3465
      Width           =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Please input start row number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2355
      TabIndex        =   6
      Top             =   2640
      Width           =   3600
   End
End
Attribute VB_Name = "FrmSERImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private Sub CmdCloseExcel_Click()
        If Dir("C:\mfg\SER.bz") <> "" Then    '由VB关闭EXCEL
            xlBook.RunAutoMacros (xlAutoClose)        '执行EXCEL关闭宏
             xlBook.Close (True)                      '关闭EXCEL工作薄
            xlApp.Quit                                '关闭EXCEL
        End If
            Set xlApp = Nothing                       '释放EXCEL对象
        End
End Sub

 
Private Sub CmdOpenExcel_Click()               '打开EXCEL过程
     If Dir("C:\mfg\SER.bz") = "" Then '判断EXCEL是否打开
         Set xlApp = CreateObject("Excel.Application")   '创建EXCEL应用类对象
         xlApp.Visible = True                            '设置EXCEL应用类对象可见
         Set xlBook = xlApp.Workbooks.Open("C:\mfg\SER.xls") '打开EXCEL工作薄
         Set xlSheet = xlBook.Worksheets(1)                         '打开EXCEL工作表1
         xlSheet.Activate '激活EXCEL工作表
         'xlSheet.Cells(1, 2) = "vvv" '给单元格行1列2赋值             测试用语句
        xlBook.RunAutoMacros (xlAutoOpen) '运行EXCEL中的启动宏
    Else
        MsgBox ("EXCEL is Opened")
 End If
End Sub

Private Sub CmdQuit_Click()
        If Dir("C:\mfg\SER.bz") <> "" Then '由VB关闭EXCEL
            xlBook.RunAutoMacros (xlAutoClose)      '执行EXCEL关闭宏
             xlBook.Close (True)                     '关闭EXCEL工作薄
            xlApp.Quit                               '关闭EXCEL
        End If
            Set xlApp = Nothing                       '释放EXCEL对象
        End
        
Unload Me
FrmEngineeringSys.Show 0
End Sub

Private Sub CmdWrite_Click()
Dim SERNO As String
Dim SgPtNO As String
Dim ImportSER As ClsSER
Set ImportSER = New ClsSER
Dim i As Integer
'Dim J As Integer
For i = val(Trim(TxtStartRow.Text)) To val(Trim(TxtEndRow.Text))
  If xlSheet.Cells(i, 1) <> "" Then            '每行第1列中不为空择开始写
   
      With ImportSER    '已经定义Public MySER As New ClsSER, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字
     SERNO = xlSheet.Cells(i, 1)
     SERNO = "SER00000" & SERNO
    .SERIndex = SERNO
    .Applicant = xlSheet.Cells(i, 2)
    .CAorA = xlSheet.Cells(i, 3)
    .Description = xlSheet.Cells(i, 5)
    .IDSO = "Open"                              '固定输入值
    .OpnDate = xlSheet.Cells(i, 6)
    .ClosDate = xlSheet.Cells(i, 6)
    .PJNOIndex = 999999                         '固定输入值
    .PjtName = xlSheet.Cells(i, 7)
    .FinsGdNO = "NA"
     SgPtNO = xlSheet.Cells(i, 4)
     SgPtNO = Replace(SgPtNO, " ", "")          '去掉中间的空格
     
    .SglPrtNO = val(SgPtNO)
    .CommtNote = xlSheet.Cells(i, 8)
    
           '判断SERIndex序号是否已经存在
                If .In_DB(val(xlSheet.Cells(i, 1))) = True Then
                   MsgBox "In" + Str(i) + " row, SER number exists, Please go next" + vbCrLf + "在第" + Str(i) + " 行, SER号重复，请进行下一个写记录", vbInformation, "System Info."
                
                Else
                   .Insert                   '添加
                    'MsgBox "In" + Str(I) + "row, Succeed to Add" + vbCrLf + "在第" + Str(I) + " 行, 添加成功", vbInformation, "System Info."    '测试单条记录用,大批写的话去掉这句
                End If

      End With
   End If
Next

End Sub


Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
End Sub
