VERSION 5.00
Begin VB.Form FrmCPCNImport 
   Caption         =   "Import CPCN data from Excel to SQL DataBase table"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7845
   StartUpPosition =   2  '屏幕中心
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
      Left            =   1590
      TabIndex        =   5
      Top             =   675
      Width           =   2415
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
      Left            =   4440
      TabIndex        =   4
      Top             =   675
      Width           =   2535
   End
   Begin VB.TextBox TxtStartRow 
      Height          =   495
      Left            =   4755
      TabIndex        =   3
      Top             =   1755
      Width           =   2220
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
      Left            =   615
      TabIndex        =   2
      Top             =   3540
      Width           =   3480
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
      Left            =   4620
      TabIndex        =   1
      Top             =   3540
      Width           =   2505
   End
   Begin VB.TextBox TxtEndRow 
      Height          =   495
      Left            =   4755
      TabIndex        =   0
      Top             =   2580
      Width           =   2220
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
      Left            =   960
      TabIndex        =   7
      Top             =   1860
      Width           =   3600
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
      Left            =   945
      TabIndex        =   6
      Top             =   2685
      Width           =   3600
   End
End
Attribute VB_Name = "FrmCPCNImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private Sub CmdCloseExcel_Click()
        If Dir("C:\mfg\CPCN.bz") <> "" Then    '由VB关闭EXCEL
            xlBook.RunAutoMacros (xlAutoClose)        '执行EXCEL关闭宏
             xlBook.Close (True)                      '关闭EXCEL工作薄
            xlApp.Quit                                '关闭EXCEL
        End If
            Set xlApp = Nothing                       '释放EXCEL对象
        End
End Sub

 
Private Sub CmdOpenExcel_Click()               '打开EXCEL过程
     If Dir("C:\mfg\CPCN.bz") = "" Then '判断EXCEL是否打开
         Set xlApp = CreateObject("Excel.Application")   '创建EXCEL应用类对象
         xlApp.Visible = True                            '设置EXCEL应用类对象可见
         Set xlBook = xlApp.Workbooks.Open("C:\mfg\CPCN.xls") '打开EXCEL工作薄
         Set xlSheet = xlBook.Worksheets(1)                         '打开EXCEL工作表1
         xlSheet.Activate '激活EXCEL工作表
         'xlSheet.Cells(1, 2) = "vvv" '给单元格行1列2赋值             测试用语句
        xlBook.RunAutoMacros (xlAutoOpen) '运行EXCEL中的启动宏
    Else
        MsgBox ("EXCEL is Opened")
 End If
End Sub

Private Sub CmdQuit_Click()
        If Dir("C:\mfg\CPCN.bz") <> "" Then '由VB关闭EXCEL
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
Dim CPCNNO As String
Dim SgPtNO As String
Dim ImportCPCN As ClsCPCN
Set ImportCPCN = New ClsCPCN
Dim i As Integer
'Dim J As Integer
For i = val(Trim(TxtStartRow.Text)) To val(Trim(TxtEndRow.Text))
  If xlSheet.Cells(i, 1) <> "" Then            '每行第1列中不为空择开始写
   
      With ImportCPCN    '已经定义Public ImportCPCN As New ClsCPCN, 类模块赋变量值  ############以下相关改成对应的控件名字,表的名字,字段名字

    .CPCNIndex = xlSheet.Cells(i, 1)
    .Applicant = xlSheet.Cells(i, 2)
    .CPCNMP = xlSheet.Cells(i, 3)
    .Description = xlSheet.Cells(i, 4)
    .IDSO = "Close"                              '固定输入值
    .OpnDate = Date
    .ClosDate = Date
    .PJNOIndex = 999999                         '固定输入值
    .PjtName = xlSheet.Cells(i, 5)
    .FinsGdNO = "NA"
    .SglPrtNO = "NA"
    .CommtNote = "NA"
    
           '判断CPCNIndex序号是否已经存在
                If .In_DB(xlSheet.Cells(i, 1)) = True Then
                   MsgBox "In" + Str(i) + " row, CPCN number exists, Please go next" + vbCrLf + "在第" + Str(i) + " 行, CPCN号重复，请进行下一个写记录", vbInformation, "System Info."
                
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
