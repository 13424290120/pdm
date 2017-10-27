VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPJNO 
   Caption         =   "PDM-Project Number Admin 工程管理子系统"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPJNO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11310
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdFresh 
      Caption         =   "Refresh"
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
      Left            =   8100
      TabIndex        =   8
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CommandButton CmdToQuery 
      Caption         =   "Search 查询"
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
      Left            =   930
      TabIndex        =   7
      Top             =   780
      Width           =   1710
   End
   Begin VB.CommandButton PageGO 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10140
      TabIndex        =   6
      Top             =   1740
      Width           =   375
   End
   Begin VB.TextBox txtPage_nd 
      Height          =   375
      Left            =   9420
      TabIndex        =   5
      Top             =   1620
      Width           =   735
   End
   Begin VB.TextBox txtPage 
      Height          =   375
      Left            =   9420
      TabIndex        =   4
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "Last page"
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
      Left            =   8100
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "First page"
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
      Left            =   930
      TabIndex        =   2
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "Previous page"
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
      Left            =   3315
      TabIndex        =   1
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next page"
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
      Left            =   5910
      TabIndex        =   0
      Top             =   1620
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4065
      Left            =   495
      TabIndex        =   9
      Top             =   2295
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   7170
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "PJNOIndex"
         Caption         =   "PJNOIndex"
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
         DataField       =   "Applicant"
         Caption         =   "Applicant"
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
      BeginProperty Column02 
         DataField       =   "Leader"
         Caption         =   "Leader"
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
      BeginProperty Column03 
         DataField       =   "Description"
         Caption         =   "Description"
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
      BeginProperty Column04 
         DataField       =   "IDSQ"
         Caption         =   "IDSQ"
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
      BeginProperty Column05 
         DataField       =   "OpnDate"
         Caption         =   "OpnDate"
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
      BeginProperty Column06 
         DataField       =   "ClosDate"
         Caption         =   "ClosDate"
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
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3209.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   7635
      Picture         =   "FrmPJNO.frx":08CA
      Top             =   6705
      Width           =   300
   End
   Begin VB.Label LblBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Return返回"
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
      Left            =   8235
      MouseIcon       =   "FrmPJNO.frx":0CE6
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6705
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   5475
      Picture         =   "FrmPJNO.frx":0FF0
      Top             =   6705
      Width           =   300
   End
   Begin VB.Label LblDelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Delete删除"
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
      Left            =   6075
      MouseIcon       =   "FrmPJNO.frx":140C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6705
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3435
      Picture         =   "FrmPJNO.frx":1716
      Top             =   6705
      Width           =   300
   End
   Begin VB.Label LblModify 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Modify修改"
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
      Left            =   4035
      MouseIcon       =   "FrmPJNO.frx":1B32
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6705
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1635
      Picture         =   "FrmPJNO.frx":1E3C
      Top             =   6705
      Width           =   300
   End
   Begin VB.Label LblAdd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add添加"
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
      Left            =   2235
      MouseIcon       =   "FrmPJNO.frx":2258
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   6705
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmPJNO.frx":2562
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2820
      TabIndex        =   10
      Top             =   525
      Width           =   3240
   End
End
Attribute VB_Name = "FrmPJNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'新功能模块中 所有的Call Refresh_PJNO(lCurrentpage)中的PJNO要统一置换为新表的名
Option Explicit
Dim lCurrentpage As Long           '定义当前页变量
Dim Conn As New ADODB.Connection   '定义一个ADO连接

Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim objrs As New ADODB.Recordset    '定义另一个记录集用于存放每一页的记录

Private Sub CmdFirst_Click()     '第1页操作
   lCurrentpage = 1
   Call Refresh_PJNO(lCurrentpage)
End Sub

Private Sub CmdFresh_Click()        '刷新操作
 Call Refresh_PJNO(lCurrentpage)
End Sub

Private Sub CmdLast_Click()          '第末页操作
   lCurrentpage = 10000
   Call Refresh_PJNO(lCurrentpage)
End Sub

Private Sub CmdNext_Click()           '下1页操作
   lCurrentpage = lCurrentpage + 1
   Call Refresh_PJNO(lCurrentpage)
End Sub

Private Sub CmdPrevious_Click()       '上1页操作
 If lCurrentpage > 1 Then
   lCurrentpage = lCurrentpage - 1
   Call Refresh_PJNO(lCurrentpage)
 End If
End Sub

Private Sub CmdToQuery_Click()
QuerytableName = "PJNO"                                  '##########告诉通用查询窗口是对哪个表进行操作

'@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能
If SystemAdmin <> "Y" Then
    MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
    FrmQuery.CmdModify.Enabled = False
    FrmQuery.CmdDel.Enabled = False

    FrmQuery.DataGrid1.AllowDelete = False
    FrmQuery.DataGrid1.AllowAddNew = False
    FrmQuery.DataGrid1.AllowUpdate = False
End If
'@@@@@@@@@@判断是否是管理员用户，否则要屏蔽一些修改删除功能

FrmQuery.Show 1 'frm.Show style Style为0是窗体是无模式的 style 为 1则窗体是模式的模式窗体时，除了模式窗体中的对象之外不能进行输入（键盘或鼠标单击）。
End Sub

Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Call ResizeForm(Me)
End Sub

Private Sub PageGO_Click()          '去到指定页
   If Not IsNumeric(txtPage_nd) Then
       MsgBox "Page No. must be Number, No letter " + vbCrLf + "请输入页码的数字编号", vbInformation, "Error Info!"
       txtPage_nd.SetFocus
   End If
   
   If val(txtPage_nd.Text) < 1 Then
   lCurrentpage = 1
   Call Refresh_PJNO(lCurrentpage)
   Exit Sub
   End If
   
   lCurrentpage = val(txtPage_nd.Text)  'val函数是字符串转换成数值
   Call Refresh_PJNO(lCurrentpage)

End Sub


Private Sub txtPage_nd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then PageGO_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rcds = Nothing
Set objrs = Nothing
If Conn.State = adStateOpen Then Conn.Close
End Sub
Private Sub Form_Load()
'Load Skin & Format Control
LoadSkin Me
ResizeInit Me

 lCurrentpage = 1           '窗口打开默认是第1页操作
 Call Refresh_PJNO(lCurrentpage)
End Sub

Private Sub LblAdd_Click()
FrmPJNOEdit.Caption = "Add One Project Number."     '##########在对应打开窗口中标题要赋值
lCurrentpage = 10000                                 '当添加记录时一般默认去第末页操作
Call Refresh_PJNO(lCurrentpage)
'标明操作为添加而非修改
FrmPJNOEdit.Modify = False                             '##########在对应打开窗口中Modify标示要赋值

'如果是添加模式要隐藏一些控件
FrmPJNOEdit.TxtIDSQ.Visible = False
FrmPJNOEdit.TxtOpnDate.Visible = False
FrmPJNOEdit.TxtClosDate.Visible = False
FrmPJNOEdit.LblOld0.Visible = False
FrmPJNOEdit.LblOld1.Visible = False
FrmPJNOEdit.LblOld2.Visible = False
FrmPJNOEdit.LblReminder.Visible = False
FrmPJNOEdit.Show 1                                     '##########对应编辑窗口打开
Call Refresh_PJNO(lCurrentpage) '添加完成后再刷新一次
End Sub

Private Sub LblBack_Click()
Set rcds = Nothing
Set objrs = Nothing
If Conn.State = adStateOpen Then Conn.Close
Unload Me
FrmEngineeringSys.Show
End Sub


Private Sub LblDelete_Click()

    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    If SystemAdmin <> "Y" Then
        MsgBox "you are not administrator, No right to delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    
    
Dim TempPJNOID As String                            '##########TempPJNOID更换成对应表中Key字段名
'保存待删除记录的ID
  TempPJNOID = objrs.Fields(0)                      '##########TempPJNOID更换成对应表中Key字段名
  
'弹出删除确认对话框 Str是数字变字符串的函数,这里如果不用Str会出错
  If MsgBox("Confirm to delete" + Str(objrs.Fields(0)) + "?" + vbCrLf + "是否删除" + Str(objrs.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete 确认删除") = vbYes Then
    
    '调用类中Delete函数删除PJNO信息
    MyPJNO.Delete (TempPJNOID)                    '##########TempPJNOID更换成对应表中Key字段名
    MsgBox "Succeed to delete, 删除成功", vbInformation, "System Info."
  End If
  '刷新胶水供应商管理界面
Call Refresh_PJNO(lCurrentpage)
End Sub


Private Sub LblModify_Click()

'保存待修改记录的原始ID
FrmPJNOEdit.OriPJNOIndex = Trim(objrs.Fields(0))           '##########对应编辑窗口变量赋值
'把待修改信息添加到编辑窗口
FrmPJNOEdit.TxtPJNOIndex = Trim(objrs.Fields(0))           '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtApplicant = Trim(objrs.Fields(1))             '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtLeader = Trim(objrs.Fields(2))                 '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtDescription = Trim(objrs.Fields(3))            '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtIDSQ = Trim(objrs.Fields(4))                   '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtOpnDate = Trim(objrs.Fields(5))                '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtClosDate = Trim(objrs.Fields(6))               '##########对应编辑窗口控件赋值
FrmPJNOEdit.TxtPJNOIndex.Enabled = False  '既然是修改，主键索引是不能改的            '##########对应编辑窗口和控件名
FrmPJNOEdit.TxtApplicant.Enabled = False       '既然是修改，申请人一般不用改        '##########对应编辑窗口和控件名
FrmPJNOEdit.Caption = "Modify One Project Number."                                           '##########对应编辑窗口标题
'标明操作为修改操作
FrmPJNOEdit.Modify = True                                     '##########对应编辑窗口变量赋值
FrmPJNOEdit.Show 1                                            '##########对应编辑窗口打开

Call Refresh_PJNO(lCurrentpage)
End Sub


Private Sub Refresh_PJNO(lPage As Long)
          Dim adoPrimaryRS     As ADODB.Recordset
          Dim lPageCount     As Long
          Dim nPageSize     As Integer
          Dim lCount     As Long
          
  '连接数据库
Conn.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(DBUser) + ";pwd=" + Trim(Password) + ";database=" + Trim(DataBase) + ""
Conn.Open

rcds.Open "select * from PJNO", Conn, adOpenKeyset, adOpenStatic  '启动一个Static类型的游标,否则记录数RecordCount总为-1  '##########对应表名字PJNO

  '如果不能查到记录
If rcds.RecordCount = 0 Then
  '修改和删除不可用
LblModify.Enabled = False
LblDelete.Enabled = False
Else
  '如果能查到记录,修改和删除可用
LblModify.Enabled = True
LblDelete.Enabled = True
End If

 
   '每页显示的记录数为15
   nPageSize = 15
   rcds.PageSize = nPageSize         '每页显示的记录数赋值给记录集属性. PageSize分页显示时每一页的记录数
' ADO PageCount 属性
'The PageCount property returns a long value that indicates the number of pages with data in a Recordset object.
'PageCount属性的作用是：返回一个长值，用于指定记录集对象中数据页面的数量。

'Tip: To divide the Recordset into a series of pages, use the PageSize property.
'提示: 你可以使用PageSize属性将记录集分割为一系列的页面?

'Note: If the last page contains fewer records than specified in PageSize, it still counts as an additional page in the PageCount property.
'注意：如果最后一页的记录数量少于在PageSize属性中指定的数量，那么它仍然被视为一页。

'Note: If this method is not supported it returns -1.
'注意：如果不支持这个方法，那么将返回-1。

'IntFix 函数返回参数的整数部分?
'语法
'Int(number)
'Fix(number)
'必要的 number 参数是 Double 或任何有效的数值表达式。如果 number 包含 Null，则返回 Null。
'说明
'Int 和 Fix 都会删除 number 的小数部份而返回剩下的整数。
'Int 和 Fix 的不同之处在于，如果 number 为负数，则 Int 返回小于或等于 number 的第一个负整数，而 Fix 则会返回大于或等于 number 的第一个负整数。例如，Int 将 -8.4 转换成 -9，而 Fix 将 -8.4 转换成 -8。
  lPageCount = rcds.PageCount
              If lCurrentpage > lPageCount Then
                  lCurrentpage = lPageCount
              End If
          rcds.AbsolutePage = lCurrentpage
          
Set objrs = Nothing  '原记录中的内容需要先清空才能写
          '添加字段名称
          For lCount = 0 To rcds.Fields.Count - 1
            If lCount = 0 Then
              objrs.Fields.Append rcds.Fields(lCount).Name, adUnsignedBigInt, rcds.Fields(lCount).DefinedSize  'adUnsignedBigInt   8字节不带符号整型
              GoTo NextLine
            End If
            objrs.Fields.Append rcds.Fields(lCount).Name, adVarChar, rcds.Fields(lCount).DefinedSize 'adVarChar其余字段用字符串
NextLine:
          Next
          
          '打开记录集
          objrs.Open
          
          '将指定记录数循环添加到objrs中
          For lCount = 1 To nPageSize   'nPageSize每页显示的记录数为10
                  If rcds.EOF = True Then
                  Exit For
                  End If
                 
                  objrs.AddNew
                  objrs!PJNOIndex = rcds!PJNOIndex                                             '##########对应表字段赋值
                  objrs!Applicant = rcds!Applicant                                                 '##########对应表字段赋值
                  objrs!Leader = rcds!Leader                                                        '##########对应表字段赋值
                  objrs!Description = rcds!Description                                             '##########对应表字段赋值
                  objrs!IDSQ = rcds!IDSQ                                                            '##########对应表字段赋值
                  objrs!OpnDate = Format(rcds!OpnDate, "YYYY/MM/DD")  '日期需要格式化后再输入       '##########对应表字段赋值
                  objrs!ClosDate = Format(rcds!ClosDate, "YYYY/MM/DD")  '日期需要格式化后再输入      '##########对应表字段赋值
                  rcds.MoveNext
          Next
          '绑定
          Set DataGrid1.DataSource = objrs
            
          '显示页数
          txtPage.Text = lPage & "/" & rcds.PageCount
Conn.Close
 
End Sub


