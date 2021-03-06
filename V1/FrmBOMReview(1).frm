VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBOMReview 
   Caption         =   "BOM Version Review"
   ClientHeight    =   8772
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13176
   LinkTopic       =   "Form1"
   ScaleHeight     =   8772
   ScaleWidth      =   13176
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "BOM Version"
      Height          =   1185
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   13035
      Begin VB.CommandButton Command2 
         Caption         =   "Export to Excel"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   11280
         TabIndex        =   15
         Top             =   660
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print BOM"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10110
         TabIndex        =   8
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   11970
         TabIndex        =   7
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdReiew 
         Caption         =   "Review BOM"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10110
         TabIndex        =   5
         Top             =   150
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "Finish Goods NO"
         Top             =   330
         Width           =   1725
      End
      Begin VB.ComboBox cmbBOMVersion 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblSubcon 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   3510
         TabIndex        =   17
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "SUBCON:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2700
         TabIndex        =   16
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   6300
         TabIndex        =   14
         Top             =   270
         Width           =   3675
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5250
         TabIndex        =   13
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lblCPCN 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   8310
         TabIndex        =   12
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "CP/CN:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   11
         Top             =   780
         Width           =   765
      End
      Begin VB.Label lblUpdateDate 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   4950
         TabIndex        =   10
         Top             =   750
         Width           =   2145
      End
      Begin VB.Label Label3 
         Caption         =   "Update Time:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3690
         TabIndex        =   9
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "FG NO:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "BOM Version Number:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   2025
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7380
      Left            =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   13035
      _ExtentX        =   22987
      _ExtentY        =   13018
      _Version        =   393216
      Rows            =   33
      Cols            =   13
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmBOMReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iAsc As Integer

Private Sub CmdExit_Click()
    Unload Me
End Sub

'Public Sub cmdReiew_Click()
'    On Error Resume Next
'    Dim myCnn As New ADODB.Connection
'    Dim myRS As New ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'    Dim rs2 As New ADODB.Recordset
'    Dim ProjectDesc As String
'    Dim i, J, x As Integer
'    myCnn.Open connString
'
'    If Len(Trim(Text1.Text)) = 0 Then
'        MsgBox "You must enter a new 12NC for the Finish Goods", vbInformation, "System Info."
'        Exit Sub
'    ElseIf Not (Len(Trim(Text1.Text)) = 12 And IsNumeric(Trim(Text1.Text))) Then
'        MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
'        Exit Sub
'    ElseIf cmbBOMVersion.Text = "" Then
'        MsgBox "Please choose the BOM Version Number.", vbInformation, "System Info."
'        Exit Sub
'    End If
'
'    StrSql = "Select IsNull(a.CPCNNmbr,'N/A'),a.UpdateDate,IsNull(b.Description,'N/A') From BOMCPCN a inner join FinsGd b on a.BOMId=b.FinsGdIndex Where a.BOMVersion=" & cmbBOMVersion.Text & " and a.BOMID=" & Text1.Text & ""
'    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
'    If myRS.RecordCount > 0 Then
'        lblUpdateDate.Caption = myRS(1)
'        lblCPCN.Caption = myRS(0)
'        ProjectDesc = Trim(myRS(2))
'    End If
'    myRS.Close
'
'
'    StrSql = "SELECT subcon FROM SUBCON WHERE FinsGDIndex=" & Text1.Text & ""
'    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
'    If myRS.RecordCount > 0 Then
'        lblSubcon.Caption = myRS(0)
'    End If
'    myRS.Close
'
'
'    If ProjectDesc = "N/A" Then
'        StrSql = "Select IsNull(b.Description,'N/A') From BOMCPCN a inner join SglPrt b on Left(a.BOMId,11)+'0'=b.SglPrtIndex Where a.BOMVersion=" & cmbBOMVersion.Text & " and a.BOMID=" & Text1.Text & ""
'        myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
'        If myRS.RecordCount > 0 Then
'            ProjectDesc = Trim(myRS(0))
'        End If
'        myRS.Close
'    End If
'    lblDesc.Caption = ProjectDesc
'
'    MSFlexGrid1.Clear
'    MSFlexGridTileInitialize
'    StrSql = "Select Distinct ParentID,ChildID,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,isNull(Family,''),isNull(CPCN,''),SeqIndex FROM SglPrt4BOMLog Where BOMVersion=" & cmbBOMVersion.Text & " And BOM=" & Text1.Text & "  Order by SeqIndex,ParentID,ChildID"
'    'Debug.Print StrSql
'    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
'
'    With MSFlexGrid1
'        i = 1
'        Do While Not myRS.EOF
'                .TextMatrix(i, 0) = i        '输入每行的行号
'                .TextMatrix(i, 1) = myRS(0)
'                .TextMatrix(i, 2) = myRS(1)
'                .TextMatrix(i, 3) = myRS(2)
'                .TextMatrix(i, 4) = myRS(3)
'                .TextMatrix(i, 5) = myRS(4)
'                .TextMatrix(i, 6) = myRS(5)
'                .TextMatrix(i, 7) = Convert_IT2KP(myRS(1), myRS(5))
'                .TextMatrix(i, 8) = myRS(6)
'                If myRS(7) = "Add-Upgrade" Then
'                    .TextMatrix(i, 9) = "Add"
'                ElseIf myRS(7) = "Delete-Upgrade" Then
'                    .TextMatrix(i, 9) = "Delete"
'                Else
'                    .TextMatrix(i, 9) = myRS(7)
'                End If
'                .TextMatrix(i, 10) = myRS(8)
'                .TextMatrix(i, 11) = myRS(9)
'                .TextMatrix(i, 12) = myRS(10)
'
'
'            i = i + 1
'            .Rows = i + 1
'
'            myRS.MoveNext
'        Loop
'        .Rows = .Rows - 1 '多一空行
'    End With
'
'    myRS.Close
'    Set myRS = Nothing
'    myCnn.Close
'    Set myCnn = Nothing
'
'    '划删除线 背景色
'    Dim AStatus As String
'    For J = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(J, 8) <> "" Then
'            AStatus = UCase(left(MSFlexGrid1.TextMatrix(J, 8), 1))
'        Else
'            AStatus = ""
'        End If
'
'        If AStatus = "D" Then
'            MSFlexGrid1.Row = J
'            For x = 0 To 11
'                MSFlexGrid1.Col = x
'                MSFlexGrid1.CellFontStrikeThrough = True
'                MSFlexGrid1.CellBackColor = &HCCCCCC
'            Next
'        ElseIf AStatus = "M" Then
'            MSFlexGrid1.Row = J
'            For x = 1 To 11
'                MSFlexGrid1.Col = x                    '从第ColNo列第0行开始
'                'MSFlexGrid1.CellFontStrikeThrough = True
'                MSFlexGrid1.CellBackColor = &HCCFFCC
'            Next
'        ElseIf AStatus = "A" Then
'            MSFlexGrid1.Row = J
'            For x = 1 To 11
'                MSFlexGrid1.Col = x
'                'MSFlexGrid1.CellFontStrikeThrough = True
'                MSFlexGrid1.CellBackColor = &HFF99CC
'            Next
'        End If
'    Next J
'End Sub

Public Sub cmdReiew_Click()
    On Error Resume Next
    Dim myCnn As New ADODB.Connection
    Dim myRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim ProjectDesc As String
    Dim i, J, x As Integer
    myCnn.Open connString
    
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "You must enter a new 12NC for the Finish Goods", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(Text1.Text)) = 12 And IsNumeric(Trim(Text1.Text))) Then
        MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
        Exit Sub
    ElseIf cmbBOMVersion.Text = "" Then
        MsgBox "Please choose the BOM Version Number.", vbInformation, "System Info."
        Exit Sub
    End If
    
    StrSql = "Select IsNull(a.CPCNNmbr,'N/A'),a.UpdateDate,IsNull(b.Description,'N/A') From BOMCPCN a inner join FinsGd b on a.BOMId=b.FinsGdIndex Where a.BOMVersion=" & cmbBOMVersion.Text & " and a.BOMID=" & Text1.Text & ""
    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
    If myRS.RecordCount > 0 Then
        lblUpdateDate.Caption = myRS(1)
        lblCPCN.Caption = myRS(0)
        ProjectDesc = Trim(myRS(2))
    End If
    myRS.Close
    
    
    StrSql = "SELECT subcon FROM SUBCON WHERE FinsGDIndex=" & Text1.Text & ""
    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
    If myRS.RecordCount > 0 Then
        lblSubcon.Caption = myRS(0)
    End If
    myRS.Close
    
    
    If ProjectDesc = "N/A" Then
        StrSql = "Select IsNull(b.Description,'N/A') From BOMCPCN a inner join SglPrt b on Left(a.BOMId,11)+'0'=b.SglPrtIndex Where a.BOMVersion=" & cmbBOMVersion.Text & " and a.BOMID=" & Text1.Text & ""
        myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
        If myRS.RecordCount > 0 Then
            ProjectDesc = Trim(myRS(0))
        End If
        myRS.Close
    End If
    lblDesc.Caption = ProjectDesc
    
    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    StrSql = "Select Distinct ParentID,ChildID,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,isNull(Family,''),isNull(CPCN,''),SeqIndex FROM SglPrt4BOMLog Where BOMVersion=" & cmbBOMVersion.Text & " And BOM=" & Text1.Text & "  Order by SeqIndex,ParentID,ChildID"
    'Debug.Print StrSql
    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic

    With MSFlexGrid1
        i = 1
        Do While Not myRS.EOF
                .TextMatrix(i, 0) = i        '输入每行的行号
                .TextMatrix(i, 1) = myRS(0)
                .TextMatrix(i, 2) = myRS(1)
                .TextMatrix(i, 3) = myRS(2)
                .TextMatrix(i, 4) = myRS(3)
                .TextMatrix(i, 5) = myRS(4)
                .TextMatrix(i, 6) = myRS(5)
                .TextMatrix(i, 7) = myRS(6)
                If myRS(7) = "Add-Upgrade" Then
                    .TextMatrix(i, 8) = "Add"
                ElseIf myRS(7) = "Delete-Upgrade" Then
                    .TextMatrix(i, 8) = "Delete"
                Else
                    .TextMatrix(i, 8) = myRS(7)
                End If
                .TextMatrix(i, 9) = myRS(8)
                .TextMatrix(i, 10) = myRS(9)
                .TextMatrix(i, 11) = myRS(10)

            
            i = i + 1
            .Rows = i + 1

            myRS.MoveNext
        Loop
        .Rows = .Rows - 1 '多一空行
    End With

    myRS.Close
    Set myRS = Nothing
    myCnn.Close
    Set myCnn = Nothing
    
    '划删除线 背景色
    Dim AStatus As String
    For J = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(J, 8) <> "" Then
            AStatus = UCase(left(MSFlexGrid1.TextMatrix(J, 8), 1))
        Else
            AStatus = ""
        End If
        
        If AStatus = "D" Then
            MSFlexGrid1.Row = J
            For x = 0 To 11
                MSFlexGrid1.Col = x
                MSFlexGrid1.CellFontStrikeThrough = True
                MSFlexGrid1.CellBackColor = &HCCCCCC
            Next
        ElseIf AStatus = "M" Then
            MSFlexGrid1.Row = J
            For x = 1 To 11
                MSFlexGrid1.Col = x                    '从第ColNo列第0行开始
                'MSFlexGrid1.CellFontStrikeThrough = True
                MSFlexGrid1.CellBackColor = &HCCFFCC
            Next
        ElseIf AStatus = "A" Then
            MSFlexGrid1.Row = J
            For x = 1 To 11
                MSFlexGrid1.Col = x
                'MSFlexGrid1.CellFontStrikeThrough = True
                MSFlexGrid1.CellBackColor = &HFF99CC
            Next
        End If
    Next J
End Sub

Private Function Convert_IT2KP(sglprt As String, ItemType As String) As String
    On Error Resume Next
    Dim myCnn As New ADODB.Connection
    Dim myRS As New ADODB.Recordset
    myCnn.Open connString
    
    StrSql = "Select KindProduct From ConvertConfig Where IndexFrom <= " & sglprt & " And IndexEnd >=" & sglprt & " And ItemType='" & ItemType & "'"
    myRS.Open StrSql, myCnn, adOpenKeyset, adLockPessimistic
    If myRS.RecordCount > 0 Then
        Convert_IT2KP = myRS(0)
    Else
        Convert_IT2KP = ""
    End If
    myRS.Close
    
End Function

Private Sub Command1_Click()
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



Private Sub Command2_Click()
    On Error Resume Next

    Dim sHeader As String
    
    Dim J As Integer
    Dim x As Integer
    Dim L As String
    Dim str1(255) As Variant
    Dim arrIT() As String
     
    
    Set xlApp = CreateObject("Excel.Application")   '创建Excel文件
    Set xlApp = New excel.Application
    xlApp.SheetsInNewWorkbook = 1                   '将新建的工作薄数量设为1
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)              '第1张工作表

    
'    '解决出现部件挂起提示
'    xlApp.OleRequestPendingTimeout = 10000   '10000毫秒后出现忙对话框
'    xlApp.OleServerBusyTimeout = 1000     '请求超时1秒
'    xlApp.OleServerBusyRaiseError = True '不显示忙对话框
    
    sHeader = "BOM Version"
    xlSheet.Cells(1, 1) = sHeader
    xlSheet.Cells(2, 1) = "BOM:" & Text1.Text
    xlSheet.Cells(2, 2) = "CP/CN:" & lblCPCN.Caption
    xlSheet.Cells(2, 3) = "SubCon:" & lblSubcon.Caption
    xlSheet.Cells(2, 4) = "Version:" & cmbBOMVersion.Text
    xlSheet.Cells(2, 5) = "Description:" & lblDesc.Caption
    xlSheet.Cells(2, 6) = "Update Time:" & lblUpdateDate.Caption
    xlSheet.Cells(2, 7) = "Table Maker:" & PDMUserName
    Dim lngRowsCount As Long, lngColumnsCount As Long, lngRow As Long, lngColumn As Long
    Dim strText As String

    lngRowsCount = MSFlexGrid1.Rows
    lngColumnsCount = MSFlexGrid1.Cols
    For lngRow = 3 To lngRowsCount + 2
            For lngColumn = 1 To lngColumnsCount

                If lngColumn > 11 Then
                    strText = MSFlexGrid1.TextMatrix(lngRow - 3, lngColumn - 1)
                    xlSheet.Cells(lngRow, lngColumn - 1) = "'" + strText
                Else
                    strText = MSFlexGrid1.TextMatrix(lngRow - 3, lngColumn - 1)
                    xlSheet.Cells(lngRow, lngColumn) = "'" + strText
                End If
            Next
    Next

    xlApp.ActiveWorkbook.Close True     '关闭工作簿并保存
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub



Private Sub Form_Load()
    'Load Skin & Format Control
    'LoadSkin Me
    
    '''Call ResizeInit(Me)
    MSFlexGrid1.Rows = 3   '设置总行数
    MSFlexGrid1.Cols = 13   '设置总列数
    MSFlexGrid1.ColWidth(0) = 12 * 25 * 2
    MSFlexGrid1.ColWidth(1) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(2) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(3) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(4) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(5) = 12 * 25 * 2.3
    MSFlexGrid1.ColWidth(6) = 12 * 25 * 6
    MSFlexGrid1.ColWidth(7) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(8) = 12 * 25 * 4.8
    MSFlexGrid1.ColWidth(9) = 12 * 25 * 4.8
    MSFlexGrid1.ColWidth(10) = 12 * 25 * 0.01
    MSFlexGrid1.ColWidth(11) = 12 * 25 * 5
    MSFlexGrid1.ColWidth(11) = 12 * 25 * 5

    
    MSFlexGrid1.ColAlignment(0) = 3     '()中为列的编号
    MSFlexGrid1.ColAlignment(1) = 3
    MSFlexGrid1.ColAlignment(2) = 1
    MSFlexGrid1.ColAlignment(3) = 1
    MSFlexGrid1.ColAlignment(4) = 3
    MSFlexGrid1.ColAlignment(5) = 1
    MSFlexGrid1.ColAlignment(6) = 1
    MSFlexGrid1.ColAlignment(7) = 1
    MSFlexGrid1.ColAlignment(8) = 1
    MSFlexGrid1.ColAlignment(9) = 1
    MSFlexGrid1.ColAlignment(10) = 1
    MSFlexGrid1.ColAlignment(11) = 1
    MSFlexGrid1.ColAlignment(12) = 1


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
    

    'Set BOM Version
    Dim i
    For i = 1 To 30
        cmbBOMVersion.AddItem i
    Next i
       
    MSFlexGridTileInitialize
    
End Sub

Private Sub MSFlexGridTileInitialize()
    MSFlexGrid1.TextMatrix(0, 0) = "Index"
    MSFlexGrid1.TextMatrix(0, 1) = "Parent12NC"
    MSFlexGrid1.TextMatrix(0, 2) = "Child12NC"
    MSFlexGrid1.TextMatrix(0, 3) = "Quantity"
    MSFlexGrid1.TextMatrix(0, 4) = "PrtUnit"
    MSFlexGrid1.TextMatrix(0, 5) = "Description"
    MSFlexGrid1.TextMatrix(0, 6) = "Item Type"
    MSFlexGrid1.TextMatrix(0, 7) = "Kind Of Product"
    MSFlexGrid1.TextMatrix(0, 8) = "SER NO."
    MSFlexGrid1.TextMatrix(0, 9) = "ChgStatus"
    MSFlexGrid1.TextMatrix(0, 10) = "Note"
    MSFlexGrid1.TextMatrix(0, 11) = "Family"
    MSFlexGrid1.TextMatrix(0, 12) = "CPCN"
End Sub


Private Sub Form_Resize()
'确保窗体改变时控件随之改变
    Resize_ALL Me
End Sub


Private Sub MSFlexGrid1_Click()
    '排序
    iAsc = iAsc + 1
    If MSFlexGrid1.Row = 1 Then
        MSFlexGrid1.Col = MSFlexGrid1.MouseCol
        MSFlexGrid1.Sort = CInt(iAsc Mod 2) + 1
    End If
End Sub

Private Sub Text1_Click()
    If Not IsNumeric(Text1.Text) Then Text1.Text = ""
End Sub

