VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTimeWriting 
   Caption         =   "Time Writing"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13305
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify TW Record"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   7560
      TabIndex        =   20
      Top             =   120
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog excel 
      Left            =   7440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Record"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11400
      TabIndex        =   17
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add TW Record"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9480
      TabIndex        =   8
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11400
      TabIndex        =   5
      Top             =   600
      Width           =   1800
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete TW Record"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9480
      TabIndex        =   4
      Top             =   600
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search综合查询"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   13185
      Begin VB.TextBox txtRFS 
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
         Left            =   10560
         TabIndex        =   22
         Top             =   330
         Width           =   1935
      End
      Begin VB.ComboBox cmbOther 
         Height          =   300
         ItemData        =   "FrmTimeWriting.frx":0000
         Left            =   7560
         List            =   "FrmTimeWriting.frx":0019
         TabIndex        =   19
         Text            =   "ALL"
         Top             =   960
         Width           =   2145
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   300
         ItemData        =   "FrmTimeWriting.frx":0055
         Left            =   1080
         List            =   "FrmTimeWriting.frx":006E
         TabIndex        =   16
         Text            =   "ALL"
         Top             =   960
         Width           =   2160
      End
      Begin VB.TextBox txtWeekTo 
         Height          =   300
         Left            =   6000
         TabIndex        =   14
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtWeekFrom 
         Height          =   300
         Left            =   4680
         TabIndex        =   12
         Top             =   960
         Width           =   800
      End
      Begin VB.ComboBox ComboPJNOIndex 
         Height          =   300
         Left            =   4680
         TabIndex        =   10
         Top             =   360
         Width           =   5000
      End
      Begin VB.ComboBox cmbEngineer 
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   2160
      End
      Begin VB.CommandButton CmdToQuery 
         Caption         =   "Search 查询"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10710
         TabIndex        =   1
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "RFS:"
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
         Left            =   10080
         TabIndex        =   21
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Other:"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "- WK"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Week:   WK"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   11
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Engineer:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   1830
      End
      Begin VB.Label Label9 
         Caption         =   "Project No:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3570
         TabIndex        =   2
         Top             =   330
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5820
      Left            =   30
      TabIndex        =   6
      Top             =   2595
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   10266
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      DefColWidth     =   80
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "RowNumber"
         Caption         =   "NO."
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
         DataField       =   "TWEngineer"
         Caption         =   "Name"
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
         DataField       =   "TWGroup"
         Caption         =   "Group"
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
         DataField       =   "TWWeek"
         Caption         =   "Week"
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
         DataField       =   "TWProjectStatus"
         Caption         =   "Project Status"
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
         DataField       =   "TWProjectNmbr"
         Caption         =   "Project No"
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
         DataField       =   "TWProjectDesc"
         Caption         =   "Project Description"
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
      BeginProperty Column07 
         DataField       =   "TWRFS"
         Caption         =   "RFS"
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
      BeginProperty Column08 
         DataField       =   "TWRFSDesc"
         Caption         =   "RFS Name"
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
      BeginProperty Column09 
         DataField       =   "TWOther"
         Caption         =   "Other"
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
      BeginProperty Column10 
         DataField       =   "TWRemark"
         Caption         =   "Remark"
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
      BeginProperty Column11 
         DataField       =   "TWActualHours"
         Caption         =   "Actual Hours"
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
      BeginProperty Column12 
         DataField       =   "TWOTHours"
         Caption         =   "OT Hours"
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
      BeginProperty Column13 
         DataField       =   "TWCreateDate"
         Caption         =   "Date"
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
      BeginProperty Column14 
         DataField       =   "TWIndex"
         Caption         =   "ID"
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
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1440
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Writing Records           项目用时"
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
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "FrmTimeWriting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim FinalSql As String


Private Function FormatProjectCode(ByVal PJNOIndex As String) As String
    Dim i As Integer
    FormatProjectCode = PJNOIndex
    For i = 1 To 6 - Len(PJNOIndex)
        FormatProjectCode = "0" & FormatProjectCode
    Next
End Function

Private Sub cmdAdd_Click()
    FrmTimeWritingEdit.formMode = "ADD"
    FrmTimeWritingEdit.TWIndex = 0
    FrmTimeWritingEdit.Show 0
End Sub

Private Sub CmdDelete_Click()
    On Error Resume Next
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    
    Dim TempTWIndex As Integer
    
    TempTWIndex = rs.Fields(14)
    
    If Err.Number = 94 Then
        MsgBox "No chose record to delete", vbInformation, "System Info."
        Exit Sub
    End If
    
    If MsgBox("Confirm to delete the record?" + vbCrLf + "是否删除这条记录?", vbYesNo + vbDefaultButton2, "Confirm to Delete 确认删除") = vbYes Then
        Conn.Execute "Delete from TimeWriting Where TWIndex =" & CStr(TempTWIndex) & ""
        MsgBox "Succeed to delete, 删除成功", vbInformation, "System Info."
    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open FinalSql, Conn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs
    
End Sub

Private Sub CmdExit_Click()
    If Conn.State = adStateOpen Then Conn.Close: Set Conn = Nothing
    Unload Me
    FrmEngineeringSys.Show 0
End Sub

Private Sub cmdExport_Click()

    Dim i As Integer
    Dim sHeader As String
    Set xlApp = CreateObject("Excel.Application")   '创建Excel文件
    Set xlApp = New excel.Application
    xlApp.SheetsInNewWorkbook = 1                   '将新建的工作薄数量设为1
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)              '第1张工作表

    xlSheet.Cells(1, 1) = "Time Writing"
    For i = 0 To DataGrid1.Columns.count - 1
        xlSheet.Cells(3, i + 1) = DataGrid1.Columns(i).Caption
    Next i
    
    xlSheet.Cells(2, i - 3) = "Table Maker:": xlSheet.Cells(2, i - 2) = PDMUserName
    xlSheet.Cells(2, i - 1) = "Print Date:": xlSheet.Cells(2, i) = Now()
        
    xlSheet.Cells(4, 1).CopyFromRecordset rs       '此行是粘贴数据

    xlApp.ActiveWorkbook.Close True     '关闭工作簿并保存
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
'Dim strLine As String
'Dim i As Integer
'
'    Screen.MousePointer = vbHourglass
'
'    Open App.Path & "\TimeWriting_" & Year(Now) & Month(Now) & Day(Now) & Minute(Now) & Second(Now) & ".txt" For Output As #1
'        Print #1, "Name" & vbTab & "Group" & vbTab & "Week" & vbTab & "Project Status" & vbTab & "Project Number" & vbTab & "Project Description" & vbTab & "Other" & vbTab & "Remark" & vbTab & "Actual Hours" & vbTab & "OT Hours" & vbTab & "Date" & vbTab & "Index" & vbTab
'        With rs
'            .MoveFirst
'            Do While Not .EOF
'                strLine = ""
'                For i = 0 To .Fields.count - 1
'                    strLine = strLine & "" & .Fields(i).Value
'                    If i < .Fields.count - 1 Then
'                        strLine = strLine & vbTab
'                    End If
'                Next i
'
'                Print #1, strLine
'                .MoveNext
'            Loop
'            .MoveFirst
'        End With
'    Close #1
'
'    Screen.MousePointer = vbDefault
'    MsgBox "Export Successfully!", vbInformation
End Sub



Private Sub cmdModify_Click()
    On Error Resume Next
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    If SystemAdmin <> "Y" And PDMUserName <> Trim(rs.Fields(1)) Then
        MsgBox "You are not administrator, No right to delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    
    Dim TempTWIndex As Integer
    
    TempTWIndex = rs.Fields(14)
    
    If Err.Number = 94 Or TempTWIndex = 0 Then
        MsgBox "No choose record to edit.", vbInformation, "System Info."
        Exit Sub
    End If
    FrmTimeWritingEdit.TWIndex = TempTWIndex
    FrmTimeWritingEdit.Show 0
End Sub

Public Sub CmdToQuery_Click()
    On Error Resume Next
    Dim StrSql, StrSql2, statmt As String
    
    If rs.State = adStateOpen Then rs.Close
    StrSql = "select ROW_NUMBER() over(Order by TWIndex) as RowNumber,TWEngineer,TWGroup,TWWeek,TWProjectStatus,TWProjectNmbr,TWProjectDesc,TWRFS,TWRFSDESC,TWOther,TWRemark,TWActualHours,TWOTHours,TWCreateDate,TWIndex from TimeWriting where 1=1"
    StrSql2 = "select NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'TOTAL:',SUM(TWActualHours),SUM(TWOTHours),NULL,NULL from TimeWriting where 1=1"
    If Trim(txtRFS.Text) <> "" Then statmt = statmt & " And (TWRFS = '" & txtRFS.Text & "' Or TWRFSDESC like '%" & txtRFS.Text & "%')"
    If cmbEngineer.Text <> "ALL" Then statmt = statmt & " And TWEngineer='" & cmbEngineer.Text & "'"
    If cmbGroup.Text <> "ALL" Then statmt = statmt & " And TWGroup='" & cmbGroup.Text & "'"
    If ComboPJNOIndex.Text <> "ALL" Then statmt = statmt & " And TWProjectNmbr='" & left(ComboPJNOIndex.Text, 6) & "'"
    If IsNumeric(txtWeekFrom.Text) And Len(txtWeekFrom.Text) = 4 Then statmt = statmt & " And TWWeek>='WK" & txtWeekFrom.Text & "'"
    If IsNumeric(txtWeekTo.Text) And Len(txtWeekTo.Text) = 4 Then statmt = statmt & " And TWWeek<='WK" & txtWeekTo.Text & "'"
    If cmbOther.Text <> "ALL" Then statmt = statmt & " And TWOther='" & cmbOther.Text & "'"
    
    FinalSql = StrSql & statmt & " UNION " & StrSql2 & statmt
    rs.Open FinalSql, Conn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs
    
    If Err Then MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmTimeWriting:CmdSearch"

End Sub

Private Sub Form_Load()
    On Error Resume Next
    'LoadSkin Me

    Conn.Open connString
    
    
    Set rs.ActiveConnection = Conn
    
    'Load Skin & Format Control
    'LoadSkin Me
    
    
    cmbEngineer.AddItem "ALL"
    StrSql = "Select [Name] from Users Order by [Name]"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        cmbEngineer.AddItem (rs(0))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
        rs.MoveNext
    Loop
    rs.Close
    cmbEngineer.Text = "ALL"


    ComboPJNOIndex.AddItem "ALL"
    StrSql = "Select PJNOIndex, Description from PJNO Order by PJNOIndex"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        ComboPJNOIndex.AddItem (FormatProjectCode(Trim(CStr(rs(0)))) & "-" & Trim(CStr(rs(1)))) 'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
        rs.MoveNext
    Loop
    rs.Close
    ComboPJNOIndex.Text = "ALL"

    
    DataGrid1.AllowDelete = False
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    

    If SystemAdmin = "Y" Then
        FinalSql = "select ROW_NUMBER() over(Order by TWIndex) as RowNumber,TWEngineer,TWGroup,TWWeek,TWProjectStatus,TWProjectNmbr,TWProjectDesc,TWRFS,TWRFSDesc,TWOther,TWRemark,TWActualHours,TWOTHours,TWCreateDate,TWIndex from TimeWriting Order by RowNumber"
    Else
        FinalSql = "select ROW_NUMBER() over(Order by TWIndex) as RowNumber,TWEngineer,TWGroup,TWWeek,TWProjectStatus,TWProjectNmbr,TWProjectDesc,TWRFS,TWRFSDesc,TWOther,TWRemark,TWActualHours,TWOTHours,TWCreateDate,TWIndex from TimeWriting Where TWEngineer='" & PDMUserName & "' Order by RowNumber"
    End If
    rs.Open FinalSql, Conn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs

End Sub

Private Sub Form_Resize()
    Resize_ALL Me
End Sub
