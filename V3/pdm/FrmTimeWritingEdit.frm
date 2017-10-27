VERSION 5.00
Begin VB.Form FrmTimeWritingEdit 
   Caption         =   "Time Writing Record"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   10965
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtRFSDesc 
      Height          =   405
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   5250
      Width           =   3000
   End
   Begin VB.TextBox txtRFS 
      Height          =   405
      Left            =   4800
      TabIndex        =   25
      Top             =   4680
      Width           =   3000
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   360
      Top             =   6960
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.TextBox TxtPjtName 
      Height          =   400
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4020
      Width           =   3000
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   400
      Left            =   4800
      TabIndex        =   10
      Top             =   3420
      Width           =   3000
   End
   Begin VB.ComboBox cmbEngineer 
      Height          =   300
      Left            =   4800
      TabIndex        =   9
      Top             =   1050
      Width           =   3000
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   300
      ItemData        =   "FrmTimeWritingEdit.frx":0000
      Left            =   4800
      List            =   "FrmTimeWritingEdit.frx":0016
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1650
      Width           =   3000
   End
   Begin VB.TextBox txtWeek 
      Height          =   400
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2130
      Width           =   3000
   End
   Begin VB.ComboBox cmbProjectStatus 
      Height          =   300
      ItemData        =   "FrmTimeWritingEdit.frx":0073
      Left            =   4800
      List            =   "FrmTimeWritingEdit.frx":007D
      TabIndex        =   6
      Text            =   "OPEN/CLOSE"
      Top             =   2670
      Width           =   3000
   End
   Begin VB.TextBox txtOTHours 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   400
      Left            =   4800
      TabIndex        =   5
      Top             =   7350
      Width           =   3000
   End
   Begin VB.TextBox txtActualHours 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   400
      Left            =   4800
      TabIndex        =   4
      Top             =   6870
      Width           =   3000
   End
   Begin VB.TextBox txtRemark 
      Height          =   400
      Left            =   4800
      TabIndex        =   3
      Top             =   6270
      Width           =   3000
   End
   Begin VB.ComboBox cmbOther 
      Height          =   300
      ItemData        =   "FrmTimeWritingEdit.frx":008E
      Left            =   4800
      List            =   "FrmTimeWritingEdit.frx":00AA
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5820
      Width           =   3000
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   7950
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   7950
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "RFS Number 工程报价名称"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1410
      MouseIcon       =   "FrmTimeWritingEdit.frx":00EA
      TabIndex        =   27
      Top             =   5220
      Width           =   3225
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "RFS Number 工程报价编号"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1410
      MouseIcon       =   "FrmTimeWritingEdit.frx":03F4
      TabIndex        =   24
      Top             =   4740
      Width           =   3225
   End
   Begin VB.Label LblPjtName 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name 项目名称描述"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1410
      MouseIcon       =   "FrmTimeWritingEdit.frx":06FE
      TabIndex        =   23
      Top             =   4170
      Width           =   4005
   End
   Begin VB.Label LblPJNOIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Number 所属项目编号"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1410
      MouseIcon       =   "FrmTimeWritingEdit.frx":0A08
      TabIndex        =   22
      Top             =   3510
      Width           =   4005
   End
   Begin VB.Label Label1 
      Caption         =   "Time Writing Record"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3120
      TabIndex        =   21
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label Label2 
      Caption         =   "Name 填写人"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      TabIndex        =   20
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label Label3 
      Caption         =   "Group 组"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1410
      TabIndex        =   19
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "Week 周"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1410
      TabIndex        =   18
      Top             =   2220
      Width           =   1965
   End
   Begin VB.Label Label5 
      Caption         =   "Project Status"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1410
      TabIndex        =   17
      Top             =   2850
      Width           =   1965
   End
   Begin VB.Label Label6 
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   16
      Top             =   5820
      Width           =   4005
   End
   Begin VB.Label Label7 
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1350
      TabIndex        =   15
      Top             =   6390
      Width           =   4005
   End
   Begin VB.Label Label8 
      Caption         =   "Actual Hours"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   14
      Top             =   6990
      Width           =   4005
   End
   Begin VB.Label Label9 
      Caption         =   "OT Hours"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1320
      TabIndex        =   13
      Top             =   7470
      Width           =   4005
   End
   Begin VB.Label Label10 
      Caption         =   "eg: WK1206"
      Height          =   225
      Left            =   8040
      TabIndex        =   12
      Top             =   2250
      Width           =   1245
   End
End
Attribute VB_Name = "FrmTimeWritingEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Public formMode As String
Public TWIndex As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Dim StrSql As String
    If cmbEngineer.Text = "" Then
        MsgBox "Choose Engineer Name", vbCritical
    ElseIf cmbGroup.Text = "" Then
        MsgBox "Choose Group", vbCritical
    ElseIf Len(txtWeek.Text) <> 6 Or val(right(txtWeek.Text, 2)) > 53 Then
        MsgBox "Input Valid Week", vbCritical
    ElseIf cmbProjectStatus.Text = "OPEN/CLOSE" Or cmbProjectStatus.Text = "" Then
        MsgBox "Choose Project Status", vbCritical
'    ElseIf TxtPJNOIndex.Text = "" Then
'        MsgBox "Input Project Number", vbCritical
'    ElseIf TxtPjtName.Text = "" Then
'        MsgBox "Input Project Name", vbCritical
'    ElseIf cmbOther.Text = "" Then
'        MsgBox "Choose Other", vbCritical
    ElseIf Not IsNumeric(txtActualHours.Text) Then
        MsgBox "Input Valid Actual Hours", vbCritical
'    ElseIf Not IsNumeric(txtOTHours.Text) Then
'        MsgBox "Input Valid OT Hours", vbCritical
    Else
        If formMode = "MODIFY" Then
            StrSql = "Update TimeWriting SET "
            StrSql = StrSql & "TWGroup='" & cmbGroup.Text
            StrSql = StrSql & "',TWWeek='" & txtWeek.Text
            StrSql = StrSql & "',TWProjectStatus='" & cmbProjectStatus.Text
            StrSql = StrSql & "',TWProjectNmbr='" & TxtPJNOIndex.Text
            StrSql = StrSql & "',TWProjectDesc='" & TxtPjtName.Text
            StrSql = StrSql & "',TWOther='" & cmbOther.Text
            StrSql = StrSql & "',TWRemark='" & txtRemark.Text
            StrSql = StrSql & "',TWActualHours=" & txtActualHours.Text
            StrSql = StrSql & ",TWOTHours=" & IIf(Trim(txtOTHours.Text) = "", 0, Trim(txtOTHours.Text))
            StrSql = StrSql & ",TWRFS='" & txtRFS.Text & "'"
            StrSql = StrSql & ",TWRFSDesc='" & txtRFSDesc.Text & "'"
            StrSql = StrSql & " WHERE TWIndex=" & TWIndex
    
            Conn.Execute StrSql
            If Err = 0 Then MsgBox "Modify Sucessfully.", vbInformation
            Call FrmTimeWriting.CmdToQuery_Click
            Conn.Close
            Set Conn = Nothing
            Unload Me
        Else
            StrSql = "Insert into TimeWriting(TWEngineer,TWGroup,TWWeek,TWProjectStatus,TWProjectNmbr,TWProjectDesc,TWOther,TWRemark,TWActualHours,TWOTHours,TWRFS,TWRFSDesc) values('"
            StrSql = StrSql & cmbEngineer.Text
            StrSql = StrSql & "','" & cmbGroup.Text
            StrSql = StrSql & "','" & txtWeek.Text
            StrSql = StrSql & "','" & cmbProjectStatus.Text
            StrSql = StrSql & "','" & TxtPJNOIndex.Text
            StrSql = StrSql & "','" & TxtPjtName.Text
            StrSql = StrSql & "','" & cmbOther.Text
            StrSql = StrSql & "','" & txtRemark.Text
            StrSql = StrSql & "'," & txtActualHours.Text
            StrSql = StrSql & "," & IIf(Trim(txtOTHours.Text) = "", 0, Trim(txtOTHours.Text))
            StrSql = StrSql & ",'" & txtRFS.Text & "'"
            StrSql = StrSql & ",'" & txtRFSDesc.Text & "'"
            StrSql = StrSql & ")"

            Conn.Execute StrSql
            If Err = 0 Then MsgBox "Submit Sucessfully.", vbInformation
            TxtPJNOIndex.Text = ""
            TxtPjtName.Text = ""
            txtRFS.Text = ""
            txtRFSDesc.Text = ""
            txtRemark.Text = ""
            txtActualHours.Text = ""
            txtOTHours.Text = ""
            Unload Me
        End If
    End If
    formMode = "ADD"
    TWIndex = 0
    Me.Show 0
End Sub

Private Sub Form_Load()
    'Load User information
    'LoadSkin Me
    If Conn.State = adStateOpen Then Conn.Close
    Conn.Open connString
    If TWIndex <> 0 Then
        formMode = "MODIFY"
        Dim StrSql As String
        StrSql = "select * from TimeWriting where TWIndex=" & CStr(TWIndex)
        
        rs.Open StrSql, Conn, adOpenStatic, adLockOptimistic
        If rs.EOF Or rs.BOF Then
            MsgBox "Database Error, Please try again."
            Unload Me
        Else
            cmbEngineer.Text = rs.Fields("TWEngineer")
            txtWeek.Text = rs.Fields("TWWeek")
            cmbGroup.Text = rs.Fields("TWGroup")
            cmbProjectStatus.Text = rs.Fields("TWProjectStatus")
            TxtPJNOIndex.Text = rs.Fields("TWProjectNmbr")
            TxtPjtName.Text = rs.Fields("TWProjectDesc")
            txtRFS.Text = rs.Fields("TWRFS")
            txtRFSDesc.Text = rs.Fields("TWRFSDesc")
            For i = 0 To cmbOther.ListCount - 1
                If Trim(cmbOther.List(i)) = Trim(rs.Fields("TWOther")) Then cmbOther.ListIndex = i
            Next i
            txtRemark.Text = rs.Fields("TWRemark")
            txtActualHours.Text = rs.Fields("TWActualHours")
            txtOTHours.Text = rs.Fields("TWOTHours")
        End If
        rs.Close
    Else
        formMode = "ADD"
        cmbEngineer.Text = PDMUserName
        cmbGroup.Text = PDMUserGroup
        cmbProjectStatus.Text = ""
        TxtPJNOIndex.Text = ""
        TxtPjtName.Text = ""
        txtRFS.Text = ""
        txtRFSDesc.Text = ""
        cmbOther.ListIndex = 0
        txtRemark.Text = ""
        txtActualHours.Text = ""
        txtOTHours.Text = ""
    End If
End Sub



Private Sub Form_Resize()
 '确保窗体改变时控件随之改变
    Resize_ALL Me
End Sub

Private Sub TxtPJNOIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNO为要查询的表名
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始


    sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '其中1UseEquel代表= 2UseLike是代表Like

    TxtPjtName.Text = Trim(UsrCtlFind(3)) 'UsrCtlFind括号中的3()是对应Description的字段序号
    Erase UsrCtlFind
    sqlUsrCtrl.MoveRecord (MoveNext)
    
    sqlUsrCtrl.CloseRS
End If
End Sub

'Private Sub TxtPjtName_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    ComboPjtName.Clear
'    ComboPJNOIndex.Clear
'    Dim sqlUsrCtrl As Control
'    Set sqlUsrCtrl = sqlSDBC1
'
'    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
'    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNO为要查询的表名
'    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始
'
'     Do While sqlUsrCtrl.IfBOForEOF = False
'       sqlUsrCtrl.FindRecord "Description", UseLike, Trim(TxtPjtName.Text)  '其中1UseEquel代表= 2UseLike是代表Like
'
'       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind括号中的0()是对应PJNOIndex的字段序号
'       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind括号中的3()是对应Description的字段序号
'       Erase UsrCtlFind
'       sqlUsrCtrl.MoveRecord (MoveNext)
'
'     Loop
'    sqlUsrCtrl.CloseRS
'End If
'End Sub
Private Sub TxtRFS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("RFSRFQ")    'PJNO为要查询的表名
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord已经取消从第一开始找，所以这里要设置到开始


    sqlUsrCtrl.FindRecord "RFSRFQIndex", UseEquel, Trim(txtRFS.Text)  '其中1UseEquel代表= 2UseLike是代表Like

    txtRFSDesc.Text = CStr(UsrCtlFind(3)) 'UsrCtlFind括号中的3()是对应Description的字段序号
    Erase UsrCtlFind
    sqlUsrCtrl.MoveRecord (MoveNext)
    
    sqlUsrCtrl.CloseRS
End If
End Sub
