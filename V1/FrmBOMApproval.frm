VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBOMApproval 
   Caption         =   "PDM-BOM Submit / Approval ���̹�����ϵͳ"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBOMApproval.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   14070
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Caption         =   "Search��ѯ"
      Height          =   795
      Left            =   420
      TabIndex        =   20
      Top             =   1200
      Width           =   13185
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5070
         TabIndex        =   24
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtBN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1740
         TabIndex        =   23
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton CmdToQuery 
         Caption         =   "Search ��ѯ"
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
         Left            =   7140
         TabIndex        =   21
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Description��"
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
         TabIndex        =   25
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label Label8 
         Caption         =   "BOM Number��"
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
         TabIndex        =   22
         Top             =   300
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmd_Assign 
      Caption         =   "Assign   ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11610
      TabIndex        =   19
      Top             =   8160
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ListBox TxtExecutName 
      Height          =   1080
      ItemData        =   "FrmBOMApproval.frx":08CA
      Left            =   7710
      List            =   "FrmBOMApproval.frx":08CC
      Style           =   1  'Checkbox
      TabIndex        =   18
      Top             =   6360
      Width           =   2985
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9210
      TabIndex        =   17
      Top             =   270
      Width           =   1800
   End
   Begin VB.CommandButton CmdReject 
      Caption         =   "Reject   ���"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11625
      TabIndex        =   16
      Top             =   7590
      Width           =   1755
   End
   Begin VB.CommandButton CmdApprove 
      Caption         =   "Approve ��׼"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11625
      TabIndex        =   15
      Top             =   7050
      Width           =   1755
   End
   Begin VB.CommandButton CmdSubmit 
      Caption         =   "Submit   �ύ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11625
      TabIndex        =   14
      Top             =   6480
      Width           =   1755
   End
   Begin VB.TextBox TxtRejectReason 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7710
      TabIndex        =   12
      Top             =   8190
      Width           =   2985
   End
   Begin VB.TextBox TxtComment 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2265
      TabIndex        =   5
      Top             =   8160
      Width           =   2985
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2265
      TabIndex        =   4
      Top             =   7350
      Width           =   2985
   End
   Begin VB.TextBox TxtFinsGdIndex 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   3
      Top             =   6510
      Width           =   2985
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Return / Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11160
      TabIndex        =   2
      Top             =   270
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3795
      Left            =   450
      TabIndex        =   0
      Top             =   2325
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   6694
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      DefColWidth     =   80
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "FinsGdIndex"
         Caption         =   "FinsGdIndex"
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
      BeginProperty Column02 
         DataField       =   "Submiter"
         Caption         =   "Submiter"
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
         DataField       =   "SubmitDate"
         Caption         =   "SubmitDate"
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
         DataField       =   "Approver"
         Caption         =   "Approver"
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
         DataField       =   "ApproveDate"
         Caption         =   "ApproveDate"
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
         DataField       =   "Rejecter"
         Caption         =   "Rejecter"
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
         DataField       =   "RejectDate"
         Caption         =   "RejectDate"
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
         DataField       =   "RejectReason"
         Caption         =   "RejectReason"
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
         DataField       =   "CommtNote"
         Caption         =   "CommtNote"
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
         DataField       =   "CheckHistory"
         Caption         =   "CheckHistory"
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
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1440
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPickerExecutDate 
      Height          =   420
      Left            =   7710
      TabIndex        =   11
      Top             =   7350
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   98107393
      CurrentDate     =   39979
   End
   Begin VB.Label Label7 
      Caption         =   "Reject  Reason"
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
      Left            =   5685
      TabIndex        =   13
      Top             =   8205
      Width           =   1785
   End
   Begin VB.Label Label6 
      Caption         =   "Execute Date"
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
      Left            =   5895
      TabIndex        =   10
      Top             =   7425
      Width           =   1560
   End
   Begin VB.Label Label5 
      Caption         =   "Executor Name"
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
      Left            =   5670
      TabIndex        =   9
      Top             =   6570
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   2385
      Left            =   450
      Top             =   6345
      Width           =   4995
   End
   Begin VB.Label Label4 
      Caption         =   "Comment"
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
      Left            =   975
      TabIndex        =   8
      Top             =   8205
      Width           =   1170
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
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
      Left            =   810
      TabIndex        =   7
      Top             =   7380
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "BOM Number"
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
      Left            =   555
      TabIndex        =   6
      Top             =   6570
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmBOMApproval.frx":08CE
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
      Left            =   420
      TabIndex        =   1
      Top             =   210
      Width           =   5070
   End
End
Attribute VB_Name = "FrmBOMApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�¹���ģ���� ���е�Call Refresh_BOMSubmitApprove(lCurrentpage)�е�BOMSubmitApproveҪͳһ�û�Ϊ�±����
Option Explicit

Dim lCurrentpage As Long           '���嵱ǰҳ����
Dim Conn As New ADODB.Connection   '����һ��ADO����

Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
Dim objrs As New ADODB.Recordset    '������һ����¼�����ڴ��ÿһҳ�ļ�¼

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
    Next
End Function


Private Sub cmd_Assign_Click()
    On Error GoTo vbErrorHandler
    
    Dim Assigner As String
    Dim AssignDate As Date

    
    Assigner = Trim(TxtExecutName)
    AssignDate = DTPickerExecutDate.Value

    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    
    '�жϼ�¼�Ƿ����
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "The BOM (Finish Goods) Number Record NOT Exist,Please Submit firstly.", vbInformation, "System Info."
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Exit Sub
    Else
        If Trim(rs("SubmitDate")) = "" Or IsNull(rs("SubmitDate")) Then
            MsgBox "The BOM (Finish Goods) Number still not be Submitted (No Submit Date), Can NOT Approve  ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
            Exit Sub
        End If
        rs("Assigner") = PDMUserName
        rs("AssignDate") = AssignDate
        rs.Update
    End If
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    MsgBox "BOM (Finish Goods) Number Record has Assigned ", vbInformation, "System Info."
    Conn.Close
    
    Call Refresh_BOMSubmitApprove   '�����ɺ���ˢ��һ��
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdAssign"
End Sub

Private Sub CmdSubmit_Click()
    On Error GoTo vbErrorHandler
    Dim Submiter As String
    Dim SubmitDate As Date
    Dim Approver As String
    Dim ApproveDate As Date
    Dim Rejecter As String
    Dim RejectDate As Date
    Dim RejectReason As String
    Dim CommtNote As String
    Dim FlagSubmitDate As Boolean
    Dim i As Integer
    
    i = 0
    Do While i < TxtExecutName.ListCount
        If TxtExecutName.Selected(i) = True Then Submiter = Submiter & "," & Trim(TxtExecutName.List(i))
        i = i + 1
    Loop
    
    Submiter = Mid(Submiter, 2, Len(Submiter) - 1)

    SubmitDate = DTPickerExecutDate.Value
    Approver = ""
    ApproveDate = 0 - 0 - 0
    Rejecter = ""
    RejectDate = 0 - 0 - 0
    RejectReason = ""
    CommtNote = TxtComment
    
    Dim Conn As New ADODB.Connection
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    If Len(Trim(TxtFinsGdIndex.Text)) = 0 Then        '�ж�TxtFinsGdIndex(����FinishGood NO)���ݵĺϷ���
        MsgBox "You must enter a new 12NC for the BOM(Finish Goods) Number", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(TxtFinsGdIndex.Text)) = 12 And Isnum(Trim(TxtFinsGdIndex.Text))) Then
        MsgBox "BOM(Finish Goods) Number is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    ElseIf Submiter = "" Then
        MsgBox "Please choose the submiter.", vbCritical
        Exit Sub
    Else        '��ʼ�ж������Finish Good NO �Ƿ������ݿ�������
        rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(TxtFinsGdIndex.Text) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The BOM(Finish Goods) Number is not registered in Finish Goods database", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
            Exit Sub
        End If
        If rs.State = adStateOpen Then rs.Close
    End If
    
    '�жϼ�¼�Ƿ����
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        If SystemAdmin = "Y" Then
            'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
            GoTo AdminGoAhead1
        End If
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        If InStr(Trim(rs("Submiter")), PDMUserName) = 0 Then   '�жϵ�ǰ�û��Ƿ���BOM��owner(Submiter)
            MsgBox "You are NOT BOM owner(Submiter),NO right to submit ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
            Exit Sub
        End If
AdminGoAhead1:
        If Not IsNull(rs("SubmitDate")) Then    '���SubmitDate��Ϊ��(��ʾ�Ѿ��ύSubmit��BOM)
            
            If MsgBox("The BOM (Finish Goods) Number Record already Submit. " & vbCrLf & "     Resubmit if You Update your BOM." & vbCrLf & "Resubmit will clear Previous Approval and Rejection Info." & vbCrLf & "     Are You Sure to Resubmit?", vbYesNo + vbDefaultButton2, "Confirm to Resubmit") = vbNo Then
                If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
                Exit Sub
            End If
            
            If MsgBox("Please Choose Yes if you only Clear Approval Info. and Get Right to update BOM." & vbCrLf & "Please Choose No if you Finish BOM update and Resumbit to wait approve ", vbYesNo + vbDefaultButton2, "Confirm Resubmit Purpose ") = vbNo Then
                rs("SubmitDate") = SubmitDate
            Else
                rs("SubmitDate") = Null
            End If
            
            FlagSubmitDate = True   'FlagSubmitDate���߹������жϵı��
            GoTo ContinueSubmit
        Else
ContinueSubmit:
            rs("Submiter") = Submiter
            
            If Not FlagSubmitDate Then
                rs("SubmitDate") = SubmitDate
            End If
            
            If Trim(rs("Approver")) <> "" Then
                rs("CheckHistory") = "#Approved:" & Trim(rs("Approver")) & "/" & Format(rs("ApproveDate"), "YYYY/MM/DD") & rs("CheckHistory")
            End If
            
            If Trim(rs("Rejecter")) <> "" Then
                rs("CheckHistory") = "#Rejected:" & Trim(rs("Rejecter")) & "/" & Format(rs("RejectDate"), "YYYY/MM/DD") & "/" & Trim(rs("RejectReason")) & rs("CheckHistory")
            End If
            rs("Approver") = Approver
            rs("ApproveDate") = Null
            rs("Rejecter") = Rejecter
            rs("RejectDate") = Null
            rs("RejectReason") = RejectReason
            rs("CommtNote") = CommtNote
            rs.Update
        End If
    Else
        '��¼�����ڲſ�ʼд��
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        rs.Open "INSERT INTO BOMSubmitApprove (FinsGdIndex,Description,Submiter,SubmitDate,CommtNote) VALUES ('" & Trim(TxtFinsGdIndex) _
        & "','" & Trim(TxtDescription) & "','" & Submiter & "','" & SubmitDate & "','" & CommtNote & "')", Conn, adOpenKeyset, adLockOptimistic
    End If
    
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    MsgBox "BOM (Finish Goods) Number Record has Submitted ", vbInformation, "System Info."
    Conn.Close
    
    Call Refresh_BOMSubmitApprove   '�����ɺ���ˢ��һ��
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdSubmit"
End Sub
Private Sub CmdApprove_Click()
    On Error GoTo vbErrorHandler
    
    Dim Approver As String
    Dim ApproveDate As Date
    Dim Rejecter As String
    Dim RejectDate As Date
    Dim RejectReason As String
    Dim CommtNote As String
    
    Approver = Trim(PDMUserName)
    ApproveDate = DTPickerExecutDate.Value
    Rejecter = ""
    RejectDate = 0 - 0 - 0
    RejectReason = ""
    CommtNote = TxtComment
    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    
    '�жϼ�¼�Ƿ����
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "The BOM (Finish Goods) Number Record NOT Exist,Please Submit firstly ", vbInformation, "System Info."
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Exit Sub
    Else
        If Trim(rs("SubmitDate")) = "" Or IsNull(rs("SubmitDate")) Then
            MsgBox "The BOM (Finish Goods) Number still not be Submitted (No Submit Date), Can NOT Approve  ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
            Exit Sub
        End If
        rs("Approver") = Approver
        rs("ApproveDate") = ApproveDate
        rs("Rejecter") = Rejecter
        rs("RejectDate") = Null
        rs("RejectReason") = RejectReason
        rs("CommtNote") = CommtNote
        rs.Update
    End If
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    MsgBox "BOM (Finish Goods) Number Record has Approved ", vbInformation, "System Info."
    Conn.Close
    
    Call Refresh_BOMSubmitApprove   '�����ɺ���ˢ��һ��
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdApprove"
End Sub


Private Sub CmdReject_Click()
    On Error GoTo vbErrorHandler
    
    Dim Approver As String
    Dim ApproveDate As Date
    Dim Rejecter As String
    Dim RejectDate As Date
    Dim RejectReason As String
    Dim CommtNote As String
    
    Approver = ""
    ApproveDate = 0 - 0 - 0
    Rejecter = Trim(PDMUserName)
    RejectDate = DTPickerExecutDate.Value
    RejectReason = TxtRejectReason
    CommtNote = TxtComment
    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    
    
    '�жϼ�¼�Ƿ����
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(TxtFinsGdIndex) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "The BOM (Finish Goods) Number Record NOT Exist,Please Submit firstly ", vbInformation, "System Info."
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Exit Sub
    Else
        If Trim(rs("SubmitDate")) = "" Or IsNull(rs("SubmitDate")) Then
            MsgBox "The BOM (Finish Goods) Number still not be Submitted (No Submit Date), Can NOT Reject  ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
            Exit Sub
        End If
        rs2.Open "Select IsNull(Max(BOMVersion),1) from BOMCPCN Where BOMID=" & Trim(TxtFinsGdIndex), Conn, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            If rs2(0) > 1 Then
                MsgBox "The BOM has been approved, unable to reject it."
                rs2.Close
                Exit Sub
            End If
        End If
        rs("Approver") = Approver
        rs("ApproveDate") = Null
        rs("Rejecter") = Rejecter
        rs("RejectDate") = RejectDate
        rs("RejectReason") = RejectReason
        rs("CommtNote") = CommtNote
        rs.Update
    End If
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    MsgBox "BOM (Finish Goods) Number Record has Rejected ", vbInformation, "System Info."
    Conn.Close
    
    Call Refresh_BOMSubmitApprove   '�����ɺ���ˢ��һ��
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdApprove"
    
End Sub

Private Sub CmdToQuery_Click()
'    QuerytableName = "BOMSubmitApprove"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
'
'    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
'    If SystemAdmin <> "Y" Then
'        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
'        FrmQuery.CmdModify.Enabled = False
'        FrmQuery.CmdDel.Enabled = False
'
'        FrmQuery.DataGrid1.AllowDelete = False
'        FrmQuery.DataGrid1.AllowAddNew = False
'        FrmQuery.DataGrid1.AllowUpdate = False
'    End If
'    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
'    Set FromForm = Me
'    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
'    FrmQuery.Caption = "PDM-BOM Approval"
    Dim StrSql As String
    On Error Resume Next
    '�������ݿ�
    If Conn.State = adStateOpen Then Conn.Close
    If objrs.State = adStateOpen Then objrs.Close
    Conn.ConnectionString = connString
    Conn.Open
    StrSql = "select * from BOMSubmitApprove where 1=1"
    If Trim(txtBN.Text) <> "" Then StrSql = StrSql & " And FinsGdIndex like '" & Trim(txtBN.Text) & "%'"
    If Trim(txtDesc.Text) <> "" Then StrSql = StrSql & " And Description like '" & Trim(txtDesc.Text) & "%'"
    StrSql = StrSql + " Order by FinsGdIndex"
    objrs.Open StrSql, Conn, adOpenStatic, adLockOptimistic  '����һ��Static���͵��α�,�����¼��RecordCount��Ϊ-1 '##########��Ӧ������BOMSubmitApprove

    Set DataGrid1.DataSource = objrs
    If Err Then MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdSearch"

End Sub

Private Sub CmdDelete_Click()
    On Error GoTo vbErrorHandler
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Dim TempFinsGdIndex As String
    
    TempFinsGdIndex = Trim(Str(objrs.Fields(0)))     '����ɾ��ȷ�϶Ի��� Str�����ֱ��ַ����ĺ���,�����������Str�����
    
    If MsgBox("Confirm to delete " + TempFinsGdIndex + "?" + vbCrLf + "�Ƿ�ɾ�� " + TempFinsGdIndex + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete ȷ��ɾ��") = vbYes Then
        rs.Open "Delete from BOMSubmitApprove Where FinsGdIndex ='" & TempFinsGdIndex & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        MsgBox "Succeed to delete, ɾ���ɹ�", vbInformation, "System Info."
    End If
    Conn.Close
    
    Call Refresh_BOMSubmitApprove
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMApproval:CmdDelete"
End Sub
Private Sub CmdExit_Click()
    On Error Resume Next
    Set objrs = Nothing
    If Conn.State = adStateOpen Then Conn.Close
    Unload Me
    If FromForm.Caption = Me.Caption Then Set FromForm = FrmEngineeringSys
    FromForm.Show 0
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i, j As Integer
    Dim ArrSubmiter() As String
    If KeyCode = 38 Then objrs.MovePrevious
    If KeyCode = 40 Then objrs.MoveNext
    If objrs.EOF Then objrs.MoveFirst
    If objrs.BOF Then objrs.MoveLast
    
    If IsNull(objrs.Fields("FinsGdIndex")) Then
        TxtFinsGdIndex = ""
    Else
        TxtFinsGdIndex = Trim(objrs.Fields("FinsGdIndex"))             '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    If IsNull(objrs.Fields("Description")) Then
        TxtDescription = ""
    Else
        TxtDescription = Trim(objrs.Fields("Description"))             '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    If IsNull(objrs.Fields("CommtNote")) Then
        TxtComment = ""
    Else
        TxtComment = Trim(objrs.Fields("CommtNote"))                 '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    
    For i = 0 To TxtExecutName.ListCount - 1
        TxtExecutName.Selected(i) = False
    Next
    ArrSubmiter = Split(Trim(objrs.Fields("Submiter")), ",")
    For i = 0 To TxtExecutName.ListCount - 1
        For j = 0 To UBound(ArrSubmiter)
            If ArrSubmiter(j) = TxtExecutName.List(i) Then TxtExecutName.Selected(i) = True
        Next
    Next
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Dim i, j As Integer
    Dim ArrSubmiter() As String
    If objrs.EOF Then objrs.MoveFirst
    If objrs.BOF Then objrs.MoveLast
    
    If IsNull(objrs.Fields("FinsGdIndex")) Then
        TxtFinsGdIndex = ""
    Else
        TxtFinsGdIndex = Trim(objrs.Fields("FinsGdIndex"))             '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    If IsNull(objrs.Fields("Description")) Then
        TxtDescription = ""
    Else
        TxtDescription = Trim(objrs.Fields("Description"))             '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    If IsNull(objrs.Fields("CommtNote")) Then
        TxtComment = ""
    Else
        TxtComment = Trim(objrs.Fields("CommtNote"))                 '##########��Ӧ�༭���ڿؼ���ֵ
    End If
    For i = 0 To TxtExecutName.ListCount - 1
        TxtExecutName.Selected(i) = False
    Next
    ArrSubmiter = Split(Trim(objrs.Fields("Submiter")), ",")
    For i = 0 To TxtExecutName.ListCount - 1
        For j = 0 To UBound(ArrSubmiter)
            If ArrSubmiter(j) = TxtExecutName.List(i) Then TxtExecutName.Selected(i) = True
        Next
    Next
End Sub


Private Sub Form_Resize()
    'ȷ������ı�ʱ�ؼ���֮�ı�
    Resize_ALL Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objrs = Nothing
    If Conn.State = adStateOpen Then Conn.Close
    Unload Me
    If FromForm.Caption = Me.Caption Then Set FromForm = FrmEngineeringSys
    FromForm.Show 0
End Sub

Private Sub Form_Load()
    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    'Load Skin & Format Control
    'LoadSkin Me
    '''Call ResizeInit(Me)
    
    DTPickerExecutDate.Value = Date
    DataGrid1.AllowDelete = False
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    

    If SystemAdmin <> "Y" Then    '����Ա��ApproveȨ��

        CmdApprove.Enabled = False          'PDM�û�û��ApproveȨ��
        
        CmdReject.Enabled = False
        
        Label7.ForeColor = &HC0C0C0
        TxtRejectReason.Enabled = False
    End If
    
    rs.Open "Select * from Users Order by Name ", Conn, adOpenKeyset, adLockOptimistic
    Do While rs.EOF = False
        TxtExecutName.AddItem Trim(rs("Name"))
        rs.MoveNext
    Loop
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    Conn.Close
    
    Call Refresh_BOMSubmitApprove
End Sub
Private Sub Refresh_BOMSubmitApprove()
    On Error Resume Next
    '�������ݿ�
    If Conn.State = adStateOpen Then Conn.Close
    If objrs.State = adStateOpen Then objrs.Close
    Conn.ConnectionString = connString
    Conn.Open
    
    objrs.Open "select * from BOMSubmitApprove Order by FinsGdIndex", Conn, adOpenStatic, adLockOptimistic  '����һ��Static���͵��α�,�����¼��RecordCount��Ϊ-1 '##########��Ӧ������BOMSubmitApprove

    Set DataGrid1.DataSource = objrs
End Sub

