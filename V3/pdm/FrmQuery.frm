VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmQuery 
   Caption         =   "General Search Window ͨ�ò�ѯ����"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   876
   StartUpPosition =   2  '��Ļ����
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8745
      Left            =   60
      TabIndex        =   21
      Top             =   1950
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   15425
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdExportExcel 
      Caption         =   "Export / Print Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   18
      Top             =   1380
      Width           =   885
   End
   Begin VB.TextBox TxtQry2 
      Height          =   345
      Left            =   7890
      TabIndex        =   17
      Top             =   480
      Width           =   2000
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12090
      TabIndex        =   14
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11070
      TabIndex        =   13
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   11
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton CmdExecQ 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      TabIndex        =   10
      Top             =   180
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   60
      TabIndex        =   4
      Top             =   1050
      Width           =   9915
      Begin VB.TextBox txtReason 
         Height          =   405
         Left            =   7350
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ComboBox CmboDate 
         Height          =   345
         ItemData        =   "FrmQuery.frx":08CA
         Left            =   450
         List            =   "FrmQuery.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   285
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   3670017
         CurrentDate     =   39974
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2445
         TabIndex        =   6
         Top             =   285
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         Format          =   3670017
         CurrentDate     =   39974
      End
      Begin VB.CheckBox ChkBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   345
         Width           =   255
      End
      Begin VB.Label lblReason 
         Caption         =   "Reason:"
         Height          =   375
         Left            =   6540
         TabIndex        =   22
         Top             =   390
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4050
         TabIndex        =   9
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1935
         TabIndex        =   8
         Top             =   345
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Condition      ���ò�ѯ����Ŀ������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9915
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   450
         Width           =   255
      End
      Begin VB.ComboBox CmboEqut2 
         Height          =   345
         Left            =   6810
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   975
      End
      Begin VB.CheckBox ChkBox3 
         Caption         =   "And"
         Height          =   375
         Left            =   6090
         TabIndex        =   15
         Top             =   420
         Width           =   765
      End
      Begin VB.TextBox TxtQry1 
         Height          =   345
         Left            =   4020
         TabIndex        =   3
         Top             =   420
         Width           =   2000
      End
      Begin VB.ComboBox CmboEqut1 
         Height          =   345
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
      Begin VB.ComboBox CmboItem 
         Height          =   345
         Left            =   450
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   2595
      End
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ô�ͨ�ò�ѯ������Ҫ����
' 1- QueryTableNameͨ�ò�ѯ���ڲ�ѯ�����ı����� ��Form Load��

Option Explicit
Dim QryItem As Long             '���嵱ǰ�ֶ�������ѭ����ѯ����
Dim QrysqlStr As String         '����SQL��ѯ����ַ�����
Dim DtGrdLen As Long            '����Gridĳһ�����ڵ�����0��1��2...���߶�Ӧ��¼���ֶε��ֶ����
Dim response As Integer         'msgbox����ֵ�ж�
Dim Qcnn As New ADODB.Connection   '����һ��ADO����
Dim QRS As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼

Private Sub Check1_Click()
    If Check1.Value = 0 Then ChkBox2.Value = 0
End Sub

Private Sub ChkBox3_Click()
    If ChkBox3.Value = 1 Then
        CmboEqut1.ListIndex = 2    '������β��ұ�־ChkBox3.Value����Ϊ1�����ֻ����>��
        Check1.Value = 1
        CmboEqut1.Enabled = False
    Else
        CmboEqut1.Enabled = True
    End If
End Sub

Private Sub CmboEqut1_Click()
    If ChkBox3.Value = 1 Then
        CmboEqut1.ListIndex = 2    '����Ϊ>����,ֻ����ʾ��
    End If
End Sub

Private Sub cmdAdd_Click()
    If QueryTableName = "SglPrt" Then
        FrmSglPrtEdit.Caption = "Add One Single Part Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmSglPrtEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        
        '��������ģʽҪ����һЩ�ؼ�
        
        FrmSglPrtEdit.TxtSglPrtVer.Visible = False
        FrmSglPrtEdit.TxtPrtUnit.Visible = False
        FrmSglPrtEdit.TxtIDSO.Visible = False
        FrmSglPrtEdit.TxtNewOldStatus.Visible = False
        FrmSglPrtEdit.TxtOpnDate.Visible = False
        FrmSglPrtEdit.TxtClosDate.Visible = False
        FrmSglPrtEdit.TxtProductLine.Visible = False
        FrmSglPrtEdit.TxtItemType.Visible = False
        FrmSglPrtEdit.TxtLocation.Visible = False
        
        FrmSglPrtEdit.LblOld0.Visible = False
        FrmSglPrtEdit.LblOld1.Visible = False
        FrmSglPrtEdit.LblOld2.Visible = False
        FrmSglPrtEdit.LblOld3.Visible = False
        FrmSglPrtEdit.LblOld4.Visible = False
        FrmSglPrtEdit.LblOld5.Visible = False
        FrmSglPrtEdit.LblOld6.Visible = False
        FrmSglPrtEdit.LblOld7.Visible = False
        FrmSglPrtEdit.LblOld8.Visible = False
        FrmSglPrtEdit.LblReminder.Visible = False
                
        FrmSglPrtEdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "RFSRFQ" Then
        FrmRFSRFQEdit.Caption = "Add One RFQ or RFS Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmRFSRFQEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        '��������ģʽҪ����һЩ�ؼ�
        FrmRFSRFQEdit.TxtIDSQ.Visible = False
        FrmRFSRFQEdit.TxtOpnDate.Visible = False
        FrmRFSRFQEdit.TxtClosDate.Visible = False
        FrmRFSRFQEdit.LblOld0.Visible = False
        FrmRFSRFQEdit.LblOld1.Visible = False
        FrmRFSRFQEdit.LblOld2.Visible = False
        FrmRFSRFQEdit.LblReminder.Visible = False
        FrmRFSRFQEdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "CNCSN" Then
        FrmCNCSNEdit.Caption = "Add One CONCESSION Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmCNCSNEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        
        '��������ģʽҪ����һЩ�ؼ�
        FrmCNCSNEdit.TxtCPCNMP.Visible = False
        FrmCNCSNEdit.TxtIDSO.Visible = False
        FrmCNCSNEdit.TxtOpnDate.Visible = False
        FrmCNCSNEdit.TxtClosDate.Visible = False
        
        FrmCNCSNEdit.LblOld0.Visible = False
        FrmCNCSNEdit.LblOld1.Visible = False
        FrmCNCSNEdit.LblOld2.Visible = False
        FrmCNCSNEdit.LblOld3.Visible = False
        FrmCNCSNEdit.LblReminder.Visible = False
                
        FrmCNCSNEdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "SER" Then
        FrmSEREdit.Caption = "Add One SER Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmSEREdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        
        '��������ģʽҪ����һЩ�ؼ�
        FrmSEREdit.TxtCAorA.Visible = False
        FrmSEREdit.TxtIDSO.Visible = False
        FrmSEREdit.TxtOpnDate.Visible = False
        FrmSEREdit.TxtClosDate.Visible = False
        FrmSEREdit.LblOld0.Visible = False
        FrmSEREdit.LblOld1.Visible = False
        FrmSEREdit.LblOld2.Visible = False
        FrmSEREdit.LblOld3.Visible = False
        FrmSEREdit.LblReminder.Visible = False
        FrmSEREdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "PJNO" Then
        FrmPJNOEdit.Caption = "Add One Project Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmPJNOEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        
        '��������ģʽҪ����һЩ�ؼ�
        FrmPJNOEdit.TxtIDSQ.Visible = False
        FrmPJNOEdit.TxtOpnDate.Visible = False
        FrmPJNOEdit.TxtClosDate.Visible = False
        FrmPJNOEdit.LblOld0.Visible = False
        FrmPJNOEdit.LblOld1.Visible = False
        FrmPJNOEdit.LblOld2.Visible = False
        FrmPJNOEdit.LblReminder.Visible = False
        FrmPJNOEdit.Show 0                                    '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "CPCN" Then
        FrmCPCNEdit.Caption = "Add One CP/CN Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmCPCNEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        '��������ģʽҪ����һЩ�ؼ�
        FrmCPCNEdit.TxtCPCNMP.Visible = False
        FrmCPCNEdit.TxtIDSO.Visible = False
        FrmCPCNEdit.TxtOpnDate.Visible = False
        FrmCPCNEdit.TxtClosDate.Visible = False
        FrmCPCNEdit.LblOld0.Visible = False
        FrmCPCNEdit.LblOld1.Visible = False
        FrmCPCNEdit.LblOld2.Visible = False
        FrmCPCNEdit.LblOld3.Visible = False
        FrmCPCNEdit.LblReminder.Visible = False
        FrmCPCNEdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    ElseIf QueryTableName = "FinsGd" Then
        FrmFinsGdEdit.Caption = "Add One Finish Goods Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ
        '��������Ϊ��Ӷ����޸�
        FrmFinsGdEdit.Modify = False                             '##########�ڶ�Ӧ�򿪴�����Modify��ʾҪ��ֵ
        
        '��������ģʽҪ����һЩ�ؼ�
        FrmFinsGdEdit.TxtIDSO.Visible = False
        FrmFinsGdEdit.TxtOpnDate.Visible = False
        FrmFinsGdEdit.TxtClosDate.Visible = False
        FrmFinsGdEdit.LblReminder.Visible = False
        FrmFinsGdEdit.TxtProductLine.Visible = False
        FrmFinsGdEdit.TxtItemType.Visible = False
        FrmFinsGdEdit.TxtLocation.Visible = False
        
        FrmFinsGdEdit.LblOld0.Visible = False
        FrmFinsGdEdit.LblOld1.Visible = False
        FrmFinsGdEdit.LblOld2.Visible = False
        FrmFinsGdEdit.LblOld3.Visible = False
        FrmFinsGdEdit.LblOld4.Visible = False
        FrmFinsGdEdit.LblOld5.Visible = False
        FrmFinsGdEdit.CmdSysDistrb.Enabled = True
        FrmFinsGdEdit.Show 0                                     '##########��Ӧ�༭���ڴ�
    End If
    CmdExecQ_Click
End Sub

'ɾ����Ԫ��(��¼)
Private Sub CmdDel_Click()
    If QRS.EOF Then
        MsgBox "No Chosed Record!"
        Exit Sub
    Else
        Dim TempID As String

        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
        If SystemAdmin <> "Y" Then
          MsgBox "you are not administrator, No right to delete", vbInformation, "System Info."
          Exit Sub
        End If
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
        
        TempID = QRS.Fields(0)                      '##########TempSglPrtID�����ɶ�Ӧ����Key�ֶ���
        '����ɾ��ȷ�϶Ի��� Str�����ֱ��ַ����ĺ���,�����������Str�����
        If MsgBox("Confirm to delete" + CStr(QRS.Fields(0)) + "?" + vbCrLf + "�Ƿ�ɾ��" + CStr(QRS.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete ȷ��ɾ��") = vbYes Then
            If QueryTableName = "SglPrt" Then
                '��������Delete����ɾ��SglPrt��Ϣ
                MySglPrt.Delete (TempID)
            ElseIf QueryTableName = "RFSRFQ" Then
                  '��������Delete����ɾ��RFQ/RFS��Ϣ
                MyRFSRFQ.Delete (TempID)
            ElseIf QueryTableName = "CNCSN" Then
                MyCNCSN.Delete (TempID)
            ElseIf QueryTableName = "PJNO" Then
                MyPJNO.Delete (TempID)
            ElseIf QueryTableName = "CPCN" Then
                MyCPCN.Delete (TempID)
            ElseIf QueryTableName = "FinsGd" Then
                MyFinsGd.Delete (TempID)
            ElseIf QueryTableName = "SER" Then
                MySER.Delete (TempID)
            End If
            MsgBox "Succeed to delete, ɾ���ɹ�", vbInformation, "System Info."
        End If
    End If
    CmdExecQ_Click
End Sub



'�޸ĵ�Ԫ��(��¼)
Private Sub cmdModify_Click()
    If QRS.EOF Then
        MsgBox "No Chosed Record!"
    Else
        If PDMUserName <> Trim(QRS.Fields("Applicant")) And SystemAdmin <> "Y" Then MsgBox "No right to modify it.", vbInformation: Exit Sub
        If QueryTableName = "SglPrt" Then
            '������޸ļ�¼��ԭʼID
            FrmSglPrtEdit.OriSglPrtIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))       '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmSglPrtEdit.TxtSglPrtIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))       '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtSglPrtVer = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))       '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtPrtUnit = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))       '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtApplicant = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtProductLine = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtIDSO = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))                   '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtNewOldStatus = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))                   '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtClosDate = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(12)), "", FormatNumber6(Trim(QRS.Fields(11))))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtPjtName = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtItemType = IIf(QRS.Fields(13) = Null, "", Trim(QRS.Fields(13)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtLocation = IIf(QRS.Fields(14) = Null, "", Trim(QRS.Fields(14)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSglPrtEdit.TxtCommtNote = IIf(QRS.Fields(15) = Null, "", Trim(QRS.Fields(15)))              '##########��Ӧ�༭���ڿؼ���ֵ
            
            FrmSglPrtEdit.TxtSglPrtIndex.Locked = True   '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmSglPrtEdit.TxtApplicant.Locked = True       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmSglPrtEdit.Caption = "Modify One Single Part Number."                                  '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmSglPrtEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            FrmSglPrtEdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "RFSRFQ" Then
            '������޸ļ�¼��ԭʼID
            FrmRFSRFQEdit.OriRFSRFQIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmRFSRFQEdit.TxtRFSRFQIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtLeader = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                 '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtIDSQ = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                   '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmRFSRFQEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########��Ӧ�༭���ڿؼ���ֵ
        
            FrmRFSRFQEdit.TxtRFSRFQIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmRFSRFQEdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmRFSRFQEdit.Caption = "Modify one RFQ/RFS Number."                                           '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmRFSRFQEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            FrmRFSRFQEdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "CNCSN" Then
            '������޸ļ�¼��ԭʼID
            FrmCNCSNEdit.OriCNCSNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))           '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmCNCSNEdit.TxtCNCSNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtCPCNMP = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))         '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                  '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCNCSNEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))            '##########��Ӧ�༭���ڿؼ���ֵ
            
            FrmCNCSNEdit.TxtCNCSNIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmCNCSNEdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmCNCSNEdit.Caption = "Modify One CONCESSION Number."                                  '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmCNCSNEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            
            FrmCNCSNEdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "PJNO" Then
            '������޸ļ�¼��ԭʼID
            FrmPJNOEdit.OriPJNOIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))        '##########��Ӧ�༭���ڱ�����ֵ
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmPJNOEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtLeader = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtIDSQ = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                 '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmPJNOEdit.TxtPJNOIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmPJNOEdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmPJNOEdit.Caption = "Modify One Project Number."                                           '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmPJNOEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            FrmPJNOEdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "SER" Then
            '������޸ļ�¼��ԭʼID
            FrmSEREdit.OriSERIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmSEREdit.TxtSERIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))        '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))           '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtCAorA = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))          '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))                  '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmSEREdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))              '##########��Ӧ�༭���ڿؼ���ֵ
            
            FrmSEREdit.TxtSERIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmSEREdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmSEREdit.Caption = "Modify One SER Number."                                  '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmSEREdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            
            FrmSEREdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "CPCN" Then
            '������޸ļ�¼��ԭʼID
            FrmCPCNEdit.OriCPCNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))         '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmCPCNEdit.TxtCPCNIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))           '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtCPCNMP = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))           '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtIDSO = IIf(IsNull(QRS.Fields(4)), "", Trim(QRS.Fields(4)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtClosDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtPjtName = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtFinsGdNO = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtSglPrtNO = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))              '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmCPCNEdit.txtReason = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))             '##########��Ӧ�༭���ڿؼ���ֵ
            
            FrmCPCNEdit.TxtCPCNIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmCPCNEdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmCPCNEdit.Caption = "Modify One CP/CN Number."                                  '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmCPCNEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            FrmCPCNEdit.Show 0                                            '##########��Ӧ�༭���ڴ�
        ElseIf QueryTableName = "FinsGd" Then
            '������޸ļ�¼��ԭʼID
            FrmFinsGdEdit.OriFinsGdIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))         '##########��Ӧ�༭���ڱ�����ֵ
            
            '�Ѵ��޸���Ϣ��ӵ��༭����
            FrmFinsGdEdit.TxtFinsGdIndex = IIf(IsNull(QRS.Fields(0)), "", Trim(QRS.Fields(0)))          '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtApplicant = IIf(IsNull(QRS.Fields(1)), "", Trim(QRS.Fields(1)))            '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtProductLine = IIf(IsNull(QRS.Fields(2)), "", Trim(QRS.Fields(2)))                '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtDescription = IIf(IsNull(QRS.Fields(3)), "", Trim(QRS.Fields(3)))        '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtIDSO = IIf(IsNull(QRS.Fields(5)), "", Trim(QRS.Fields(5)))                  '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtOpnDate = IIf(IsNull(QRS.Fields(6)), "", Trim(QRS.Fields(6)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtClosDate = IIf(IsNull(QRS.Fields(7)), "", Trim(QRS.Fields(7)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtPJNOIndex = IIf(IsNull(QRS.Fields(8)), "", Trim(QRS.Fields(8)))             '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtPjtName = IIf(IsNull(QRS.Fields(9)), "", Trim(QRS.Fields(9)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtItemType = IIf(IsNull(QRS.Fields(10)), "", Trim(QRS.Fields(10)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtLocation = IIf(IsNull(QRS.Fields(11)), "", Trim(QRS.Fields(11)))               '##########��Ӧ�༭���ڿؼ���ֵ
            FrmFinsGdEdit.TxtCommtNote = IIf(IsNull(QRS.Fields(12)), "", Trim(QRS.Fields(12)))                '##########��Ӧ�༭���ڿؼ���ֵ
            
            FrmFinsGdEdit.TxtFinsGdIndex.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
            FrmFinsGdEdit.TxtApplicant.Enabled = False       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
            FrmFinsGdEdit.Caption = "Modify One Finish Goods Number."                                  '##########��Ӧ�༭���ڱ���
            '��������Ϊ�޸Ĳ���
            FrmFinsGdEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ
            FrmFinsGdEdit.CmdSysDistrb.Enabled = False
            FrmFinsGdEdit.Show 0                                            '##########��Ӧ�༭���ڴ�

        Else
        End If
    End If
    CmdExecQ_Click
End Sub
'��ͨ�ò�ѯ�����˳�
Private Sub CmdExit_Click()
    On Error Resume Next
    If QRS.State = adStateOpen Then QRS.Close
    Set QRS = Nothing
    If Qcnn.State = adStateOpen Then Qcnn.Close
    Set Qcnn = Nothing
    Unload Me
    FromForm.Show 0
End Sub

Private Sub CmdExportExcel_Click()
    On Error Resume Next

    Dim i As Integer
    Dim sHeader As String
    Set xlApp = CreateObject("Excel.Application")   '����Excel�ļ�
    Set xlApp = New excel.Application
    
'        '������ֲ���������ʾ
'    xlApp.OleRequestPendingTimeout = 10000   '10000��������æ�Ի���
'    xlApp.OleServerBusyTimeout = 1000     '����ʱ1��
'    xlApp.OleServerBusyRaiseError = True '����ʾæ�Ի���
    
    
    xlApp.SheetsInNewWorkbook = 1                   '���½��Ĺ�����������Ϊ1
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)              '��1�Ź�����
    If QueryTableName = "SglPrt" Then
        sHeader = "Single Part Number"
    ElseIf QueryTableName = "RFSRFQ" Then
        sHeader = "RFS/RFQ Number"
    ElseIf QueryTableName = "CNCSN" Then
        sHeader = "CONCESSION Number"
    ElseIf QueryTableName = "PJNO" Then
        sHeader = "Project Number"
    ElseIf QueryTableName = "SER" Then
        sHeader = "SER Number"
    ElseIf QueryTableName = "CPCN" Then
        sHeader = "CPCN Number"
    ElseIf QueryTableName = "FinsGd" Then
        sHeader = "Finish Goods Number"
    End If
    xlSheet.Cells(1, 1) = sHeader
    For i = 0 To DataGrid1.Columns.count - 1
        xlSheet.Cells(3, i + 1) = DataGrid1.Columns(i).Caption
    Next i
    
    xlSheet.Cells(2, i - 3) = "Table Maker:": xlSheet.Cells(2, i - 2) = PDMUserName
    xlSheet.Cells(2, i - 1) = "Print Date:": xlSheet.Cells(2, i) = Now()
        
    xlSheet.Cells(4, 1).CopyFromRecordset Qcnn.Execute(QrysqlStr)       '������ճ������
    xlSheet.Columns("K").NumberFormat = "############"

    xlApp.ActiveWorkbook.Close True     '�رչ�����������
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If QRS.State = adStateOpen Then QRS.Close
    Set QRS = Nothing
    If Qcnn.State = adStateOpen Then Qcnn.Close
    Set Qcnn = Nothing
    Unload Me
    FromForm.Show 0
End Sub


Private Sub DataGrid1_Error(ByVal dataerror As Integer, response As Integer)
response = 0
MsgBox "check the Input Data type or Length", vbInformation, "Error Info!"
End Sub
Private Sub DataGrid1_colEdit(ByVal colindex As Integer)

On Error Resume Next
DataGrid1.SetFocus
DataGrid1.SelStart = 0
DataGrid1.SelLength = Len(QRS.Fields(DtGrdLen))   'DataGrid��Ԫ���Ӧ��¼���ֶε��ֶ���ŵĳ���
 
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)  'ע�������ڲ���������,���˱������
    
    Select Case DataGrid1.Col
        Case 0                             '0������Ϊ���ɸ���  �����1, 2, 3 ��������µ��л���������1, 2, 3...
            DataGrid1.AllowUpdate = False
        Case Else
            DataGrid1.AllowUpdate = True
            
                 If SystemAdmin <> "Y" Then             '��һ���жϣ��������sa�û��Ļ����ǲ����޸�
                 DataGrid1.AllowUpdate = False
                 End If

            DtGrdLen = DataGrid1.Col      'DataGrid1.Col�ǵ��DataGrid��Ԫ�񷵻ظõ�Ԫ�����ڵ�����0��1��2...
    ' TxtTest.Text = DataGrid1.Col  ������������ʹ��ʱ�ڴ���������Ϊ�ɼ�visible = true
    End Select
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me

If QueryTableName = "SglPrt" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "RFSRFQ" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "CNCSN" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "PJNO" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "CPCN" Then
    CmdExportExcel.Visible = True
    lblReason.Visible = True
    txtReason.Visible = True
ElseIf QueryTableName = "SER" Then
    CmdExportExcel.Visible = True
ElseIf QueryTableName = "FinsGd" Then
    CmdExportExcel.Visible = False
ElseIf QueryTableName = "BOMSubmitApprove" Then
    CmdExportExcel.Visible = False
    cmdAdd.Visible = False
    cmdModify.Visible = False
    CmdDel.Visible = False
End If

Qcnn.ConnectionString = connString
Qcnn.Open


'SERҪ������Ѿ����������
If QueryTableName = "SER" Then
    QRS.Open "select * from " + QueryTableName + " where applicant<>''", Qcnn, adOpenKeyset, adOpenStatic  ' ���ݱ������ Public QueryTableName As String
Else
    QRS.Open "select * from " + QueryTableName + " where 1>2", Qcnn, adOpenKeyset, adOpenStatic  ' ���ݱ������ Public QueryTableName As String
End If
Set DataGrid1.DataSource = QRS
DataGrid1.Columns(0).Width = 150  '��һ������ʾ���п�Ȳ��������⣬���Ե��������п�
'Call AutoFitWidth(DataGrid1)

          For QryItem = 0 To QRS.Fields.count - 1          '��ѯ��Ŀֵ��ʼ��
          
          Check1.Value = 1                                '��һ����ѯ��Ĭ�Ͽ�����
          
             If InStr(QRS.Fields(QryItem).Name, "Date") <> 0 Then  'instrָ��һ�ַ�������һ�ַ��������ȳ��ֵ�λ��
               ChkBox2.Value = 0                                '�ڶ�����ѯ��Ҳ������
               CmboDate.AddItem (QRS.Fields(QryItem).Name)
               CmboDate.ListIndex = 0                           '��һ��ĿĬ��ѡ��
           GoTo NextLine                                    '�����Date�ֶ�ֵ����뵽��2����ѯ��
             End If
        
          CmboItem.AddItem (QRS.Fields(QryItem).Name)
NextLine:
          Next
          CmboItem.ListIndex = 0                           '��һ��ĿĬ��ѡ��
          
          CmboEqut1.AddItem ("=")                           '��ѯ��ʽֵ��ʼ��
          CmboEqut1.AddItem ("like")
          CmboEqut1.AddItem (">")
          CmboEqut1.AddItem ("<")
          
          CmboEqut1.ListIndex = 0                           '��һ��ĿĬ��ѡ��
          
          CmboEqut2.AddItem ("<")
          CmboEqut2.ListIndex = 0                           '��һ��ĿĬ��ѡ��
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub


Private Sub CmdExecQ_Click()
On Error Resume Next
'SQL��ѯ����ַ�����ʽ  SELECT   Field_1   FROM   Table_A   WHERE Field_1 = 'DDD'  ORDER  BY  Field_1
'����:  SqlStmt = "SELECT Glue12NC FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"

QrysqlStr = "Select * from" + " "                                     '��Ҫ�������Ĳ�ѯ�ֶ�
QrysqlStr = QrysqlStr + QueryTableName + " Where 1=1"                          '�����ѯ�ı��� ���ݱ������ Public QueryTableName As String

If Check1.Value = 1 Then
'    For QryItem = 0 To Qrs.Fields.Count - 1
'        '�����0�ֶε�����������(12NC)���ҵ�0�ֶε��ֶ�����Combo��ѡ��Ҫ��ѯ
'        If Qrs.Fields(QryItem).Name = CmboItem.Text Then    'adBigInt(ֵΪ20)8�ֽڴ���������,adUnsignedBigInt(ֵΪ21)8�ֽڲ�����������
'           If (Qrs.Fields(QryItem).Type = 20) And (Not IsNumeric(TxtQry1.Text)) Then   '�������ѯ�����ݶ�ǡ�ò�������(12NC)
'              MsgBox " Input query type is not matching, check if it should be Number", vbInformation, "Error Info!"
'              Exit Sub
'           Else
                If CmboEqut1.Text = "like" Then                   '����SQL������Likeƥ���ѯ���ַ���Ҫ��%�ţ������=����һ��
                    QrysqlStr = QrysqlStr + " And " & CmboItem.Text & " like '%" + Trim(TxtQry1.Text) + "%'"
                Else
                    QrysqlStr = QrysqlStr + " And " & CmboItem.Text & CmboEqut1.Text & "'" & Trim(TxtQry1.Text) & "'"
                End If
'           End If
'        End If
'    Next
    If ChkBox3.Value = 1 Then
       '�ַ��������ʽ��and XXDate Between #2009-01-20# and #2009-05-15#  ע��#������Access���ݿ��
       'Select * from GlueSupplier Where (Glue12NC Between XXXXXXXA and XXXXXXXB)  And (RdDate Between '2007-04-10' and '2008-11-10')Order By Glue12NC
        QrysqlStr = QrysqlStr + " And " & CmboItem.Text & CmboEqut2.Text & "'" & Trim(TxtQry2.Text) & "'"
       'TxtTest.Text = QrysqlStr  '������������ʹ��ʱ�ڴ���������Ϊ�ɼ�visible = true
    End If
End If


' Format(DTPicker1.Value, "YYYY/MM/DD hh:mm:ss") ʱ������ʽ�����ַ�����
If ChkBox2.Value = 1 Then                                               '�ڶ�����ѯ�����ʱ��Ĭ���д���ѯ��Ŀ
        QrysqlStr = QrysqlStr + " And (" + CmboDate.Text + " Between " + "'" + Format(DTPicker1.Value, "YYYY/MM/DD") + "'" + " and " + "'" + Format(DTPicker2.Value, "YYYY/MM/DD") + "'" + ")" '�������ڵĲ�ѯ�ַ���
       '�����Format(DTPicker1.Value, "YYYY/MM/DD")������Ҫ '���Ÿ������������ ע�� '������SQL���ݿ��
End If

If QueryTableName = "CPCN" Then QrysqlStr = QrysqlStr + " And Reason like '%" & txtReason.Text & "%'"

If Check1.Value = 1 Then QrysqlStr = QrysqlStr + " Order By" + " " + CmboItem.Text
Debug.Print QrysqlStr
Set QRS = Nothing  'ԭ��¼�е�������Ҫ����ղ���д
QRS.Open QrysqlStr, Qcnn, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = QRS
DataGrid1.Columns(1).Width = 150  '��һ������ʾ���п�Ȳ��������⣬���Ե��������п�

End Sub

Private Sub TxtQry1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then CmdExecQ_Click
End Sub

Public Sub AutoFitWidth(ByRef dg As DataGrid)

Dim tr As ADODB.Recordset
Dim r As ADODB.Recordset
Set tr = dg.DataSource
If tr Is Nothing Then Exit Sub
If tr.State = 0 Then Exit Sub
If tr.RecordCount = 0 Then Exit Sub

Set r = tr.Clone

Dim m
Dim Width

For m = 0 To dg.Columns.count - 1
    Width = Len(dg.Columns(m).Caption)
    r.MoveFirst
   While Not r.EOF
      If Len(r(m)) > Width Then
          Width = Len(r(m))
      End If
      r.MoveNext
   Wend
   dg.Columns(m).Width = Width * 229
Next m

End Sub
