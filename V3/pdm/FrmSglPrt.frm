VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSglPrt 
   Caption         =   "PDM-Single Part Number Admin ���̹�����ϵͳ"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSglPrt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   14460
   StartUpPosition =   2  '��Ļ����
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4070
      Left            =   255
      TabIndex        =   14
      Top             =   2580
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   7170
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      DefColWidth     =   80
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "SglPrtIndex"
         Caption         =   "SglPrtIndex"
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
         DataField       =   "SglPrtVer"
         Caption         =   "Ver"
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
         DataField       =   "PrtUnit"
         Caption         =   "Unit"
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
      BeginProperty Column04 
         DataField       =   "ProductLine"
         Caption         =   "ProductLine"
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
      BeginProperty Column06 
         DataField       =   "IDSO"
         Caption         =   "IDSO"
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
         DataField       =   "NewOldStatus"
         Caption         =   "O/N"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "PjtName"
         Caption         =   "PjtName"
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
         DataField       =   "ItemType"
         Caption         =   "ItemType"
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
         DataField       =   "Location"
         Caption         =   "Location"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   299.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
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
      Left            =   9720
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CmdToQuery 
      Caption         =   "Search ��ѯ"
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
      Left            =   2550
      TabIndex        =   7
      Top             =   975
      Width           =   1695
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
      Left            =   11760
      TabIndex        =   6
      Top             =   1935
      Width           =   375
   End
   Begin VB.TextBox txtPage_nd 
      Height          =   375
      Left            =   11040
      TabIndex        =   5
      Top             =   1815
      Width           =   735
   End
   Begin VB.TextBox txtPage 
      Height          =   375
      Left            =   11040
      TabIndex        =   4
      Top             =   1215
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
      Left            =   9720
      TabIndex        =   3
      Top             =   1815
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
      Left            =   2550
      TabIndex        =   2
      Top             =   1815
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
      Left            =   4950
      TabIndex        =   1
      Top             =   1815
      Width           =   1425
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
      Left            =   7530
      TabIndex        =   0
      Top             =   1815
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   9150
      Picture         =   "FrmSglPrt.frx":08CA
      Top             =   7170
      Width           =   300
   End
   Begin VB.Label LblBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Return����"
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
      Left            =   9750
      MouseIcon       =   "FrmSglPrt.frx":0CE6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   6990
      Picture         =   "FrmSglPrt.frx":0FF0
      Top             =   7170
      Width           =   300
   End
   Begin VB.Label LblDelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Deleteɾ��"
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
      Left            =   7590
      MouseIcon       =   "FrmSglPrt.frx":140C
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4950
      Picture         =   "FrmSglPrt.frx":1716
      Top             =   7170
      Width           =   300
   End
   Begin VB.Label LblModify 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Modify�޸�"
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
      Left            =   5550
      MouseIcon       =   "FrmSglPrt.frx":1B32
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3150
      Picture         =   "FrmSglPrt.frx":1E3C
      Top             =   7170
      Width           =   300
   End
   Begin VB.Label LblAdd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add���"
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
      Left            =   3750
      MouseIcon       =   "FrmSglPrt.frx":2258
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7170
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSglPrt.frx":2562
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
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   4020
   End
End
Attribute VB_Name = "FrmSglPrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�¹���ģ���� ���е�Call Refresh_SglPrt(lCurrentpage)�е�SglPrtҪͳһ�û�Ϊ�±����
Option Explicit
Dim lCurrentpage As Long           '���嵱ǰҳ����
Dim Conn As New ADODB.Connection   '����һ��ADO����

Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
Dim objrs As New ADODB.Recordset    '������һ����¼�����ڴ��ÿһҳ�ļ�¼

Private Sub CmdFirst_Click()     '��1ҳ����
   lCurrentpage = 1
   Call Refresh_SglPrt(lCurrentpage)
End Sub

Private Sub CmdFresh_Click()        'ˢ�²���
 Call Refresh_SglPrt(lCurrentpage)
End Sub

Private Sub CmdLast_Click()          '��ĩҳ����
   lCurrentpage = 10000
   Call Refresh_SglPrt(lCurrentpage)
End Sub

Private Sub CmdNext_Click()           '��1ҳ����
   lCurrentpage = lCurrentpage + 1
   Call Refresh_SglPrt(lCurrentpage)
End Sub

Private Sub CmdPrevious_Click()       '��1ҳ����
 If lCurrentpage > 1 Then
   lCurrentpage = lCurrentpage - 1
   Call Refresh_SglPrt(lCurrentpage)
 End If
End Sub

Private Sub CmdToQuery_Click()
QuerytableName = "SglPrt"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���

    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.CmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False

        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
'���Ϻ������޹ص���Ŀ����
FrmQuery.DataGrid1.Columns(15).Visible = False
FrmQuery.DataGrid1.Columns(16).Visible = False
FrmQuery.DataGrid1.Columns(17).Visible = False
FrmQuery.DataGrid1.Columns(18).Visible = False
FrmQuery.DataGrid1.Columns(19).Visible = False
FrmQuery.DataGrid1.Columns(20).Visible = False
FrmQuery.DataGrid1.Columns(21).Visible = False
FrmQuery.DataGrid1.Columns(22).Visible = False
FrmQuery.DataGrid1.Columns(23).Visible = False
FrmQuery.DataGrid1.Columns(24).Visible = False
    
FrmQuery.Show 1 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Call ResizeForm(Me)
End Sub

Private Sub PageGO_Click()          'ȥ��ָ��ҳ
   If Not IsNumeric(txtPage_nd) Then
       MsgBox "Page No. must be Number, No letter " + vbCrLf + "������ҳ������ֱ��", vbInformation, "Error Info!"
       txtPage_nd.SetFocus
   End If
   
   If val(txtPage_nd.Text) < 1 Then
   lCurrentpage = 1
   Call Refresh_SglPrt(lCurrentpage)
   Exit Sub
   End If
   
   lCurrentpage = val(txtPage_nd.Text)  'val�������ַ���ת������ֵ
   Call Refresh_SglPrt(lCurrentpage)

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

 lCurrentpage = 1           '���ڴ�Ĭ���ǵ�1ҳ����
 Call Refresh_SglPrt(lCurrentpage)
End Sub

Private Sub LblAdd_Click()
FrmSglPrtEdit.Caption = "Add One Single Part Number."     '##########�ڶ�Ӧ�򿪴����б���Ҫ��ֵ

lCurrentpage = 10000                                 '����Ӽ�¼ʱһ��Ĭ��ȥ��ĩҳ����
Call Refresh_SglPrt(lCurrentpage)

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
        

FrmSglPrtEdit.Show 1                                     '##########��Ӧ�༭���ڴ�
Call Refresh_SglPrt(lCurrentpage) '�����ɺ���ˢ��һ��
End Sub

Private Sub LblBack_Click()
Set rcds = Nothing
Set objrs = Nothing
If Conn.State = adStateOpen Then Conn.Close
Unload Me
FrmEngineeringSys.Show
End Sub


Private Sub LblDelete_Click()

    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    If SystemAdmin <> "Y" Then
        MsgBox "you are not administrator, No right to delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    
    
Dim TempSglPrtID As String                            '##########TempSglPrtID�����ɶ�Ӧ����Key�ֶ���
'�����ɾ����¼��ID
  TempSglPrtID = objrs.Fields(0)                      '##########TempSglPrtID�����ɶ�Ӧ����Key�ֶ���
  
'����ɾ��ȷ�϶Ի��� Str�����ֱ��ַ����ĺ���,�����������Str�����
  If MsgBox("Confirm to delete" + Str(objrs.Fields(0)) + "?" + vbCrLf + "�Ƿ�ɾ��" + Str(objrs.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete ȷ��ɾ��") = vbYes Then
    
    '��������Delete����ɾ��SglPrt��Ϣ
    MySglPrt.Delete (TempSglPrtID)                    '##########TempSglPrtID�����ɶ�Ӧ����Key�ֶ���
    MsgBox "Succeed to delete, ɾ���ɹ�", vbInformation, "System Info."
  End If
  'ˢ�½�ˮ��Ӧ�̹������
Call Refresh_SglPrt(lCurrentpage)
End Sub


Private Sub LblModify_Click()

'������޸ļ�¼��ԭʼID
FrmSglPrtEdit.OriSglPrtIndex = Trim(objrs.Fields(0))           '##########��Ӧ�༭���ڱ�����ֵ

'�Ѵ��޸���Ϣ��ӵ��༭����
FrmSglPrtEdit.TxtSglPrtIndex = Trim(objrs.Fields(0))           '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtSglPrtVer = Trim(objrs.Fields(1))           '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtPrtUnit = Trim(objrs.Fields(2))           '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtDescription = Trim(objrs.Fields(3))            '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtApplicant = Trim(objrs.Fields(4))             '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtProductLine = Trim(objrs.Fields(5))                 '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtIDSO = Trim(objrs.Fields(6))                   '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtNewOldStatus = Trim(objrs.Fields(7))                   '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtOpnDate = Trim(objrs.Fields(8))                '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtClosDate = Trim(objrs.Fields(9))               '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtPJNOIndex = FormatNumber6(Trim(objrs.Fields(10)))               '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtPjtName = Trim(objrs.Fields(11))               '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtItemType = Trim(objrs.Fields(12))               '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtLocation = Trim(objrs.Fields(13))               '##########��Ӧ�༭���ڿؼ���ֵ
FrmSglPrtEdit.TxtCommtNote = Trim(objrs.Fields(14))               '##########��Ӧ�༭���ڿؼ���ֵ

FrmSglPrtEdit.TxtSglPrtIndex.Locked = True   '��Ȼ���޸ģ����������ǲ��ܸĵ�            '##########��Ӧ�༭���ںͿؼ���
FrmSglPrtEdit.TxtApplicant.Locked = True       '��Ȼ���޸ģ�������һ�㲻�ø�        '##########��Ӧ�༭���ںͿؼ���
FrmSglPrtEdit.Caption = "Modify One Single Part Number."                                  '##########��Ӧ�༭���ڱ���
'��������Ϊ�޸Ĳ���
FrmSglPrtEdit.Modify = True                                     '##########��Ӧ�༭���ڱ�����ֵ

FrmSglPrtEdit.Show 1                                            '##########��Ӧ�༭���ڴ�

Call Refresh_SglPrt(lCurrentpage)
End Sub


Private Sub Refresh_SglPrt(lPage As Long)
          Dim adoPrimaryRS     As ADODB.Recordset
          Dim lPageCount     As Long
          Dim nPageSize     As Integer
          Dim lCount     As Long
          
  '�������ݿ�
Conn.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(DBUser) + ";pwd=" + Trim(Password) + ";database=" + Trim(DataBase) + ""
Conn.Open

'����һ��Static���͵��α�,�����¼��RecordCount��Ϊ-1 '##########��Ӧ������SglPrt
rcds.Open "select SglPrtIndex,SglPrtVer,PrtUnit,Description,Applicant,ProductLine,IDSO,NewOldStatus,OpnDate,ClosDate,PJNOIndex,PjtName,ItemType,Location,CommtNote from SglPrt", Conn, adOpenKeyset, adOpenStatic

  '������ܲ鵽��¼
If rcds.RecordCount = 0 Then
  '�޸ĺ�ɾ��������
LblModify.Enabled = False
LblDelete.Enabled = False
Else
  '����ܲ鵽��¼,�޸ĺ�ɾ������
LblModify.Enabled = True
LblDelete.Enabled = True
End If

 
   'ÿҳ��ʾ�ļ�¼��Ϊ15
   nPageSize = 15
   rcds.PageSize = nPageSize         'ÿҳ��ʾ�ļ�¼����ֵ����¼������. PageSize��ҳ��ʾʱÿһҳ�ļ�¼��
' ADO PageCount ����
'The PageCount property returns a long value that indicates the number of pages with data in a Recordset object.
'PageCount���Ե������ǣ�����һ����ֵ������ָ����¼������������ҳ���������

'Tip: To divide the Recordset into a series of pages, use the PageSize property.
'��ʾ: �����ʹ��PageSize���Խ���¼���ָ�Ϊһϵ�е�ҳ��?

'Note: If the last page contains fewer records than specified in PageSize, it still counts as an additional page in the PageCount property.
'ע�⣺������һҳ�ļ�¼����������PageSize������ָ������������ô����Ȼ����Ϊһҳ��

'Note: If this method is not supported it returns -1.
'ע�⣺�����֧�������������ô������-1��

'IntFix �������ز�������������?
'�﷨
'Int(number)
'Fix(number)
'��Ҫ�� number ������ Double ���κ���Ч����ֵ���ʽ����� number ���� Null���򷵻� Null��
'˵��
'Int �� Fix ����ɾ�� number ��С�����ݶ�����ʣ�µ�������
'Int �� Fix �Ĳ�֮ͬ�����ڣ���� number Ϊ�������� Int ����С�ڻ���� number �ĵ�һ������������ Fix ��᷵�ش��ڻ���� number �ĵ�һ�������������磬Int �� -8.4 ת���� -9���� Fix �� -8.4 ת���� -8��
  lPageCount = rcds.PageCount
              If lCurrentpage > lPageCount Then
                  lCurrentpage = lPageCount
              End If
          rcds.AbsolutePage = lCurrentpage
          
Set objrs = Nothing  'ԭ��¼�е�������Ҫ����ղ���д
          '����ֶ�����
          For lCount = 0 To rcds.Fields.Count - 1
            If lCount = 0 Or lCount = 1 Or lCount = 10 Then                           ' ############## ���ڴ����ֵ��ֶ���Ҫ����������ֶ����
              objrs.Fields.Append rcds.Fields(lCount).Name, adUnsignedBigInt, rcds.Fields(lCount).DefinedSize  'adUnsignedBigInt   8�ֽڲ�����������
              GoTo NextLine
            End If
            objrs.Fields.Append rcds.Fields(lCount).Name, adVarChar, rcds.Fields(lCount).DefinedSize 'adVarChar�����ֶ����ַ���
NextLine:
          Next
          
          '�򿪼�¼��
          objrs.Open
          
          '��ָ����¼��ѭ����ӵ�objrs��
          For lCount = 1 To nPageSize   'nPageSizeÿҳ��ʾ�ļ�¼��Ϊ10
                  If rcds.EOF = True Then
                  Exit For
                  End If
                 
                  objrs.AddNew
                  objrs!SglPrtIndex = rcds!SglPrtIndex                                             '##########��Ӧ���ֶθ�ֵ
                  objrs!SglPrtVer = rcds!SglPrtVer                                                 '##########��Ӧ���ֶθ�ֵ
                  objrs!PrtUnit = rcds!PrtUnit                                                     '##########��Ӧ���ֶθ�ֵ
                  objrs!Description = rcds!Description                                             '##########��Ӧ���ֶθ�ֵ
                  objrs!Applicant = rcds!Applicant                                                 '##########��Ӧ���ֶθ�ֵ
                  objrs!ProductLine = rcds!ProductLine                                             '##########��Ӧ���ֶθ�ֵ
                  objrs!IDSO = rcds!IDSO                                                            '##########��Ӧ���ֶθ�ֵ
                  objrs!NewOldStatus = rcds!NewOldStatus                                            '##########��Ӧ���ֶθ�ֵ
                  objrs!OpnDate = Format(rcds!OpnDate, "YYYY/MM/DD")  '������Ҫ��ʽ����������       '##########��Ӧ���ֶθ�ֵ
                  objrs!ClosDate = Format(rcds!ClosDate, "YYYY/MM/DD")  '������Ҫ��ʽ����������      '##########��Ӧ���ֶθ�ֵ
                  objrs!PJNOIndex = rcds!PJNOIndex                                                   '##########��Ӧ���ֶθ�ֵ
                  objrs!PJTName = rcds!PJTName                                                      '##########��Ӧ���ֶθ�ֵ
                  objrs!ItemType = rcds!ItemType                                                    '##########��Ӧ���ֶθ�ֵ
                  objrs!Location = rcds!Location                                                    '##########��Ӧ���ֶθ�ֵ
                  objrs!CommtNote = rcds!CommtNote                                                  '##########��Ӧ���ֶθ�ֵ
                  
                  rcds.MoveNext
          Next
          '��
          Set DataGrid1.DataSource = objrs
            
          '��ʾҳ��
          txtPage.Text = lPage & "/" & rcds.PageCount
Conn.Close
 
End Sub






