VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmGlueSupplier 
   Caption         =   "PDM-Gule/Electro Supplier Admin ���̹�����ϵͳ"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "FrmGlueSupplier.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   9390
   StartUpPosition =   2  '��Ļ����
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
      Left            =   6435
      TabIndex        =   14
      Top             =   1020
      Width           =   900
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
      Left            =   1080
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1560
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
      Left            =   8160
      TabIndex        =   12
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtPage_nd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtPage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   960
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
      Left            =   6120
      TabIndex        =   9
      Top             =   1560
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1560
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
      Left            =   2730
      TabIndex        =   7
      Top             =   1560
      Width           =   1395
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
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4070
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   7335
      _ExtentX        =   12938
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Glue12NC"
         Caption         =   "Glue/Electro Part 12NC"
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
         DataField       =   "SupplierName"
         Caption         =   "SupplierName"
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
         DataField       =   "SupplierPN"
         Caption         =   "SupplierPN"
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
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2234.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   6840
      Picture         =   "FrmGlueSupplier.frx":08CA
      Top             =   6360
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
      Left            =   7440
      MouseIcon       =   "FrmGlueSupplier.frx":0CE6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   4680
      Picture         =   "FrmGlueSupplier.frx":0FF0
      Top             =   6360
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
      Left            =   5280
      MouseIcon       =   "FrmGlueSupplier.frx":140C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2640
      Picture         =   "FrmGlueSupplier.frx":1716
      Top             =   6360
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
      Left            =   3240
      MouseIcon       =   "FrmGlueSupplier.frx":1B32
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   840
      Picture         =   "FrmGlueSupplier.frx":1E3C
      Top             =   6360
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
      Left            =   1440
      MouseIcon       =   "FrmGlueSupplier.frx":2258
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmGlueSupplier.frx":2562
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2745
      TabIndex        =   0
      Top             =   600
      Width           =   4005
   End
End
Attribute VB_Name = "FrmGlueSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lCurrentpage As Long           '���嵱ǰҳ����
Dim Conn As New ADODB.Connection   '����һ��ADO����

Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
Dim objrs As New ADODB.Recordset    '������һ����¼�����ڴ��ÿһҳ�ļ�¼

Private Sub CmdFirst_Click()     '��1ҳ����
   lCurrentpage = 1
   Call Refresh_GlueSupplier(lCurrentpage)
End Sub

Private Sub CmdFresh_Click()
 Call Refresh_GlueSupplier(lCurrentpage)
End Sub

Private Sub CmdLast_Click()          '��ĩҳ����
   lCurrentpage = 10000
   Call Refresh_GlueSupplier(lCurrentpage)
End Sub

Private Sub CmdNext_Click()           '��1ҳ����
   lCurrentpage = lCurrentpage + 1
   Call Refresh_GlueSupplier(lCurrentpage)
End Sub

Private Sub CmdPrevious_Click()       '��1ҳ����
 If lCurrentpage > 1 Then
   lCurrentpage = lCurrentpage - 1
   Call Refresh_GlueSupplier(lCurrentpage)
 End If
End Sub

Private Sub CmdToQuery_Click()
QuerytableName = "GlueSupplier"   '����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub PageGO_Click()          'ȥ��ָ��ҳ
   If Not IsNumeric(txtPage_nd) Then
       MsgBox "Page No. must be Number, No letter " + vbCrLf + "������ҳ������ֱ��", vbInformation, "Error Info!"
       txtPage_nd.SetFocus
   End If
   
   If val(txtPage_nd.Text) < 1 Then
   lCurrentpage = 1
   Call Refresh_GlueSupplier(lCurrentpage)
   Exit Sub
   End If
   
   lCurrentpage = val(txtPage_nd.Text)  'val�������ַ���ת������ֵ
   Call Refresh_GlueSupplier(lCurrentpage)

End Sub


Private Sub txtPage_nd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then PageGO_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rcds = Nothing
Set objrs = Nothing
If Conn.State = adStateOpen Then Conn.Close
Unload Me
FrmEngineeringSys.Show 0
End Sub
Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
 lCurrentpage = 1           '���ڴ�Ĭ���ǵ�1ҳ����
 Call Refresh_GlueSupplier(lCurrentpage)
End Sub

Private Sub LblAdd_Click()
FrmGlueSupplierEdit.Caption = "Add Glue/Electro Supplier Info."

'��������Ϊ��Ӷ����޸�
FrmGlueSupplierEdit.Modify = False

'ˢ�½�ˮ��Ӧ�̹������
lCurrentpage = 10000           '����Ӽ�¼ʱһ��Ĭ��ȥ��ĩҳ����
Call Refresh_GlueSupplier(lCurrentpage)

'��ʾ��ˮ��Ӧ����Ϣ�༭����
FrmGlueSupplierEdit.Show 1
Call Refresh_GlueSupplier(lCurrentpage) '�����ɺ���ˢ��һ��
End Sub

Private Sub LblBack_Click()
Set rcds = Nothing
Set objrs = Nothing
If Conn.State = adStateOpen Then Conn.Close
Unload Me
FrmEngineeringSys.Show 0
End Sub


Private Sub LblDelete_Click()
Dim TempGlueSupplierID As String
'�����ɾ����¼��ID
  TempGlueSupplierID = objrs.Fields(0)
  
'����ɾ��ȷ�϶Ի��� Str�����ֱ��ַ����ĺ���,�����������Str�����
  If MsgBox("Confirm to delete" + Str(objrs.Fields(0)) + "?" + vbCrLf + "�Ƿ�ɾ��" + Str(objrs.Fields(0)) + "?", vbYesNo + vbDefaultButton2, "Confirm to Delete ȷ��ɾ��") = vbYes Then
    
    '��������Delete����ɾ����ˮ��Ӧ����Ϣ
    MyGlueSupplier.Delete (TempGlueSupplierID)
    MsgBox "Succeed to delete, ɾ���ɹ�", vbInformation, "System Info."
  End If
  'ˢ�½�ˮ��Ӧ�̹������
Call Refresh_GlueSupplier(lCurrentpage)
End Sub


Private Sub LblModify_Click()

'������޸ļ�¼��ԭʼID
FrmGlueSupplierEdit.OriGlue12NC = Trim(objrs.Fields(0))

'�Ѵ��޸���Ϣ��ӵ��༭����
FrmGlueSupplierEdit.TxtGlue12NC = Trim(objrs.Fields(0))
FrmGlueSupplierEdit.TxtSupplierName = Trim(objrs.Fields(1))
FrmGlueSupplierEdit.TxtSupplierPN = Trim(objrs.Fields(2))

FrmGlueSupplierEdit.TxtGlue12NC.Enabled = False  '��Ȼ���޸ģ����������ǲ��ܸĵ�
FrmGlueSupplierEdit.Caption = "Modify Glue/Electro Supplier Info."
'��������Ϊ�޸Ĳ���
FrmGlueSupplierEdit.Modify = True
'��ʾ��ˮ��Ӧ�̱༭����
FrmGlueSupplierEdit.Show 1
'ˢ�½�ˮ��Ӧ�̹������
Call Refresh_GlueSupplier(lCurrentpage)
End Sub


Private Sub Refresh_GlueSupplier(lPage As Long)
          Dim adoPrimaryRS     As ADODB.Recordset
          Dim lPageCount     As Long
          Dim nPageSize     As Integer
          Dim lCount     As Long
          
  '�������ݿ�
Conn.ConnectionString = connString
Conn.Open

rcds.Open "select * from GlueSupplier", Conn, adOpenKeyset, adOpenStatic  '����һ��Static���͵��α�,�����¼��RecordCount��Ϊ-1

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
          For lCount = 0 To rcds.Fields.count - 1
            If lCount = 0 Then
              objrs.Fields.Append rcds.Fields(lCount).Name, adUnsignedBigInt, rcds.Fields(lCount).DefinedSize  'adUnsignedBigInt   8�ֽڲ�����������
              GoTo NextLine
            End If
            objrs.Fields.Append rcds.Fields(lCount).Name, adVarChar, rcds.Fields(lCount).DefinedSize  'adVarChar�����ֶ����ַ���
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
                  objrs!Glue12NC = rcds!Glue12NC
                  objrs!SupplierName = rcds!SupplierName
                  objrs!SupplierPN = rcds!SupplierPN
                  rcds.MoveNext
          Next
          '��
          Set DataGrid1.DataSource = objrs
            
          '��ʾҳ��
          txtPage.Text = lPage & "/" & rcds.PageCount
Conn.Close
 
End Sub


