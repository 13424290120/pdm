VERSION 5.00
Begin VB.Form FrmSglPrtImport 
   Caption         =   "Import Single Part data from Excel to SQL DataBase table"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSglPrtImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10470
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox TxtEndRow 
      Height          =   495
      Left            =   5910
      TabIndex        =   6
      Top             =   3075
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
      Left            =   6375
      TabIndex        =   4
      Top             =   4950
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
      Left            =   1650
      TabIndex        =   3
      Top             =   4950
      Width           =   3480
   End
   Begin VB.TextBox TxtStartRow 
      Height          =   495
      Left            =   5910
      TabIndex        =   2
      Top             =   2250
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
      Left            =   6225
      TabIndex        =   1
      Top             =   960
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
      Left            =   2010
      TabIndex        =   0
      Top             =   960
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
      Left            =   2100
      TabIndex        =   7
      Top             =   3180
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
      Left            =   2115
      TabIndex        =   5
      Top             =   2355
      Width           =   3600
   End
End
Attribute VB_Name = "FrmSglPrtImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private Sub CmdCloseExcel_Click()
        If Dir("C:\mfg\SinglePart.bz") <> "" Then '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)      'ִ��EXCEL�رպ�
             xlBook.Close (True)                     '�ر�EXCEL������
            xlApp.Quit                               '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
End Sub

 
Private Sub CmdOpenExcel_Click()               '��EXCEL����
     If Dir("C:\mfg\SinglePart.bz") = "" Then '�ж�EXCEL�Ƿ��
         Set xlApp = CreateObject("Excel.Application")   '����EXCELӦ�������
         xlApp.Visible = True                            '����EXCELӦ�������ɼ�
         Set xlBook = xlApp.Workbooks.Open("C:\mfg\SinglePart.xls") '��EXCEL������
         Set xlSheet = xlBook.Worksheets(1)                         '��EXCEL������1
         xlSheet.Activate '����EXCEL������
         'xlSheet.Cells(1, 2) = "vvv" '����Ԫ����1��2��ֵ             ���������
        xlBook.RunAutoMacros (xlAutoOpen) '����EXCEL�е�������
    Else
        MsgBox ("EXCEL is Opened")
 End If
End Sub

Private Sub CmdQuit_Click()
        If Dir("C:\mfg\SinglePart.bz") <> "" Then '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)      'ִ��EXCEL�رպ�
             xlBook.Close (True)                     '�ر�EXCEL������
            xlApp.Quit                               '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
        
Unload Me
End Sub

Private Sub CmdWrite_Click()
Dim ImportSglPrt As ClsSglPrt
Set ImportSglPrt = New ClsSglPrt
Dim i As Integer
'Dim J As Integer
For i = val(Trim(TxtStartRow.Text)) To val(Trim(TxtEndRow.Text))
  If xlSheet.Cells(i, 1) <> "" Then            'ÿ�е�1���в�Ϊ����ʼд
   
      With ImportSglPrt    '�Ѿ�����Public MySglPrt As New ClsSglPrt, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .SglPrtIndex = val(left(xlSheet.Cells(i, 1), 11) + "0")
    .SglPrtVer = val(right(xlSheet.Cells(i, 1), 1))
    .PrtUnit = xlSheet.Cells(i, 3)
    .Applicant = "NA"                           '�̶�����ֵ
    .Description = xlSheet.Cells(i, 2)
    .IDSO = "Open"                              '�̶�����ֵ
    .NewOldStatus = "Old"                       '�̶�����ֵ
    .OpnDate = Date
    .ClosDate = Date
    .PJNOIndex = 999999                         '�̶�����ֵ
    .PjtName = "NA"                             '�̶�����ֵ
    .ProductLine = "5000"                       '�̶�����ֵ
    .ItemType = xlSheet.Cells(i, 4)
    .Location = "TR-AV"                         '�̶�����ֵ
    .CommtNote = "NA"                           '�̶�����ֵ
    
           '�ж�SglPrtIndex����Ƿ��Ѿ�����
                If .In_DB(val(left(xlSheet.Cells(i, 1), 11) + "0")) = True Then
                   MsgBox "In" + Str(i) + " row, Single Part number exists, Please go next" + vbCrLf + "�ڵ�" + Str(i) + " ��, Single Part���ظ����������һ��д��¼", vbInformation, "System Info."
                
                Else
                   .Insert                   '���
                    'MsgBox "In" + Str(I) + "row, Succeed to Add" + vbCrLf + "�ڵ�" + Str(I) + " ��, ��ӳɹ�"  '���Ե�����¼��,����д�Ļ�ȥ�����
                End If

      End With
   End If
Next

End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
End Sub
