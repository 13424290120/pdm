VERSION 5.00
Begin VB.Form FrmSERImport 
   Caption         =   "Import SER data from Excel to SQL DataBase table"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10530
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox TxtEndRow 
      Height          =   495
      Left            =   6150
      TabIndex        =   5
      Top             =   3360
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
      Left            =   6480
      TabIndex        =   4
      Top             =   5220
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
      Left            =   1755
      TabIndex        =   3
      Top             =   5220
      Width           =   3480
   End
   Begin VB.TextBox TxtStartRow 
      Height          =   495
      Left            =   6150
      TabIndex        =   2
      Top             =   2535
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
      Left            =   6420
      TabIndex        =   1
      Top             =   1245
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
      Left            =   2145
      TabIndex        =   0
      Top             =   1245
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
      Left            =   2340
      TabIndex        =   7
      Top             =   3465
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
      Left            =   2355
      TabIndex        =   6
      Top             =   2640
      Width           =   3600
   End
End
Attribute VB_Name = "FrmSERImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private Sub CmdCloseExcel_Click()
        If Dir("C:\mfg\SER.bz") <> "" Then    '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)        'ִ��EXCEL�رպ�
             xlBook.Close (True)                      '�ر�EXCEL������
            xlApp.Quit                                '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
End Sub

 
Private Sub CmdOpenExcel_Click()               '��EXCEL����
     If Dir("C:\mfg\SER.bz") = "" Then '�ж�EXCEL�Ƿ��
         Set xlApp = CreateObject("Excel.Application")   '����EXCELӦ�������
         xlApp.Visible = True                            '����EXCELӦ�������ɼ�
         Set xlBook = xlApp.Workbooks.Open("C:\mfg\SER.xls") '��EXCEL������
         Set xlSheet = xlBook.Worksheets(1)                         '��EXCEL������1
         xlSheet.Activate '����EXCEL������
         'xlSheet.Cells(1, 2) = "vvv" '����Ԫ����1��2��ֵ             ���������
        xlBook.RunAutoMacros (xlAutoOpen) '����EXCEL�е�������
    Else
        MsgBox ("EXCEL is Opened")
 End If
End Sub

Private Sub CmdQuit_Click()
        If Dir("C:\mfg\SER.bz") <> "" Then '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)      'ִ��EXCEL�رպ�
             xlBook.Close (True)                     '�ر�EXCEL������
            xlApp.Quit                               '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
        
Unload Me
FrmEngineeringSys.Show 0
End Sub

Private Sub CmdWrite_Click()
Dim SERNO As String
Dim SgPtNO As String
Dim ImportSER As ClsSER
Set ImportSER = New ClsSER
Dim i As Integer
'Dim J As Integer
For i = val(Trim(TxtStartRow.Text)) To val(Trim(TxtEndRow.Text))
  If xlSheet.Cells(i, 1) <> "" Then            'ÿ�е�1���в�Ϊ����ʼд
   
      With ImportSER    '�Ѿ�����Public MySER As New ClsSER, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
     SERNO = xlSheet.Cells(i, 1)
     SERNO = "SER00000" & SERNO
    .SERIndex = SERNO
    .Applicant = xlSheet.Cells(i, 2)
    .CAorA = xlSheet.Cells(i, 3)
    .Description = xlSheet.Cells(i, 5)
    .IDSO = "Open"                              '�̶�����ֵ
    .OpnDate = xlSheet.Cells(i, 6)
    .ClosDate = xlSheet.Cells(i, 6)
    .PJNOIndex = 999999                         '�̶�����ֵ
    .PjtName = xlSheet.Cells(i, 7)
    .FinsGdNO = "NA"
     SgPtNO = xlSheet.Cells(i, 4)
     SgPtNO = Replace(SgPtNO, " ", "")          'ȥ���м�Ŀո�
     
    .SglPrtNO = val(SgPtNO)
    .CommtNote = xlSheet.Cells(i, 8)
    
           '�ж�SERIndex����Ƿ��Ѿ�����
                If .In_DB(val(xlSheet.Cells(i, 1))) = True Then
                   MsgBox "In" + Str(i) + " row, SER number exists, Please go next" + vbCrLf + "�ڵ�" + Str(i) + " ��, SER���ظ����������һ��д��¼", vbInformation, "System Info."
                
                Else
                   .Insert                   '���
                    'MsgBox "In" + Str(I) + "row, Succeed to Add" + vbCrLf + "�ڵ�" + Str(I) + " ��, ��ӳɹ�", vbInformation, "System Info."    '���Ե�����¼��,����д�Ļ�ȥ�����
                End If

      End With
   End If
Next

End Sub


Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
End Sub
