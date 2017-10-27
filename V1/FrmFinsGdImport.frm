VERSION 5.00
Begin VB.Form FrmFinsGdImport 
   Caption         =   "Import Finish Goods data from Excel to SQL DataBase table"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10485
   Icon            =   "FrmFinsGdImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10485
   StartUpPosition =   2  '��Ļ����
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
      Left            =   2055
      TabIndex        =   5
      Top             =   1260
      Width           =   2415
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
      Left            =   6330
      TabIndex        =   4
      Top             =   1260
      Width           =   2535
   End
   Begin VB.TextBox TxtStartRow 
      Height          =   495
      Left            =   6060
      TabIndex        =   3
      Top             =   2550
      Width           =   2220
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
      Left            =   1665
      TabIndex        =   2
      Top             =   5235
      Width           =   3480
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
      Left            =   6390
      TabIndex        =   1
      Top             =   5235
      Width           =   2505
   End
   Begin VB.TextBox TxtEndRow 
      Height          =   495
      Left            =   6060
      TabIndex        =   0
      Top             =   3375
      Width           =   2220
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
      Left            =   2265
      TabIndex        =   7
      Top             =   2655
      Width           =   3600
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
      Left            =   2250
      TabIndex        =   6
      Top             =   3480
      Width           =   3600
   End
End
Attribute VB_Name = "FrmFinsGdImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private Sub CmdCloseExcel_Click()
        If Dir("C:\mfg\FinishGoods.bz") <> "" Then    '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)        'ִ��EXCEL�رպ�
             xlBook.Close (True)                      '�ر�EXCEL������
            xlApp.Quit                                '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
End Sub

 
Private Sub CmdOpenExcel_Click()               '��EXCEL����
     If Dir("C:\mfg\FinishGoods.bz") = "" Then '�ж�EXCEL�Ƿ��
         Set xlApp = CreateObject("Excel.Application")   '����EXCELӦ�������
         xlApp.Visible = True                            '����EXCELӦ�������ɼ�
         Set xlBook = xlApp.Workbooks.Open("C:\mfg\FinishGoods.xls") '��EXCEL������
         Set xlSheet = xlBook.Worksheets(1)                         '��EXCEL������1
         xlSheet.Activate '����EXCEL������
         'xlSheet.Cells(1, 2) = "vvv" '����Ԫ����1��2��ֵ             ���������
        xlBook.RunAutoMacros (xlAutoOpen) '����EXCEL�е�������
    Else
        MsgBox ("EXCEL is Opened")
 End If
End Sub

Private Sub CmdQuit_Click()
        If Dir("C:\mfg\FinishGoods.bz") <> "" Then '��VB�ر�EXCEL
            xlBook.RunAutoMacros (xlAutoClose)      'ִ��EXCEL�رպ�
             xlBook.Close (True)                     '�ر�EXCEL������
            xlApp.Quit                               '�ر�EXCEL
        End If
            Set xlApp = Nothing                       '�ͷ�EXCEL����
        End
        
Unload Me
End Sub

Private Sub CmdWrite_Click()
Dim ImportFinsGd As ClsFinsGd
Set ImportFinsGd = New ClsFinsGd
Dim i As Integer
'Dim J As Integer
For i = val(Trim(TxtStartRow.Text)) To val(Trim(TxtEndRow.Text))
  If xlSheet.Cells(i, 1) <> "" Then            'ÿ�е�1���в�Ϊ����ʼд
   
      With ImportFinsGd    '�Ѿ�����Public MyFinsGd As New ClsFinsGd, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .FinsGdIndex = val(xlSheet.Cells(i, 1))
    .Applicant = "NA"                           '�̶�����ֵ
    .ProductLine = xlSheet.Cells(i, 3)                       '�̶�����ֵ
    .Description = xlSheet.Cells(i, 2)
    .IDSO = "Open"                              '�̶�����ֵ
    .OpnDate = Date
    .ClosDate = Date
    .PJNOIndex = 999999                         '�̶�����ֵ
    .PjtName = "NA"                             '�̶�����ֵ
    .ItemType = "400"
    .Location = "AV/CAR"                         '�̶�����ֵ
    .CommtNote = "NA"                           '�̶�����ֵ
    
           '�ж�FinsGdIndex����Ƿ��Ѿ�����
                If .In_DB(val(xlSheet.Cells(i, 1))) = True Then
                   MsgBox "In" + Str(i) + " row, Finish Goods number exists, Please go next" + vbCrLf + "�ڵ�" + Str(i) + " ��, Finish Goods���ظ����������һ��д��¼", vbInformation, "System Info."
                
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
