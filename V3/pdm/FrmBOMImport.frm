VERSION 5.00
Begin VB.Form FrmBOMImport 
   Caption         =   "Import BOM from An Excel File"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7620
   ControlBox      =   0   'False
   Icon            =   "FrmBOMImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7620
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
      Left            =   1035
      TabIndex        =   5
      Top             =   315
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
      Left            =   4365
      TabIndex        =   4
      Top             =   315
      Width           =   2535
   End
   Begin VB.TextBox TxtStartRow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4455
      TabIndex        =   3
      Text            =   "2"
      Top             =   1335
      Width           =   2220
   End
   Begin VB.CommandButton CmdWrite 
      Caption         =   "Start to write BOM from Excel"
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
      Left            =   390
      TabIndex        =   2
      Top             =   3285
      Width           =   3540
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
      Left            =   4680
      TabIndex        =   1
      Top             =   3285
      Width           =   2505
   End
   Begin VB.TextBox TxtEndRow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4455
      TabIndex        =   0
      Top             =   2295
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmBOMImport.frx":08CA
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
      Left            =   645
      TabIndex        =   7
      Top             =   1305
      Width           =   3600
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmBOMImport.frx":0905
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
      Left            =   645
      TabIndex        =   6
      Top             =   2265
      Width           =   3600
   End
End
Attribute VB_Name = "FrmBOMImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Isnum(Str As String) As Boolean     '�ж�һ���ַ������Ƿ�������  ��IsNumeric�ж�0000d031Ϊ��(����double������)
  Isnum = True
  Dim i  As Integer
  For i = 1 To Len(Str)
      Select Case Mid(Str, i, 1)
          Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          ' Isnum = True  ����дIsnum = True�ͳ���,��Ϊ����м�����ĸfalse�˺��������ֵĻ��ֳ�Ϊtrue��
          Case Else
            Isnum = False
      End Select
  Next
End Function
 
Private Sub CmdOpenExcel_Click()
    On Error Resume Next
    Dim strBOMFileName  As String
    strBOMFileName = App.Path + "\BOMTemplate.xls"            '��Ҫ�򿪵��ļ���·�����ļ���
    
    ModulesOpenExcel (strBOMFileName)
End Sub
Private Sub CmdCloseExcel_Click()
    On Error Resume Next
    Dim strBOMFileName  As String
    strBOMFileName = App.Path + "\BOMTemplate.xls"            '��Ҫ�رյ��ļ���·�����ļ���
    
    ModulesCloseExcel (strBOMFileName)
End Sub


Private Sub CmdQuit_Click()
    On Error Resume Next
    Dim strBOMFileName  As String
    strBOMFileName = App.Path + "\BOMTemplate.xls"            '��Ҫ�رյ��ļ���·�����ļ���
    
    ModulesCloseExcel (strBOMFileName)
    
    Unload Me
    FromForm2.Show 0
End Sub

Private Sub CmdWrite_Click()
Screen.MousePointer = vbHourglass   '����ʱ��ϳ�����Ҫ�������״̬
On Error GoTo vbErrorHandler
    Dim TempFinsGdIndex As String
    Dim TempDescription As String
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn

Dim i As Integer
Dim FlagString As String
Dim ParentString As String
Dim ChildString As String
Dim QtyString As String

TempFinsGdIndex = ""
TempDescription = ""

For i = CInt(Trim(TxtStartRow.Text)) To CInt(Trim(TxtEndRow.Text))
 
FlagString = xlSheet.Cells(i, 1)
ParentString = xlSheet.Cells(i, 2)
ChildString = xlSheet.Cells(i, 3)
QtyString = xlSheet.Cells(i, 4)

    If Trim(FlagString) = "Y" Then
        If Len(Trim(ParentString)) = 0 Then        '�ж�����Parent Item���ݵĺϷ���
            MsgBox "In" + Str(i) + " row, You must enter a 12NC for the Parent Item", vbInformation, "System Info."
            Exit Sub
        ElseIf Not (Len(Trim(ParentString)) = 12 And Isnum(Trim(ParentString))) Then
               MsgBox "In" + Str(i) + " row, Parent Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
               Exit Sub
            Else        '��ʼ�ж������Parent Item�Ƿ���Finish Goods���ݿ�������
               rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(ParentString) & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
                    MsgBox "In" + Str(i) + " row, The Finish goods NO(Parent Item). is not existing in database", vbInformation, "System Info."
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
               Else
                   If TempFinsGdIndex = "" And TempDescription = "" Then
                   TempFinsGdIndex = ParentString
                   TempDescription = rs("Description")
                   End If
               End If
                    If rs.State = adStateOpen Then rs.Close
        End If
        GoTo chckChildStrg
    Else
        If Trim(FlagString) <> "N" Then Exit Sub
             If Len(Trim(ParentString)) = 0 Then        '�ж�����Parent Item���ݵĺϷ���
             MsgBox "In" + Str(i) + " row, You must enter a 12NC for the Parent Item", vbInformation, "System Info."
             Exit Sub
        ElseIf Not (Len(Trim(ParentString)) = 12 And Isnum(Trim(ParentString))) Then
               MsgBox "In" + Str(i) + " row, Parent Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
               Exit Sub
            Else        '��ʼ�ж������Parent Item�Ƿ���Single Part���ݿ�������
               rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(ParentString), 1, Len(Trim(ParentString)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
               MsgBox "In" + Str(i) + " row, The Single part NO(Parent Item). is not existing in database", vbInformation, "System Info."
                     If rs.State = adStateOpen Then rs.Close
               Exit Sub
               End If
                     If rs.State = adStateOpen Then rs.Close
        End If
chckChildStrg:
        If Len(Trim(ChildString)) = 0 Then        '�ж�����Child Item���ݵĺϷ���
              MsgBox "In" + Str(i) + " row, You must enter a 12NC for the Child Item", vbInformation, "System Info."
              Exit Sub
        ElseIf Not (Len(Trim(ChildString)) = 12 And Isnum(Trim(ChildString))) Then
               MsgBox "In" + Str(i) + " row, Child Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
               Exit Sub
            Else        '��ʼ�ж������Child Item�Ƿ���Single Part���ݿ�������
                rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(ChildString), 1, Len(Trim(ChildString)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                MsgBox "In" + Str(i) + " row, The Single part NO(Child Item). is not existing in database", vbInformation, "System Info."
                       If rs.State = adStateOpen Then rs.Close
                Exit Sub
                End If
                       If rs.State = adStateOpen Then rs.Close
        End If
        
        If Len(Trim(QtyString)) = 0 Then        '�ж�����quantity���ݵĺϷ���
             MsgBox "In" + Str(i) + " row, You must enter a number for the quantity", vbInformation, "System Info."
             Exit Sub
        ElseIf Not IsNumeric(val(Trim(QtyString))) Then
            MsgBox "In" + Str(i) + " row, Quantity Item is Number, no Letter" + vbCrLf + "����������,����ĸ", vbInformation, "System Info."
            Exit Sub
        End If
        
    End If
        
         '����ж���BOMOrigData��Parent Item + Child Item�����ļ�¼�Ƿ����
         rs.Open "Select * from BOMOrigData Where ChildID='" & Trim(ChildString) & "'" & " and  ParentID ='" & Trim(ParentString) & "'", Conn, adOpenKeyset, adLockOptimistic

         If rs.RecordCount > 0 Then
             MsgBox "In" + Str(i) + " row, The record has already existed, Not repeat to add.", vbInformation, "System Info"
             If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
             GoTo ContinueNext
         Else
                 If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
             rs.Open "INSERT INTO BOMOrigData (ParentID, ChildID, Quantity) VALUES (" & Trim(ParentString) & "," & Trim(ChildString) & "," & Round(Trim(QtyString), 7) & ")", Conn, adOpenKeyset, adLockOptimistic
                 If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
         End If
ContinueNext:
Next i
    
    Screen.MousePointer = vbDefault                  '�ָ����״̬
    MsgBox "You have finished importing one BOM from Excel file", vbInformation, "System Info."
    
    '�Ǽ�BOM�����߻��ύ��
    '���жϼ�¼�Ƿ����
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & TempFinsGdIndex & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
         If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
         Exit Sub
    Else
         If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
         rs.Open "INSERT INTO BOMSubmitApprove (FinsGdIndex,Description,Submiter) VALUES ('" & TempFinsGdIndex & "','" & TempDescription & "','" & PDMUserName & "')", Conn, adOpenKeyset, adLockOptimistic
         If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    End If
    Set xlSheet = Nothing
    Conn.Close
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMImport:CmdWrite"
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
End Sub
