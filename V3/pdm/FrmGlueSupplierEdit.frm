VERSION 5.00
Begin VB.Form FrmGlueSupplierEdit 
   Caption         =   "Glue/Electro SupplierInfo. ��ˮ�๩Ӧ����Ϣ"
   ClientHeight    =   3732
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   6936
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGlueSupplierEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   6936
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox TxtSupplierPN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2355
      Width           =   2775
   End
   Begin VB.TextBox TxtSupplierName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1635
      Width           =   2775
   End
   Begin VB.TextBox TxtGlue12NC 
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
      Left            =   3720
      TabIndex        =   3
      Top             =   915
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "ע��:��������12NC��Ҫ�汾��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   390
      TabIndex        =   8
      Top             =   210
      Width           =   6090
   End
   Begin VB.Label LblOK 
      BackStyle       =   0  'Transparent
      Caption         =   "OK ȷ ��"
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
      Left            =   1920
      MouseIcon       =   "FrmGlueSupplierEdit.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3090
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1320
      Picture         =   "FrmGlueSupplierEdit.frx":0BD4
      Top             =   3096
      Width           =   240
   End
   Begin VB.Label LblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel ȡ ��"
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
      Left            =   4200
      MouseIcon       =   "FrmGlueSupplierEdit.frx":0FF0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3090
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   3600
      Picture         =   "FrmGlueSupplierEdit.frx":12FA
      Top             =   3096
      Width           =   240
   End
   Begin VB.Label LblSupplierPN 
      BackStyle       =   0  'Transparent
      Caption         =   "SupplierPN��Ӧ���Ϻ�"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmGlueSupplierEdit.frx":1716
      TabIndex        =   2
      Top             =   2355
      Width           =   3255
   End
   Begin VB.Label LblSupplierName 
      BackStyle       =   0  'Transparent
      Caption         =   "SupplierName��Ӧ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmGlueSupplierEdit.frx":1A20
      TabIndex        =   1
      Top             =   1635
      Width           =   3255
   End
   Begin VB.Label LblGlue12NC 
      BackStyle       =   0  'Transparent
      Caption         =   "Part 12NC ����12NC��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmGlueSupplierEdit.frx":1D2A
      TabIndex        =   0
      Top             =   915
      Width           =   3255
   End
End
Attribute VB_Name = "FrmGlueSupplierEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriGlue12NC As String

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'TxtApplicant.Text = PDMUserName
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
Private Function Check() As Boolean
If Trim(TxtGlue12NC) = "" Then
    MsgBox "Please input Glue/Electro 12NC" + vbCrLf + "�����뽺ˮ/��������12NC��", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
  End If
If Not (Len(Trim(TxtGlue12NC)) = 12 And IsNumeric(Trim(TxtGlue12NC))) Then
    MsgBox "Glue/Electro 12NC is 12 Number, No letter " + vbCrLf + "������12λ������,����ĸ", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
End If
If right(Trim(TxtGlue12NC), 1) = 0 Then
    MsgBox "Last number(Version Number) can NOT be 0 " + vbCrLf + "���Ϻ������һλ�汾�Ų���Ϊ0", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
End If
If Trim(TxtSupplierName) = "" Then
    MsgBox "Please input SupplierName" + vbCrLf + "�����빩Ӧ����", vbInformation, "System Info."
    TxtSupplierName.SetFocus
    Check = False
    Exit Function
   End If
If Trim(TxtSupplierPN) = "" Then
    MsgBox "Please input SupplierPN" + vbCrLf + "�����빩Ӧ���Ϻ�", vbInformation, "System Info."
    TxtSupplierPN.SetFocus
    Check = False
    Exit Function
   End If
   Check = True
End Function
Private Sub lblOk_Click()
    
   '�ж�Ҫ�༭��Ϣ�Ƿ�����
   If Check = False Then
    Exit Sub
   End If
   
   With MyGlueSupplier              '�Ѿ�����Public MyGlueSupplier As New ClsGlueSupplier, ��ģ�鸳����ֵ
    .Glue12NC = TxtGlue12NC.Text
    .SupplierName = TxtSupplierName.Text
    .SupplierPN = TxtSupplierPN.Text
   
            '�жϲ�������ӻ����޸�
       If Modify = False Then         '�ж�Ϊ��Ӳ���
     
           '�ж�Glue12NC�Ƿ��Ѿ�����
                If .In_DB(TxtGlue12NC.Text) = True Then
                   MsgBox "Glue 12NC exists, Please re-input" + vbCrLf + "��ˮ12NC���ظ�������������", vbInformation, "System Info."
                   TxtGlue12NC.SetFocus
                   TxtGlue12NC.SelStart = 0
                   TxtGlue12NC.SelLength = Len(TxtGlue12NC)
                   Exit Sub
                Else
                   .Insert                   '���
                    MsgBox "Succeed to Add" + vbCrLf + "��ӳɹ�", vbInformation, "System Info."
                End If
       Else  '�ж�Ϊ�޸Ĳ���
        .Update (OriGlue12NC)
         MsgBox "Succeed to Modify" + vbCrLf + "�޸ĳɹ�", vbInformation, "System Info."
       End If
       
    End With
    
    Unload Me
    
End Sub
