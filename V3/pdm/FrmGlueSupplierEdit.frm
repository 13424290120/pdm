VERSION 5.00
Begin VB.Form FrmGlueSupplierEdit 
   Caption         =   "Glue/Electro SupplierInfo. 胶水类供应商信息"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtSupplierPN 
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "注意:以下物料12NC需要版本号"
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
      Caption         =   "OK 确 定"
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
      Caption         =   "Cancel 取 消"
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
      Caption         =   "SupplierPN供应商料号"
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
      Caption         =   "SupplierName供应商名"
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
      Caption         =   "Part 12NC 物料12NC号"
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
    MsgBox "Please input Glue/Electro 12NC" + vbCrLf + "请输入胶水/电子物料12NC号", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
  End If
If Not (Len(Trim(TxtGlue12NC)) = 12 And IsNumeric(Trim(TxtGlue12NC))) Then
    MsgBox "Glue/Electro 12NC is 12 Number, No letter " + vbCrLf + "请输入12位的数字,无字母", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
End If
If right(Trim(TxtGlue12NC), 1) = 0 Then
    MsgBox "Last number(Version Number) can NOT be 0 " + vbCrLf + "物料号中最后一位版本号不能为0", vbInformation, "System Info."
    TxtGlue12NC.SetFocus
    Check = False
    Exit Function
End If
If Trim(TxtSupplierName) = "" Then
    MsgBox "Please input SupplierName" + vbCrLf + "请输入供应商名", vbInformation, "System Info."
    TxtSupplierName.SetFocus
    Check = False
    Exit Function
   End If
If Trim(TxtSupplierPN) = "" Then
    MsgBox "Please input SupplierPN" + vbCrLf + "请输入供应商料号", vbInformation, "System Info."
    TxtSupplierPN.SetFocus
    Check = False
    Exit Function
   End If
   Check = True
End Function
Private Sub lblOk_Click()
    
   '判断要编辑信息是否完整
   If Check = False Then
    Exit Sub
   End If
   
   With MyGlueSupplier              '已经定义Public MyGlueSupplier As New ClsGlueSupplier, 类模块赋变量值
    .Glue12NC = TxtGlue12NC.Text
    .SupplierName = TxtSupplierName.Text
    .SupplierPN = TxtSupplierPN.Text
   
            '判断操作是添加还是修改
       If Modify = False Then         '判断为添加操作
     
           '判断Glue12NC是否已经存在
                If .In_DB(TxtGlue12NC.Text) = True Then
                   MsgBox "Glue 12NC exists, Please re-input" + vbCrLf + "胶水12NC号重复，请重新设置", vbInformation, "System Info."
                   TxtGlue12NC.SetFocus
                   TxtGlue12NC.SelStart = 0
                   TxtGlue12NC.SelLength = Len(TxtGlue12NC)
                   Exit Sub
                Else
                   .Insert                   '添加
                    MsgBox "Succeed to Add" + vbCrLf + "添加成功", vbInformation, "System Info."
                End If
       Else  '判断为修改操作
        .Update (OriGlue12NC)
         MsgBox "Succeed to Modify" + vbCrLf + "修改成功", vbInformation, "System Info."
       End If
       
    End With
    
    Unload Me
    
End Sub
