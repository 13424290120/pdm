VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSglPrtEdit 
   Caption         =   "Single Part Number Edit.   Single Part����༭"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmSglPrtEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   12390
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdSysDistrb 
      Caption         =   "System Distribute"
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
      Left            =   8970
      TabIndex        =   63
      Top             =   90
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seek PjtNO. from PjtName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   8910
      TabIndex        =   60
      Top             =   5910
      Width           =   3345
      Begin VB.ComboBox ComboPjtName 
         Height          =   345
         Left            =   105
         TabIndex        =   62
         Text            =   "ComboPjtName"
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox ComboPJNOIndex 
         Height          =   345
         Left            =   105
         TabIndex        =   61
         Text            =   "ComboPJNOIndex"
         Top             =   285
         Width           =   3135
      End
   End
   Begin VB.ComboBox CombSglPrtVer 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   765
      Width           =   1410
   End
   Begin VB.ComboBox CombNewOldStatus 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   4125
      Width           =   1410
   End
   Begin VB.TextBox TxtNewOldStatus 
      Height          =   375
      Left            =   5955
      TabIndex        =   52
      Top             =   4125
      Width           =   1380
   End
   Begin VB.ComboBox CombPrtUnit 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   1485
      Width           =   1410
   End
   Begin VB.TextBox TxtPrtUnit 
      Height          =   375
      Left            =   5955
      TabIndex        =   47
      Top             =   1470
      Width           =   1380
   End
   Begin VB.TextBox TxtSglPrtVer 
      Height          =   375
      Left            =   5955
      TabIndex        =   45
      Top             =   750
      Width           =   1380
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   5955
      TabIndex        =   15
      Top             =   2745
      Width           =   2775
   End
   Begin VB.TextBox TxtProductLine 
      Height          =   375
      Left            =   5955
      TabIndex        =   14
      Top             =   7545
      Width           =   1380
   End
   Begin VB.TextBox TxtSglPrtIndex 
      Height          =   375
      Left            =   5955
      TabIndex        =   13
      Top             =   90
      Width           =   2775
   End
   Begin VB.TextBox TxtIDSO 
      Height          =   375
      Left            =   5955
      TabIndex        =   12
      Top             =   3435
      Width           =   1380
   End
   Begin VB.TextBox TxtOpnDate 
      Height          =   375
      Left            =   5955
      TabIndex        =   11
      Top             =   4815
      Width           =   1380
   End
   Begin VB.TextBox TxtClosDate 
      Height          =   375
      Left            =   5955
      TabIndex        =   10
      Top             =   5490
      Width           =   1380
   End
   Begin VB.ComboBox CombIDSO 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3435
      Width           =   1410
   End
   Begin VB.TextBox TxtApplicant 
      Height          =   375
      Left            =   5970
      TabIndex        =   8
      Top             =   2100
      Width           =   2775
   End
   Begin VB.TextBox TxtPJNOIndex 
      Height          =   375
      Left            =   5955
      TabIndex        =   7
      Top             =   6210
      Width           =   2775
   End
   Begin VB.TextBox TxtPjtName 
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   6855
      Width           =   2775
   End
   Begin VB.TextBox TxtItemType 
      Height          =   375
      Left            =   5955
      TabIndex        =   5
      Top             =   8220
      Width           =   1380
   End
   Begin VB.TextBox TxtLocation 
      Height          =   375
      Left            =   5955
      TabIndex        =   4
      Top             =   8880
      Width           =   1380
   End
   Begin VB.TextBox TxtCommtNote 
      Height          =   375
      Left            =   5955
      TabIndex        =   3
      Top             =   9525
      Width           =   2775
   End
   Begin VB.ComboBox CombProductLine 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7560
      Width           =   1410
   End
   Begin VB.ComboBox CombItemType 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8235
      Width           =   1410
   End
   Begin VB.ComboBox CombLocation 
      Height          =   345
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   8910
      Width           =   1410
   End
   Begin ERP.sqlSDBC sqlSDBC1 
      Left            =   8250
      Top             =   10125
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComCtl2.DTPicker DTPickerClosDate 
      Height          =   420
      Left            =   7350
      TabIndex        =   16
      Top             =   5475
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39979
   End
   Begin MSComCtl2.DTPicker DTPickerOpnDate 
      Height          =   420
      Left            =   7350
      TabIndex        =   17
      Top             =   4800
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   39989
   End
   Begin VB.Label LblNew8 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   59
      Top             =   8625
      Width           =   390
   End
   Begin VB.Label LblOld8 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   58
      Top             =   8625
      Width           =   285
   End
   Begin VB.Label LblNew7 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   56
      Top             =   7950
      Width           =   390
   End
   Begin VB.Label LblOld7 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   55
      Top             =   7950
      Width           =   285
   End
   Begin VB.Label LblNewOldStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status New/Old"
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
      Left            =   3630
      MouseIcon       =   "FrmSglPrtEdit.frx":08CA
      TabIndex        =   54
      Top             =   4140
      Width           =   2235
   End
   Begin VB.Label LblNew6 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   51
      Top             =   7290
      Width           =   390
   End
   Begin VB.Label LblOld6 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   50
      Top             =   7290
      Width           =   285
   End
   Begin VB.Label LblPrtUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Unit ��Ʒ��λ"
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
      Left            =   3225
      MouseIcon       =   "FrmSglPrtEdit.frx":0BD4
      TabIndex        =   48
      Top             =   1500
      Width           =   2625
   End
   Begin VB.Label LblSglPrtVer 
      BackStyle       =   0  'Transparent
      Caption         =   "Version Number �汾���"
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
      Left            =   2220
      MouseIcon       =   "FrmSglPrtEdit.frx":0EDE
      TabIndex        =   46
      Top             =   780
      Width           =   3630
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
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
      Height          =   315
      Left            =   4185
      MouseIcon       =   "FrmSglPrtEdit.frx":11E8
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   10215
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3585
      Picture         =   "FrmSglPrtEdit.frx":14F2
      Top             =   10185
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
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
      Height          =   315
      Left            =   6465
      MouseIcon       =   "FrmSglPrtEdit.frx":190E
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   10215
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5865
      Picture         =   "FrmSglPrtEdit.frx":1C18
      Top             =   10185
      Width           =   300
   End
   Begin VB.Label LblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description ��Ʒ����"
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
      Left            =   2865
      MouseIcon       =   "FrmSglPrtEdit.frx":2034
      TabIndex        =   42
      Top             =   2745
      Width           =   2985
   End
   Begin VB.Label LblProductLine 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Line ��Ʒ�߱��"
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
      Left            =   2355
      MouseIcon       =   "FrmSglPrtEdit.frx":233E
      TabIndex        =   41
      Top             =   7545
      Width           =   3495
   End
   Begin VB.Label LblSglPrtIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Single Part NO.    Single Part ���"
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
      Left            =   1035
      MouseIcon       =   "FrmSglPrtEdit.frx":2648
      TabIndex        =   40
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label LblIDSO 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Open/Close"
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
      Left            =   3180
      MouseIcon       =   "FrmSglPrtEdit.frx":2952
      TabIndex        =   39
      Top             =   3450
      Width           =   2685
   End
   Begin VB.Label LblOpnDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date ��ʼ����"
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
      Left            =   3015
      MouseIcon       =   "FrmSglPrtEdit.frx":2C5C
      TabIndex        =   38
      Top             =   4830
      Width           =   2865
   End
   Begin VB.Label LblClosDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Close Date ��������"
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
      Left            =   2985
      MouseIcon       =   "FrmSglPrtEdit.frx":2F66
      TabIndex        =   37
      Top             =   5505
      Width           =   2895
   End
   Begin VB.Label LblOld0 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   36
      Top             =   465
      Width           =   285
   End
   Begin VB.Label LblNew0 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   35
      Top             =   465
      Width           =   390
   End
   Begin VB.Label LblOld1 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   34
      Top             =   1185
      Width           =   285
   End
   Begin VB.Label LblNew1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   33
      Top             =   1185
      Width           =   390
   End
   Begin VB.Label LblOld2 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   32
      Top             =   3165
      Width           =   285
   End
   Begin VB.Label LblNew2 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   31
      Top             =   3165
      Width           =   390
   End
   Begin VB.Label LblApplicant 
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant ������"
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
      Left            =   3510
      MouseIcon       =   "FrmSglPrtEdit.frx":3270
      TabIndex        =   30
      Top             =   2130
      Width           =   2355
   End
   Begin VB.Label LblReminder 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FrmSglPrtEdit.frx":357A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      Left            =   9330
      TabIndex        =   29
      Top             =   2805
      Width           =   2715
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   3390
      Shape           =   4  'Rounded Rectangle
      Top             =   10050
      Width           =   4740
   End
   Begin VB.Label LblPJNOIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Number ������Ŀ���"
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
      Left            =   1650
      MouseIcon       =   "FrmSglPrtEdit.frx":35BF
      TabIndex        =   28
      Top             =   6180
      Width           =   4200
   End
   Begin VB.Label LblPjtName 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name ��Ŀ��������"
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
      Left            =   1920
      MouseIcon       =   "FrmSglPrtEdit.frx":38C9
      TabIndex        =   27
      Top             =   6870
      Width           =   3915
   End
   Begin VB.Label LblItemType 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type Number. ��Ʒ�����"
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
      Left            =   1200
      MouseIcon       =   "FrmSglPrtEdit.frx":3BD3
      TabIndex        =   26
      Top             =   8220
      Width           =   4680
   End
   Begin VB.Label LblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Number. ��Ʒ���ͱ��"
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
      Left            =   1365
      MouseIcon       =   "FrmSglPrtEdit.frx":3EDD
      TabIndex        =   25
      Top             =   8880
      Width           =   4500
   End
   Begin VB.Label LblCommtNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment  Note. ע�ͺͱ�ע"
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
      Left            =   1905
      MouseIcon       =   "FrmSglPrtEdit.frx":41E7
      TabIndex        =   24
      Top             =   9525
      Width           =   3960
   End
   Begin VB.Label LblNew3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   23
      Top             =   3840
      Width           =   390
   End
   Begin VB.Label LblOld3 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   22
      Top             =   3840
      Width           =   285
   End
   Begin VB.Label LblOld4 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   21
      Top             =   4530
      Width           =   285
   End
   Begin VB.Label LblNew4 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   20
      Top             =   4530
      Width           =   390
   End
   Begin VB.Label LblOld5 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   19
      Top             =   5220
      Width           =   285
   End
   Begin VB.Label LblNew5 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7845
      TabIndex        =   18
      Top             =   5220
      Width           =   390
   End
End
Attribute VB_Name = "FrmSglPrtEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean
Public OriSglPrtIndex As String                       '############�����ĳɶ�Ӧ�ı��ֶ�����

Private Sub CmdSysDistrb_Click()
FrmSglPrtNOSection.ModifyFm = Modify          '�ѵ�ǰ���ڵ�״̬�̳и�����һ������
FrmSglPrtNOSection.Show 1
End Sub



Private Sub ComboPJNOIndex_Click()
TxtPJNOIndex.Text = ComboPJNOIndex.Text
TxtPjtName.Text = ComboPjtName.List(ComboPJNOIndex.ListIndex)
End Sub

Private Sub ComboPjtName_Click()
TxtPjtName.Text = ComboPjtName.Text
TxtPJNOIndex.Text = ComboPJNOIndex.List(ComboPjtName.ListIndex)
End Sub
Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Resize_ALL Me
End Sub

Private Sub TxtPJNOIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNOΪҪ��ѯ�ı���
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord�Ѿ�ȡ���ӵ�һ��ʼ�ң���������Ҫ���õ���ʼ

    Do While sqlUsrCtrl.IfBOForEOF = False
        sqlUsrCtrl.FindRecord "PJNOIndex", UseEquel, Trim(TxtPJNOIndex.Text)  '����1UseEquel����= 2UseLike�Ǵ���Like

       ComboPJNOIndex.AddItem (FormatNumber6(CStr(UsrCtlFind(0))))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
    
    Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub TxtPjtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ComboPjtName.Clear
    ComboPJNOIndex.Clear
    Dim sqlUsrCtrl As Control
    Set sqlUsrCtrl = sqlSDBC1

    sqlUsrCtrl.OpenConnection DBUser, Password, Server, DataBase
    sqlUsrCtrl.OpenRecordset ("PJNO")    'PJNOΪҪ��ѯ�ı���
    sqlUsrCtrl.MoveRecord (MoveFirst) 'sqlUsrCtrl.FindRecord�Ѿ�ȡ���ӵ�һ��ʼ�ң���������Ҫ���õ���ʼ

     Do While sqlUsrCtrl.IfBOForEOF = False
       sqlUsrCtrl.FindRecord "Description", UseLike, Trim(TxtPjtName.Text)  '����1UseEquel����= 2UseLike�Ǵ���Like

       ComboPJNOIndex.AddItem (UsrCtlFind(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
       ComboPjtName.AddItem (UsrCtlFind(3))  'UsrCtlFind�����е�3()�Ƕ�ӦDescription���ֶ����
       Erase UsrCtlFind
       sqlUsrCtrl.MoveRecord (MoveNext)
  
     Loop
    sqlUsrCtrl.CloseRS
End If
End Sub

Private Sub Form_Load()               '############������ظĳɶ�Ӧ�Ŀؼ�,����ֶ�����
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me
TxtApplicant.Text = PDMUserName
CombSglPrtVer.AddItem ("0")
CombSglPrtVer.AddItem ("1")
CombSglPrtVer.AddItem ("2")
CombSglPrtVer.AddItem ("3")
CombSglPrtVer.AddItem ("4")
CombSglPrtVer.AddItem ("5")
CombSglPrtVer.AddItem ("6")
CombSglPrtVer.AddItem ("7")
CombSglPrtVer.AddItem ("8")
CombSglPrtVer.AddItem ("9")
CombSglPrtVer.ListIndex = 1

CombPrtUnit.AddItem ("Piece")
CombPrtUnit.AddItem ("Gram")
CombPrtUnit.AddItem ("Meter")
CombPrtUnit.ListIndex = 0

TxtApplicant.Text = PDMUserName

CombIDSO.AddItem ("Open")
CombIDSO.AddItem ("Close")
CombIDSO.ListIndex = 0

CombNewOldStatus.AddItem ("New")
CombNewOldStatus.AddItem ("Old")
CombNewOldStatus.ListIndex = 0


CombProductLine.AddItem ("5000")
CombProductLine.ListIndex = 0

CombItemType.AddItem ("100")
CombItemType.AddItem ("110")
CombItemType.AddItem ("060")
CombItemType.AddItem ("030")
CombItemType.AddItem ("020")
CombItemType.AddItem ("080")
CombItemType.AddItem ("090")
CombItemType.AddItem ("200")
CombItemType.AddItem ("050")
CombItemType.AddItem ("010")
CombItemType.AddItem ("070")
CombItemType.AddItem ("300")
CombItemType.AddItem ("040")
CombItemType.AddItem ("???")
CombItemType.ListIndex = 0

CombLocation.AddItem ("TR-AV")
CombLocation.ListIndex = 0

DTPickerOpnDate.Value = Date
DTPickerClosDate.Value = Date
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub
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

Private Function Check() As Boolean                        '############������ظĳɶ�Ӧ�Ŀؼ�,����ֶ�����
If Trim(TxtSglPrtIndex) = "" Then
    MsgBox "Please input Single part Number" + vbCrLf + "������Single part��", vbInformation, "System Info."
    TxtSglPrtIndex.SetFocus
    Check = False
    Exit Function
  End If
If Not (Len(TxtSglPrtIndex) = 12 And Isnum(TxtSglPrtIndex)) Then  '����Left() Right()�Ǵ���ߺ��ұ߽�ȡ�ַ���
    MsgBox "Single part Series Number is 12 Number, no Letter" + vbCrLf + "Single part��12λ���ֵı��,����ĸ", vbInformation, "System Info."
    TxtSglPrtIndex.SetFocus
    Check = False
    Exit Function
  End If
If Not (right(TxtSglPrtIndex, 1) = "0") And Not (left(TxtSglPrtIndex, 1) = "8") Then '����Left() Right()�Ǵ���ߺ��ұ߽�ȡ�ַ���
    MsgBox "Single part last number must be 0, Not set version number here" + vbCrLf + "Single part���ĩλ����Ϊ0,��Ҫ����������汾��", vbInformation, "System Info."
    TxtSglPrtIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtApplicant) = "" Then
    MsgBox "Please input Applicant Name" + vbCrLf + "��������������", vbInformation, "System Info."
    TxtApplicant.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtDescription) = "" Then
    MsgBox "Please input Description" + vbCrLf + "��������Ʒ����", vbInformation, "System Info."
    TxtDescription.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPJNOIndex) = "" Or (Not Isnum(TxtPJNOIndex)) Or (Len(TxtPJNOIndex) <> 6) Then
    MsgBox "Please input Project Number, 6 number" + vbCrLf + "�������漰��Ŀ���, 6λ������", vbInformation, "System Info."
    TxtPJNOIndex.SetFocus
    Check = False
    Exit Function
  End If
If Trim(TxtPjtName) = "" Then
    MsgBox "Please input Project Name" + vbCrLf + "�������漰��Ŀ����", vbInformation, "System Info."
    TxtPjtName.SetFocus
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
     
   With MySglPrt    '�Ѿ�����Public MySglPrt As New ClsSglPrt, ��ģ�鸳����ֵ  ############������ظĳɶ�Ӧ�Ŀؼ�����,�������,�ֶ�����
    .SglPrtIndex = TxtSglPrtIndex.Text
    .SglPrtVer = CombSglPrtVer.Text
    .PrtUnit = CombPrtUnit.Text
    .Applicant = TxtApplicant.Text
    .Description = TxtDescription.Text
    .IDSO = CombIDSO.Text
    .NewOldStatus = CombNewOldStatus.Text
    .OpnDate = DTPickerOpnDate.Value
    .ClosDate = DTPickerClosDate.Value
    .PJNOIndex = TxtPJNOIndex.Text
    .PjtName = TxtPjtName.Text
    .ProductLine = CombProductLine.Text
    .ItemType = CombItemType.Text
    .Location = CombLocation.Text
    .CommtNote = TxtCommtNote.Text
    
    '�������Ϻţ�ItemType����
    If FrmSglPrtEdit.Modify = False Then
        If .ItemType = "???" Then
            MsgBox "Item Type MUST be chosed." & vbCrLf & "��Ʒ������Ǳ�ѡ��."
            Exit Sub
        End If
    End If
   
            '�жϲ�������ӻ����޸�
       If Modify = False Then         '�ж�Ϊ��Ӳ���
     
           '�ж�SglPrtIndex����Ƿ��Ѿ�����
                If .In_DB(TxtSglPrtIndex.Text) = True Then
                   MsgBox "Single Part number exists, Please re-input" + vbCrLf + "Single Part���ظ�������������", vbInformation, "System Info."
                   TxtSglPrtIndex.SetFocus
                   TxtSglPrtIndex.SelStart = 0
                   TxtSglPrtIndex.SelLength = Len(TxtSglPrtIndex)
                   Exit Sub
                Else
                   .Insert                   '���
                    MsgBox "Succeed to Add" + vbCrLf + "��ӳɹ�", vbInformation, "System Info."
                End If
       Else  '�ж�Ϊ�޸Ĳ���
        .Update (OriSglPrtIndex)
        MsgBox "Succeed to Modify" + vbCrLf + "�޸ĳɹ�", vbInformation, "System Info."
       End If
    End With
    Unload Me    '�ر�������
End Sub




