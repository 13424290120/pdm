VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBOMAdmin 
   Caption         =   "PDM-BOM Admin"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBOMAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10950
   ScaleWidth      =   13590
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frameMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   3300
      TabIndex        =   58
      Top             =   4260
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1455
         Left            =   60
         TabIndex        =   59
         Top             =   60
         Width           =   7935
      End
   End
   Begin VB.Frame FrmPaste 
      Caption         =   "Paste Code"
      Height          =   1455
      Left            =   4920
      TabIndex        =   53
      Top             =   3330
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton OKButton 
         Caption         =   "Yes"
         Height          =   375
         Left            =   2280
         TabIndex        =   56
         Top             =   870
         Width           =   885
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "No"
         Height          =   375
         Left            =   3150
         TabIndex        =   55
         Top             =   870
         Width           =   885
      End
      Begin VB.TextBox txtNewCode 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   54
         Top             =   870
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Would you like to copy all of its childs to paste under the following item?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   57
         Top             =   270
         Width           =   4035
      End
   End
   Begin VB.TextBox cmbBOMVersion 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1710
      Width           =   345
   End
   Begin VB.Frame FrmUpgrade 
      Caption         =   "Upgrade Version for Single Part"
      Height          =   1635
      Left            =   5070
      TabIndex        =   41
      Top             =   4440
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtSglParent 
         Height          =   405
         Left            =   150
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2760
         TabIndex        =   49
         Top             =   900
         Width           =   1000
      End
      Begin VB.CommandButton cmdUpgrade 
         Caption         =   "Save"
         Height          =   400
         Left            =   2760
         TabIndex        =   48
         Top             =   420
         Width           =   1000
      End
      Begin VB.ComboBox cmbSglVer2 
         Height          =   345
         ItemData        =   "FrmBOMAdmin.frx":08CA
         Left            =   1890
         List            =   "FrmBOMAdmin.frx":08E9
         TabIndex        =   46
         Top             =   960
         Width           =   675
      End
      Begin VB.ComboBox cmbSglVer1 
         BackColor       =   &H8000000F&
         Height          =   345
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox txt12NC 
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Version:"
         Height          =   405
         Left            =   120
         TabIndex        =   47
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "=>"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   10.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   45
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "12NC:"
         Height          =   405
         Left            =   150
         TabIndex        =   42
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdReview 
      Caption         =   "Review BOM Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1890
      TabIndex        =   39
      Top             =   10440
      Width           =   1755
   End
   Begin VB.CommandButton cmdBOMSave 
      Caption         =   "Save BOM Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   60
      TabIndex        =   38
      Top             =   10440
      Width           =   1785
   End
   Begin VB.ComboBox txtSERlocate 
      Height          =   345
      Left            =   9180
      TabIndex        =   37
      Top             =   2100
      Width           =   3075
   End
   Begin VB.ComboBox txtNodeDrwlocate 
      Height          =   345
      Left            =   8340
      TabIndex        =   36
      Top             =   1200
      Width           =   3915
   End
   Begin VB.ComboBox txtSubCon 
      Height          =   345
      Left            =   6960
      TabIndex        =   35
      Text            =   "SUBCON"
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11100
      TabIndex        =   34
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3750
      TabIndex        =   2
      Top             =   10440
      Width           =   1425
   End
   Begin VB.CommandButton CmdApprove 
      Caption         =   "Submit Approve"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9540
      TabIndex        =   32
      Top             =   10440
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Standard Part Lib"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5220
      TabIndex        =   31
      Top             =   10440
      Width           =   1635
   End
   Begin VB.TextBox txtSERNO 
      Height          =   360
      Left            =   9990
      TabIndex        =   30
      Top             =   1710
      Width           =   2265
   End
   Begin VB.TextBox txtCPCNNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   29
      Top             =   1710
      Width           =   1440
   End
   Begin VB.TextBox txtCPCNlocate 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   28
      Top             =   2100
      Width           =   2940
   End
   Begin VB.CommandButton CmdSERView 
      Caption         =   "See SER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12315
      TabIndex        =   25
      Top             =   2160
      Width           =   1125
   End
   Begin VB.CommandButton CmdSERPathAdd 
      Caption         =   "Add Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12330
      TabIndex        =   24
      Top             =   1800
      Width           =   1125
   End
   Begin VB.CommandButton CmdCPCNView 
      Caption         =   "See CP/CN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   23
      Top             =   2160
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3765
      Top             =   10635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExportBOM 
      Caption         =   "Export BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8280
      TabIndex        =   19
      Top             =   10440
      Width           =   1185
   End
   Begin VB.CommandButton CmdImportBOM 
      Caption         =   "Import BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6960
      TabIndex        =   18
      Top             =   10440
      Width           =   1260
   End
   Begin VB.CommandButton CmdSearchSglPrt 
      Caption         =   "Search SglPrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12015
      TabIndex        =   17
      Top             =   465
      Width           =   1425
   End
   Begin VB.CommandButton CmdSearchFinsGd 
      Caption         =   "Search FinsGd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10470
      TabIndex        =   16
      Top             =   465
      Width           =   1500
   End
   Begin VB.CommandButton CmdBuildFirstBOM 
      Caption         =   "Build / Initialize BOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8460
      MaskColor       =   &H80000010&
      TabIndex        =   15
      Top             =   450
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      MaxLength       =   12
      TabIndex        =   13
      Text            =   "Single Part 12NC"
      Top             =   450
      Width           =   1515
   End
   Begin VB.CommandButton CmdDrwPathAdd 
      Caption         =   "Add Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12330
      TabIndex        =   12
      Top             =   945
      Width           =   1125
   End
   Begin VB.Timer tmrDragTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   10350
   End
   Begin VB.CommandButton CmdDrwView 
      Caption         =   "See Drw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12330
      TabIndex        =   7
      Top             =   1260
      Width           =   1125
   End
   Begin VB.TextBox txtNodePrtUnit 
      Height          =   360
      Left            =   7590
      TabIndex        =   6
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtNodeSglPrt12NC 
      Height          =   360
      Left            =   3780
      TabIndex        =   5
      Top             =   1200
      Width           =   1410
   End
   Begin VB.TextBox txtNodeDescription 
      Height          =   360
      Left            =   5250
      TabIndex        =   4
      Top             =   1200
      Width           =   2310
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12390
      TabIndex        =   1
      Top             =   10440
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3780
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "Finish Goods NO"
      Top             =   450
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13230
      Top             =   10305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":0908
            Key             =   "NEW"
            Object.Tag             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":0A1A
            Key             =   "LOCKED"
            Object.Tag             =   "LOCKED"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":0E6C
            Key             =   "FILE"
            Object.Tag             =   "FILE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":311E
            Key             =   "CHILD"
            Object.Tag             =   "CHILD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":3438
            Key             =   "FOLDER"
            Object.Tag             =   "FOLDER"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":3752
            Key             =   "DELETE"
            Object.Tag             =   "DELETE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":3864
            Key             =   "OPENFOLDER"
            Object.Tag             =   "OPENFOLDER"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":3B7E
            Key             =   "SETTINGS"
            Object.Tag             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":3E98
            Key             =   "PREVIOUS"
            Object.Tag             =   "PREVIOUS"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":41EA
            Key             =   "NEXT"
            Object.Tag             =   "NEXT"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":498E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":4DE0
            Key             =   "BAS"
            Object.Tag             =   "BAS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5132
            Key             =   "CLS"
            Object.Tag             =   "CLS"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":53F4
            Key             =   "VB"
            Object.Tag             =   "VB"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5746
            Key             =   "VIEWBOOKMARKS"
            Object.Tag             =   "VIEWBOOKMARKS"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5A60
            Key             =   "ADDBOOKMARK"
            Object.Tag             =   "ADDBOOKMARK"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5D7A
            Key             =   "OPEN"
            Object.Tag             =   "OPEN"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5E8C
            Key             =   "PRINT"
            Object.Tag             =   "PRINT"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBOMAdmin.frx":5F9E
            Key             =   "FIND"
            Object.Tag             =   "FIND"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvCodeItems 
      Height          =   10275
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   18124
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox MSFlexGrid1EditText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6030
      TabIndex        =   20
      Text            =   "MsFleGrdTxt"
      Top             =   11160
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7545
      Left            =   3810
      TabIndex        =   40
      Top             =   2850
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   13309
      _Version        =   393216
      Rows            =   33
      Cols            =   12
      AllowUserResizing=   1
   End
   Begin VB.Label LblDRWPathSeek 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11370
      TabIndex        =   61
      Top             =   990
      Width           =   855
   End
   Begin VB.Label cmdLock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNLOCK"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7440
      MouseIcon       =   "FrmBOMAdmin.frx":60B0
      TabIndex        =   60
      Top             =   1770
      Width           =   1005
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000080FF&
      Height          =   705
      Left            =   3720
      Top             =   900
      Width           =   9810
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Version"
      Height          =   315
      Left            =   5910
      TabIndex        =   52
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label Lblwarning 
      BackColor       =   &H0000FFFF&
      Caption         =   "Warning: Only Green hightlight on first row means approved BOM ����: ��1����ɫͻ����ʾ��BOM����׼�Ĳ�����ʽʹ��"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   2580
      Width           =   9780
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      Height          =   915
      Left            =   8640
      Top             =   1650
      Width           =   4905
   End
   Begin VB.Label LblSERPathSeek 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8670
      TabIndex        =   27
      Top             =   2175
      Width           =   480
   End
   Begin VB.Label LblCPCNPathSeek 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   26
      Top             =   2145
      Width           =   480
   End
   Begin VB.Label LblSER 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SER Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8670
      TabIndex        =   22
      Top             =   1785
      Width           =   1260
   End
   Begin VB.Label LblCPCN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CP/CN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3810
      TabIndex        =   21
      Top             =   1770
      Width           =   585
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      Caption         =   "A BOM need one single part as a child at least, Please Input following content and Click Build Button to establish"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3810
      TabIndex        =   14
      Top             =   180
      Width           =   9615
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   3720
      Top             =   120
      Width           =   9795
   End
   Begin VB.Label LblDrw 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Drawing location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8370
      TabIndex        =   11
      Top             =   975
      Width           =   2940
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   10
      Top             =   975
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5250
      TabIndex        =   9
      Top             =   975
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "  Selected 12NC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3795
      TabIndex        =   8
      Top             =   975
      Width           =   1440
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   900
      Left            =   3735
      Top             =   1650
      Width           =   4860
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCode 
         Caption         =   "&Add New Code Here"
      End
      Begin VB.Menu mnuDeleteCode 
         Caption         =   "&Delete Selected Code"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename Code Here"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpgradeVer 
         Caption         =   "&Upgrade Version"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Code"
      End
      Begin VB.Menu mnuUncopy 
         Caption         =   "&Undo Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste Code"
      End
   End
End
Attribute VB_Name = "FrmBOMAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecursionFlag As Boolean
Private RowNum As Integer
Private FinishGoodsNO As String
Private SourceNodeParent As String
Private Const TreeRootTag As String = "Root"  '���ڵ���
Private SourceNode As Object  '����ڵ���ҷ��Դ�ڵ�
Private targetNode As Object  '����ڵ���ҷ��Ŀ��ڵ�
Private AddNodeOk As Boolean   '����ӽڵ�ɹ���ı��
Private ApprovalStatus As Boolean   '����BOM�Ƿ���׼�ı��
Private OpennerSubmiter As Boolean   '����BOM�����Ƿ�����(�ύ��)�ı��
Private NotDeleteChildTree As Boolean   '����BOM�Ƿ�ɾ�����е�����
Private miScrollDir As Integer  '����TreeView�����ķ���
Private miClipBoardFormat As Integer
Private scr As Object         '�����ר�õı��ʽ�����Ļ����������ò���
Private sNodeText As String '�������ٽڵ㽹��
Private Conn As New ADODB.Connection
Private StrSql As String
Private CurVersion, LastVersion As Integer
Private CPCN As String
Private ChgCPCN, ChgMass As Boolean
Private oldCode As String
Private J As Integer
Public CopyNodeSource, PasteNodeTarget As Node
Public IsCopy, bNotSave1stVer As Boolean
Public sChilds As String
Public CurNode As Node
Private Family() As String
Public bCopyRoot As Boolean
Private isApproved, isRejected As Boolean
Private OrientCurNodeKey, OrientParentNodeKey As String
Private Action As String
Private BOMString As String
Private temp_tb_SglPrt4BOMLog, temp_tb_BOMOrigData As String
Private arrBOM() As String
Public BOMLock As Boolean, BOMLocker As String

Private Sub CancelButton_Click()
    mnuCopy = True
    mnuPaste = False
    mnuUncopy = False
    IsCopy = False
    CopyNodeSource = ""
    Unload Me
End Sub


Private Sub cmdBOMSave_Click()
    On Error Resume Next
    
    Me.Enabled = False
    frameMsg.Visible = True
    DoEvents
    
    Dim i, J
    Dim rs As New ADODB.Recordset
    
    i = 0: J = 0
    Conn.BeginTrans
    StrSql = "Select IsSave From BOMCPCN Where BOMID=" & FinishGoodsNO & " And BOMVersion=" & CurVersion & " And CPCNNmbr='" & CPCN & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockPessimistic
    If CurVersion = 1 Then
        If rs.RecordCount > 0 Then
            If CBool(rs(0)) Then
                If MsgBox("The 1st version had been saved, do you want to overwrite it?", vbYesNo) = vbYes Then
                    'Conn.Execute ("Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CurVersion)
                    Conn.Execute ("Update BOMCPCN Set UpdateDate=Getdate() Where BOMID=" & FinishGoodsNO & " And CPCNNmbr='" & CPCN & "' And BOMVersion=" & CurVersion)
                Else
                    Exit Sub
                End If
            End If
        Else
            'Add new BOMCPCN
            Conn.Execute ("Insert into BOMCPCN(BOMID,CPCNNmbr,CPCNLocate,BOMVersion,isSave) Values(" & FinishGoodsNO & ",'" & CPCN & "','" & Trim(txtCPCNlocate.Text) & "'," & CurVersion & ",1)")
        End If
    Else
        If rs.RecordCount = 0 Then
            'Add new BOMCPCN
            Conn.Execute ("Insert into BOMCPCN(BOMID,CPCNNmbr,CPCNLocate,BOMVersion,isSave) Values(" & FinishGoodsNO & ",'" & CPCN & "','" & Trim(txtCPCNlocate.Text) & "'," & CurVersion & ",1)")
        Else
            'Update BOMCPCN
            If Trim(txtCPCNlocate.Text) = "" Then
                Conn.Execute ("Update BOMCPCN Set isSave=1,UpdateDate=Getdate() Where BOMID=" & FinishGoodsNO & " And CPCNNmbr='" & CPCN & "' And BOMVersion=" & CurVersion)
            Else
                Conn.Execute ("Update BOMCPCN Set isSave=1,CPCNLocate='" & Trim(txtCPCNlocate.Text) & "',UpdateDate=Getdate() Where BOMID=" & FinishGoodsNO & " And CPCNNmbr='" & CPCN & "' And BOMVersion=" & CurVersion)
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing

    
    If Err Then
        Conn.RollbackTrans
        MsgBox "Save BOM Version Failed, Something Error, Please contact system admin.", vbInformation, "System Info"
    Else
'        Call UpdateBOMVerQtyDesc(FinishGoodsNO, CStr(CurVersion), MSFlexGrid1)
        Me.Enabled = True
        frameMsg.Visible = False
        '##########ѯ���û�Ҫ��Ҫ����BOM#######
        If BOMLock Then
            If MsgBox("Would you like to unlock the BOM?", vbYesNo, "System Info") = vbYes Then
                StrSql = "UPDATE BOMCPCN SET IsLocked=0 WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "'"
                Conn.Execute StrSql
            End If
        End If
        
        Me.Enabled = False
        frameMsg.Visible = True
        DoEvents
        
        
        Conn.CommitTrans
        If CurVersion <> 1 Then
            Call SaveBOMData
        Else
            StrSql = "DELETE FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=1"
            Conn.Execute StrSql
            With MSFlexGrid1
                For i = 2 To .Rows - 2
                    If Trim(.TextMatrix(i, 2)) <> "" Then
                        '�������µ��޸���־
                        StrSql = "Insert into  SglPrt4BOMLog  (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,CommtNote,Family) Values("
                        StrSql = StrSql & "" & FinishGoodsNO
                        StrSql = StrSql & "," & i + J
                        StrSql = StrSql & "," & .TextMatrix(i, 2)
                        StrSql = StrSql & "," & .TextMatrix(i, 3)
                        StrSql = StrSql & ",1"
                        StrSql = StrSql & ",'" & .TextMatrix(i, 4)
                        StrSql = StrSql & "','" & .TextMatrix(i, 5)
                        StrSql = StrSql & "','" & Replace(.TextMatrix(i, 6), "'", "''")
                        StrSql = StrSql & "','" & .TextMatrix(i, 7)
                        StrSql = StrSql & "','" & .TextMatrix(i, 8)
                        StrSql = StrSql & "','" & .TextMatrix(i, 10)
                        StrSql = StrSql & "','" & .TextMatrix(i, 11) & "')"
                        Conn.Execute StrSql
                    End If
                Next i
            End With
        End If
        bNotSave1stVer = True

        ChgMass = False
        ChgCPCN = False
        LastVersion = CurVersion
        cmbBOMVersion.Text = CurVersion
    End If
    Me.Enabled = True
    frameMsg.Visible = False
End Sub

Private Sub CmdBuildFirstBOM_Click()
    On Error Resume Next
    Dim TempDescription, SglPrtVersion As String
    Dim NeedCreateAssPart As Boolean
    Dim rs As New ADODB.Recordset
        
    Conn.BeginTrans
    NeedCreateAssPart = False
    If Not (Len(Trim(Text1.Text)) = 12 And Isnum(Trim(Text1.Text))) Then
        MsgBox "Finish Goods is 12 Number, no Letter." + vbCrLf + "������12λ���ֵı��,����ĸ.", vbInformation, "System Info."
        Text1.SetFocus
        Exit Sub
    '������ĺ���ֻ��Ҫ����������Σ�2441XXXX,9041XXX,4341 078 6XXXX
    ElseIf left(Text1.Text, 4) <> "2441" And left(Text1.Text, 4) <> "9041" And left(Text1.Text, 8) <> "43410786" Then
        MsgBox "You must input a new valid 12NC for the Finish Goods.", vbInformation, "System Info."
        Text1.SetFocus
        Exit Sub
    ElseIf txtSubCon.Text = "" Or txtSubCon.Text = "SUBCON" Then
        MsgBox "Please input the SUBCON in the text box."
        txtSubCon.SetFocus
        Exit Sub
    Else
        '��ʼ�ж������Finish Good NO �Ƿ������ݿ�������
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(Text1.Text) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            '�ж��Ƿ�����װ��
            If rs.State = adStateOpen Then rs.Close
            StrSql = "Select * from AssemblyPrtList where PrefixNo = Left('" & Trim(Text1.Text) & "',Len(PrefixNo)) And Enable=1"
            rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Finish goods NO. is not existing in database, or the Assembly Part Number is disable.", vbInformation, "System Info."
                Exit Sub
            Else
                '����һ����װ�Ϻ�
                NeedCreateAssPart = True
                TempDescription = "Assembly Part"
            End If
        Else
            TempDescription = rs("Description")
        End If
        If rs.State = adStateOpen Then rs.Close
    End If
    
    If Len(Trim(Text2.Text)) = 0 Then        '�ж�Text2(����SinglePart NO)���ݵĺϷ���
        MsgBox "You must enter a new 12NC for the Single Part", vbInformation, "System Info."
        Text2.SetFocus
        Exit Sub
    ElseIf Not (Len(Trim(Text2.Text)) = 12 And Isnum(Trim(Text2.Text))) Then
        MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Text2.SetFocus
        Exit Sub
    Else
        '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
        rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(Text2.Text), 1, Len(Trim(Text2.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The Single part NO. is not existing in database", vbInformation, "System Info."
            Text2.SetFocus
            If rs.State = adStateOpen Then rs.Close
            Exit Sub
        Else
            SglPrtVersion = rs.Fields("SglPrtVer")
        End If
        If rs.State = adStateOpen Then rs.Close
    End If

    
    'д��ϵͳ����
    StrSql = "IF NOT EXISTS(Select * from SysVar where itemtype='SUBCON' and itemvalue='" & Trim(txtSubCon.Text) & "' and creator='" & PDMUserName & "') Insert into SysVar values ('SUBCON','" & Trim(txtSubCon.Text) & "','" & PDMUserName & "')"
    Conn.Execute StrSql
    'Ass Part����FG
    If NeedCreateAssPart Then
        StrSql = "SELECT * FROM SglPrt WHERE LEFT(SglPrtIndex,11)='" & left(Trim(Text1.Text), 11) & "'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The FG Number is invalid, please create new Single Part at first."
            Text1.SetFocus
            NeedCreateAssPart = False
            rs.Close
            Set rs = Nothing
            Exit Sub
        Else
            Conn.Execute "Insert into FinsGd (FinsGdIndex,Applicant,ProductLine,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,ItemType,Location,CommtNote,IsAssemblyPart) SELECT " & Trim(Text1.Text) & ",'" & PDMUserName & "',ProductLine,Description,IDSO,getdate(),getdate(),PJNOIndex,PjtName,ItemType,Location,CommtNote,1 FROM SglPrt WHERE LEFT(SglPrtIndex,11)='" & left(Trim(Text1.Text), 11) & "'"
        End If
        NeedCreateAssPart = False
        rs.Close
        Set rs = Nothing
    End If
    
    '�ж�һ�α���¼�Ƿ����
    If rs.State = adStateOpen Then rs.Close
    StrSql = "Select * from BOMOrigData Where ParentID ='" & Trim(Text1.Text) & "' AND LEFT(ChildID,11) = '" & left(Trim(Text2.Text), 11) & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        MsgBox "The Parent/Child Record already exist in BOM Database", vbInformation, "System Info."
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Exit Sub
    Else
        rs.Close
    End If
    '��¼�����ڲſ�ʼд��

    Conn.Execute "INSERT INTO BOMOrigData (ParentID, ChildID, Quantity) VALUES ('" & Trim(Text1.Text) & "','" & left(Trim(Text2.Text), 11) & SglPrtVersion & "',1)"
    Conn.Execute "INSERT INTO SUBCON (FinsGDIndex,SUBCON) Values (" & Trim(Text1.Text) & ",'" & Replace(txtSubCon.Text, "'", "''") & "')"
    
    
    '�Ǽ�BOM������(�ύ��)
    '���жϼ�¼�Ƿ����
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(Text1.Text) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    Else
        Conn.Execute "INSERT INTO BOMSubmitApprove (FinsGdIndex,Description,Submiter,SubmitDate) VALUES ('" & Trim(Text1.Text) & "','" & Replace(TempDescription, "'", "''") & "','" & PDMUserName & "','" & Now() & " ')"
    End If
    
    If Err > 0 Then
        Conn.RollbackTrans
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdBuildFirstBOM"
    Else
        Conn.CommitTrans
        MsgBox "You have built / initialized one BOM successfully.", vbInformation, "System Info."
        Command1_Click
    End If
    
End Sub


Private Sub cmdLock_Click()
    On Error Resume Next
    If (Text1.Text <> "" Or Not IsNumeric(Text1.Text)) Then
        If cmdLock.Caption = "UNLOCK" Then
            '########################����BOM############################
            StrSql = "IF EXISTS(SELECT * FROM BOMCPCN WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "') UPDATE BOMCPCN SET IsLocked=1,Locker='" & PDMUserName & "' WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "' ELSE INSERT INTO BOMCPCN([BOMID],[CPCNNmbr],[CPCNLocate],[BOMVersion],[IsSave],[UpdateDate],[IsLocked],[Locker]) VALUES('" & FinishGoodsNO & "','" & txtCPCNNO.Text & "',''," & cmbBOMVersion.Text & ",0,getdate(),1,'" & PDMUserName & "')"
            Conn.Execute StrSql
            Shape3.BackColor = &HC0C0FF
            cmdLock.Caption = "LOCKED"
            cmdLock.ForeColor = &HFF&
            txtCPCNNO.Enabled = False
        ElseIf cmdLock.Caption = "LOCKED" And Not IsBOMLocked Then
             '########################����BOM############################
            StrSql = "UPDATE BOMCPCN SET IsLocked=0,Locker='" & PDMUserName & "' WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "'"
            Conn.Execute StrSql
            Shape3.BackColor = &HFFC0C0
            cmdLock.Caption = "UNLOCK"
            cmdLock.ForeColor = &HFF0000
            txtCPCNNO.Enabled = True
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim i As Long, J As Long
    Dim rtMargin As RECT, rtCell As RECT, rtText As RECT

    If MsgBox("Are you sure that the default printer has set up Horizontal printing?", vbYesNo, "ERP") = vbNo Then Exit Sub
    '���ô�ӡ��Ϣ
    Printer.PaperSize = vbPRPSA4
    Printer.DrawMode = vbPixels
    SetRect rtMargin, 100, 100, 100, 100 'ҳ�߾�
    '��ʼ��ӡ
    Printer.CurrentX = rtMargin.left
    Printer.CurrentY = rtMargin.top
    Printer.Print "" '��ֽ
    SetRect rtCell, rtMargin.left, rtMargin.top, 0, 0
    With MSFlexGrid1
        For i = 0 To .Rows - 1
            .Row = i
            'ȷ���Ƿ�Ҫ��ҳ
            If Printer.ScaleHeight - .RowHeight(i) <= rtMargin.bottom Then
                Printer.NewPage
                rtCell.top = rtMargin.top
            End If
            For J = 0 To .Cols - 1
                .Col = J
                '��ӡ��Ԫ��߿�
                rtCell.right = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
                rtCell.bottom = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
                Rectangle Printer.hDC, rtCell.left, rtCell.top, rtCell.right + 1, rtCell.bottom + 1
                '���õ�Ԫ������
                Printer.FontName = .CellFontName
                Printer.FontSize = .CellFontSize
                Printer.FontBold = .CellFontBold
                Printer.FontItalic = .CellFontItalic
                Printer.FontUnderline = .CellFontUnderline
                '��ӡ��Ԫ�����֣������ڱ߾�Ϊ4��
                SetRect rtText, rtCell.left + 4, rtCell.top + 4, rtCell.right - 4, rtCell.bottom - 4
                DrawText Printer.hDC, .TextMatrix(i, J), LenB(StrConv(.TextMatrix(i, J), vbFromUnicode)), rtText, _
                DT_SINGLELINE Or GetAlign(.CellAlignment)
                rtCell.left = rtCell.left + .CellWidth \ Printer.TwipsPerPixelX
            Next
            rtCell.left = rtMargin.left
            rtCell.top = rtCell.top + .RowHeight(i) \ Printer.TwipsPerPixelY
        Next
    End With
    '��ӡ���
    Printer.EndDoc
End Sub


Private Sub cmdReview_Click()
    FrmBOMReview.Show 0
    FrmBOMReview.cmbBOMVersion.Text = cmbBOMVersion.Text
    FrmBOMReview.Text1.Text = FinishGoodsNO
    Call FrmBOMReview.cmdReiew_Click
End Sub

Private Sub GetTopBOM(ByVal ChildId As String)
    Dim StrSql  As String
    Dim objrs As New ADODB.Recordset
    'If temp_tb_BOMOrigData <> "" Then
    '    StrSql = "Select ParentId From " & temp_tb_BOMOrigData & " Where LEFT(ChildId,11)='" & left(ChildId, 11) & "'"
    'Else
        StrSql = "Select ParentId From  BOMOrigData Where LEFT(ChildId,11)='" & left(ChildId, 11) & "'"
    'End If
    'Debug.Print StrSql
    objrs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If objrs.RecordCount > 0 Then
        Do While Not objrs.EOF
            Call GetTopBOM(objrs(0))
            objrs.MoveNext
        Loop
    Else
        If InStr(BOMString, ChildId) = 0 Then
            BOMString = BOMString & "," & ChildId
        End If
    End If
    objrs.Close
    Set objrs = Nothing
End Sub

Private Sub CmdSearchSglPrt_Click()
    QueryTableName = "SglPrt"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    Set FromForm = FrmBOMAdmin
    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������

End Sub

Private Sub cmdUpgrade_Click()
    On Error Resume Next
    Dim SglPrtNo1, SglPrtNo2, StrSql As String
    Dim strBOM As String
    Dim rs As New ADODB.Recordset
    Dim i, othBOMVer As Integer
    SglPrtNo1 = CStr(left(txt12NC.Text, 11) & cmbSglVer1.Text)
    SglPrtNo2 = CStr(left(txt12NC.Text, 11) & cmbSglVer2.Text)
    If CInt(cmbSglVer2.Text) <= CInt(cmbSglVer1.Text) Then
        MsgBox "Invalid Version Number, Please choose the valid number.", vbCritical
        Exit Sub
    Else
        
        BOMString = ""
        Call GetTopBOM(SglPrtNo1)
        arrBOM = Split(Mid(BOMString, 2), ",")
        If BOMString <> "" Then
            
            'msgbox ���ʾ1024�ַ�
            If MsgBox("The 12NC had used in the following BOMs: " & vbCrLf & vbCrLf & Mid(BOMString, 2) & vbCrLf & vbCrLf & "Are you sure to upgrade the 12NC version for the above BOMs?", vbYesNo) = vbYes Then
                Screen.MousePointer = 11
                Conn.BeginTrans
                '����SglPrt�汾
                If CurVersion = 1 Then
                    StrSql = "Update SglPrt Set SglPrtVer=" & cmbSglVer2.Text & " Where SglPrtIndex = '" & txt12NC.Text & "'"
                Else
                    StrSql = "INSERT INTO PartVar([BOM],[CPCN],[PartIndex],[PartValue],[TableName],[FieldName]) VALUES('" & FinishGoodsNO & "','" & txtCPCNNO.Text & "','" & SglPrtNo1 & "','" & right(SglPrtNo2, 1) & "','SglPrt','SglPrtVer')"
                End If
                Conn.Execute StrSql
                StrSql = "Update " & temp_tb_BOMOrigData & " Set ChildId='" & SglPrtNo2 & "' Where ChildId='" & SglPrtNo1 & "'"
                Conn.Execute StrSql
                StrSql = "Update " & temp_tb_BOMOrigData & "  Set ParentID='" & SglPrtNo2 & "' Where ParentID='" & SglPrtNo1 & "'"
                Conn.Execute StrSql
                
'                '��շ��޸�����
'                StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CStr(CurVersion) & " And ChgStatus=''"
'                Conn.Execute StrSql
                
                '��д��־ ����Version��Ҫ��¼CPCN
                StrSql = "Select ParentID From  " & temp_tb_BOMOrigData & "   Where ChildID='" & SglPrtNo2 & "' Order By ParentID"
                
                rs.Open StrSql, Conn, adOpenKeyset, adLockReadOnly
                Do While Not rs.EOF
                    If CheckIsInBOM(CStr(rs(0)), MSFlexGrid1) Then
'                        'ɾ���ϼ��޸ĵľ���־
'                        StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO
'                        'StrSql = StrSql & " And ParentID=" & rs(0)
'                        StrSql = StrSql & " And Left(ChildID,11)=" & left(SglPrtNo1, 11)
'                        StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                        StrSql = "UPDATE  " & temp_tb_SglPrt4BOMLog & " SET ChildID=" & SglPrtNo1 & ", chgStatus='Delete-Upgrade',CPCN='" & txtCPCNNO.Text & "'  Where BOM=" & FinishGoodsNO
                        StrSql = StrSql & " And ParentID=" & rs(0)
                        StrSql = StrSql & " And ChildID=" & SglPrtNo1
                        StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                        Conn.Execute StrSql
                        
                        StrSql = "UPDATE " & temp_tb_SglPrt4BOMLog & " SET ParentID = " & SglPrtNo2 & " WHERE BOM = " & FinishGoodsNO
                        StrSql = StrSql & " And ParentID=" & SglPrtNo1
                        StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                        Conn.Execute StrSql
                        
'                        'Debug.Print StrSql
'                        '�������µ��޸���־
'                        StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,CPCN,Family,IsMultiBOM) Values("
'                        StrSql = StrSql & FinishGoodsNO
'                        StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1
'                        StrSql = StrSql & "," & rs(0)
'                        StrSql = StrSql & "," & SglPrtNo1
'                        StrSql = StrSql & "," & CurVersion
'                        StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4)
'                        StrSql = StrSql & ",'" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 5)
'                        StrSql = StrSql & "','" & Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 6), "'", "''")
'                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 7)
'                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 8)
'                        StrSql = StrSql & "','" & "Delete-Upgrade"
'                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 10)
'                        StrSql = StrSql & "','" & txtCPCNNO.Text
'                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 11)
'                        StrSql = StrSql & "',0)"
'                        Conn.Execute (StrSql)
'                        'Debug.Print StrSql
                        
                        StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,CPCN,Family,IsMultiBOM) Values("
                        StrSql = StrSql & FinishGoodsNO
                        StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1 'ֻ�ܼ�1����2��������
                        StrSql = StrSql & "," & rs(0)
                        StrSql = StrSql & "," & SglPrtNo2
                        StrSql = StrSql & "," & CurVersion
                        StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4)
                        StrSql = StrSql & ",'" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 5)
                        StrSql = StrSql & "','" & Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 6), "'", "''")
                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 7)
                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 8)
                        StrSql = StrSql & "','" & "Add-Upgrade"
                        StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 10)
                        StrSql = StrSql & "','" & txtCPCNNO.Text
                        StrSql = StrSql & "','" & Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 11), SglPrtNo1, SglPrtNo2)
                        StrSql = StrSql & "',0)"
                        Conn.Execute (StrSql)
                        'Debug.Print StrSql
                    End If
                rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
                
                
                If Err Then
                    Conn.RollbackTrans
                Else
                    Conn.CommitTrans
                    ChgMass = True
                    OrientCurNodeKey = Mid(OrientCurNodeKey, 1, Len(OrientCurNodeKey) - 1) & cmbSglVer2.Text
                    FrmUpgrade.Visible = False
                    Refresh_FlexGrid_TreeView False
                End If
                Screen.MousePointer = 0
            End If
        End If
    End If
End Sub


Private Sub Command5_Click()
    FrmUpgrade.Visible = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 6 Then
        If Len(Trim(Text2.Text)) = 0 Then        '�ж�Text2(����SinglePart NO)���ݵĺϷ���
            MsgBox "You must enter a new 12NC for the Single Part", vbInformation, "System Info."
            Exit Sub
        ElseIf Not (Len(Trim(Text2.Text)) = 12 And Isnum(Trim(Text2.Text))) Then
            MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
            Exit Sub
        Else
            Call GetTopBOM(Text2.Text)
            WriteLOG (Text2.Text & " is used to " & vbCrLf & BOMString)
        End If
    End If
End Sub

Private Sub Form_Resize()
    'ȷ������ı�ʱ�ؼ���֮�ı�
    Resize_ALL Me
End Sub
'Mid(myNode, 2, Len(myNode))    'ȥ�������һ���ַ�
'Mid(myNode, 1, Len(myNode)-1)  'ȥ�����ұ�һ���ַ�

Private Sub MSFlexGridColumnColorChange(MSFlexGridName As Object, ByVal ColNo As Integer, ByVal RowSum As Integer, Optional ByVal ColColor As Long = &HC0E0FF)      '&HC0E0FFΪǳ�ۺ�ɫ
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Col = ColNo                    '�ӵ�ColNo�е�0�п�ʼ
    MSFlexGridName.Row = 0                        '�ӵ�ColNo�е�0�п�ʼ
    MSFlexGridName.RowSel = RowSum - 1            '����ѡ��ֱ�����һ��RowSum
    MSFlexGridName.CellBackColor = ColColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub
Private Sub MSFlexGridRowColorChange(MSFlexGridName As Object, ByVal RowNo As Integer, ByVal ColSum As Integer, Optional ByVal RowColor As Long = &HFFFF&)          '&HFFFFΪ��ɫ
    
    MSFlexGridName.FillStyle = flexFillRepeat
    MSFlexGridName.Row = RowNo                    '�ӵ�RowNo�е�0�п�ʼ
    MSFlexGridName.Col = 1                        '�ӵ�RowNo�е�0�п�ʼ
    MSFlexGridName.ColSel = ColSum - 1            '����ѡ��ֱ�����һ��ColSum
    MSFlexGridName.CellBackColor = RowColor
    MSFlexGridName.FillStyle = flexFillSingle
    
End Sub
'��һ���򼸸��ַ�����һ���ַ�����ĳ��λ�õĺ���
Private Function InsertStr(ByVal strSource As String, ByVal strIn As String, ByVal intPos As Integer) As String
    'strSourceԴ�ַ���   strIn�����ַ���     intPos��Ҫ�����λ��
    '���þ�����a = InsertStr("1234567", "aaa", 5)
    InsertStr = left(strSource, intPos - 1) & strIn & Mid(strSource, intPos)  'Mid �������ʡ�Ի�length�����ı����ַ��������� start �����ַ������������ַ����д� start ��β�˵������ַ���
End Function

'�ж�һ���ַ������Ƿ�������  ��IsNumeric�ж�0000d031Ϊ��(����double������)
Private Function Isnum(Str As String) As Boolean
    Isnum = True
    Dim i  As Integer
    For i = 1 To Len(Str)
        Select Case Mid(Str, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            'Isnum = True  ����дIsnum = True�ͳ���,��Ϊ����м�����ĸfalse�˺��������ֵĻ��ֳ�Ϊtrue��
        Case Else
            Isnum = False
        End Select
    Next
End Function

'�жϽڵ�A�ַ����Ƿ�Ϊ�ڵ�B��ֱ��Rootǰ���ڵ�����&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeEldership(NodeAString As String, NodeB As Node) As Boolean
    If NodeB Is Nothing Then Exit Function
    If LeftcutStrg(NodeB.Key) = NodeAString Then
        AddNodeEldership = True
        RecursionFlag = True
        Exit Function
    Else
        AddNodeEldership = AddNodeEldership(NodeAString, NodeB.Parent)  '�ݹ����
    End If
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�A�ַ����Ƿ�Ϊ�ڵ�B��ͬ���ӽڵ�(�ֵܽڵ�)����&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeBrothership(NodeAString As String, NodeB As Node) As Boolean
    '����TreeView�����нڵ�ı�����
    Dim nodEachChild As Node
    For Each nodEachChild In tvCodeItems.Nodes
        
        If nodEachChild.Parent Is Nothing Then GoTo NextNode  'Root���ڵ�ParentΪNothing, ���жϳ����������ѭ������
        If nodEachChild.Parent = NodeB And nodEachChild = NodeAString Then
            AddNodeBrothership = True
            RecursionFlag = True
            Exit Function
        End If
NextNode:
        If RecursionFlag Then Exit Function
    Next nodEachChild
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�A�ͽڵ�B�ĵ�һ���ӽڵ�(���Ҫ�Ǻϲ����Ϊ�ֵܽڵ�)�����Ƿ�����ͬ&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeChildIsBrothership(NodeA As Node, NodeB As Node) As Boolean
    '����TreeView�����нڵ�ı�����
    Dim nodEachAChild As Node
    For Each nodEachAChild In tvCodeItems.Nodes
        
        If nodEachAChild.Parent Is Nothing Then GoTo NextNodeAChild  'Root���ڵ�ParentΪNothing, ���жϳ����������ѭ������
        If nodEachAChild.Parent = NodeA Then
            
            Dim nodEachBChild As Node
            For Each nodEachBChild In tvCodeItems.Nodes
                
                If nodEachBChild.Parent Is Nothing Then GoTo NextNodeBChild  'Root���ڵ�ParentΪNothing, ���жϳ����������ѭ������
                If nodEachBChild.Parent = NodeB Then
                    If nodEachAChild = nodEachBChild Then
                        AddNodeChildIsBrothership = True
                        RecursionFlag = True
                        Exit Function
                    End If
                End If
NextNodeBChild:
                If RecursionFlag Then Exit Function
            Next nodEachBChild
            If RecursionFlag Then Exit Function
            
        End If
NextNodeAChild:
        If RecursionFlag Then Exit Function
    Next nodEachAChild
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�A�����ַ�����TreeView���Ƿ���ͬ�ڵ����Ľڵ���ڲ��һ����ӽڵ�&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeKeyNameNodeExist(NodeAString As String) As Node
    '����TreeView�����нڵ�ı�����
    Dim nodEachChild As Node
    For Each nodEachChild In tvCodeItems.Nodes
        If nodEachChild = NodeAString And nodEachChild.Children > 0 Then
            Set AddNodeKeyNameNodeExist = nodEachChild
            If RecursionFlag Then Exit Function
            Exit Function
        End If
        If RecursionFlag Then Exit Function
    Next nodEachChild
    If RecursionFlag Then Exit Function
End Function
'�жϽڵ�A�Ƿ�Ϊ�ڵ�B��ǰ���ڵ�
Private Function isEldershipNode(NodeA As Node, NodeB As Node) As Boolean
    '  SourceNode��ӦNodeA, TargetNode��ӦNodeB
    If Not (NodeB.Parent Is Nothing) Then           '����ڵ�B�и��ڵ�(���Ǹ��ڵ�)
        If LeftcutStrg(NodeB.Parent.Key) = LeftcutStrg(NodeA.Key) Then
            isEldershipNode = True
            RecursionFlag = True
            Exit Function
        Else
            isEldershipNode = isEldershipNode(NodeA, NodeB.Parent)  '�ݹ����
        End If
    End If
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�B�Ƿ�Ϊ�ڵ�A���ӱ��ڵ�
Private Function isYoungershipNode(NodeA As Node, NodeB As Node) As Boolean
    '  SourceNode��ӦNodeA, TargetNode��ӦNodeB
    Dim nodEachChild As Node
    Set nodEachChild = NodeA.Child         'ǰ���Set����ȥ��
    Do While Not nodEachChild Is Nothing
        If LeftcutStrg(nodEachChild.Key) = LeftcutStrg(NodeB.Key) Then
            isYoungershipNode = True
            RecursionFlag = True
            Exit Function
        Else
            isYoungershipNode = isYoungershipNode(nodEachChild, NodeB)  '�ݹ����
        End If
        Set nodEachChild = nodEachChild.Next
        If RecursionFlag Then Exit Function
    Loop
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�B�Ƿ�Ϊ�ڵ�A�Ķ��ӱ��ڵ� (�����Լ����Ӻ󶼲�����)
Private Function isSonshipNode(NodeA As Node, NodeB As Node) As Boolean
    '  SourceNode��ӦNodeA, TargetNode��ӦNodeB
    Dim nodEachChild As Node
    
    For Each nodEachChild In tvCodeItems.Nodes
        
        If nodEachChild.Parent Is Nothing Then GoTo NextSonNode     'Root���ڵ�ParentΪNothing, ���жϳ����������ѭ������
        If nodEachChild.Parent = NodeA And nodEachChild = NodeB Then
            isSonshipNode = True
            RecursionFlag = True
            Exit Function
        End If
NextSonNode:
        If RecursionFlag Then Exit Function
    Next nodEachChild
    If RecursionFlag Then Exit Function
End Function

'�ж�B��Ϊ�ڵ����Ƿ�ͽڵ�A���ӱ��ڵ�ͬ��
Private Function isYoungershipNameNode(NodeA As Node, NodeBString As String) As Boolean
    Dim nodEachChild As Node
    Set nodEachChild = NodeA.Child         'ǰ���Set����ȥ��
    Do While Not nodEachChild Is Nothing
        If LeftcutStrg(nodEachChild.Key) = NodeBString Then
            isYoungershipNameNode = True
            RecursionFlag = True
            Exit Function
        Else
            isYoungershipNameNode = isYoungershipNameNode(nodEachChild, NodeBString)  '�ݹ����
        End If
        Set nodEachChild = nodEachChild.Next
        If RecursionFlag Then Exit Function
    Loop
    If RecursionFlag Then Exit Function
End Function

'�жϽڵ�A�ı����ӱ��ڵ��Ƿ�Ϊ�ڵ�B��ֱ��Rootǰ���ڵ�
Private Function isElderYoungershipNode(NodeA As Node, NodeB As Node) As Boolean
    '  SourceNode��ӦNodeA, TargetNode��ӦNodeB
    Dim nodEachChild As Node
    Dim nodEachParent As Node
    Set nodEachParent = NodeB.Parent       'ǰ���Set����ȥ��
    
    Do While Not nodEachParent Is Nothing
        Set nodEachChild = NodeA.Child         'ǰ���Set����ȥ��
        Do While Not nodEachChild Is Nothing
            If LeftcutStrg(nodEachChild.Key) = LeftcutStrg(nodEachParent.Key) Then
                isElderYoungershipNode = True
                RecursionFlag = True
                Exit Function
            Else
                isElderYoungershipNode = isYoungershipNode(nodEachChild, nodEachParent)    '�ݹ����
            End If
            Set nodEachChild = nodEachChild.Next
            If RecursionFlag Then Exit Function
        Loop
        Set nodEachParent = nodEachParent.Parent    'ǰ���Set����ȥ��
        If RecursionFlag Then Exit Function
    Loop
    If RecursionFlag Then Exit Function
End Function

Private Function LeftcutStrg(cutstrg As String) As String
    Dim i  As Long
    
    LeftcutStrg = ""
    
    For i = 1 To Len(cutstrg)
        If Asc(Mid(cutstrg, i, 1)) >= 48 And Asc(Mid(cutstrg, i, 1)) <= 57 Then        'Asc 48-57��Ӧ0,1,2,....9
            LeftcutStrg = LeftcutStrg & Mid(cutstrg, i, 1)
        End If
    Next
End Function

Private Sub CmdApprove_Click()
    FrmBOMApproval.TxtFinsGdIndex = MSFlexGrid1.TextMatrix(1, 3)
    FrmBOMApproval.TxtDescription = MSFlexGrid1.TextMatrix(1, 6)
    Set FromForm = FrmBOMAdmin
    FrmBOMApproval.Show
End Sub


Private Sub GeneralPathAdd(ByVal InputPathName As String, ByVal InputField As String)
    On Error GoTo vbErrorHandler
    Dim DocPathName As String
    
        
    DocPathName = Trim(InputPathName)
    If DocPathName = "" Then
        MsgBox "The Document Path/name is Null", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Mid(DocPathName, 1, 3) = "P:\") Then
        MsgBox "The Document Path must be formal Released or Shared(P:\)", vbInformation, "System Info."
        Exit Sub
    ElseIf Not OpnFileExist(DocPathName) Then
        MsgBox "The Document Path/Name is NOT existing, Please Check Path/Name", vbInformation, "System Info."
        Exit Sub
   End If

    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    If Len(Trim(txtNodeSglPrt12NC.Text)) = 0 Then        '�ж�txtNodeSglPrt12NC(����SinglePart NO Or FinishGood NO )���ݵĺϷ���
      MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
      Exit Sub
    ElseIf Not (Len(Trim(txtNodeSglPrt12NC.Text)) = 12 And Isnum(Trim(txtNodeSglPrt12NC.Text))) Then
          MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
          Exit Sub
    Else
       '��ʼ�ж������Finish Goods NO �Ƿ������ݿ�������
       rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(txtNodeSglPrt12NC.Text) & "'", Conn, adOpenKeyset, adLockOptimistic
       If rs.RecordCount = 0 Then
            'MsgBox "The Finish Goods NO. is not existing in database", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
             '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(txtNodeSglPrt12NC.Text), 1, Len(Trim(txtNodeSglPrt12NC.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
              MsgBox "The Selected Item 12NC is not existing in Database", vbInformation, "System Info."
              If rs.State = adStateOpen Then rs.Close
              Exit Sub
            Else
              rs(InputField) = DocPathName
              rs.Update
            End If
            Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & left(Trim(txtNodeSglPrt12NC.Text), 11) & "0" & " AND ITEMVALUE='" & DocPathName & "' AND CREATOR='" & InputField & "') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & left(Trim(txtNodeSglPrt12NC.Text), 11) & "0" & ",'" & DocPathName & "','" & InputField & "')")
            MsgBox "The Item Drawing(Document) Path/Name has been Added successfully ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
       Else
            rs(InputField) = DocPathName
            rs.Update

            Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & Trim(txtNodeSglPrt12NC.Text) & " AND ITEMVALUE='" & DocPathName & "' AND CREATOR='" & InputField & "') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & left(Trim(txtNodeSglPrt12NC.Text), 11) & "0" & ",'" & DocPathName & "','" & InputField & "')")
            MsgBox "The Item Drawing(Document) Path/Name has been Added successfully ", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
       End If
       If rs.State = adStateOpen Then rs.Close
    End If
Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdDrwPathAdd"
End Sub
Private Sub ClearPathAdd(ByVal InputField As String)
    On Error GoTo vbErrorHandler
    'Dim DocPathName As String
    
    If MsgBox("Confirm to Clear the Path/Name?" + vbCrLf + "ȷ���Ƿ����·��?", vbYesNo + vbDefaultButton2, "Confirm to Clear ȷ�����") = vbNo Then
        Exit Sub
    End If
    
    
    Dim rs As New ADODB.Recordset
    
    
    If Len(Trim(txtNodeSglPrt12NC.Text)) = 0 Then        '�ж�txtNodeSglPrt12NC(����SinglePart NO Or FinishGood NO )���ݵĺϷ���
        MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(txtNodeSglPrt12NC.Text)) = 12 And Isnum(Trim(txtNodeSglPrt12NC.Text))) Then
        MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    Else
        '��ʼ�ж������Finish Goods NO �Ƿ������ݿ�������
        rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(txtNodeSglPrt12NC.Text) & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            'MsgBox "The Finish Goods NO. is not existing in database", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
            GoTo TrySglPartPath
        Else
            rs(InputField) = ""
            rs.Update
        End If
        If rs.State = adStateOpen Then rs.Close
        GoTo FinishItemPathClear
TrySglPartPath:
        '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
        rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(txtNodeSglPrt12NC.Text), 1, Len(Trim(txtNodeSglPrt12NC.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The Selected Item 12NC is not existing in Database", vbInformation, "System Info."
            If rs.State = adStateOpen Then rs.Close
            Exit Sub
        Else
            rs(InputField) = ""
            rs.Update
        End If
        If rs.State = adStateOpen Then rs.Close
    End If
FinishItemPathClear:
    MsgBox "The Item Drawing(Document) Path/Name has been Cleared successfully ", vbInformation, "System Info."
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:ClearPathAdd"
End Sub
Private Sub GeneralDocView(ByVal InputPathName As String)
    Dim OpnDocPathName As String
    OpnDocPathName = Trim(InputPathName)
    If OpnDocPathName = "" Then
        MsgBox "The Drawing(Document) Path/name is Null", vbInformation, "System Info."
        Exit Sub
    End If
    OpnShllExcFile (OpnDocPathName)
End Sub

Private Sub CmdCPCNPathAdd_Click()
    GeneralPathAdd txtCPCNlocate.Text, "CPCNLocate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim xConn As New ADODB.Connection
    xConn.Open connString
    
    If ChgMass And CurVersion > 1 Then
        If MsgBox("No Save BOM, would you like to save it?.", vbYesNo) = vbYes Then
            Call cmdBOMSave_Click
        Else
            ChgMass = False
        End If
    End If
    
    If BOMLock Then
        If BOMLocker = PDMUserName Or SystemAdmin = "Y" Then
            If MsgBox("The BOM is locked, would you like to unlock it now?", vbYesNo, "PDM") = vbYes Then
                StrSql = "UPDATE BOMCPCN SET IsLocked=0 WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "'"
                xConn.Execute StrSql
                cmdLock.Caption = "UNLOCK"
                txtCPCNNO.Enabled = True
                BOMLock = False
            End If
        End If
    End If
    
    '####�˳�bomҪ�����ʱ��#########
    If FinishGoodsNO <> "" Then DropTempTable

    If xConn.State = adStateOpen Then xConn.Close: Set xConn = Nothing
    FrmEngineeringSys.Show 0
    Unload Me
End Sub

Private Sub LblCPCN_Click()
    ClearPathAdd "CPCNlocate"
End Sub

Private Sub LblCPCNPathSeek_Click()
    On Error GoTo vbErrorHandler
    Dim Filename As String, GetFilePath As String, Lines As String, RowsNum As Long
    Dim pId As Long, pHnd As Long ' �ֱ����� Process Id �� Process Handle ����
    
    If Len(Trim(CPCN)) = 0 Then
        MsgBox "You must Enter a CPCN Number into the TextBox after CP/CN Number", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(CPCN)) = 12) Then
        MsgBox "The Input Content is 12 Letter+Number,Such As HK1FS0001575" + vbCrLf + "������12λ��ĸ�����ֵı��,����HK1FS0001575", vbInformation, "System Info."
        Exit Sub
    End If
    
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"              '�ж�һ��Ŀ¼�Ƿ���ڲ���������
    If Len(Dir("C:\Temp\Search.txt")) > 0 Then Kill "C:\Temp\Search.txt"  '�ж�һ���ļ��Ƿ����,������ɾ��
    Filename = Trim(CPCN)
    'FileName = Mid(FileName, 1, Len(FileName) - 1) 'ȥ�����ұ�һ���ַ�
    Filename = "*" & Filename & "*"  'ǰ����Ǻ�
    GetFilePath = InputBox("Please input Directory Path", "Prompt Info ��������·��", "P:\Shenzhen\PssDoc\CP-CN\", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
    
    'Shell "CMD /K " & Left(GetFilePath, 2) '����Ҫ������Ӧ���������̷���,����ȡ2λ�ַ�
    'Shell "CMD /K CD " & GetFilePath       '�ٽ�����ӦĿ¼��
    'Shell "CMD /C DIR " & FileName & " /A/L/B/S >C:\Temp\Search.txt", 1 '(/LСд��ĸ,/S������Ŀ¼,/B��û��headingTitle��Summary /A��ʾ�ر������ļ�  1 VbNormalFocus ���ھ��н���,�һỹԭ��ԭ���Ĵ�Сλ�� )
    
    If Len(Dir("C:\Temp\Share.bat")) > 0 Then                    '�鿴�������ļ����ڷ�
        Kill "C:\Temp\Share.bat"                                 '������ɾ��
    End If
    Open "C:\Temp\Share.bat" For Output As #2                '���ļ�׼��д��
    Print #2, left(GetFilePath, 2)
    Print #2, "CD " & GetFilePath
    Print #2, "DIR " & Filename & " /A/B/S >C:\Temp\Search.txt"
    Close #2
    'Shell "C:\Temp\Share.bat"
    pId = Shell("C:\Temp\Share.bat", 0)        ' Shell ���� Process Id
    pHnd = OpenProcess(SYNCHRONIZE, 0, pId)    ' ȡ�� Process Handle
    If pHnd <> 0 Then
        Call WaitForSingleObject(pHnd, INFINITE)   ' ���޵ȴ���ֱ���������
        Call CloseHandle(pHnd)
    End If
    Kill "C:\Temp" & "\Share.bat"
    
    Open "C:\Temp\Search.txt" For Input As #1
    '���ֻ��һ�����ݶ�ȡ����ʹ����� Line Input #1, GetFilePath, ����ж������ݶ�ȡ��������ѭ��
    RowsNum = 0
    Do While Not EOF(1)  'EOF(filenumber) ����һ��Boolean ֵ�������Ƿ��Ѿ�����Ϊ Random ��˳�� Input �򿪵��ļ��Ľ�β��
        On Error Resume Next
        Line Input #1, GetFilePath 'Line Input #filenumber, varname ���Ѿ��򿪵��ļ�˳���ȡһ�в����������String����
        Lines = Lines & GetFilePath & Chr(13) & Chr(10)   'chr(13)�س�. Chr(10)���� chr(32)�ո�
        RowsNum = RowsNum + 1
    Loop
    
    If RowsNum = 0 Then
        MsgBox "No matching record found.", vbInformation, "System Info."
        Close #1
        Exit Sub
    End If
    
    If RowsNum = 1 Then   '���ֻ��һ�����ݶ�ȡ����ֱ�Ӹ�ֵ��Textbox: txtSERlocate
        txtCPCNlocate.Text = Replace(Lines, vbCr + vbLf, "")   'ȥ���س����з���
        Close #1
        Exit Sub
    Else
        Close #1              '����ж�������,�ʹ�txt�ļ��˹�ѡ��
        OpnShllExcFile ("C:\Temp\Search.txt")
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:LblCPCNPathSeek_Click"
End Sub
Private Sub CmdCPCNView_Click()
    GeneralDocView (txtCPCNlocate.Text)
End Sub
'####################################
Private Sub CmdSERPathAdd_Click()
    GeneralPathAdd txtSERlocate.Text, "SERlocate"
End Sub

Private Sub LblSER_Click()
    ClearPathAdd "SERlocate"
End Sub

Private Sub LblSERPathSeek_Click()
    On Error GoTo vbErrorHandler
    Dim Filename As String, GetFilePath As String, Lines As String, RowsNum As Long
    Dim pId As Long, pHnd As Long ' �ֱ����� Process Id �� Process Handle ����
    
    If Len(Trim(txtNodeSglPrt12NC.Text)) = 0 Then
        MsgBox "You must Enter a 12NC into the TextBox under Selected 12NC", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(txtNodeSglPrt12NC.Text)) = 12 And Isnum(Trim(txtNodeSglPrt12NC.Text))) Then
        MsgBox "The Input Content is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    End If
    
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"              '�ж�һ��Ŀ¼�Ƿ���ڲ���������
    If Len(Dir("C:\Temp\Search.txt")) > 0 Then Kill "C:\Temp\Search.txt"  '�ж�һ���ļ��Ƿ����,������ɾ��
    Filename = Trim(txtNodeSglPrt12NC.Text)
    Filename = Mid(Filename, 1, Len(Filename) - 1) 'ȥ�����ұ�һ���ַ�
    Filename = "*" & Filename & "*"  'ǰ����Ǻ�
    GetFilePath = InputBox("Please input Directory Path", "Prompt Info ��������·��", "P:\Shenzhen\PssDoc\SER\", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
    
    'Shell "CMD /K " & Left(GetFilePath, 2) '����Ҫ������Ӧ���������̷���,����ȡ2λ�ַ�
    'Shell "CMD /K CD " & GetFilePath       '�ٽ�����ӦĿ¼��
    'Shell "CMD /C DIR " & FileName & " /A/L/B/S >C:\Temp\Search.txt", 1 '(/LСд��ĸ,/S������Ŀ¼,/B��û��headingTitle��Summary /A��ʾ�ر������ļ�  1 VbNormalFocus ���ھ��н���,�һỹԭ��ԭ���Ĵ�Сλ�� )
    
    If Len(Dir("C:\Temp\Share.bat")) > 0 Then                    '�鿴�������ļ����ڷ�
        Kill "C:\Temp\Share.bat"                                 '������ɾ��
    End If
    Open "C:\Temp\Share.bat" For Output As #2                '���ļ�׼��д��
    Print #2, left(GetFilePath, 2)
    Print #2, "CD " & GetFilePath
    Print #2, "DIR " & Filename & " /A/B/S >C:\Temp\Search.txt"
    Close #2
    'Shell "C:\Temp\Share.bat"
    pId = Shell("C:\Temp\Share.bat", 0)        ' Shell ���� Process Id
    pHnd = OpenProcess(SYNCHRONIZE, 0, pId)    ' ȡ�� Process Handle
    If pHnd <> 0 Then
        Call WaitForSingleObject(pHnd, INFINITE)   ' ���޵ȴ���ֱ���������
        Call CloseHandle(pHnd)
    End If
    Kill "C:\Temp" & "\Share.bat"
    
    Open "C:\Temp\Search.txt" For Input As #1
    '���ֻ��һ�����ݶ�ȡ����ʹ����� Line Input #1, GetFilePath, ����ж������ݶ�ȡ��������ѭ��
    RowsNum = 0
    Do While Not EOF(1)  'EOF(filenumber) ����һ��Boolean ֵ�������Ƿ��Ѿ�����Ϊ Random ��˳�� Input �򿪵��ļ��Ľ�β��
        On Error Resume Next
        Line Input #1, GetFilePath 'Line Input #filenumber, varname ���Ѿ��򿪵��ļ�˳���ȡһ�в����������String����
        Lines = Lines & GetFilePath & Chr(13) & Chr(10)   'chr(13)�س�. Chr(10)���� chr(32)�ո�
        txtSERlocate.AddItem Trim(GetFilePath)
        RowsNum = RowsNum + 1
    Loop
    
    If RowsNum = 0 Then
        MsgBox "No matching record found.", vbInformation, "System Info."
        Close #1
        Exit Sub
    End If

    If RowsNum = 1 Then   '���ֻ��һ�����ݶ�ȡ����ֱ�Ӹ�ֵ��Textbox: txtSERlocate
        txtSERlocate.Text = Replace(Lines, vbCr + vbLf, "")   'ȥ���س����з���
        Close #1
        Exit Sub
    Else
        Close #1              '����ж�������,�ʹ�txt�ļ��˹�ѡ��
        OpnShllExcFile ("C:\Temp\Search.txt")
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:LblSERPathSeek_Click"
End Sub
Private Sub CmdSERView_Click()
    GeneralDocView (txtSERlocate.Text)
End Sub
'####################################
Private Sub CmdDrwPathAdd_Click()
    GeneralPathAdd txtNodeDrwlocate.Text, "Drwlocate"
End Sub

Private Sub LblDrw_Click()
    ClearPathAdd "Drwlocate"
End Sub

Private Sub LblDRWPathSeek_Click()
    On Error GoTo vbErrorHandler
    Dim Filename As String, GetFilePath As String, Lines As String, RowsNum As Long
    Dim pId As Long, pHnd As Long ' �ֱ����� Process Id �� Process Handle ����
    
    If Len(Trim(txtNodeSglPrt12NC.Text)) = 0 Then
        MsgBox "You must Enter a 12NC into the TextBox under Selected 12NC", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(txtNodeSglPrt12NC.Text)) = 12 And Isnum(Trim(txtNodeSglPrt12NC.Text))) Then
        MsgBox "The Input Content is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Exit Sub
    End If
    
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"              '�ж�һ��Ŀ¼�Ƿ���ڲ���������
    If Len(Dir("C:\Temp\Search.txt")) > 0 Then Kill "C:\Temp\Search.txt"  '�ж�һ���ļ��Ƿ����,������ɾ��
    Filename = Trim(txtNodeSglPrt12NC.Text)
    Filename = Mid(Filename, 1, Len(Filename) - 1) 'ȥ�����ұ�һ���ַ�
    Filename = InsertStr(Filename, "*", 5)    '�м�ո���*�Ŵ���
    Filename = InsertStr(Filename, "*", 9)    '�м�ո���*�Ŵ���
    Filename = "*" & Filename & "*"  'ǰ����Ǻ�
    GetFilePath = InputBox("Please input Directory Path", "Prompt Info ��������·��", "P:\Shenzhen\PssDoc\DRAWING\", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
    
    'Shell "CMD /K " & Left(GetFilePath, 2) '����Ҫ������Ӧ���������̷���,����ȡ2λ�ַ�
    'Shell "CMD /K CD " & GetFilePath       '�ٽ�����ӦĿ¼��
    'Shell "CMD /C DIR " & FileName & " /A/L/B/S >C:\Temp\Search.txt", 1 '(/LСд��ĸ,/S������Ŀ¼,/B��û��headingTitle��Summary /A��ʾ�ر������ļ�  1 VbNormalFocus ���ھ��н���,�һỹԭ��ԭ���Ĵ�Сλ�� )
    
    If Len(Dir("C:\Temp\Share.bat")) > 0 Then                    '�鿴�������ļ����ڷ�
        Kill "C:\Temp\Share.bat"                                 '������ɾ��
    End If
    Open "C:\Temp\Share.bat" For Output As #2                '���ļ�׼��д��
    Print #2, left(GetFilePath, 2)
    Print #2, "CD " & GetFilePath
    Print #2, "DIR " & Filename & " /A/B/S >C:\Temp\Search.txt"
    Close #2
    'Shell "C:\Temp\Share.bat"
    pId = Shell("C:\Temp\Share.bat", 0)        ' Shell ���� Process Id
    pHnd = OpenProcess(SYNCHRONIZE, 0, pId)    ' ȡ�� Process Handle
    If pHnd <> 0 Then
        Call WaitForSingleObject(pHnd, INFINITE)   ' ���޵ȴ���ֱ���������
        Call CloseHandle(pHnd)
    End If
    Kill "C:\Temp" & "\Share.bat"
    
    Open "C:\Temp\Search.txt" For Input As #1
    '���ֻ��һ�����ݶ�ȡ����ʹ����� Line Input #1, GetFilePath, ����ж������ݶ�ȡ��������ѭ��
    RowsNum = 0
    Do While Not EOF(1)  'EOF(filenumber) ����һ��Boolean ֵ�������Ƿ��Ѿ�����Ϊ Random ��˳�� Input �򿪵��ļ��Ľ�β��
        On Error Resume Next
        Line Input #1, GetFilePath 'Line Input #filenumber, varname ���Ѿ��򿪵��ļ�˳���ȡһ�в����������String����
        Lines = Lines & GetFilePath & Chr(13) & Chr(10)   'chr(13)�س�. Chr(10)���� chr(32)�ո�
        txtNodeDrwlocate.AddItem Trim(GetFilePath)
        RowsNum = RowsNum + 1
    Loop
    
    If RowsNum = 0 Then
        MsgBox "No matching record found.", vbInformation, "System Info."
        Close #1
        Exit Sub
    End If

    If RowsNum = 1 Then   '���ֻ��һ�����ݶ�ȡ����ֱ�Ӹ�ֵ��Textbox: txtSERlocate
        txtNodeDrwlocate.Text = Replace(Lines, vbCr + vbLf, "")   'ȥ���س����з���
        Close #1
        Exit Sub
    Else
        Close #1              '����ж�������,�ʹ�txt�ļ��˹�ѡ��
        OpnShllExcFile ("C:\Temp\Search.txt")
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:LblDRWPathSeek_Click"
End Sub
Private Sub CmdDrwView_Click()
    GeneralDocView (txtNodeDrwlocate.Text)
End Sub
'####################################
Private Sub CmdExportBOM_Click()
    'On Error Resume Next
    

    
    If Not ApprovalStatus Then
        MsgBox "The BOM is NOT Approved, Please do NOT use it Formally(Offically)", vbInformation, "System Info."
    End If
    
    If MsgBox("You are going to Export BOM Data to an Excel File, Continue��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
        'ExportFlexDataToExcel MSFlexGrid1, CommonDialog1, "BOM"
        'FlexGrd_SaveToExcel MSFlexGrid1, "The   Header  Sample", "The   Footer", 1, 16, "", , , , , True
        
        Dim i, J As Integer

        Set xlApp = CreateObject("Excel.Application")   '����Excel�ļ�
        Set xlApp = New excel.Application
        
        
        '������ֲ���������ʾ
        xlApp.OleRequestPendingTimeout = 10000   '10000��������æ�Ի���
        xlApp.OleServerBusyTimeout = 1000     '����ʱ1��
        xlApp.OleServerBusyRaiseError = True '����ʾæ�Ի���
    
    
        xlApp.SheetsInNewWorkbook = 1                   '���½��Ĺ�����������Ϊ1
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)              '��1�Ź�����
        xlSheet.Cells(1, 1) = "BOM"
        If FinishGoodsNO <> "" Then xlSheet.Cells(2, 1) = "Finish Goods No:": xlSheet.Cells(2, 2) = "'" + FinishGoodsNO
        If txtSubCon.Text <> "" Then xlSheet.Cells(2, 3) = "SUBCON:": xlSheet.Cells(2, 4) = txtSubCon.Text
        
        For i = 0 To MSFlexGrid1.Cols - 2
            xlSheet.Cells(3, i + 1) = MSFlexGrid1.TextMatrix(0, i)
        Next i
        xlSheet.Cells(2, i - 4) = "Table Maker:": xlSheet.Cells(2, i - 3) = PDMUserName
        xlSheet.Cells(2, i - 2) = "Print Date:": xlSheet.Cells(2, i - 1) = Now()
        
        For J = 1 To MSFlexGrid1.Rows - 1
                For i = 0 To MSFlexGrid1.Cols - 2
                    xlSheet.Cells(J + 3, i + 1) = "'" + MSFlexGrid1.TextMatrix(J, i)
                Next i
        Next J
        'xlSheet.Cells(4, 1).CopyFromRecordset Conn.Execute(strSql)       '������ճ������
    
        xlApp.ActiveWorkbook.Close True     '�رչ�����������
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
    End If
End Sub

Private Sub CmdImportBOM_Click()
    Set FromForm2 = FrmBOMAdmin
    FrmBOMImport.Show 0
End Sub

Private Sub CmdSearchFinsGd_Click()
    QueryTableName = "FinsGd"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    Set FromForm = FrmBOMAdmin
    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
End Sub

Private Sub CmdSearch_SglPrt_Click()
    MousePointer = vbHourglass   '����ʱ��ϳ�����Ҫ�������״̬
    QueryTableName = "SglPrt"                                  '##########����ͨ�ò�ѯ�����Ƕ��ĸ�����в���
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    If SystemAdmin <> "Y" Then
        MsgBox "You are not Administrator, Some Access Right is NOT workable ", vbInformation, "System Info."
        FrmQuery.cmdModify.Enabled = False
        FrmQuery.CmdDel.Enabled = False
        
        FrmQuery.DataGrid1.AllowDelete = False
        FrmQuery.DataGrid1.AllowAddNew = False
        FrmQuery.DataGrid1.AllowUpdate = False
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û�������Ҫ����һЩ�޸�ɾ������
    Set FromForm = FrmBOMAdmin
    FrmQuery.Show 0 'frm.Show style StyleΪ0�Ǵ�������ģʽ�� style Ϊ 1������ģʽ��ģʽ����ʱ������ģʽ�����еĶ���֮�ⲻ�ܽ������루���̻���굥������
    MousePointer = vbDefault                  '�ָ����״̬
End Sub


Private Sub Command1_Click()
    On Error GoTo vbErrorHandler
    Dim i As Integer
    If ChgMass And CurVersion <> 1 Then
        If MsgBox("No Save BOM, would you like to save it?.", vbYesNo) = vbYes Then
            Call cmdBOMSave_Click
        Else
            ChgMass = False
        End If
    End If
    
    Me.Enabled = False
    frameMsg.Visible = True
    
    DoEvents
    
    MSFlexGrid1EditText.Visible = False
    tvCodeItems.Nodes.Clear
    
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "You must enter a new 12NC for the Finish Goods", vbInformation, "System Info."
        Exit Sub
    ElseIf Not (Len(Trim(Text1.Text)) = 12 And Isnum(Trim(Text1.Text))) Then
        MsgBox "Finish Goods is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Me.Enabled = True
        frameMsg.Visible = False
        Exit Sub
    Else
        FinishGoodsNO = Trim(Text1.Text)
    End If
    
    '���CPCN�汾��¼
    Dim rs As New ADODB.Recordset
    StrSql = "Select Top 1 BOMVersion,isNull(CPCNNmbr,''),isNull(CPCNLocate,''),isSave From BOMCPCN Where BOMID =" & FinishGoodsNO & " Order by BOMVersion Desc"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        CPCN = ""
        CurVersion = 1
        txtCPCNNO.Text = ""
        txtCPCNlocate.Text = ""
        bNotSave1stVer = False
    Else
        rs.MoveFirst
        CPCN = rs(1)
        txtCPCNNO.Text = rs(1)
        txtCPCNlocate.Text = rs(2)
        CurVersion = rs(0)
        bNotSave1stVer = False
    End If
    If rs.State = adStateOpen Then rs.Close
    
    cmbBOMVersion.Text = ""
    isApproved = CheckIsApproved(FinishGoodsNO)
    isRejected = CheckIsRejected(FinishGoodsNO)
    
    If Not isApproved Then
        If isRejected = False Then
            txtCPCNNO.Text = ""
            txtCPCNlocate.Text = ""
            txtCPCNNO.BackColor = &H8000000F
            txtCPCNlocate.BackColor = &H8000000F
            cmbBOMVersion.BackColor = &H8000000F
            txtCPCNNO.Enabled = False
        Else
            txtCPCNNO.BackColor = &HFFF
            txtCPCNlocate.BackColor = &HFFFFFF
            cmbBOMVersion.BackColor = &HFFFFFF
            txtCPCNNO.Enabled = True
        End If
    Else
        txtCPCNNO.BackColor = &HFFF
        txtCPCNlocate.BackColor = &HFFFFFF
        cmbBOMVersion.BackColor = &HFFFFFF
        txtCPCNNO.Enabled = True
    End If
    
    cmbBOMVersion.Text = CurVersion
    LastVersion = CurVersion
    '�ж�BOM��¼�Ƿ�Ǽǲ����Ѿ���׼
    rs.Open "Select * from BOMSubmitApprove Where FinsGdIndex ='" & Trim(FinishGoodsNO) & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If Trim(rs("Approver")) <> "" Then
            ApprovalStatus = True
        Else
            ApprovalStatus = False
        End If
        'ͬʱ�ж�BOM�����Ƿ�BOM����(�ύ��)
        If InStr(Trim(rs("Submiter")), PDMUserName) Then
            OpennerSubmiter = True
        Else
            OpennerSubmiter = False
        End If
    End If
    
    If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    'Call DropTempTable '####ɾ����ʱ��####�����溯��ɾ���ˣ��ظ�#####
    Call buildInit4Version '#########�����µ���ʱ��########

    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    Refresh_FlexGrid_TreeView False
    
    If temp_tb_SglPrt4BOMLog <> "sglprt4bomlog" Then
    
        StrSql = "SELECT * FROM SglPrt4BOMLog  WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            StrSql = "INSERT INTO " & temp_tb_SglPrt4BOMLog & " SELECT * FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
            Conn.Execute StrSql
            rs.Close
        Else
        '##############�ѱ������д����־��ʱ��##################
            With MSFlexGrid1
                For i = 2 To .Rows - 2
                    If Trim(.TextMatrix(i, 2)) <> "" Then
                        '�������µ��޸���־
                        StrSql = "IF NOT EXISTS(SELECT * FROM  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CStr(CurVersion) & " And ParentID='" & .TextMatrix(i, 2) & "' And ChildID='" & .TextMatrix(i, 3) & "' And Family='" & .TextMatrix(i, 11) & "' And ChgStatus='" & .TextMatrix(i, 9) & "') "
                        StrSql = StrSql & "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,CommtNote,Family) Values("
                        StrSql = StrSql & "" & FinishGoodsNO
                        StrSql = StrSql & "," & i + J
                        StrSql = StrSql & "," & .TextMatrix(i, 2)
                        StrSql = StrSql & "," & .TextMatrix(i, 3)
                        StrSql = StrSql & "," & CStr(CurVersion)
                        StrSql = StrSql & ",'" & .TextMatrix(i, 4)
                        StrSql = StrSql & "','" & .TextMatrix(i, 5)
                        StrSql = StrSql & "','" & Replace(.TextMatrix(i, 6), "'", "''")
                        StrSql = StrSql & "','" & .TextMatrix(i, 7)
                        StrSql = StrSql & "','" & .TextMatrix(i, 8)
                        StrSql = StrSql & "','" & .TextMatrix(i, 10)
                        StrSql = StrSql & "','" & .TextMatrix(i, 11) & "')"
                        Conn.Execute StrSql
                    End If
                Next i
            End With
        End If
    End If
    '##########��֤BOM��û�б���ס##########
    If IsBOMLocked Then
        mnuAddCode.Enabled = False
        mnuPaste.Enabled = False
        mnuCopy.Enabled = False
        mnuDeleteCode.Enabled = False
        mnuUpgradeVer.Enabled = False
        mnuRename.Enabled = False
        txtCPCNNO.Enabled = False
        
        cmdBOMSave.Enabled = False
        CmdImportBOM.Enabled = False
        CmdExportBOM.Enabled = False
        CmdDrwPathAdd.Enabled = False
        CmdSERPathAdd.Enabled = False
        
        lblAlert.Caption = "This Bom is being locked by " & getBOMLocker & ", you can't edit it before unlock."
        lblAlert.ForeColor = &HC0C0FF
        'Msflexgrid������click�¼���
    Else
        mnuAddCode.Enabled = True
        mnuPaste.Enabled = True
        mnuCopy.Enabled = True
        mnuDeleteCode.Enabled = True
        mnuUpgradeVer.Enabled = True
        mnuRename.Enabled = True
        txtCPCNNO.Enabled = True
        
        cmdBOMSave.Enabled = True
        CmdImportBOM.Enabled = True
        CmdExportBOM.Enabled = True
        CmdDrwPathAdd.Enabled = True
        CmdSERPathAdd.Enabled = True
    End If
    
    '############��������############
    If BOMLock Then
        Shape3.BackColor = &HFFC0C0
        cmdLock.Caption = "LOCKED"
    Else
        Shape3.BackColor = &HFFC0C0
        cmdLock.Caption = "UNLOCK"
    End If
    Me.Enabled = True
    frameMsg.Visible = False
    Exit Sub

vbErrorHandler:
        Me.Enabled = True
        frameMsg.Visible = False
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:Command1_Click"
End Sub

Private Sub Command3_Click()
    Set FromForm = FrmBOMAdmin
    FrmStdPrtLibStructr.Show 0
End Sub


Private Sub mnuCopy_Click()
    '����code������code�Լ��ṹ
    Dim oclick As Long
    Dim oNode As Node
    sChilds = ""
'    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
'    If SystemAdmin = "Y" Or OpennerSubmiter Then
        'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
        
        If tvCodeItems.SelectedItem.Key = "ROOT" Then
            If MsgBox("Would you like to copy all children under the ROOT ?", vbYesNo) = vbYes Then
                bCopyRoot = True
            Else
                Exit Sub     '����Ǹ��ڵ����������
            End If
        End If
        
        Set oNode = tvCodeItems.SelectedItem
        Set CurNode = tvCodeItems.SelectedItem
        
        Set CopyNodeSource = tvCodeItems.SelectedItem
        
        mnuPaste.Enabled = True
        mnuCopy.Enabled = False
        mnuUncopy.Enabled = True
        
        If Not bCopyRoot Then
            traval oNode
            sChilds = oNode.Parent & sChilds
        End If
'    Else
'
'
'        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
'        Exit Sub
'
'    End If
End Sub
Public Sub CopySubNode(ByVal SubNode As String, ByVal ParentCode As String)
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    '�������ڵ���ӽڵ㸴�Ƶ������ڵ���
    StrSql = "Select * from BOMOrigData Where ParentID ='" & Trim(SubNode) & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            '���ӹ�ϵ
            StrSql = "INSERT INTO BOMOrigData(ParentID,ChildID,Quantity,ChgStatus) VALUES(" & ParentCode & "," & rs.Fields("ChildID") & ",1,'Add')"
            Conn.Execute StrSql
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    IsCopy = False
    CopyNodeSource = ""
    PasteNodeTarget = ""
    sChilds = ""
    mnuPaste.Enabled = False
    mnuUncopy.Enabled = False
    Refresh_FlexGrid_TreeView False
End Sub
Public Sub CopyNodeData(ByVal NewCode As String)
    On Error GoTo vbErrorHandler
    Dim rs As New ADODB.Recordset
'    '�����µĸ��Ӷ�Ӧ��ϵ
'    StrSql = "INSERT INTO BOMOrigData(ParentID,ChildID,Quantity,ChgStatus) VALUES(" & PasteNodeTarget & "," & NewCode & ",1,'Add')"
'    Conn.Execute StrSql
    StrSql = "Select * From " & temp_tb_BOMOrigData & " Where ParentID='" & CopyNodeSource & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Conn.Execute "IF NOT EXISTS(SELECT * FROM " & temp_tb_BOMOrigData & "  Where ParentID='" & NewCode & "' And ChildID='" & rs.Fields("ChildID") & "') INSERT INTO " & temp_tb_BOMOrigData & "  (ParentID, ChildID, Quantity,ChgStatus) VALUES ('" & NewCode & "','" & rs.Fields("ChildID") & "','" & rs.Fields("Quantity") & "','Add')"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing

          
    InsertBOMLog4Copy NewCode, MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1, GetParent(PasteNodeTarget.Key)
    IsCopy = False
    mnuPaste.Enabled = False
    mnuUncopy.Enabled = False
    Refresh_FlexGrid_TreeView False


    Exit Sub
vbErrorHandler:

    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmBOMAdmin:CopyNodeCode"
End Sub



Private Sub mnuPaste_Click()
    On Error Resume Next
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
        GoTo AdminGoAhead1
    Else
        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����

    
    
AdminGoAhead1:
    If Trim(CPCN) = "" And isApproved Then
        MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
        Exit Sub
    End If
    
    Set CurNode = tvCodeItems.SelectedItem
    Set PasteNodeTarget = tvCodeItems.SelectedItem
    OrientCurNodeKey = PasteNodeTarget.Key
    Action = "COPY"
    If bCopyRoot Then
        If PasteNodeTarget.Key = "ROOT" Then
        
            If PasteNodeTarget.Children > 1 Then
                MsgBox "Only Allow to Paste to New BOM.", vbInformation
                Exit Sub
            End If
            
            Dim rs As New ADODB.Recordset
            Conn.BeginTrans
            StrSql = "Select * From " & temp_tb_BOMOrigData & " Where ParentID='" & CopyNodeSource & "'"
            rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                    Conn.Execute "IF NOT EXISTS(SELECT * FROM " & temp_tb_BOMOrigData & " Where ParentID='" & PasteNodeTarget & "' And ChildID='" & rs.Fields("ChildID") & "') INSERT INTO " & temp_tb_BOMOrigData & " (ParentID, ChildID, Quantity,ChgStatus) VALUES ('" & PasteNodeTarget & "','" & rs.Fields("ChildID") & "','" & rs.Fields("Quantity") & "','Add')"
                    rs.MoveNext
                Loop
            End If
            rs.Close
            Set rs = Nothing
            If Err Then
                Conn.RollbackTrans
            Else
                Conn.CommitTrans
            End If
            Refresh_FlexGrid_TreeView False
            IsCopy = False
            Set CopyNodeSource = Nothing
            Set PasteNodeTarget = Nothing
            sChilds = ""
            bCopyRoot = False
            mnuCopy.Enabled = True
            mnuPaste.Enabled = False
            mnuUncopy.Enabled = False
        Else
            MsgBox "Only Allow to Paste ROOT to ROOT.", vbCritical
            Exit Sub
        End If
    Else

        If InStr(sChilds, left(CStr(PasteNodeTarget), 11)) > 0 Then
            MsgBox "Unable to copy under this code.", vbCritical
            IsCopy = False
            Set CopyNodeSource = Nothing
            Set PasteNodeTarget = Nothing
            sChilds = ""
            mnuCopy.Enabled = True
            mnuPaste.Enabled = False
            mnuUncopy.Enabled = False
            Exit Sub
        End If
        
        IsCopy = True
        FrmPaste.Visible = True
        txtNewCode.Text = PasteNodeTarget
    End If
End Sub

Private Function CheckIsChild(s As String, t As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    StrSql = "select childID from BOMOrigData where parentID =" & s
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            If CStr(rs.Fields("childID")) = t Then
                CheckIsChild = True
                Exit Function
            Else
                CheckIsChild rs.Fields("childID"), t
            End If
            rs.MoveNext
        Loop
    Else
        CheckIsChild = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub mnuuncopy_Click()
    sChilds = ""
    bCopyRoot = False
    mnuCopy.Enabled = True
    mnuPaste.Enabled = False
    mnuUncopy.Enabled = False
End Sub

Private Sub mnuUpgradeVer_Click()
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        End If
        
    
        Set CurNode = tvCodeItems.SelectedItem
        OrientCurNodeKey = CurNode.Key
        Action = "UPG"
        If tvCodeItems.SelectedItem.Key = "ROOT" Then
            MsgBox "The root node can't upgrade version."
            Exit Sub     '����Ǹ��ڵ����������
        End If
        
        '##############�����FG����������############
        If getIsFG(Replace(OrientCurNodeKey, "C", "")) Then Exit Sub
        
        '��ʼ�汾����������
        If CurVersion = 1 Then
            If checkModifyPermission(CurNode.Text) = False Then Exit Sub
        End If
        
        '����������ʾ
        FrmUpgrade.Visible = True
        txt12NC.Text = left(tvCodeItems.SelectedItem, 11) & "0"
        txtSglParent.Text = tvCodeItems.SelectedItem.Parent
        cmbSglVer1.Text = right(tvCodeItems.SelectedItem, 1)
        If cmbSglVer1.Text < 9 Then cmbSglVer2.Text = CInt(cmbSglVer1.Text) + 1
    Else

        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If

End Sub

Private Sub OKButton_Click()
    '�±��������12λ����
    If (Len(Trim(txtNewCode.Text)) <> 12 Or Not IsNumeric(txtNewCode.Text)) Then
        MsgBox ("The code MUST be made up of 12 numeric.")
    Else
        '�±�������Ǵ���û�г��ֹ���
        Dim Conn As New ADODB.Connection
        Dim StrSql As String
        Conn.Open connString
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        StrSql = "Select * from SglPrt where SglPrtIndex='" & left(txtNewCode.Text, 11) & "0'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The new code is not existing, please first apply the code", vbCritical, "ERP"
            Exit Sub
        Else
            Call CopyNodeData(left(txtNewCode.Text, 11) & CStr(rs("SglPrtVer")))
            mnuCopy = True
            mnuPaste = False
            mnuUncopy = False
            IsCopy = False
            FrmPaste.Visible = False
        End If
    End If
End Sub

Private Sub Text1_Click()
    If Not IsNumeric(Text1.Text) Then Text1.Text = ""
End Sub

Private Sub Text2_Click()
    If Not IsNumeric(Text2.Text) Then Text2.Text = ""
End Sub


Private Sub tmrDragTimer_Timer()
    Dim nHitNode As Node
    Static lCount As Long
    
    ' This timer has two functions :
    ' 1 - It will scroll the TreeView when the user is dragging
    ' 2 - It will auto-expand a node when the user drags over it for more than half a second.
    
    
    If SourceNode Is Nothing Then         '���û��ѡ�нڵ����ʱ��������
        tmrDragTimer.Enabled = False
        Exit Sub
    End If
    
    lCount = lCount + 1        '����ĳ��δչ���ڵ���ͣ��ʱ�䵽����(half a second)��չ��
    If lCount > 10 Then
        
        Set nHitNode = tvCodeItems.DropHighlight
        If nHitNode Is Nothing Then Exit Sub
        
        If nHitNode.Expanded = False Then
            nHitNode.Expanded = True
        End If
        lCount = 0
    End If
    
    If miScrollDir <> 0 Then
        If miScrollDir = -1 Then
            SendMessageLong tvCodeItems.hWnd, WM_VSCROLL, 0, 0  '���¹���
        Else
            SendMessageLong tvCodeItems.hWnd, WM_VSCROLL, 1, 0  '���Ϲ���
        End If
    End If
End Sub


'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'OLE �ϷŲ�����ʼʱ(Check OK)
Private Sub tvCodeItems_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Dim byt() As Byte
    ' Place the key of the dragged item into the clipboard in our own format declared in GetClipboardFormat api
    
    AllowedEffects = vbDropEffectMove    'VbDropEffectMove=2 �Ž��������Ҫ����Դ�Ƶ���Դ�������С��ƶ�����ԴҪɾ�����ݡ�
    
    If SourceNode Is Nothing Then Exit Sub
    
    byt = SourceNode.Key
    
    ' ��Formload�ж�������
    'miClipBoardFormat = RegisterClipboardFormat("VBCodeLibTree")
    
    Data.SetData byt, miClipBoardFormat    'SetData������ָ�������ݸ�ʽ�����ݲ��� DataObject ����
    '�﷨  object.SetData [data], [format]
    'data ��ѡ�ı����ͣ�����Ҫ���͵� DataObject ��������ݡ�
    'format  ��ѡ�ĳ�����ֵ���涨���������ݵĸ�ʽ���硰����ֵ����������
    
End Sub


'OLE �ϷŲ����ӹ��ڵ�ʱ(Check OK)
Private Sub tvCodeItems_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    'ComctlLib.TreeView 5.0 ǰ�����ComctlLib.   6.0 ǰ�����  'MSComctlLib.
    'object (tvCodeItems)������ʽ����ֵ�ǡ�Ӧ���ڡ��б��е�һ������
    'data DataObject ���󣬰���Դ�ṩ�ĸ�ʽ������Ҳ���ܰ�����Щ��ʽ�����ݡ��� DataObject ���������ݣ��򵱿ؼ����� GetData ����ʱ�ṩ���ݡ�SetData �� Clear ���������������
    'effect Դ����������õĳ�������������ʶ����֧�ֵ�Ч�������¼������У��˲������뱻Ŀ�겿����ȷ�����á�effect ֵ�����л��Ч����������ֵ���߼� Or ȷ����Ŀ�겿��Ӧ�����ЩЧ���Լ�����������ȷ���ĸ������ʺ�������Ȼ��Ѵ˲�������Ϊ�����Ч��֮һ����Դ���涨�����Ա�ȷ������ѡ�������ִ���ĸ����������ܵ�ȡֵ���ڡ�����ֵ���С�
    'button ����������������ʱ�������״̬���Ӧ�����Ϊλ 0���Ҽ�Ϊλ 1���м�Ϊλ 2����Щλ��Ӧ��ֵ�ֱ�Ϊ 1��2 �� 4����������������״̬������������λ�еĲ��֡�ȫ������������ã���Ӧ�ر������֡�ȫ�����������»�û�а������¡�
    'shift ������������ SHIFT��CTRL �� ALT ��ʱ������Щ��״̬���Ӧ��SHIFT ��Ϊλ 0��CTRL ��Ϊλ 1��ALT ��Ϊλ 2����Щλ��Ӧ��ֵ�ֱ�Ϊ 1��2 �� 4��shift ������������Щ����״̬������������λ�еĲ��֡�ȫ������������ã���Ӧ�ر������֡�ȫ�����������»�û�а������¡����磬ͬʱ���� CTRL �� ALT ����shift ֵΪ 6��
    'x,y ��Ŀ�괰���ؼ��У��涨��ǰ���ָ��ˮƽx�ʹ�ֱyλ�õ�����x��yֵ�ɶ���� ScaleHeight��ScaleWidth��ScaleLeft �� ScaleTop �������õ�����ϵͳ��ʽ����ʾ��
    'state ��������Ӧ�ڿؼ���ת��״̬���˿ؼ������Ϸŵ�������ص�Ŀ�괰���ؼ��С����ܵ�ȡֵ���ڡ�����ֵ���С�
    
    'effect ��������
    '   ����         ֵ                           ����
    'vbDropEffectNone 0                           ��Ŀ�겻�������ݡ�
    'VbDropEffectCopy 1                           �Ž�������ڴ�Դ��Ŀ������ݿ����С���ʼ����û�б��ϷŲ����ı䡣
    'VbDropEffectMove 2                           �Ž��������Ҫ���Ϸ�Դ�Ƶ���Դ�������С��ƶ����Ϸ�ԴҪɾ�����ݡ�
    'vbDropEffectScroll -2147483648#(&H80000000) ��Ŀ�겿���У��������ڻ�Ҫ��������ֵ������ֵ��ͬʹ�á�ע�� �����ڲ�����ִ���Լ��Ĺ���ʱ����Ӧ�á�
    
    Dim sTmpStr As String
    Dim nTargetNode As Node
    Dim highlight As Boolean  '�ñ���������ҷ��Ŀ��ڵ��Ƿ���Ч
    
    On Error Resume Next
    sTmpStr = Data.GetFormat(miClipBoardFormat)   'GetFormat ����,��� DataObject����(Data As MSComctlLib.DataObject)�е�����涨��ʽƥ�䣬GetFormat �������� True�����򷵻� False��
    ' First check that we allow this type of data to be dropped here
    If Err Or sTmpStr = "False" Then
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    Set nTargetNode = tvCodeItems.HitTest(x, y)   '���������ָλ�õĽڵ㸳ֵ��һ����ʱĿ��ڵ����nTargetNode
    
    If nTargetNode.Key = SourceNode.Key Then
        Set tvCodeItems.DropHighlight = Nothing
        Effect = vbDropEffectNone
    Else
        Set targetNode = nTargetNode
    End If
    
    highlight = True
    If Not targetNode Is Nothing Then  'And Not SourceNode Is Nothing
        '�������¼�������ſ���ҷ��
        '1��Դ�ڵ㲻����Ŀ��ڵ�
        '2��  Դ�ڵ㲻��Ŀ��ڵ��ȫ������ǰ���ڵ� Ҳ����˵���ܰ�һ���ڵ��ϵ������漶��Ľڵ���
        '3��Դ�ڵ㲻��Ŀ��ڵ� �Ķ��Ӽ����ӽڵ� (�Ѿ����Ӽ�¼������Ӿͻ���)
        '4��Ŀ��ڵ㲻��Դ�ڵ� ��ȫ�������ӽڵ�  (���Դ�ڵ��Ŀ��ڵ���ͬһ����֦�Ļ�����1��2���Ѿ��жϳ����� ����ר�����ڲ���ͬһ����֦�����)
        'Ŀ��ڵ㲻��Դ�ڵ� ��ȫ������ǰ���ڵ�(���ڲ���ͬһ����֦�����) �������������֣���Ϊ����ˢ��(Refresh_FlexGrid_TreeView)BOM��һ������ͬһ��֦��
        '5��Ŀ��ڵ�ֱ��Root������ǰ���ڵ㲻��Դ�ڵ��ȫ�������ӽڵ�
        '6��Ŀ��ڵ��Դ�ڵ�ĵ�һ���ӽڵ㲻����ͬ���� (�����ͬ���Ļ����ϲ�������ͬ�ֵ�������BOM����)
        If targetNode = SourceNode Then      '1��Դ�ڵ㲻����Ŀ��ڵ�      AddNodeChildIsBrothership
            highlight = False
            GoTo highlightGotvalue
        End If
        RecursionFlag = False
        If isEldershipNode(SourceNode, targetNode) Then     '2��Դ�ڵ㲻��Ŀ��ڵ��ȫ������ǰ���ڵ�
            highlight = False
            GoTo highlightGotvalue
        End If
        RecursionFlag = False
        If isSonshipNode(targetNode, SourceNode) Then   '3��Դ�ڵ㲻��Ŀ��ڵ�Ķ��Ӽ����ӽڵ�
            highlight = False
            GoTo highlightGotvalue
        End If
        RecursionFlag = False
        If isYoungershipNode(SourceNode, targetNode) Then   '4��Ŀ��ڵ㲻��Դ�ڵ��ȫ�������ӽڵ�
            highlight = False
            GoTo highlightGotvalue
        End If
        RecursionFlag = False
        If isElderYoungershipNode(SourceNode, targetNode) Then   '5��Ŀ��ڵ�ֱ��Root������ǰ���ڵ㲻��Դ�ڵ��ȫ�������ӽڵ�
            highlight = False
            GoTo highlightGotvalue
        End If
        RecursionFlag = False
        If AddNodeChildIsBrothership(SourceNode, targetNode) Then  '6��Ŀ��ڵ��Դ�ڵ�ĵ�һ���ӽڵ㲻����ͬ����
            highlight = False
            GoTo highlightGotvalue
        End If
    End If
    
highlightGotvalue:
    If highlight Then
        '��ҷ��Ч��Ŀ��ڵ�ͻ����ʾ��������ʾ��
        'DropHighlight ����һ�����ϷŲ������� HitTest �������á��ڹ���϶���Node������ʱ��HitTest�������ض��κα��ϵ��Ķ�������á����ţ�DropHighlight ���Ա�����Ϊ���еĶ���ͬʱ���ض��Ǹ���������á�Ȼ�� DropHighlight ���Ծ���ϵͳͻ����ɫͻ�����еĶ���
        Set tvCodeItems.DropHighlight = targetNode
    Else
        Set tvCodeItems.DropHighlight = Nothing
    End If
    
    
    If y > 0 And y < 300 Then
        miScrollDir = -1
    ElseIf (y < tvCodeItems.Height) And y > (tvCodeItems.Height - 500) Then
        miScrollDir = 1
    Else
        miScrollDir = 0
    End If
    
End Sub

'OLE �ϷŲ�������ʱ(Check OK)
Private Sub tvCodeItems_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        Else
            GoTo AdminGoAhead1
        End If
    Else
        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    
AdminGoAhead1:
    If Not (tvCodeItems.DropHighlight Is Nothing) Then  '��ͻ����ʾ�Ľڵ㣬��Ŀ��ڵ���ҷ��Чʱ
        If MsgBox("Ensure Drag and Drop here��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
            SourceNodeParent = SourceNode.Parent               '������ԭ����Դ�ڵ�ĸ��ڵ�û�ı�֮ǰ��ֵ,SourceNode.Parent��Դ�ڵ�ĸ��ڵ���(������ǰ��C��12NC)
            Set SourceNode.Parent = tvCodeItems.DropHighlight  'TreeView��,��Դ�ڵ�ĸ��ڵ�ı�Ϊ���������ָĿ��ڵ� tvCodeItems.DropHighlight��Drop��������ʾ�Ľڵ���(������ǰ��C��12NC)
            DragSave tvCodeItems.DropHighlight.Key, SourceNode.Key  '������ҷ�����ݿ��еı䶯
        End If
        Set tvCodeItems.DropHighlight = Nothing  'ȡ��ͻ����ʾ
    End If
    ChgMass = True
    Set SourceNode = Nothing
    Set targetNode = Nothing
End Sub

Private Sub tvCodeItems_OLECompleteDrag(Effect As Long)  '(Check OK)
    'OLECompleteDrag �¼��� OLE �ϷŲ��������õ��¼��������󱻷ŵ�Ŀ�겿��ʱ�����¼�֪ͨԴ������ִ�еĶ�����
    'Ŀ��ͨ�� OLEDragDrop �¼��� effect �������ô�ֵ�����ڴˣ�Դ�ɾ������ȡ���ʵ����������磬�������Ƶ�Ŀ�� (vbDropEffectMove)���ƶ���Դ��Ҫ������ɾ���ö���
    Screen.MousePointer = vbDefault
    tmrDragTimer.Enabled = False
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


'�������ĳ�ڵ�ʱ���ڴ�������ʾ�ýڵ��ֵ (Check OK)
Private Sub tvCodeItems_NodeClick(ByVal myNode As Node)     'myNode�Ǵ�tvCodeItems_MouseDown���͹����ڵ���(������ǰ��C��12NC)
    Dim r As Integer       'MSFlexGrid1����ѭ������
    Dim P As Integer       'myNode.keyǰ����N��C��ѭ������
    Dim k As Integer       'N��C��N�������ѭ������
    Dim tempNodekey As String
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    txtNodeSglPrt12NC = ""            '�����ԭ��������
    txtNodeDescription = ""
    txtNodePrtUnit = ""
    txtNodeDrwlocate = ""
    txtSERNO = ""
    txtSERlocate = ""
    If myNode.Index = 1 Then                'myNode.Index = 1 ��ʾ��ȡ���Ǹ��ڵ�
        tempNodekey = myNode
        rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(tempNodekey) & "'", Conn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            txtNodeSglPrt12NC = rs("FinsGdIndex")
            txtNodeDescription = Trim(rs("Description")) & ""
            txtNodePrtUnit = "Piece"
            If IsNull(rs("Drwlocate")) Then
                txtNodeDrwlocate = ""
            Else
                txtNodeDrwlocate = Trim(rs("Drwlocate")) & ""
                Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & Trim(tempNodekey) & " AND ITEMVALUE='" & Trim(rs("Drwlocate")) & "' AND CREATOR='drwlocate') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & Trim(tempNodekey) & ",'" & Trim(rs("Drwlocate")) & "','drwlocate')")
            
            End If
            
            If IsNull(rs("SERNmbr")) Then
                txtSERNO = ""
            Else
                txtSERNO = rs("SERNmbr") & ""
            End If
            
            If IsNull(rs("SERlocate")) Then
                txtSERlocate = ""
            Else
                txtSERlocate = Trim(rs("SERlocate")) & ""
                Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & Trim(tempNodekey) & " AND ITEMVALUE='" & Trim(rs("SERlocate")) & "' AND CREATOR='SERlocate') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & Trim(tempNodekey) & ",'" & Trim(rs("SERlocate")) & "','SERlocate')")

            End If
            
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
        
    Else
        
        tempNodekey = LeftcutStrg(myNode.Key)     'myNode.Key�Ǵ�tvCodeItems_MouseDown���͹����ڵ�key(ǰ�����ַ�Cϵ��)LeftcutStrgȥ��������ַ�Cϵ��
        tempNodekey = Mid(tempNodekey, 1, (Len(tempNodekey) - 1)) & "0"
        rs.Open "Select * from SglPrt where SglPrtIndex ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            
            If rs("SglPrtVer") <> val(right(myNode.Key, 1)) Then    '������Ҫ��һ��val����������Ⱥ���������֣��ұ����ַ������ǲ����
                If MsgBox("Version in Single Part Database is not same as Version in BOM Database" & vbCrLf & "Do you want to align Version same as BOM��", vbYesNo + vbInformation, "SystemInfo") = vbYes Then
                    rs("SglPrtVer") = CInt(right(myNode.Key, 1))
                    rs.Update
                End If
            End If
            
            txtNodeSglPrt12NC = left(rs("SglPrtIndex"), 11) & CStr(CInt(right(rs("SglPrtIndex"), 1)) + rs("SglPrtVer"))
            txtNodeDescription = rs("Description") & ""
            txtNodePrtUnit = rs("PrtUnit") & ""
            
            If IsNull(rs("Drwlocate")) Then
                txtNodeDrwlocate = ""
            Else
                txtNodeDrwlocate = Trim(rs("Drwlocate")) & ""
                Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & Trim(tempNodekey) & " AND ITEMVALUE='" & Trim(rs("Drwlocate")) & "' AND CREATOR='drwlocate') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & Trim(tempNodekey) & ",'" & Trim(rs("Drwlocate")) & "','drwlocate')")

            End If
            
            If IsNull(rs("SERNmbr")) Then
                txtSERNO = ""
            Else
                txtSERNO = rs("SERNmbr") & ""
            End If
            
            If IsNull(rs("SERlocate")) Then
                txtSERlocate = ""
            Else
                txtSERlocate = Trim(rs("SERlocate")) & ""
                Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & Trim(tempNodekey) & " AND ITEMVALUE='" & Trim(rs("SERlocate")) & "' AND CREATOR='SERlocate') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & Trim(tempNodekey) & ",'" & Trim(rs("SERlocate")) & "','SERlocate')")

            End If
            
            
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Set rs = Nothing
    End If
    
    tempNodekey = myNode    '����Ҫ�������Ϊǰ����tempNodekey = Mid(tempNodekey, 1, (Len(tempNodekey) - 1)) & "0",�������һλ��Ϊ0
    If myNode.Key = "ROOT" Then
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 0                         '�ӵ�R�е�0�п�ʼ
        MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1   '����ѡ��ֱ�����һ��
        Exit Sub
    Else
        P = Len(myNode.Key) - Len(LeftcutStrg(myNode.Key))   '�ҳ�myNode.Keyǰ���м���C
        k = 0
        For r = 1 To RowNum
            If MSFlexGrid1.TextMatrix(r, 3) = tempNodekey Then
                k = k + 1
                If k = P Then
                    MSFlexGrid1.Row = r
                    MSFlexGrid1.Col = 0                         '�ӵ�R�е�0�п�ʼ
                    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1   '����ѡ��ֱ�����һ��
                    Exit Sub
                End If
            End If
        Next r
    End If
    Err.Clear
End Sub

'��������갴��ʱ��ȡ��Դ�ڵ� (Check OK)
Private Sub tvCodeItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����ָ���µĶ���ֵ��Դ�ڵ�
    Set SourceNode = tvCodeItems.HitTest(x, y)  'HitTest����������������ض�λ��x��y�����Node���������
    '�ڴ�������ʾ�ýڵ��ֵ
    If Not (SourceNode Is Nothing) Then             'SourceNode�õ����ǽڵ���(ȥ��ǰ��C��12NC)
        Call tvCodeItems_NodeClick(SourceNode)
    End If
    tvCodeItems.SelectedItem = SourceNode
    tvCodeItems.DropHighlight = SourceNode
End Sub

'���ɿ���갴��ʱ (Check OK)
Private Sub tvCodeItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sKey As String
    Dim bIsRoot As Boolean
    
    ' Show Popup Menu
    
    If Button = vbRightButton Then
        '�ж�����Ǹ��ڵ�Ļ��������ɾ��������, tvCodeItems.SelectedItem�ǵ��еĽڵ���(������ǰ��C��12NC)
        If tvCodeItems.SelectedItem Is Nothing Then Exit Sub
        bIsRoot = (StrComp(tvCodeItems.SelectedItem.Key, "ROOT", vbTextCompare) = 0) 'vbTextCompare ֵΪ1ִ��һ������ԭ�ĵıȽϡ�string1 ���� string2����ֵΪ0
        If Not IsBOMLocked Then
            mnuRename.Enabled = Not (bIsRoot)
            mnuDeleteCode.Enabled = Not (bIsRoot)
        End If
        PopupMenu mnuEdit
    End If
    
End Sub

'���ƶ����ʱ (Check OK)
Private Sub tvCodeItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If SourceNode Is Nothing Then Exit Sub
    
    If Button = vbLeftButton Then
        If SourceNode.Key <> "ROOT" Then
            ' Start Dragging !
            Set tvCodeItems.SelectedItem = SourceNode
            tmrDragTimer.Interval = 100
            tmrDragTimer.Enabled = True
            'tvCodeItems.DragIcon = ImageList1.ListImages(10).Picture        '��Ϊʵ��ʹ�õ���OLEDrag,���Ա������ͨDragû��
            tvCodeItems.OLEDrag              '������ OLEDrag ����ʱ�������� OLEStartDrag �¼�������������Ŀ�겿���ṩ����
        End If
    Else
        Set SourceNode = Nothing
        Set targetNode = Nothing
        Set tvCodeItems.DropHighlight = Nothing         'ȡ��Drop��ͻ����ʾ
    End If
    
End Sub


'������ҷ������� (Check OK)
Private Sub DragSave(ParentNodeKey As String, ChildNodeKey As String)
    On Error GoTo vbErrorHandler
    
    
    Dim rs As New ADODB.Recordset
    
    
    '����Դ�ڵ�ĸ��ڵ�·�������ڵ�   ������ק��Դ�ڵ�����ChildNodeKey�ͱ���ĸ��ڵ�������BOMOrigData�ҳ�һ����¼��
    rs.Open "Select * from BOMOrigData Where ChildID='" & LeftcutStrg(ChildNodeKey) & "'" & " and  ParentID ='" & SourceNodeParent & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        Exit Sub
    End If
    If ParentNodeKey = "ROOT" Then ParentNodeKey = tvCodeItems.Nodes("ROOT")
    rs("ParentID") = LeftcutStrg(ParentNodeKey)      '���ҳ�������һ����¼�еĸ��ڵ�������Ŀ��ڵ�tvCodeItems.DropHighlight�����滻��
    rs.Update
    rs.Close
    
    Set rs = Nothing
    Refresh_FlexGrid_TreeView False
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:DragSave"
    
End Sub


'ɾ��һ���ڵ�Ĳ���(Check OK)
Private Sub mnuDeleteCode_Click()
    ' Delete the selected CodeItem and all it's children
    
    On Error GoTo vbErrorHandler
    Dim LastRcd As Boolean
    Dim sKey As String
    Dim oNode As Node
    Dim oParentNode As Node
    Dim sMessage As String
    Dim oWait As CWaitCursor
    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
        GoTo AdminGoAhead1
    Else
        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    
    
    
AdminGoAhead1:
    If Trim(CPCN) = "" And isApproved Then
        MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
        Exit Sub
    End If
    

    Set oNode = tvCodeItems.SelectedItem
    Set CurNode = tvCodeItems.SelectedItem
    OrientCurNodeKey = CurNode.Key
    OrientParentNodeKey = CurNode.Parent.Key
    Action = "DEL"
    sKey = oNode
    

    
    If sKey = "InputNew12NC" Or sKey = "CInputNew12NC" Then      '�����InputNew12NC,��ʾʵ��ֻ����ɾ��һ��treeview����ʱ�ڵ�,��BOM�и����޼�¼��
        Refresh_FlexGrid_TreeView False
        Exit Sub
    End If
    
    If oNode.Key = "ROOT" Then Exit Sub     '����Ǹ��ڵ����˳�ɾ������
    
    If oNode.Parent.Key = "ROOT" And oNode.Parent.Children = 1 Then       '����ǽڵ㸸�ڵ���ROOT����ROOTֻ��һ��Child(��ѡ�е�����ڵ�)
        MsgBox "Delete Final Record in BOM, BOM will not Exist", vbInformation, "System Info"
        LastRcd = True
        GoTo DeleteGoAhead
    End If
    
    BOMString = ""
    'Screen.MousePointer = 11
    Call GetTopBOM(CurNode.Parent.Text)
    arrBOM = Split(Mid(BOMString, 2), ",")
    If BOMString <> "" Then
        'msgbox ���ʾ1024�ַ�
        'MsgBox BOMString
        If UBound(arrBOM) > 0 Then
            If CurVersion = 1 Then
                If CheckIsApprovalForAll(arrBOM) Then
                    MsgBox "The Assembly Part can't change, because it used to other formal BOMs.", vbCritical
                    Exit Sub
                End If
            End If
            
            If MsgBox("The 12NC had used in the following BOMs: " & vbCrLf & vbCrLf & Mid(BOMString, 2) & vbCrLf & vbCrLf & "Are you sure to delete it?", vbYesNo) = vbYes Then
                If oNode.Children > 0 Then    '����нڵ�ѡ�в������ӽڵ���Ҫ�в�ͬ��ʾ��Children�ǽڵ���ӽڵ�����
                    sMessage = sMessage & "The code includes the children, Are you sure to delete it?"
                    If MsgBox(sMessage, vbYesNo) = vbYes Then
                        NotDeleteChildTree = False
                        GoTo DeleteGoAhead
                    Else
                        Exit Sub
                    End If
                Else
                    sMessage = "No child, Are you sure to delete it?"
                    If MsgBox(sMessage, vbYesNo) = vbYes Then
                        NotDeleteChildTree = True
                        GoTo DeleteGoAhead
                    Else
                        Exit Sub
                    End If
                End If
            Else
                'Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            If oNode.Children > 0 Then    '����нڵ�ѡ�в������ӽڵ���Ҫ�в�ͬ��ʾ��Children�ǽڵ���ӽڵ�����
                sMessage = sMessage & "The code includes the children, Are you sure to delete it?"
                If MsgBox(sMessage, vbYesNo) = vbYes Then
                    NotDeleteChildTree = False
                    GoTo DeleteGoAhead
                Else
                    Exit Sub
                End If
            Else
                sMessage = "No child, Are you sure to delete it?"
                If MsgBox(sMessage, vbYesNo) = vbYes Then
                    NotDeleteChildTree = True
                    GoTo DeleteGoAhead
                Else
                    Exit Sub
                End If
            End If
        End If
        'Screen.MousePointer = 0
    End If
DeleteGoAhead:
    Set oParentNode = oNode.Parent  '��ǰѡ�е�Ҫɾ���Ľڵ�ĸ��ڵ㸳ֵ
    SourceNodeParent = oParentNode
    
    Set oWait = New CWaitCursor
    oWait.SetCursor
    
    DeleteCodeItem SourceNodeParent, sKey
    
    If Not LastRcd Then          '����������һ���ڵ���Refresh_FlexGrid_TreeView
        ChgMass = True
        Refresh_FlexGrid_TreeView False
        Set oWait = Nothing
    Else
        tvCodeItems.Nodes.Clear      '��������һ���ڵ����������TreeView������
    End If
    Exit Sub

vbErrorHandler:
        Set oWait = Nothing
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:DeleteCodeItem"
End Sub

'ɾ��һ���ڵ�Ĳ���(Check OK)
Private Sub DeleteCodeItem(ParentNodeKey As String, ChildNodeKey As String)
    
    On Error GoTo vbErrorHandler
    Dim rs As New ADODB.Recordset
    Dim StrSql As String
    Dim P, k As Integer

    
    If NotDeleteChildTree Then
        'ɾ��Դ�ڵ㱾�����һ������, ��ɾ�����
        If IsNumeric(ChildNodeKey) Then
            '������־

            If CurVersion > 1 Then
            
'                'ɾ���ϼ��޸ĵľ���־
'                StrSql = "Delete From " & temp_tb_SglPrt4BOMLog & " Where BOM=" & FinishGoodsNO
'                StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
'                StrSql = StrSql & " And Left(ChildID,11)='" & left(ChildNodeKey, 11) & "'" '����
'                StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                StrSql = "SELECT * FROM " & temp_tb_SglPrt4BOMLog & " Where BOM=" & FinishGoodsNO
                StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
                StrSql = StrSql & " And Left(ChildID,11)='" & left(ChildNodeKey, 11) & "'"
                StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                StrSql = StrSql & " And chgStatus not like 'Delete%'"
                rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Or Not rs.BOF Then
                    StrSql = "UPDATE " & temp_tb_SglPrt4BOMLog & " SET chgStatus='Delete-'+chgStatus Where BOM=" & FinishGoodsNO
                    StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
                    StrSql = StrSql & " And Left(ChildID,11)='" & left(ChildNodeKey, 11) & "'" '����
                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                    StrSql = StrSql & " And chgStatus not like 'Delete%'"
                    Conn.Execute StrSql
                        
                Else
                
                    StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,CPCN) Values("
                    StrSql = StrSql & FinishGoodsNO
                    StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1
                    StrSql = StrSql & ",'" & ParentNodeKey
                    StrSql = StrSql & "','" & ChildNodeKey
                    StrSql = StrSql & "'," & CurVersion
                    StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4)
                    StrSql = StrSql & ",'" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 5)
                    StrSql = StrSql & "','" & Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 6), "'", "''")
                    StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 7)
                    StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 8)
                    StrSql = StrSql & "','" & "Delete"
                    StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 10)
                    StrSql = StrSql & "','" & txtCPCNNO.Text & "')"
                    Conn.Execute (StrSql)
                End If
                rs.Close
            End If
            
            StrSql = "Delete  " & temp_tb_BOMOrigData & "    Where ChildID='" & ChildNodeKey & "' and  ParentID ='" & ParentNodeKey & "'"
            Conn.Execute StrSql
        End If
    Else
        '������ɾ��Դ�ڵ���ӽڵ�����
        If ParentNodeKey = SourceNodeParent Then
            StrSql = "Delete " & temp_tb_BOMOrigData & "  Where ChildID='" & ChildNodeKey & "' and  ParentID ='" & ParentNodeKey & "'"
            Conn.Execute StrSql
        End If
        'Դ�ڵ��ӽڵ���ɾ�����
        'ɾ���ϼ��޸ĵľ���־
'        StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO
'        StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
'        StrSql = StrSql & " And ChildID='" & ChildNodeKey & "'"
'        StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
        StrSql = "SELECT * FROM " & temp_tb_SglPrt4BOMLog & " Where BOM=" & FinishGoodsNO
        StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
        StrSql = StrSql & " And Left(ChildID,11)='" & left(ChildNodeKey, 11) & "'"
        StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
        StrSql = StrSql & " And chgStatus not like 'Delete%'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Or Not rs.BOF Then
            StrSql = "UPDATE " & temp_tb_SglPrt4BOMLog & " SET chgStatus='Delete-'+chgStatus Where BOM=" & FinishGoodsNO
            StrSql = StrSql & " And ParentID='" & ParentNodeKey & "'"
            StrSql = StrSql & " And ChildID='" & ChildNodeKey & "'" '����
            StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
            StrSql = StrSql & " And chgStatus not like 'Delete%'"
            Conn.Execute StrSql
        Else
            StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,CPCN) Values("
            StrSql = StrSql & FinishGoodsNO
            StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1
            StrSql = StrSql & ",'" & ParentNodeKey
            StrSql = StrSql & "','" & ChildNodeKey
            StrSql = StrSql & "'," & CurVersion
            StrSql = StrSql & "," & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4)
            StrSql = StrSql & ",'" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 5)
            StrSql = StrSql & "','" & Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 6), "'", "''")
            StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 7)
            StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 8)
            StrSql = StrSql & "','" & "Delete"
            StrSql = StrSql & "','" & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 10)
            StrSql = StrSql & "','" & txtCPCNNO.Text & "')"
            Conn.Execute (StrSql)
        End If
        rs.Close

        If rs.State = adStateOpen Then rs.Close
        StrSql = "Select * from  " & temp_tb_BOMOrigData & "   Where ParentID='" & ChildNodeKey & "'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                DeleteCodeItem ChildNodeKey, rs("ChildID")     '�ݹ�����ҳ����в㼶���ӽڵ�����
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    End If

    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:DeleteCodeItem"
End Sub


'����һ���ڵ�Ĳ���(Check OK)
Private Sub mnuAddCode_Click()
    On Error Resume Next
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        End If
        Set CurNode = tvCodeItems.SelectedItem
        
        BOMString = ""
        Call GetTopBOM(CurNode.Text)
        arrBOM = Split(Mid(BOMString, 2), ",")
        '��ʼ�汾�������޸�New Assembly
        If CurVersion = 1 Then
            If Not checkModifyPermission(CurNode.Text) Then Exit Sub
        Else
            If CurNode.Text <> FinishGoodsNO Then
                Screen.MousePointer = 11
                If BOMString <> "" Then
                    'msgbox ���ʾ1024�ַ�
                    If MsgBox("The Parent had used in the following BOMs: " & vbCrLf & vbCrLf & Mid(BOMString, 2) & vbCrLf & vbCrLf & "Are you sure to add it to the above BOMs?", vbYesNo) = vbNo Then
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                Screen.MousePointer = 0
            End If
        End If
        
        OrientCurNodeKey = CurNode.Key
        Action = "ADD"
        oldCode = "InputNew12NC"
        AddCode "InputNew12NC"
    Else
        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����

End Sub

Private Sub CopyCode()
    On Error GoTo vbErrorHandler
    Dim AddNodeParentImage As String
    Dim nNode As Node
    Dim nParentNode As Node
    Dim sParentKey As String
    
    Set nNode = tvCodeItems.SelectedItem   'TreeView��ѡ�еĽڵ㸳ֵ��nNode
    
    
    If nNode.Key <> "ROOT" Then            '�ж�������Ǹ��ڵ�Ļ�
        Set nParentNode = tvCodeItems.Nodes(nNode.Key)        'Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ㿪ʼ��һ�����ڵ�
        SourceNodeParent = nParentNode
        AddNodeParentImage = nParentNode.Image       '����Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ��ԭͼ��
        nParentNode.Image = "FOLDER"
        nParentNode.ExpandedImage = "FOLDER"
        'ExpandedImage���Է��ػ������ڹ�����ImageList�ؼ��е�ListImage������������ֵ����Node����չ��ʱ��ʾ ListImage ����
    End If
    
    Set nNode = tvCodeItems.Nodes.Add(tvCodeItems.SelectedItem, tvwChild, "CC" & CopyNodeSource, CopyNodeSource, "CHILD")
    Set tvCodeItems.SelectedItem = nNode       '�任ѡ�еĽڵ�(������ʾ)�ӱ����ӵĽڵ� ��Ϊ�ո���ӵĽڵ�
    nNode.EnsureVisible
    
    AddNodeOk = True                   '�ȼ����ǿ��Ա༭OK��
    tvCodeItems.StartLabelEdit    'StartLabelEdit���������û��༭��ǩ��
    ' �� LabelEdit ��������Ϊ 1���ֶ���ʱ�������� StartLabelEdit ����������һ��ǩ�༭������
    ' ��һ�����ϵ��� StartLabelEdit ����ʱ��BeforeLabelEdit �¼�Ҳͬʱ������
    
    If Not AddNodeOk Then
        tvCodeItems.Nodes.Remove nNode.Key
        nParentNode.Image = AddNodeParentImage
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmBOMAdmin:CopyCode"
    
End Sub


'����һ���ڵ�Ĳ���(Check OK)
Private Sub AddCode(NCode As String)
    
    On Error GoTo vbErrorHandler
    Dim AddNodeParentImage As String
    Dim nNode As Node
    Dim nParentNode As Node
    Dim sParentKey As String
    Dim NCodeText As String
    
    Set nNode = tvCodeItems.SelectedItem   'TreeView��ѡ�еĽڵ㸳ֵ��nNode
    
    If nNode.Key <> "ROOT" Then            '�ж�������Ǹ��ڵ�Ļ�
        Set nParentNode = tvCodeItems.Nodes(nNode.Key)        'Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ㿪ʼ��һ�����ڵ�
        SourceNodeParent = nParentNode
        AddNodeParentImage = nParentNode.Image       '����Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ��ԭͼ��
        nParentNode.Image = "FOLDER"
        nParentNode.ExpandedImage = "FOLDER"
        'ExpandedImage���Է��ػ������ڹ�����ImageList�ؼ��е�ListImage������������ֵ����Node����չ��ʱ��ʾ ListImage ����
    End If
    
    If NCode = "InputNew12NC" Then
        NCodeText = NCode
    Else
        NCodeText = Replace(NCode, "C", "")
    End If
        
    Set nNode = tvCodeItems.Nodes.Add(tvCodeItems.SelectedItem, tvwChild, "C" & "InputNew12NC", "C" & NCodeText, "CHILD")
    Set tvCodeItems.SelectedItem = nNode       '�任ѡ�еĽڵ�(������ʾ)�ӱ����ӵĽڵ� ��Ϊ�ո���ӵĽڵ�
    nNode.EnsureVisible
    
    AddNodeOk = True                   '�ȼ����ǿ��Ա༭OK��
    tvCodeItems.StartLabelEdit    'StartLabelEdit���������û��༭��ǩ��
    ' �� LabelEdit ��������Ϊ 1���ֶ���ʱ�������� StartLabelEdit ����������һ��ǩ�༭������
    ' ��һ�����ϵ��� StartLabelEdit ����ʱ��BeforeLabelEdit �¼�Ҳͬʱ������
    
    
    If Not AddNodeOk Then
        tvCodeItems.Nodes.Remove nNode.Key
        nParentNode.Image = AddNodeParentImage
    Else
        ChgMass = True
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmBOMAdmin:AddCode"

End Sub

'����һ���ڵ�(��)�Ĳ���(Check OK)
Private Sub mnuRename_Click()
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
    If SystemAdmin = "Y" Or OpennerSubmiter Then
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        Else
            GoTo AdminGoAhead1
        End If
    Else
        MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
        Exit Sub
    End If
    

    
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����

AdminGoAhead1:
    ' Change the Label - remember, we only allow 12 Characters
    Set CurNode = tvCodeItems.SelectedItem
    
    '��ʼ�汾�������޸�Assembly
    If CurVersion = 1 Then
        If Not CheckCanBeRename(CurNode.Text) Then
            MsgBox "The Assembly can't rename.", vbCritical
        End If
    End If
    
    OrientCurNodeKey = CurNode.Key
    OrientParentNodeKey = CurNode.Parent.Key
    Action = "REN"
    tvCodeItems.StartLabelEdit
    ChgMass = True
End Sub

'����һ���ڵ��ǩ��ǰԤ������(Check OK)
Private Sub tvCodeItems_BeforeLabelEdit(Cancel As Integer)
    Dim lEditHWND As Long
    ' Limit the text entry size to 12 characters (as defined in our database
    ' Get the handle of the Edit Box on the treeview
    lEditHWND = SendMessageLong(tvCodeItems.hWnd, TVM_GETEDITCONTROL, 0, 0)
    ' Now limit the size to 12 characters
    If lEditHWND > 0 Then
        SendMessageLong lEditHWND, EM_LIMITTEXT, 12, 0
    End If
    
End Sub

'����һ���ڵ��ǩ�ĺ�������(Check OK)
Private Sub tvCodeItems_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo vbErrorHandler
    Dim sKey As String
    Dim sKeyb4Rename As String
    Dim oNode As Node
    Dim oParentNode As Node
    Dim sMessage As String
    Dim oWait As CWaitCursor
    
    Dim rs As New ADODB.Recordset
    
    Set oNode = tvCodeItems.SelectedItem      'tvCodeItems.SelectedItem�����������һ����Ҫ�����Ľڵ㣬��һ���������ӵĽڵ�
    sKeyb4Rename = oNode
    oldCode = oNode
    
    If Len(NewString) = 0 Then
        MsgBox "You must enter a new 12NC for the new item", vbInformation, "System Info."
        Cancel = True
        If Trim(sKeyb4Rename) = "CInputNew12NC" Then tvCodeItems.Nodes.Remove "CInputNew12NC"
        Exit Sub
    ElseIf Not (Len(NewString) = 12 And Isnum(NewString)) Then
        MsgBox "Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
        Cancel = True
        If Trim(sKeyb4Rename) = "CInputNew12NC" Then tvCodeItems.Nodes.Remove "CInputNew12NC"
        Exit Sub
'    ElseIf Left(Trim(NewString), 11) <> Left(Trim(sKeyb4Rename), 11) And Trim(sKeyb4Rename) <> "CInputNew12NC" Then
'        MsgBox "The operation is only vaild for renaming the added New 12NC.", vbInformation, "System Info."
'        Cancel = True
'        Exit Sub
    End If
    
    
    If oNode.Key = "ROOT" Then Exit Sub     '����Ǹ��ڵ����˳���������
    
    If oNode Is Nothing Then      '�����û�нڵ�ѡ������ʾ���˳���������
        MsgBox "No Selected Record", vbInformation, "System Info"
        Exit Sub
    End If
    
    Set oParentNode = oNode.Parent  '��ǰѡ�е�Ҫ�����Ľڵ�ĸ��ڵ㸳ֵ
    SourceNodeParent = oParentNode
    
    '���ж��Ƿ���Finish Good , ��װ�ϼ�����Finish Good
    StrSql = "Select * From FinsGd Where FinsGdIndex=" & Trim(NewString) & " And (IsAssemblyPart=0 Or IsAssemblyPart=Null)"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        'Single Part�������������
        rs.Close
        StrSql = "Select * from SglPrt Where SglPrtIndex='" & left(Trim(NewString), 11) & "0'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The New Code is not existing, please first apply the code.", vbInformation, "System Info."
            rs.Close
            Cancel = True
            tvCodeItems.Nodes.Remove "CInputNew12NC"
            Exit Sub
        Else
            '��ȡ���°汾��
            NewString = left(NewString, 11) & CStr(rs("SglPrtVer"))
        End If
    Else
        NewString = Trim(NewString)
    End If
    rs.Close
    
    '�����Ƿ��ظ����
    StrSql = "Select * from " & temp_tb_BOMOrigData & " where ChildID like '" & left(NewString, 11) & "%' and ParentID='" & CStr(oParentNode.Text) & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        MsgBox "Unable to duplicate add the code on the SAME node.", vbCritical
        rs.Close
        Cancel = True
        tvCodeItems.Nodes.Remove "CInputNew12NC"
        Exit Sub
    End If
        
    
    If Not (Isnum(sKeyb4Rename)) Then GoTo handleAddcode   'ֱ��ȥ�������ڵ�Ĳ���
    StrSql = "Select * from  " & temp_tb_BOMOrigData & "  Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then   '������жϸ����Ľڵ���Add code�ո����ӵ�(ֻ��TreeView�пɼ�,��BOMOrigDataû�м�¼��)����˵����BOMOrigData(�м�¼��)
        sMessage = "Rename selected Code " & "and all child records ?"
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        If MsgBox(sMessage, vbYesNo + vbExclamation, "Rename Code Record") = vbNo Then
            Exit Sub
        End If
        
        
        sKey = NewString
        '�¸����Ľڵ����ֲ��ܺ�(ͬ������)���ֵܽڵ�����ͬ��
        RecursionFlag = False
        If AddNodeBrothership(sKey, oParentNode) Then
            MsgBox "This new name Can NOT be same name as it's Brother.", vbInformation, "System Info"
            NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
            Exit Sub
        End If
        '�¸����Ľڵ����ֲ��ܺͽ�Ҫ��Ϊ�ĸ����Լ�ֱ��Root���ڵ��ǰ���ڵ�ͬ��
        RecursionFlag = False
        If AddNodeEldership(sKey, oParentNode) Then
            MsgBox "This new name Can NOT be current Item's Parentship name.", vbInformation, "System Info"
            NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
            Exit Sub
        End If
        '�ж�Ҫ�ĳɵĽڵ�������Դ(������)�ڵ��µ��κ�һ���ӽڵ�(ͬһ��֦�ṹ�µ��ж�)
        If oNode.Children > 0 Then    '���ж�Ҫ�����Ľڵ���û�ӽڵ�
            RecursionFlag = False
            If isYoungershipNameNode(oNode, sKey) Then
                MsgBox "This new name Can NOT be current Item's Childship name.", vbInformation, "System Info"
                NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
                Exit Sub
            End If
        End If
        
        If Not (AddNodeKeyNameNodeExist(sKey) Is Nothing) Then '�ж������NewString�ڴ�BOM(TreeView)���Ƿ�Ϊһ���Ѿ����ڵĽڵ���ͬʱ�����ӽڵ�
            'AddNodeChildIsBrothership�жϽڵ�AddNodeKeyNameNodeExist(sKey)�ͽڵ�oNode��һ���ӽڵ�(���Ҫ�Ǻϲ����Ϊ�ֵܽڵ�)�����Ƿ�����ͬ&
            RecursionFlag = False
            If AddNodeChildIsBrothership(AddNodeKeyNameNodeExist(sKey), oNode) Then
                MsgBox "This New Name already has children and make same brother Name, Can NOT Rename.", vbInformation, "System Info"
                NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
                Exit Sub
            End If
            
            'isYoungershipNode�жϽڵ�oNode�Ƿ�Ϊ�ڵ�AddNodeKeyNameNodeExist(sKey)�ı����ӱ��ڵ�(BOMȷʵ�ǿ���ˢ����������,�������oNodeԭ���кܶ��ӽڵ���֦���ڸ����ϲ����ӽڵ���֦ȫ����ʧ,ʧȥ�����ϲ��ı���)
            RecursionFlag = False
            If isYoungershipNode(AddNodeKeyNameNodeExist(sKey), oNode) Then
                MsgBox "This New Name already has children which is Parentship to current Item, Can NOT Rename.", vbInformation, "System Info"
                NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
                Exit Sub
            End If
            
            'isElderYoungershipNode�жϽڵ�AddNodeKeyNameNodeExist(sKey)�ı����ӱ��ڵ��Ƿ�Ϊ�ڵ�oNode��ֱ��Rootǰ���ڵ�
            RecursionFlag = False
            If isElderYoungershipNode(AddNodeKeyNameNodeExist(sKey), oNode) Then
                MsgBox "This New Name already has children and which is Parentship to current Item's all Parent class, Can NOT Rename.", vbInformation, "System Info"
                NewString = sKeyb4Rename  '�ָ�ԭ���Ľڵ�����
                Exit Sub
            End If
        End If
        
        
        Conn.BeginTrans
        'update rename
        Conn.Execute "Update  " & temp_tb_BOMOrigData & "  Set ChildID='" & Trim(NewString) & "' Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'"
        
        'If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        '��������bom��İ汾
        If MsgBox("The code is being used the other BOM, are you sure to upgrade the new code for other BOM too?", vbYesNo, "ERP") = vbYes Then
            'BOM��ChildID����
            Conn.Execute "Update  " & temp_tb_BOMOrigData & "  Set ChildID='" & Trim(NewString) & "' Where ChildID='" & sKeyb4Rename & "'"
            'BOM��ChildID����������
            Conn.Execute "Update  " & temp_tb_BOMOrigData & "  Set  ParentID='" & Trim(NewString) & "' Where ParentID ='" & sKeyb4Rename & "'"

            Conn.CommitTrans
        Else
            Conn.RollbackTrans
            Exit Sub
        End If
        
        'If rs.State = adStateOpen Then rs.Close
        'Set rs = Nothing
        
        
        
        
'        rs.Open "Select * from BOMOrigData Where ParentID ='" & sKeyb4Rename & "'", Conn, adOpenKeyset, adLockOptimistic
'        If rs.RecordCount > 0 Then    '����(����)Ҫ�����Ľڵ����ӽڵ���Ҫ��������(��һ��)�ӽڵ�, ���¼������ӽڵ㲻�ø��¸Ķ�
'
'            Set oWait = New CWaitCursor
'            oWait.SetCursor
'
'            rs.MoveFirst
'            Do While Not rs.EOF
'                rs("ParentID") = Trim(NewString)
'                rs.Update
'                rs.MoveNext
'            Loop
'
'            If rs.State = adStateOpen Then rs.Close
'
'            Refresh_FlexGrid_TreeView
'            Set oWait = Nothing
'
'            Exit Sub
'
'        End If
        
        
        
        If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
        Refresh_FlexGrid_TreeView False
        Exit Sub
        
    Else
handleAddcode:        '�����Add code��TreeView�ո����ӵ�InputNew12NC�ڵ�
        sKey = NewString
        '�¼ӵĽڵ����ֲ��ܺ�(ͬ������)���ֵܽڵ�����ͬ��
        RecursionFlag = False
        If AddNodeBrothership(sKey, oParentNode) Then
            MsgBox "This new Item Can NOT be same name as it's Brother.", vbInformation, "System Info"
            Cancel = True
            tvCodeItems.Nodes.Remove "CInputNew12NC"
            Exit Sub
        End If
        '�¼ӵĽڵ����ֲ��ܺͽ�Ҫ��Ϊ�ĸ����Լ�ֱ��Root���ڵ��ǰ���ڵ�ͬ��
        RecursionFlag = False
        If AddNodeEldership(sKey, oParentNode) Then
            MsgBox "This new Item to be itself Parent class, Can NOT add.", vbInformation, "System Info"
            Cancel = True
            tvCodeItems.Nodes.Remove "CInputNew12NC"
            Exit Sub
        End If
        
        If Not (AddNodeKeyNameNodeExist(sKey) Is Nothing) Then '�ж������NewString�ڴ�BOM(TreeView)���Ƿ�Ϊһ���Ѿ����ڵĽڵ���ͬʱ�����ӽڵ�
            'isEldershipNode�жϽڵ�AddNodeKeyNameNodeExist(sKey)�Ƿ�Ϊ�ڵ�oParentNode��ǰ���ڵ�. ���ж�(���ӽڵ��)�������ж�(�����Ƿ���ӽڵ��)��ʵ�����Ѿ���,���Բ��õ�)
            'RecursionFlag = False
            'If isEldershipNode(AddNodeKeyNameNodeExist(sKey), oParentNode) Then
            'MsgBox "This New Item has children and make Parent-Child relation wrong, Can NOT add.", vbInformation, "System Info"
            'tvCodeItems.Nodes.Remove "CInputNew12NC"
            'Exit Sub
            'End If
            
            'isYoungershipNode�ж��¼ӽڵ㸸�ڵ�oParentNode�Ƿ�Ϊ�ڵ�AddNodeKeyNameNodeExist(sKey)�ı����ӱ��ڵ�
            RecursionFlag = False
            If isYoungershipNode(AddNodeKeyNameNodeExist(sKey), oParentNode) Then
                MsgBox "This New Item already has children which is Parentship to current Item, Can NOT Add.", vbInformation, "System Info"
                Cancel = True
                tvCodeItems.Nodes.Remove "CInputNew12NC"
                Exit Sub
            End If
            
            'isElderYoungershipNode�жϽڵ�AddNodeKeyNameNodeExist(sKey)�ı����ӱ��ڵ��Ƿ�Ϊ�¼ӽڵ㸸�ڵ�oParentNode��ֱ��Rootǰ���ڵ�
            RecursionFlag = False
            If isElderYoungershipNode(AddNodeKeyNameNodeExist(sKey), oParentNode) Then
                MsgBox "This New Item already has children and which is Parentship to current Item's all Parent class, Can NOT Add.", vbInformation, "System Info"
                Cancel = True
                tvCodeItems.Nodes.Remove "CInputNew12NC"
                Exit Sub
            End If
        End If
    
        
        Conn.Execute "INSERT INTO  " & temp_tb_BOMOrigData & "  (ParentID, ChildID, Quantity,ChgStatus) VALUES ('" & SourceNodeParent & "','" & sKey & "','" & "1" & "','')"
        
        Refresh_FlexGrid_TreeView False
        DoEvents
        
        If Not IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)) Then
            InsertBOMLog4Add MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4), sKey, SourceNodeParent, 1, GetParent(oNode.Key) & sKey, oNode
        Else
            InsertBOMLog4Add MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4), sKey, SourceNodeParent, MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) + 1, GetParent(oNode.Key) & sKey, oNode
        End If
        
    End If
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmBOMAdmin:tvCodeItems_AfterLabelEdit", , App.ProductName
End Sub
Private Sub InsertBOMLog4Copy(ByVal ParentKey As String, ByVal rowIndex As Integer, ByVal GrandpaKey As String)
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim PrtUnit, Description, ItemType, SERLocate, SERNmbr, TempSER, CommtNote As String
    Dim iQuantity As Integer
    

    '�����������ϼ�һ�����
    If rs.State = adStateOpen Then rs.Close
    StrSql = "Select * From " & temp_tb_BOMOrigData & " Where ParentID='" & ParentKey & "'"
    rs.Open StrSql, Conn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
        
            If rs2.State = adStateOpen Then rs2.Close
            StrSql = "Select * from SglPrt Where SglPrtIndex ='" & left(rs("ChildID"), 11) & "0" & "' Order By SglPrtIndex"
            rs2.Open StrSql, Conn, adOpenStatic, adLockReadOnly
            If rs2.RecordCount > 0 Then
                PrtUnit = Trim(rs2.Fields("PrtUnit"))
                Description = Trim(rs2.Fields("Description"))
                ItemType = Trim(rs2.Fields("ItemType"))
                iQuantity = Trim(rs.Fields("Quantity"))
        
                If Not IsNull(rs2.Fields("SERLocate")) Then
                    TempSER = Mid(Replace(Trim(rs2.Fields("SERLocate")), "----", ""), 32, 4)
                    If TempSER = "EASE" Then
                        TempSER = "RELEASREPORT"
                    Else
                        TempSER = "SER00000" & TempSER
                    End If
                Else
                    TempSER = ""
                End If
        
                If IsNull(rs2.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ objrs3.Fields("SERNmbr") = Null
                    If TempSER <> "" Then
                        SERNmbr = TempSER
                    Else
                        SERNmbr = ""
                    End If
                Else
                    If TempSER <> "" Then
                        SERNmbr = TempSER
                    Else
                        SERNmbr = rs2.Fields("SERNmbr")
                    End If
                End If
        
                If IsNull(rs2.Fields("CommtNote")) Then
                    CommtNote = ""
                Else
                    CommtNote = Trim(rs2.Fields("CommtNote"))
                End If
                rs2.Close
                Set rs2 = Nothing
                
        
                If CurVersion > 1 Then
'                    'ɾ���ϼ��޸ĵľ���־
'                    StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO
'                    StrSql = StrSql & " And ParentID='" & rs("ParentID") & "'"
'                    StrSql = StrSql & " And ChildID='" & rs("ChildID") & "'"
'                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                    
                    StrSql = "UPDATE " & temp_tb_SglPrt4BOMLog & " SET chgStatus='Delete'   Where BOM=" & FinishGoodsNO
                    StrSql = StrSql & " And ParentID='" & rs("ParentID") & "'"
                    StrSql = StrSql & " And ChildID='" & rs("ChildID") & "'"
                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
                    Conn.Execute StrSql
                
                    StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,Family,CPCN) Values("
                    StrSql = StrSql & FinishGoodsNO
                    StrSql = StrSql & "," & rowIndex
                    StrSql = StrSql & ",'" & rs("ParentID")
                    StrSql = StrSql & "','" & rs("ChildID")
                    StrSql = StrSql & "'," & CurVersion
                    StrSql = StrSql & "," & CStr(iQuantity)
                    StrSql = StrSql & ",'" & PrtUnit
                    StrSql = StrSql & "','" & Replace(Description, "'", "''")
                    StrSql = StrSql & "','" & ItemType
                    StrSql = StrSql & "','" & SERNmbr
                    StrSql = StrSql & "','" & "Add"
                    StrSql = StrSql & "','" & CommtNote
                    StrSql = StrSql & "','" & GrandpaKey & rs("ParentID") & ">" & rs("ChildID")
                    StrSql = StrSql & "','" & txtCPCNNO.Text & "'"
                    StrSql = StrSql & ")"
                    Conn.Execute (StrSql)
                End If
            End If
            InsertBOMLog4Copy rs("ChildID"), rowIndex, GrandpaKey & rs("ParentID") & ">"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Private Sub InsertBOMLog4Add(ByVal Qty As String, ByVal SglPrtKey As String, ByVal ParentKey As String, ByVal rowIndex As Integer, ByVal GrandpaKey As String, ByVal xNode As Node)
    Dim rs As New ADODB.Recordset
    Dim PrtUnit, Description, ItemType, SERLocate, SERNmbr, TempSER, CommtNote As String
    Dim i As Integer
    
    If rs.State = adStateOpen Then rs.Close
    StrSql = "Select * from SglPrt Where SglPrtIndex ='" & left(SglPrtKey, 11) & "0" & "' Order By SglPrtIndex"
    rs.Open StrSql, Conn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        PrtUnit = Trim(rs.Fields("PrtUnit"))
        Description = Trim(rs.Fields("Description"))
        ItemType = Trim(rs.Fields("ItemType"))

        If Not IsNull(rs.Fields("SERLocate")) Then
            TempSER = Mid(Replace(Trim(rs.Fields("SERLocate")), "----", ""), 32, 4)
            If TempSER = "EASE" Then
                TempSER = "RELEASREPORT"
            Else
                TempSER = "SER00000" & TempSER
            End If
        Else
            TempSER = ""
        End If

        If IsNull(rs.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ objrs3.Fields("SERNmbr") = Null
            If TempSER <> "" Then
                SERNmbr = TempSER
            Else
                SERNmbr = ""
            End If
        Else
            If TempSER <> "" Then
                SERNmbr = TempSER
            Else
                SERNmbr = rs.Fields("SERNmbr")
            End If
        End If

        If IsNull(rs.Fields("CommtNote")) Then
            CommtNote = ""
        Else
            CommtNote = Trim(rs.Fields("CommtNote"))
        End If
        rs.Close
        Set rs = Nothing
        

        If CurVersion > 1 Then
            If UCase(Action) = "COPY" Then
                If SglPrtKey <> CopyNodeSource Then
''                    'ɾ���ϼ��޸ĵľ���־
''                    StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO
''                    StrSql = StrSql & " And ParentID='" & ParentKey & "'"
''                    StrSql = StrSql & " And ChildID='" & SglPrtKey & "'"
''                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
'
'                    StrSql = "UPDATE  " & temp_tb_SglPrt4BOMLog & " SET chgStatus='Delete', CPCN='" & txtCPCNNO.Text & "'  Where BOM=" & FinishGoodsNO
'                    StrSql = StrSql & " And ParentID='" & ParentKey & "'"
'                    StrSql = StrSql & " And ChildID='" & SglPrtKey & "'"
'                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
'                    Conn.Execute StrSql
                
                
                    StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,Family,CPCN) Values("
                    StrSql = StrSql & FinishGoodsNO
                    StrSql = StrSql & "," & rowIndex
                    StrSql = StrSql & ",'" & ParentKey
                    StrSql = StrSql & "','" & SglPrtKey
                    StrSql = StrSql & "'," & CurVersion
                    StrSql = StrSql & "," & Qty
                    StrSql = StrSql & ",'" & PrtUnit
                    StrSql = StrSql & "','" & Replace(Description, "'", "''")
                    StrSql = StrSql & "','" & ItemType
                    StrSql = StrSql & "','" & SERNmbr
                    StrSql = StrSql & "','" & "Add"
                    StrSql = StrSql & "','" & CommtNote
                    StrSql = StrSql & "','" & Replace(GrandpaKey, SglPrtKey, ParentKey)
                    StrSql = StrSql & "','" & txtCPCNNO.Text & "'"
                    StrSql = StrSql & ")"
                    Conn.Execute (StrSql)
                End If
            ElseIf UCase(Action) = "ADD" Then
''                    'ɾ���ϼ��޸ĵľ���־
''                    StrSql = "Delete From  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO
''                    StrSql = StrSql & " And ParentID='" & ParentKey & "'"
''                    StrSql = StrSql & " And ChildID='" & SglPrtKey & "'"
''                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
'                    StrSql = "UPDATE  " & temp_tb_SglPrt4BOMLog & " SET chgStatus='Delete', CPCN='" & txtCPCNNO.Text & "'  Where BOM=" & FinishGoodsNO
'                    StrSql = StrSql & " And ParentID='" & ParentKey & "'"
'                    StrSql = StrSql & " And ChildID='" & SglPrtKey & "'"
'                    StrSql = StrSql & " And BOMVersion=" & CStr(CurVersion)
'                    Conn.Execute StrSql
                    
                    With MSFlexGrid1
                        For i = 2 To .Rows - 2
                            If Trim(.TextMatrix(i, 2)) = ParentKey And Trim(.TextMatrix(i, 3)) = SglPrtKey Then
                                StrSql = "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,ChgStatus,CommtNote,Family,CPCN) Values("
                                StrSql = StrSql & FinishGoodsNO
                                StrSql = StrSql & "," & rowIndex
                                StrSql = StrSql & ",'" & ParentKey
                                StrSql = StrSql & "','" & SglPrtKey
                                StrSql = StrSql & "'," & CurVersion
                                StrSql = StrSql & ",1"
                                StrSql = StrSql & ",'" & PrtUnit
                                StrSql = StrSql & "','" & Replace(Description, "'", "''")
                                StrSql = StrSql & "','" & ItemType
                                StrSql = StrSql & "','" & SERNmbr
                                StrSql = StrSql & "','" & "Add"
                                StrSql = StrSql & "','" & CommtNote
                                StrSql = StrSql & "','" & .TextMatrix(i, 11)
                                StrSql = StrSql & "','" & txtCPCNNO.Text & "')"
                                Conn.Execute (StrSql)
                            End If
                        Next i
                    End With
            End If
        End If
        '�����������ϼ�һ�����
        If rs.State = adStateOpen Then rs.Close
        StrSql = "Select * From " & temp_tb_BOMOrigData & " Where ParentID='" & SglPrtKey & "'"
        rs.Open StrSql, Conn, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                InsertBOMLog4Add rs("Quantity"), rs("ChildID"), rs("ParentID"), rowIndex, GetParent(xNode.Key) & rs("ParentID") & ">" & rs("ChildID"), xNode
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If
    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
End Sub


'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub FillTree()
    On Error Resume Next
    Dim J As Integer
    'Dim lCount As Long
    Dim sParent As String
    Dim sKey As String
    Dim sText As String
    Dim nNode As Node              '�������������
    Dim sData()  As String
    Dim SameChild()  As String
    Dim SameParent()  As String
    Dim lList  As Long
    
    '  ��TreeView����Ҫ����
    
    Set tvCodeItems.ImageList = Nothing   '���
    Set tvCodeItems.ImageList = ImageList1   'TreeView��ImageList1����һ��
    
    If RowNum <= 1 Then '����ǵ�����(�޼�¼)����ֻ����һ��Root
        tvCodeItems.Nodes.Add , , "ROOT", MSFlexGrid1.TextMatrix(1, 3), "SETTINGS"
        BoldTreeNode tvCodeItems.Nodes("ROOT")   'BoldTreeNode��һ��Sub������Make a tree node bold
        Exit Sub
    End If
    
    TreeRedraw tvCodeItems.hWnd, False   '��ʱ��ֹTreeView���Զ��ػ�.TreeRedraw��һ��Sub������
    
    Set tvCodeItems.ImageList = Nothing     '���
    Set tvCodeItems.ImageList = ImageList1  'TreeView��ImageList1����һ��
    '
    ' Populate the TreeView Nodes
    '
    
    With tvCodeItems.Nodes    '������ʽ��ʼ���ڵ�
        .Clear
        If IsBOMLocked Then
            '.Add , , "ROOT", MSFlexGrid1.TextMatrix(1, 3), "SETTINGS"
            .Add , , "ROOT", MSFlexGrid1.TextMatrix(1, 3), "LOCKED"
        Else
            .Add , , "ROOT", MSFlexGrid1.TextMatrix(1, 3), "SETTINGS"  '����MSFlexGrid1.TextMatrix(1, 3)�Ǹ��ڵ����ֿ����滻
        End If
        '
        ' Make our Root Item BOLD
        '
        BoldTreeNode tvCodeItems.Nodes("ROOT") 'ע��ڵ�ؼ���ָ����ʽ������˫���ż�����
        '
        ' Now add all nodes into TreeView, but under the root item.��ʼ�����нڵ㵽���ڵ���
        ' We reparent the nodes in the next step
        '
        
        lList = 1
        ReDim sData(1 To lList)
        ReDim SameChild(1 To lList)
        ReDim SameParent(1 To lList)
        'Set myNod=TreeView�ؼ���.Nodes.Add(a,b,key,Text,Image)
        'a: �������key��Ҳ����˵Ҫ���ĸ��ڵ�����������ݣ�a�����ĸ��ڵ��keyֵ
        'b: �Ͳ�����Ĺ�ϵ������Ͳ��յĽڵ���ƽ���ľ�д"tvwNext"������ǲ�������ӽڵ��д"tvwChild"
        'key: �ڵ�Ĺؼ��֣�����˵�ǽڵ�����֣������ظ�
        'text: �ڵ�����ʾ������
        'image���ڵ��ͼ��
        J = 2
        Do Until J = RowNum         '��Ѱÿ����¼ֱ����β,�Ȱ�ÿ����¼��һ���ڵ�(���ֲ㼶)
            
            'If UCase(Left(Trim(MSFlexGrid1.TextMatrix(J, 12)), 1)) <> "D" Then '�ų���ɾ��������LOG
                lList = lList + 1
                ReDim Preserve SameParent(1 To lList)     '��ʼ�ж��Ƿ����ظ��Ľڵ�key
                ReDim Preserve SameChild(1 To lList)      '��ʼ�ж��Ƿ����ظ��Ľڵ�key
                SameParent(lList) = MSFlexGrid1.TextMatrix(J, 2)
                SameChild(lList) = MSFlexGrid1.TextMatrix(J, 3)
                Dim s As Integer, PrefixStrgParent As String
                Dim t As Integer, PrefixStrgChild As String
                PrefixStrgParent = ""
                PrefixStrgChild = ""
                
                
                For s = 1 To lList - 1
                    If SameChild(s) = MSFlexGrid1.TextMatrix(J, 3) Then
                        PrefixStrgChild = PrefixStrgChild & "C"
                    End If
                    If SameParent(s) = MSFlexGrid1.TextMatrix(J, 2) And SameChild(s) = MSFlexGrid1.TextMatrix(J, 3) Then
                        PrefixStrgParent = PrefixStrgParent & "C"
                    End If
                Next s
                sParent = PrefixStrgParent & MSFlexGrid1.TextMatrix(J, 2)
                sKey = PrefixStrgChild & MSFlexGrid1.TextMatrix(J, 3)
                
                sText = MSFlexGrid1.TextMatrix(J, 3)
                
                Set nNode = .Add("ROOT", tvwChild, "C" & sKey, sText, "FOLDER")

                '
                ' Record parent ID
                '
                ReDim Preserve sData(1 To lList)  'Preserve ��ѡ�ġ��ؼ��֣����ı�ԭ��������ĩά�Ĵ�Сʱ��ʹ�ô˹ؼ��ֿ��Ա���������ԭ��������
                sData(lList) = "C" & sParent
                
            'End If
            J = J + 1
        Loop
        
    End With
    
    ' Here's where we rebuild the structure of the nodes ÿ����¼�����ṹ(�ֲ㼶)
    
    Dim vNode  As Node
    For Each vNode In tvCodeItems.Nodes
        vNode.Expanded = True
    Next
    
    lList = 0
    For Each nNode In tvCodeItems.Nodes
        lList = lList + 1
        sParent = sData(lList)  'sData(1)�ǿ�ֵ,��Ϊ����ÿ����¼��һ���ڵ�(���ֲ㼶)ʱlList�Ǵ�1��ʼ
        If Len(sParent) <= 0 Or Len(nNode) <= 0 Then      ' Don't try and reparent the ROOT !
            GoTo NextNode
        End If
        If sParent = "C" & Trim(MSFlexGrid1.TextMatrix(1, 3)) Then
            sParent = "ROOT"
        End If
        Set nNode.Parent = tvCodeItems.Nodes(sParent)
NextNode:
    Next nNode
    '
    ' Now setup the images for each node in the treeview & set each node to
    ' be sorted if it has children
    '
    For Each nNode In tvCodeItems.Nodes
        If nNode.Children = 0 Then
            nNode.Image = "CHILD"
        Else
            nNode.Sorted = True  'Sorted���Է��ػ�����һֵ����ֵȷ�� Node ������ӽڵ��Ƿ���ĸ˳�����С�
        End If
    Next nNode
    
    '
    ' Expand the Root Node
    '
    tvCodeItems.Nodes("ROOT").Sorted = True
    tvCodeItems.Nodes("ROOT").Expanded = True
    
    TreeRedraw tvCodeItems.hWnd, True

    Select Case Action
    Case "ADD"
        tvCodeItems.Nodes(OrientCurNodeKey).Selected = True
        tvCodeItems.Nodes(OrientCurNodeKey).EnsureVisible
    Case "DEL"
        tvCodeItems.Nodes(OrientParentNodeKey).Selected = True
        tvCodeItems.Nodes(OrientParentNodeKey).EnsureVisible
    Case "UPG"
        tvCodeItems.Nodes(OrientCurNodeKey).Selected = True
        tvCodeItems.Nodes(OrientCurNodeKey).EnsureVisible
    Case "COPY"
        tvCodeItems.Nodes(OrientCurNodeKey).Selected = True
        tvCodeItems.Nodes(OrientCurNodeKey).EnsureVisible
    Case Else
        tvCodeItems.Nodes(OrientCurNodeKey).Selected = True
        tvCodeItems.Nodes(OrientCurNodeKey).EnsureVisible
    End Select
End Sub


Private Sub BoldTreeNode(nNode As Node)
    ' Make a tree node bold
    ' Many thanks to VBNet for this code
    
    On Error GoTo vbErrorHandler
    
    Dim TVI As TVITEM
    Dim lRet As Long
    Dim hItemTV As Long
    Dim lHwnd As Long
    
    Set tvCodeItems.SelectedItem = nNode  'SelectedItem����һ��ֵ�������ؼ���ѡ�еļ�¼����ǩ��
    
    lHwnd = tvCodeItems.hWnd
    hItemTV = SendMessageLong(lHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    '  ��ģ��Modulel��������Declare
    '  Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
    '  Public Const TVGN_CARET As Long = &H9
    '  Public Const TV_FIRST As Long = &H1100
    
    If hItemTV > 0 Then
        '  ��ģ��Modulel��������Declare
        '   Public Const TVIF_STATE As Long = &H8
        '   Public Const TVIS_BOLD As Long = &H10
        '   Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
        '   Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
        With TVI
            .hItem = hItemTV
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            lRet = SendMessageAny(lHwnd, TVM_GETITEM, 0&, TVI)
            .State = TVIS_BOLD
        End With
        lRet = SendMessageAny(lHwnd, TVM_SETITEM, 0&, TVI)
    End If
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "frmCodeLib::BoldTreeNode"
    
End Sub

Private Sub TreeRedraw(ByVal lHwnd As Long, ByVal bRedraw As Boolean)
    '
    ' Utility Routine for TreeRedraw on/off
    '
    SendMessageLong lHwnd, WM_SETREDRAW, bRedraw, 0
    '��ģ��CodeModule��������Declare
    'Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'hWnd������ľ����ϣ������Ϣ���͸��ĸ�����.��Text1.hWnd��Form1.hWnd�ֱ���Եõ�Text1��Form1�ľ����
    'wMsg�������͵���Ϣ�����ݾ�������Ͳ�ͬ�Ķ��󣬽���ͬ����Ϣ��Ϊʵ�δ��ͣ��Բ���Ԥ�ڵĶ�����
    'wParam��lParam�����ӵ���Ϣ��Ϣ���������ǿ�ѡ�Ĳ����������ṩ����wMsg��Ϣ�������Ϣ����ͬ��wMsg����ʹ�������������е�0��1��2�����������Ҫ�ĸ����Ӳ�������ʵ�θ�ΪNULL����VB�и�Ϊ0����
    
    '�����б����ӻ�ɾ���ַ���ʱ.�б�ؼ����Զ������Ĵ��ں����ػ�.�����������ַ�����Ҫ����.
    '�����ϣ���������ַ����������ǰ��ʱ��ֹ�б���Զ��ػ�.��Ҫ:
    'SendMessage (hwndList, WM_SETREDRAW, FALSE, 0) ;
    '��������ɺ��ٻָ��б�ؼ����Զ��ػ��Ϳ�����:
    'SendMessage (hwndList, WM_SETREDRAW, TRUE, 0) ;
End Sub
Public Sub Get_BOM_FlexGrid(ByVal parentid As String, ByVal TempLevel As Integer)
    '//��ʼBOM
    On Error Resume Next
    Dim StrSql As String
    Dim objConn As New ADODB.Connection
    objConn.Open connString
    Dim objrs As New ADODB.Recordset
    Dim objRS3 As New ADODB.Recordset
    Dim TempSER, SERNmbr As String
    Dim bFinsGd As Boolean
    Dim PrtUnit, Description, ItemType, SERNumber, SERLocation, Commnote As Variant
    

    StrSql = "SELECT * FROM " & temp_tb_BOMOrigData & " WHERE ParentId='" + parentid + "' Order By ChildId"

    If objrs.State = adStateOpen Then objrs.Close
    objrs.Open StrSql, objConn, adOpenStatic, adLockOptimistic
    If objrs.RecordCount > 0 Then
        Do While Not objrs.EOF
            PrtUnit = ""
            Description = ""
            ItemType = ""
            SERNumber = ""
            SERLocation = ""
            Commnote = ""
            
            MSFlexGrid1.TextMatrix(J, 0) = J        '����ÿ�е��к�
            MSFlexGrid1.TextMatrix(J, 1) = TempLevel
            MSFlexGrid1.TextMatrix(J, 2) = objrs.Fields("ParentID")
            MSFlexGrid1.TextMatrix(J, 3) = objrs.Fields("ChildID")
            MSFlexGrid1.TextMatrix(J, 4) = objrs.Fields("Quantity")
            
            If IsNull(objrs.Fields("ChgStatus")) Then          '������IsNull�����ж�,������ objrs2.Fields("ChgStatus") = Null
                MSFlexGrid1.TextMatrix(J, 9) = ""
            Else
                If CurVersion > 1 Then MSFlexGrid1.TextMatrix(J, 9) = objrs.Fields("ChgStatus")
            End If
                    
            '���ϼ��п�������װ�ϼ�
            StrSql = "Select * from FinsGd Where FinsGdIndex='" & objrs.Fields("ChildID") & "'"
            objRS3.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
            If objRS3.EOF Or objRS3.BOF Then
                bFinsGd = False
                objRS3.Close
                StrSql = "Select Top 1 * from SglPrt Where SglPrtIndex ='" & left(objrs.Fields("ChildID"), 11) & "0" & "' Order By SglPrtIndex"
                objRS3.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
                If Not objRS3.EOF Then
                    PrtUnit = Trim(objRS3.Fields("PrtUnit"))
                    Description = Trim(objRS3.Fields("Description"))
                    ItemType = Trim(objRS3.Fields("ItemType"))
                    SERLocation = IIf(IsNull(objRS3.Fields("SERLocate")), "", Trim(objRS3.Fields("SERLocate")))
                    SERNumber = IIf(IsNull(objRS3.Fields("SERNmbr")), "", Trim(objRS3.Fields("SERNmbr")))
                    Commnote = IIf(IsNull(objRS3.Fields("CommtNote")), "", Trim(objRS3.Fields("CommtNote")))
                    If temp_tb_BOMOrigData <> "BOMOrigData" Then
                        SERNumber = getPartValue("SglPrt", "SERNmbr", left(objrs.Fields("ChildID"), 11) & "0", IIf(IsNull(objRS3.Fields("SERNmbr")), "", SERNumber))
                        Commnote = getPartValue("SglPrt", "CommtNote", left(objrs.Fields("ChildID"), 11) & "0", IIf(IsNull(objRS3.Fields("CommtNote")), "", Commnote))
                    End If
                End If
            Else
                PrtUnit = Trim(objRS3.Fields("PrtUnit"))
                Description = Trim(objRS3.Fields("Description"))
                ItemType = Trim(objRS3.Fields("ItemType"))
                SERLocation = IIf(IsNull(objRS3.Fields("SERLocate")), "", Trim(objRS3.Fields("SERLocate")))
                SERNumber = IIf(IsNull(objRS3.Fields("SERNmbr")), "", Trim(objRS3.Fields("SERNmbr")))
                Commnote = IIf(IsNull(objRS3.Fields("CommtNote")), "", Trim(objRS3.Fields("CommtNote")))
                If temp_tb_BOMOrigData <> "BOMOrigData" Then
                    SERNumber = getPartValue("FinsGd", "SERNmbr", CStr(objrs.Fields("ChildID")), IIf(IsNull(objRS3.Fields("SERNmbr")), "", SERNumber))
                    Commnote = getPartValue("FinsGd", "CommtNote", CStr(objrs.Fields("ChildID")), IIf(IsNull(objRS3.Fields("CommtNote")), "", Commnote))
                End If
            End If
                
            MSFlexGrid1.TextMatrix(J, 5) = PrtUnit
            MSFlexGrid1.TextMatrix(J, 6) = Description
            MSFlexGrid1.TextMatrix(J, 7) = ItemType


            
            If SERLocation <> "" Then
                TempSER = Mid(Replace(SERLocation, "----", ""), 32, 5)
                If TempSER = "EASE " Then
                    TempSER = "RELEASREPORT"
                Else
                    If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "��" Then
                        TempSER = "SER00000" & left(TempSER, 4)
                    Else
                        TempSER = "SER0000" & TempSER
                    End If
                End If
            Else
                TempSER = ""
            End If

            
            If SERNumber = "" Or IsNull(SERNumber) Then  '������IsNull�����ж�,������ objrs3.Fields("SERNmbr") = Null
                If TempSER <> "" Then
                    MSFlexGrid1.TextMatrix(J, 8) = TempSER
                Else
                    MSFlexGrid1.TextMatrix(J, 8) = ""
                End If
                SERNmbr = ""
            Else
                If TempSER <> "" Then
                    MSFlexGrid1.TextMatrix(J, 8) = TempSER
                Else
                    MSFlexGrid1.TextMatrix(J, 8) = SERNumber
                End If
                SERNmbr = SERNumber
            End If
            
            '����SER
            If TempSER <> "" Then
                If SERNmbr <> TempSER Then
                    If bFinsGd Then
                        Call UpdateSERNmbr(TempSER, Trim(MSFlexGrid1.TextMatrix(J, 3)), "FinsGd")
                    Else
                        Call UpdateSERNmbr(TempSER, Mid(Trim(MSFlexGrid1.TextMatrix(J, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(J, 3))) - 1) & "0", "SglPrt")
                    End If
                End If
            End If
            
            If IsNull(Commnote) Then
                MSFlexGrid1.TextMatrix(J, 10) = ""
            Else
                MSFlexGrid1.TextMatrix(J, 10) = Commnote
            End If
            ReDim Preserve Family(1 To TempLevel)
            Family(TempLevel) = CStr(objrs.Fields("ChildID"))
            Dim i As Integer
            MSFlexGrid1.TextMatrix(J, 11) = FinishGoodsNO
            For i = 1 To TempLevel
                MSFlexGrid1.TextMatrix(J, 11) = MSFlexGrid1.TextMatrix(J, 11) & ">" & Family(i)
            Next
            
            J = J + 1
            MSFlexGrid1.Rows = J + 1
            Call Get_BOM_FlexGrid(objrs.Fields("ChildID"), TempLevel + 1)
            
            If objRS3.State = adStateOpen Then objRS3.Close
            objrs.MoveNext
        Loop
    End If
    objrs.Close
    Set objrs = Nothing
    objConn.Close
    Set objConn = Nothing
End Sub
Public Sub Refresh_FlexGrid_TreeView(bSaveVersion As Boolean)
    'On Error Resume Next
    Dim strCPCN, strStatus As String
    Dim BgColor As Long
    Dim TempSER, SERNmbr As String
    Dim TempLevel As Integer
    
    Dim myCnn As New ADODB.Connection
    myCnn.Open connString
    
    '//����������¼��  rstCX��¼����Ӧ��ʱ��  rstCX2��¼����Ӧ��ʱ����һ����¼����Ķ�Ӧ�Ӽ�¼
    Dim rstCX As New ADODB.Recordset
    Dim rstCX3 As New ADODB.Recordset
    MSFlexGrid1.Clear
    MSFlexGridTileInitialize
    
    '//�ж������ͼ���Ƿ�ײ���� �����û�и���ĵײ���������ʾ���˳�
    StrSql = "SELECT * FROM " & temp_tb_BOMOrigData & " WHERE ParentID = '" + FinishGoodsNO + "'"
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
    If Not rstCX.RecordCount > 0 Then
        MsgBox " This item is not assembly, has no Child", vbInformation, "System Info."
        rstCX.Close
        Exit Sub
    End If
    rstCX.Close
    
    '//���subcon��Ϣ
    StrSql = "SELECT * FROM SUBCON Where FinsGDIndex=" & FinishGoodsNO
    rstCX.Open StrSql, myCnn, adOpenStatic, adLockOptimistic
    If rstCX.RecordCount > 0 Then
        If Not IsNull(rstCX.Fields("SubCon")) Then
            txtSubCon.Text = Trim(rstCX.Fields("SubCon"))
        Else
            txtSubCon.Text = ""
        End If
    Else
        txtSubCon.Text = ""
    End If
    rstCX.Close
    '//��MSFlexGrid����������Ϣ
    
    '����ͻ�������ͼ
    Dim Msg As String

    
    '//���BOM TOP��
    rstCX3.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(FinishGoodsNO) & "'", myCnn, adOpenKeyset, adLockOptimistic
    If rstCX3.RecordCount > 0 Then
        MSFlexGrid1.TextMatrix(1, 0) = 1        '����ÿ�е��к�
        MSFlexGrid1.TextMatrix(1, 1) = 0
        MSFlexGrid1.TextMatrix(1, 2) = "Top"
        MSFlexGrid1.TextMatrix(1, 3) = Trim(FinishGoodsNO)
        MSFlexGrid1.TextMatrix(1, 4) = 1            'Root���ڵ�(Finish Goods ��������Ϊ1)
        MSFlexGrid1.TextMatrix(1, 5) = "Piece"     '���ڸ���Ŀ��Unit,����Ϊ"Piece"
        MSFlexGrid1.TextMatrix(1, 6) = Trim(rstCX3.Fields("Description"))
        MSFlexGrid1.TextMatrix(1, 7) = Trim(rstCX3.Fields("ItemType"))
        
        If Not IsNull(rstCX3.Fields("SERLocate")) Then
            TempSER = Mid(Replace(rstCX3.Fields("SERLocate"), "----", ""), 32, 5)
            If TempSER = "EASE " Then
                TempSER = "RELEASREPORT"
            Else
                If right(TempSER, 1) = "-" Or right(TempSER, 1) = "(" Or right(TempSER, 1) = "��" Then
                    TempSER = "SER00000" & left(TempSER, 4)
                Else
                    TempSER = "SER0000" & TempSER
                End If
            End If
        Else
            TempSER = ""
        End If

        
        If IsNull(rstCX3.Fields("SERNmbr")) Then    '������IsNull�����ж�,������ rstCX3.Fields("SERNmbr") = Null
            If TempSER <> "" Then
                MSFlexGrid1.TextMatrix(1, 8) = TempSER
            Else
                MSFlexGrid1.TextMatrix(1, 8) = ""
            End If
            SERNmbr = ""
        Else
            If TempSER <> "" Then
                MSFlexGrid1.TextMatrix(1, 8) = TempSER
            Else
                MSFlexGrid1.TextMatrix(1, 8) = rstCX3.Fields("SERNmbr")
            End If
            SERNmbr = Trim(rstCX3.Fields("SERNmbr"))
        End If
        
        '����SER
        If TempSER <> "" Then
            If SERNmbr <> TempSER Then Call UpdateSERNmbr(TempSER, Trim(MSFlexGrid1.TextMatrix(1, 3)), "FinsGd")
        End If
        
        If IsNull(rstCX3.Fields("CommtNote")) Then
            MSFlexGrid1.TextMatrix(1, 10) = ""
        Else
            MSFlexGrid1.TextMatrix(1, 10) = Trim(rstCX3.Fields("CommtNote"))
        End If

    End If
    If rstCX3.State = adStateOpen Then rstCX3.Close
    If rstCX.State = adStateOpen Then rstCX.Close
    MSFlexGrid1.Rows = 3
    J = 2
    Call Get_BOM_FlexGrid(FinishGoodsNO, 1)
    
    RowNum = J
    'MSFlexGrid1.Rows = RowNum '#########�������¶��������������п���############
    MSFlexGridColumnColorChange MSFlexGrid1, 4, J            '����quantity��(��4��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 8, J            '����SER��(��8��)Ϊǳ�ۺ�ɫ
    MSFlexGridColumnColorChange MSFlexGrid1, 9, J           '����ChangeStatus��(��9��)Ϊǳ�ۺ�ɫ
    'MSFlexGrid_ChgStatus_HightlightRow (10)                   '�Ե�10���������ݵ�������Ϊ��ɫ
    MSFlexGrid_ApproveStatus_HightlightRow (ApprovalStatus)   '�Ե�1������Ϊ��ɫ������Ѿ���׼��BOM
    MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 0                 '���õ�Ԫ��λ��ȡ������ı亯���е�ĳ�и�����ʾ
    
    FillTree
    
    tvCodeItems.HideSelection = False
    If bSaveVersion Then cmdBOMSave_Click
    
End Sub
Private Sub MSFlexGrid_ChgStatus_HightlightRow(ByVal CheckColNO As Integer)
    Dim RowSumVar As Integer
    For RowSumVar = 1 To RowNum
        If MSFlexGrid1.TextMatrix(RowSumVar, CheckColNO) <> "" Then
            MSFlexGridRowColorChange MSFlexGrid1, RowSumVar, MSFlexGrid1.Cols
        End If
    Next RowSumVar
End Sub

Private Sub MSFlexGrid_ActionStatus_HightlightRow(ByVal CheckRowNO As Integer, ByVal BgColor As Long)
    Dim ColSumVar As Integer
    For ColSumVar = 1 To MSFlexGrid1.Cols - 1
        MSFlexGridRowColorChange MSFlexGrid1, CheckRowNO, ColSumVar, BgColor
    Next ColSumVar
End Sub

Private Sub MSFlexGrid_ApproveStatus_HightlightRow(ByVal ApproverOK As Boolean)
    If ApproverOK Then
        MSFlexGridRowColorChange MSFlexGrid1, 2, MSFlexGrid1.Cols, &H80FF80     '&H80FF80Ϊ��ɫ
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    Dim ColNoExitable As Integer
    
    If IsBOMLocked Then Exit Sub
    
    ColNoExitable = MSFlexGrid1.Col
    Select Case ColNoExitable
    Case 4, 9                            '��4�� quantity,��9��ChangeStatus
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        If SystemAdmin = "Y" Or OpennerSubmiter Then
            'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
            GoTo AdminGoAhead1
        Else
            MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
            Exit Sub
        End If
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        
AdminGoAhead1:
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        End If
        
        If MSFlexGrid1.Row = 1 Then Exit Sub
        BOMString = ""
        If CurVersion = 1 And MSFlexGrid1.Col = 4 Then
            If checkModifyPermission(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) = False Then Exit Sub
        End If
        
        If MSFlexGrid1.Row = 1 Then
            MSFlexGrid1.Row = 2
            MsgBox "The Root(Top) Item Quantity/ChangeStatus is not Editable" & vbCrLf & "Please Edit Quantity/ChangeStatus from 2nd Row", vbInformation, "System Info."
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(FinishGoodsNO)
    Case 8, 10                              '��8��SER,��10��Notes
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����
        If SystemAdmin = "Y" Or OpennerSubmiter Then
            'MsgBox "You are Administrator, Full right to Go Ahead", vbInformation, "System Info."
            GoTo AdminGoAhead2
        Else
            MsgBox "You are not the BOM Author, No Right to update", vbInformation, "System Info."
            Exit Sub
        End If
        '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û��������ֱ����ת����

AdminGoAhead2:
        If Trim(CPCN) = "" And isApproved Then
            MsgBox "You can't modify BOM yet, please input the CPCN Number at first.", vbCritical
            Exit Sub
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(FinishGoodsNO)
        
    Case Else
        'MsgBox " This column is not editable", vbInformation, "System Info."
        Exit Sub
    End Select
    
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid1EditText_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo vbErrorHandler
    'Set scr = CreateObject("MSScriptControl.ScriptControl")
    'scr.Language = "vbscript"
    Dim RowNoTemp As Integer
    Dim ColNoTemp As Integer
    
    
    Dim rs As New ADODB.Recordset
    
    If MSFlexGrid1.Col <> 4 And MSFlexGrid1.Col <> 8 And MSFlexGrid1.Col <> 9 And MSFlexGrid1.Col <> 10 Then Exit Sub
    ColNoTemp = MSFlexGrid1.Col
    RowNoTemp = MSFlexGrid1.Row
    'MsgBox "Col:" & ColNoTemp & ",Row:" & RowNoTemp & ",Keycode:" & KeyCode
    
'    BOMString = ""
'    Call GetTopBOM(MSFlexGrid1.TextMatrix(RowNoTemp, 3))
'    arrBOM = Split(Mid(BOMString, 2), ",")
    
    If KeyCode = 27 Then
        MSFlexGrid1EditText.Visible = False
        MSFlexGrid1.SetFocus
        Exit Sub
    ElseIf KeyCode = 13 Then
        'MSFlexGrid1.Text = scr.Eval(FinishGoodsNO)                                 '��ScriptControl������������ʽ
        
        MSFlexGrid1.Text = Trim(MSFlexGrid1EditText.Text)
        
        Select Case ColNoTemp
        Case 4
            'MsgBox MSFlexGrid1EditText.Text
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            If EvaluateExpr(Trim(MSFlexGrid1EditText.Text)) <> 0 Then
                MSFlexGrid1.Text = EvaluateExpr(Trim(MSFlexGrid1EditText.Text))          'EvaluateExpr��GeneralFuncģ���ж���ĺ���
            Else
                MSFlexGrid1.Text = ""
            End If
            rs.Open "Select * from " & temp_tb_BOMOrigData & " Where ParentID ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 2)) & "' and ChildID = '" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                MsgBox "The Record is not existing in BOM database", vbInformation, "System Info."
                If rs.State = adStateOpen Then rs.Close
                Exit Sub
            Else
                rs("Quantity") = Round(Trim(MSFlexGrid1.Text), 7)    '����7λС��������
                If CurVersion > 1 Then rs("ChgStatus") = "Modify" 'update QTY
                rs.Update
                
                'д���޸���־
                StrSql = "UPDATE " & temp_tb_SglPrt4BOMLog & "  SET chgStatus='Modify',CPCN='" & txtCPCNNO.Text & "',Quantity= " & Round(Trim(MSFlexGrid1.Text), 7) & " Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CStr(CurVersion) & " And ParentID='" & MSFlexGrid1.TextMatrix(RowNoTemp, 2) & "' And ChildID='" & MSFlexGrid1.TextMatrix(RowNoTemp, 3) & "' And Family='" & MSFlexGrid1.TextMatrix(RowNoTemp, 11) & "' And (chgStatus Like 'Add%' OR chgStatus IS NULL OR chgStatus='Modify')"
                Conn.Execute StrSql
                
                ChgMass = True '�޸ı�־
            End If
            If rs.State = adStateOpen Then rs.Close
            
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 8
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                If rs.State = adStateOpen Then rs.Close
                rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                    MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
                Else
                    If CurVersion = 1 Then
                        rs("SERNmbr") = Mid(Trim(MSFlexGrid1.Text), 1, 12)  'ǿ�ƽ�ȡ12λ
                        rs.Update
                    Else
                        StrSql = "INSERT INTO PartVar([BOM],[CPCN],[PartIndex],[PartValue],[TableName],[FieldName]) VALUES('" & FinishGoodsNO & "','" & CPCN & "','" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "','" & Mid(Trim(MSFlexGrid1.Text), 1, 12) & "','FinsGd','SERNmbr')"
                        Conn.Execute StrSql
                    End If
                    ChgMass = True
                End If
            Else
                If CurVersion = 1 Then
                    rs("SERNmbr") = Mid(Trim(MSFlexGrid1.Text), 1, 12)  'ǿ�ƽ�ȡ12λ
                    rs.Update
                Else
                    StrSql = "INSERT INTO PartVar([BOM],[CPCN],[PartIndex],[PartValue],[TableName],[FieldName]) VALUES('" & FinishGoodsNO & "','" & CPCN & "','" & Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0" & "','" & Mid(Trim(MSFlexGrid1.Text), 1, 12) & "','SglPrt','SERNmbr')"
                    Conn.Execute StrSql
                End If
                ChgMass = True
            End If
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 9
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            rs.Open "Select * from " & temp_tb_BOMOrigData & " Where ChildID ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "' and  ParentID ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 2)) & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                rs("ChgStatus") = Trim(MSFlexGrid1.Text)
                rs.Update
                ChgMass = True
                
                If Trim(MSFlexGrid1.Text) <> "" Then
                    MSFlexGridRowColorChange MSFlexGrid1, MSFlexGrid1.Row, MSFlexGrid1.Cols
                End If
                MSFlexGrid1.Row = RowNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
                MSFlexGrid1.Col = ColNoTemp         '��Ϊ�����ж�MSFlexGrid1�ĵ�Ԫ��Ĳ���������Ҫ�ָ�ԭ��������
            Else
                MsgBox "Failed to Write the Record into Bom Form", vbInformation, "System Info."
            End If
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 10
            If MSFlexGrid1.Row = RowNum Then
                MSFlexGrid1EditText.Visible = False
                MSFlexGrid1.Text = ""
                Exit Sub
            End If
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount = 0 Then
                If rs.State = adStateOpen Then rs.Close
                rs.Open "Select * from FinsGd Where FinsGdIndex ='" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "'", Conn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                    MsgBox "The Record is not existing in Database", vbInformation, "System Info."
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
                Else
                    If CurVersion = 1 Then
                        rs("CommtNote") = Trim(MSFlexGrid1.Text)
                        rs.Update
                    Else
                        StrSql = "INSERT INTO PartVar([BOM],[CPCN],[PartIndex],[PartValue],[TableName],[FieldName]) VALUES('" & FinishGoodsNO & "','" & CPCN & "','" & Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)) & "','" & Trim(MSFlexGrid1.Text) & "','FinsGd','CommtNote')"
                        Conn.Execute StrSql
                    End If
                    ChgMass = True
                End If
            Else
                If CurVersion = 1 Then
                    rs("CommtNote") = Trim(MSFlexGrid1.Text)
                    rs.Update
                Else
                    StrSql = "INSERT INTO PartVar([BOM],[CPCN],[PartIndex],[PartValue],[TableName],[FieldName]) VALUES('" & FinishGoodsNO & "','" & CPCN & "','" & Mid(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(RowNoTemp, 3))) - 1) & "0" & "','" & Trim(MSFlexGrid1.Text) & "','SglPrt','CommtNote')"
                    Conn.Execute StrSql
                End If
                ChgMass = True
            End If
            If rs.State = adStateOpen Then rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case Else
        End Select
        
        If MSFlexGrid1.Row < RowNum Then                                         '��Ҫ��һ���ж��Ƿ񳬳����ֵ
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        Else
            MSFlexGrid1EditText.Visible = False
            Exit Sub
        End If
        MSFlexGrid1EditText.Visible = True
        MSFlexGrid1EditText.Width = MSFlexGrid1.CellWidth
        MSFlexGrid1EditText.Height = MSFlexGrid1.CellHeight
        MSFlexGrid1EditText.left = MSFlexGrid1.CellLeft + MSFlexGrid1.left
        MSFlexGrid1EditText.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
        MSFlexGrid1EditText.SetFocus
        MSFlexGrid1EditText.Text = MSFlexGrid1.Text
        MSFlexGrid1EditText.SelStart = 0
        MSFlexGrid1EditText.SelLength = Len(FinishGoodsNO)
        Refresh_FlexGrid_TreeView False
        MSFlexGrid1EditText.Visible = False
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmBOMAdmin:CmdDrwPathAdd"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Command1_Click
End Sub
Private Sub Command2_Click()
    Dim xConn As New ADODB.Connection
    xConn.Open connString
    
    If ChgMass And CurVersion > 1 Then
        If MsgBox("No Save BOM, would you like to save it?.", vbYesNo) = vbYes Then
            Call cmdBOMSave_Click
        Else
            ChgMass = False
        End If
    End If
    
    If BOMLock Then
        If BOMLocker = PDMUserName Or SystemAdmin = "Y" Then
            If MsgBox("The BOM is locked, would you like to unlock it now?", vbYesNo, "PDM") = vbYes Then
                StrSql = "UPDATE BOMCPCN SET IsLocked=0 WHERE BOMID='" & FinishGoodsNO & "' AND CPCNNmbr='" & txtCPCNNO.Text & "'"
                xConn.Execute StrSql
                cmdLock.Caption = "UNLOCK"
                txtCPCNNO.Enabled = True
                BOMLock = False
            End If
        End If
    End If
    
    '####�˳�bomҪ�����ʱ��#########
    If FinishGoodsNO <> "" Then DropTempTable

    If xConn.State = adStateOpen Then xConn.Close: Set xConn = Nothing
    FrmEngineeringSys.Show 0
    Unload Me
End Sub

Private Sub Form_Load()
    'Load Skin & Format Control
    '''LoadSkin Me
    
    lblMsg.Caption = vbCrLf & "�����ĵȴ������ݴ�����....." & vbCrLf & "Data processing, please wait a moment."

    If Conn.State = adStateOpen Then Conn.Close
    Conn.Open connString
    
    ' Register Our New Clipboard Format
    miClipBoardFormat = RegisterClipboardFormat("VBCodeLibTree")
    
    MSFlexGrid1.Rows = 3   '����������
    MSFlexGrid1.Cols = 12   '����������
    MSFlexGrid1.ColWidth(0) = 12 * 25 * 2
    MSFlexGrid1.ColWidth(1) = 12 * 25 * 1.8
    MSFlexGrid1.ColWidth(2) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(3) = 12 * 25 * 4.5
    MSFlexGrid1.ColWidth(4) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(5) = 12 * 25 * 2.3
    MSFlexGrid1.ColWidth(6) = 12 * 25 * 6
    MSFlexGrid1.ColWidth(7) = 12 * 25 * 3
    MSFlexGrid1.ColWidth(8) = 12 * 25 * 4.8
    MSFlexGrid1.ColWidth(9) = 12 * 25 * 4.8
    MSFlexGrid1.ColWidth(10) = 12 * 25 * 3.4
    MSFlexGrid1.ColWidth(11) = 12 * 25 * 0.01
    
    MSFlexGrid1.ColAlignment(0) = 3     '()��Ϊ�еı��
    MSFlexGrid1.ColAlignment(1) = 3
    MSFlexGrid1.ColAlignment(2) = 1
    MSFlexGrid1.ColAlignment(3) = 1
    MSFlexGrid1.ColAlignment(4) = 3
    MSFlexGrid1.ColAlignment(5) = 1
    MSFlexGrid1.ColAlignment(6) = 1
    MSFlexGrid1.ColAlignment(7) = 1
    MSFlexGrid1.ColAlignment(8) = 1
    MSFlexGrid1.ColAlignment(9) = 1
    MSFlexGrid1.ColAlignment(10) = 1
    MSFlexGrid1.ColAlignment(11) = 1

    'flexAlignLeftTop 0 ��Ԫ��������󡢶������롣
    'flexAlignLeftCenter 1 �ַ�����ȱʡ���뷽ʽ����Ԫ��������󡢾��ж��롣
    'flexAlignLeftBottom 2 ��Ԫ��������󡢵ײ����롣
    'flexAlignCenterTop 3 ��Ԫ������ݾ��С��������롣
    'flexAlignCenterCenter 4 ��Ԫ������ݾ��С����ж��롣
    'flexAlignCenterBottom 5 ��Ԫ������ݾ��С��ײ����롣
    'flexAlignRightTop 6 ��Ԫ��������ҡ��������롣
    'flexAlignRightCenter 7 ��ֵ��ȱʡ���뷽ʽ����Ԫ��������ҡ����ж��롣
    'flexAlignRightBottom 8 ��Ԫ��������ҡ��ײ����롣
    'flexAlignGeneral 9 ��Ԫ������ݰ�һ�㷽ʽ���ж��롣�ַ��������󡢾��С���ʾ�����ְ����ҡ����С���ʾ��
    
    'Load Subcon
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "Select ItemValue From SysVar Where Creator ='" & PDMUserName & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        txtSubCon.AddItem (rs(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
        rs.MoveNext
    Loop
    rs.Close
       
    MSFlexGridTileInitialize
    If IsCopy Then
        mnuCopy.Enabled = False
        mnuPaste.Enabled = True
        mnuUncopy.Enabled = True
    Else
        mnuCopy.Enabled = True
        mnuPaste.Enabled = False
        mnuUncopy.Enabled = False
    End If
    
End Sub

Private Sub MSFlexGridTileInitialize()
    MSFlexGrid1.TextMatrix(0, 0) = "Index"
    MSFlexGrid1.TextMatrix(0, 1) = "Level"
    MSFlexGrid1.TextMatrix(0, 2) = "Parent12NC"
    MSFlexGrid1.TextMatrix(0, 3) = "Child12NC"
    MSFlexGrid1.TextMatrix(0, 4) = "Quantity"
    MSFlexGrid1.TextMatrix(0, 5) = "PrtUnit"
    MSFlexGrid1.TextMatrix(0, 6) = "Description"
    MSFlexGrid1.TextMatrix(0, 7) = "ItemType"
    MSFlexGrid1.TextMatrix(0, 8) = "SER NO."
    MSFlexGrid1.TextMatrix(0, 9) = "ChgStatus"
    MSFlexGrid1.TextMatrix(0, 10) = "Note"
    MSFlexGrid1.TextMatrix(0, 11) = "Family"
End Sub

Private Sub txtCPCNNO_Change()
    If ChgMass Then
        MsgBox "No save version after modification, please save it then to continue.", vbCritical
        txtCPCNNO.Text = CPCN
    End If
End Sub

Private Sub txtCPCNNO_Click()
    If Not isApproved Then
        MsgBox "The BOM has not been approved yet.", vbInformation
        txtCPCNNO.Enabled = False
    End If
End Sub

Private Sub txtCPCNNO_LostFocus()
    On Error Resume Next
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    StrSql = "Select * from CPCN where CPCNIndex='" & txtCPCNNO.Text & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = 0 Then
        If MsgBox("Invaild CPCN Number, do you want to input it again, or Click No to Exit.", vbYesNo) = vbYes Then
            txtCPCNNO.SetFocus
            rs.Close
            Set rs = Nothing
            Exit Sub
        Else
            Call Command2_Click
        End If
    Else
        rs.Close
        StrSql = "Select * from BOMCPCN where CPCNNmbr='" & txtCPCNNO.Text & "' and BOMID='" & FinishGoodsNO & "' Order By BOMVersion Desc"
        rs.Open StrSql, Conn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount = 0 Then
            cmbBOMVersion.Text = LastVersion + 1
            CurVersion = cmbBOMVersion.Text
            ChgCPCN = True
        Else
            rs.MoveFirst
            If CurVersion <> rs("BOMVersion") Then
                MsgBox "The CPCN Number had saved as Version " & rs("BOMVersion") & ", save nothing for it.", vbCritical
                Exit Sub
            Else
                ChgCPCN = True
            End If
        End If
        rs.Close
                
        Call buildInit4Version '#########�����µ���ʱ��########
        StrSql = "SELECT * FROM SglPrt4BOMLog  WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
            rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                StrSql = "INSERT INTO " & temp_tb_SglPrt4BOMLog & " SELECT * FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
                Conn.Execute StrSql
                rs.Close
            Else
            '##############�ѱ������д����־��ʱ��##################
                With MSFlexGrid1
                    For i = 2 To .Rows - 2
                        If Trim(.TextMatrix(i, 2)) <> "" Then
                            '�������µ��޸���־
                            StrSql = "IF NOT EXISTS(SELECT * FROM  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CStr(CurVersion) & " And ParentID='" & .TextMatrix(i, 2) & "' And ChildID='" & .TextMatrix(i, 3) & "' And Family='" & .TextMatrix(i, 11) & "' And ChgStatus='" & .TextMatrix(i, 9) & "') "
                            StrSql = StrSql & "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,CommtNote,Family) Values("
                            StrSql = StrSql & "" & FinishGoodsNO
                            StrSql = StrSql & "," & i + J
                            StrSql = StrSql & "," & .TextMatrix(i, 2)
                            StrSql = StrSql & "," & .TextMatrix(i, 3)
                            StrSql = StrSql & "," & CStr(CurVersion)
                            StrSql = StrSql & ",'" & .TextMatrix(i, 4)
                            StrSql = StrSql & "','" & .TextMatrix(i, 5)
                            StrSql = StrSql & "','" & Replace(.TextMatrix(i, 6), "'", "''")
                            StrSql = StrSql & "','" & .TextMatrix(i, 7)
                            StrSql = StrSql & "','" & .TextMatrix(i, 8)
                            StrSql = StrSql & "','" & .TextMatrix(i, 10)
                            StrSql = StrSql & "','" & .TextMatrix(i, 11) & "')"
                            Conn.Execute StrSql
                        End If
                    Next i
                End With
            End If
    '#############д�����##############
    End If
    CPCN = Trim(txtCPCNNO.Text)
    Set rs = Nothing
    
    '#####auto lock the bom##########
    Call cmdLock_Click
    MsgBox "The BOM is locked now, please unlock it after you finish to modify the BOM.", vbInformation, "PDM"
    
End Sub

Private Sub txtNewCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OKButton_Click
End Sub



Private Sub txtNodeSglPrt12NC_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Dim str12NC As String
        
        str12NC = LeftcutStrg(txtNodeSglPrt12NC.Text)     'myNode.Key�Ǵ�tvCodeItems_MouseDown���͹����ڵ�key(ǰ�����ַ�Cϵ��)LeftcutStrgȥ��������ַ�Cϵ��
        str12NC = Mid(str12NC, 1, (Len(str12NC) - 1)) & "0"
        
        
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        
        StrSql = "select sernmbr, serlocate,drwlocate,Description,prtUnit from SglPrt where SglPrtIndex='" & str12NC & "'"

        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        
        If rs.RecordCount = 0 Then
            txtNodeDrwlocate.Text = ""
            txtSERNO.Text = ""
            txtSERlocate.Text = ""
            txtNodeDescription.Text = ""
            txtNodePrtUnit.Text = ""
            MsgBox "Invalid Single Part Number!"
            Exit Sub
        Else
            Conn.Execute ("IF NOT EXISTS (SELECT * FROM SYSVAR WHERE ITEMTYPE=" & str12NC & " AND ITEMVALUE='" & Trim(IIsNull(rs.Fields("drwlocate"))) & "' AND CREATOR='drwlocate') INSERT INTO SYSVAR(ITEMTYPE,ITEMVALUE,CREATOR) VALUES(" & str12NC & ",'" & Trim(rs.Fields("drwlocate")) & "','drwlocate')")
            txtNodeDrwlocate.Clear
            If rs2.State = adStateOpen Then rs2.Close
            rs2.Open "SELECT * FROM SYSVAR WHERE ITEMTYPE=" & left(str12NC, 11) & "0" & " And  CREATOR='drwlocate'", Conn, adOpenKeyset, adLockOptimistic
            Do While Not rs2.EOF
                txtNodeDrwlocate.AddItem rs2("ItemValue")
                rs2.MoveNext
            Loop
            rs2.Close
            txtSERNO.Text = IIsNull(rs.Fields("sernmbr"))
            txtSERlocate.Text = Trim(IIsNull(rs.Fields("serlocate")))
            txtNodeDescription.Text = IIsNull(rs.Fields("Description"))
            txtNodePrtUnit.Text = IIsNull(rs.Fields("prtUnit"))
        End If
        Set rs2 = Nothing
        rs.Close
        Set rs = Nothing
        
        FindData MSFlexGrid1, 3, txtNodeSglPrt12NC.Text
    End If
End Sub
Private Sub FindData(MshGrid As Object, gCol As Integer, TxtText As String)
    On Error Resume Next
    
    Dim gRows     As Integer
    
    For gRows = 1 To MshGrid.Rows - 1
        If MshGrid.TextMatrix(gRows, gCol) = TxtText Then Exit For
    Next gRows
    If gRows = MshGrid.Rows Then MsgBox "δ�ҵ� ", vbInformation + vbOKOnly, "��ʾ ": Exit Sub
    
    MshGrid.TopRow = gRows
    MshGrid.Row = gRows
    MshGrid.Col = 0
    MshGrid.ColSel = 0
    MshGrid.ColSel = MshGrid.Cols - 1
End Sub
Private Sub UpdateSERNmbr(ByVal SERNO As String, PartNo As String, TableName As String)
    On Error Resume Next
    Dim StrSql As String
    If LCase(TableName) = "sglprt" Then
        StrSql = "Update " & TableName & "  Set SERNmbr='" & SERNO & "' where SglPrtIndex='" & PartNo & "'"
    ElseIf LCase(TableName) = "finsgd" Then
        StrSql = "Update " & TableName & "  Set SERNmbr='" & SERNO & "' where FinsGdIndex='" & PartNo & "'"
    End If
    
    Conn.Execute StrSql
End Sub
Private Function getCPCN(ByVal bom As String, ByVal bVer As Integer) As String
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    StrSql = "Select isNull(CPCNNmbr,'') From BOMCPCN Where BOMID=" & FinishGoodsNO & " And BOMVersion=" & CStr(bVer)
    
    rs.Open StrSql, Conn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        getCPCN = rs(0)
    Else
        getCPCN = ""
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub txtSubCon_Click()
 If txtSubCon.Text = "SUBCON" Then txtSubCon.Text = ""
End Sub

Private Sub traval(ByVal nodex As Node)
        Dim i, count As Integer
        Dim ChildNode As Node
        count = nodex.Children
        sChilds = sChilds & "," & nodex.Text
        If count > 0 Then
                Set ChildNode = nodex.Child
                traval ChildNode
                
                For i = 2 To count
                    Set ChildNode = ChildNode.Next
                    traval ChildNode
                Next
        End If
End Sub

Private Sub FindNode(ByVal sKey As String, ByVal sParentKey As String)
    On Error Resume Next
    Dim i As Integer
    Dim bFind As Boolean
    For i = 1 To tvCodeItems.Nodes.count
        MsgBox tvCodeItems.Nodes(i).Key & ":" & sKey
        If left(Replace(tvCodeItems.Nodes(i).Key, "C", ""), 11) = left(Replace(sKey, "C", ""), 11) Then
            tvCodeItems.Nodes(i).Selected = True
            bFind = True
            Exit Sub
        End If
    Next
    If Not bFind Then
        For i = 1 To tvCodeItems.Nodes.count
            MsgBox tvCodeItems.Nodes(i).Key & ":" & sParentKey
            If tvCodeItems.Nodes(i).Key = sParentKey Then
                tvCodeItems.Nodes(i).Selected = True
                Exit Sub
            End If
        Next
    End If
End Sub

Private Function GetParent(pKey As String) As String
    On Error Resume Next
    Dim i As Integer
    Dim vFamily As String
    With tvCodeItems
    For i = 1 To .Nodes.count
        If pKey = "ROOT" Then Exit Function
        If .Nodes(i).Key = pKey Then
            vFamily = .Nodes(i).Parent.Text & ">" & vFamily
            GetParent = GetParent(.Nodes(i).Parent.Key) & vFamily
        End If
    Next
    End With
End Function

Private Function CheckIsApproved(ByVal FinsGdNO As String) As Boolean
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    CheckIsApproved = False
    StrSql = "Select Approver,ApproveDate From BOMSubmitApprove  Where FinsGdIndex=" & FinsGdNO & " And (ApproveDate<>'' OR ApproveDate<>NULL) And (Approver<>'' OR Approver<>NULL)"
    rs.Open StrSql, Conn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs.Fields("Approver") <> "" Or rs.Fields("Approver") <> Null Then
            CheckIsApproved = True
        Else
            CheckIsApproved = False
        End If
    Else
        CheckIsApproved = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function CheckIsApprovalForAll(aBOM() As String)
    Dim i As Integer
    CheckIsApprovalForAll = False
    For i = 0 To UBound(aBOM)
        If CheckIsApproved(aBOM(i)) Then
            CheckIsApprovalForAll = True
            Exit For
        Else
            CheckIsApprovalForAll = False
        End If
    Next
End Function


Private Function CheckIsRejected(ByVal FinsGdNO As String) As Boolean
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    CheckIsRejected = False
    StrSql = "Select Rejecter,RejectDate From BOMSubmitApprove  Where FinsGdIndex=" & FinsGdNO
    rs.Open StrSql, Conn, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        If rs.Fields("Rejecter") <> "" Or rs.Fields("Rejecter") <> Null Then
            If rs.Fields("RejectDate") <> "" Or rs.Fields("RejectDate") <> Null Then
                CheckIsRejected = True
            Else
                CheckIsRejected = False
            End If
        Else
            CheckIsRejected = False
        End If
    Else
        CheckIsRejected = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function IsNewAssembly(ByVal code As String) As Boolean
    Dim i As Integer
    
    BOMString = ""
    Call GetTopBOM(code)
    arrBOM = Split(Mid(BOMString, 2), ",")
    If UBound(arrBOM) = 0 Then
        IsNewAssembly = True
    Else
        For i = 0 To UBound(arrBOM) - 1
            If arrBOM(i) <> FinishGoodsNO Then
                IsNewAssembly = False
                Exit For
            Else
                IsNewAssembly = True
            End If
        Next
    End If
End Function


Private Function CheckCanBeRename(ByVal code As String) As Boolean
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    BOMString = ""
    Call GetTopBOM(code)
    arrBOM = Split(Mid(BOMString, 2))
    StrSql = "Select * From BOMOrigData Where ParentID='" & code & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        CheckCanBeRename = False
    Else
        If UBound(arrBOM) > 0 Then
            CheckCanBeRename = False
        Else
            CheckCanBeRename = True
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function
Private Function getLastVersion(ByVal bom As String) As Integer
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    StrSql = "Select isnull(max(bomversion),1) From BOMCPCN Where BOMID='" & bom & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockReadOnly
    getLastVersion = rs(0)
    rs.Close
    Set rs = Nothing
End Function


Private Function GetBOM(ByVal ChildId As String)
    Dim StrSql  As String
    Dim objrs As New ADODB.Recordset
    StrSql = "Select ParentId From " & temp_tb_BOMOrigData & " Where ChildId='" & ChildId & "'"
    objrs.Open StrSql, Conn, adOpenKeyset, adLockReadOnly
    If objrs.RecordCount > 0 Then
        Do While Not objrs.EOF
            GetBOM = GetBOM(objrs(0))
        objrs.MoveNext
        Loop
    Else
        GetBOM = ChildId
    End If
    objrs.Close
    Set objrs = Nothing
End Function

Private Function CheckIsInBOM(ByVal parentid As String, MF As MSFlexGrid) As Boolean
    Dim i As Integer
    With MF
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 2)) = parentid Then
                CheckIsInBOM = True
                Exit Function
            Else
                CheckIsInBOM = False
            End If
        Next
    End With
End Function

Private Sub txtSubCon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Conn.Execute "IF NOT EXISTS(Select * from SysVar where itemtype='SUBCON' and itemvalue='" & Trim(txtSubCon.Text) & "' and creator='" & PDMUserName & "') Insert into SysVar values ('SUBCON','" & Trim(txtSubCon.Text) & "','" & PDMUserName & "')"
        Conn.Execute "IF EXISTS (SELECT * FROM SUBCON WHERE FinsGDIndex=" & FinishGoodsNO & ") UPDATE SUBCON SET SUBCON='" & txtSubCon.Text & "' Where FinsGDIndex=" & FinishGoodsNO & " ELSE INSERT INTO SUBCON(FinsGDIndex,SUBCON) Values(" & FinishGoodsNO & ",'" & txtSubCon.Text & "')"
    
        'Load Subcon
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        StrSql = "Select ItemValue From SysVar Where Creator ='" & PDMUserName & "'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        txtSubCon.Clear
        Do While Not rs.EOF
            txtSubCon.AddItem (rs(0))  'UsrCtlFind�����е�0()�Ƕ�ӦPJNOIndex���ֶ����
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End Sub


Private Sub UpdateBOMVerQtyDesc(ByVal bom As String, version As Integer, MF As MSFlexGrid)
    Dim i As Integer
    Dim StrSql As String
    With MF
        For i = 2 To .Rows - 1
            If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 3) <> "" Then
                StrSql = "UPDATE  " & temp_tb_SglPrt4BOMLog & "  SET Quantity=" & (.TextMatrix(i, 4)) & ", CPCN='" & txtCPCNNO.Text & "', Description='" & (.TextMatrix(i, 6)) & "' WHERE BOM=" & bom & " AND BOMVersion=" & version & " AND ParentID='" & (.TextMatrix(i, 2)) & "' AND ChildID='" & (.TextMatrix(i, 3)) & "'"
                Conn.Execute StrSql
            End If
        Next
    End With
End Sub

Private Sub InsertBOMOrigData(ByVal bom As String)
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "SELECT [ParentID],[ChildID],[Quantity],[ChgStatus],[SubCon] FROM BOMOrigData WHERE ParentID='" & bom & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            InsertBOMOrigData rs(1)
            StrSql = "IF NOT EXISTS(SELECT * FROM " & temp_tb_BOMOrigData & " WHERE ParentID='" & CStr(rs(0)) & "' AND ChildID='" & CStr(rs(1)) & "')  INSERT INTO " & temp_tb_BOMOrigData & "([ParentID],[ChildID],[Quantity],[ChgStatus],[SubCon]) VALUES('" & CStr(rs(0)) & "','" & CStr(rs(1)) & "'," & CStr(rs(2)) & ",'" & CStr(IIf(IsNull(rs(3)), "", rs(3))) & "','" & CStr(IIf(IsNull(rs(4)), "", rs(4))) & "')"
            Conn.Execute StrSql
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub DeleteBOMOrigData(ByVal bom As String)
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = " SELECT [ParentID],[ChildID] FROM BOMOrigData WHERE ParentID='" & bom & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            DeleteBOMOrigData rs(1)
            StrSql = "UPDATE BOMOrigData SET Mark='Delete' WHERE ParentID='" & rs(0) & "' AND ChildID='" & rs(1) & "'"
            Conn.Execute StrSql
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub SaveBOMData()
    On Error Resume Next
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    If CurVersion = 1 Then Exit Sub
    Set rs.ActiveConnection = Conn
    Set rs2.ActiveConnection = Conn
    Conn.BeginTrans
    '#########����BOMOrigData#############
    '##########1.��ɾ�����########
    Call DeleteBOMOrigData(FinishGoodsNO)
    
    '##########2.��ʽɾ����Ǽ�¼##########
    StrSql = "DELETE FROM BOMOrigData WHERE Mark='Delete'"
    Conn.Execute StrSql
    StrSql = "DELETE FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
    '& " AND (CPCN IS Null OR CPCN='')"
    Conn.Execute StrSql

    '#########3.����ʱ������޸ĺ�ļ�¼#############
    StrSql = "INSERT INTO BOMOrigData([ParentID],[ChildID],[Quantity],[ChgStatus],[SubCon]) SELECT [ParentID],N''+[ChildID]+'',[Quantity],[ChgStatus],[SubCon] FROM " & temp_tb_BOMOrigData & ""
    Conn.Execute StrSql
        
    '##########4.д��BOM Version Record#############
'    StrSql = "INSERT INTO SglPrt4BOMLog([BOM],[SeqIndex],[ParentID],[ChildID],[BOMVersion],[Quantity],[PrtUnit],[Description],[ItemType],[SERNmbr],[ChgStatus],[CommtNote],[UpdateDate],[Family],[CPCN],[IsMultiBOM]) SELECT N''+[BOM]+'',[SeqIndex],N''+[ParentID]+'',N''+[ChildID]+'',[BOMVersion],[Quantity],[PrtUnit],[Description],[ItemType],[SERNmbr],[ChgStatus],[CommtNote],[UpdateDate],[Family],[CPCN],[IsMultiBOM] FROM " & temp_tb_SglPrt4BOMLog
'    StrSql = StrSql & " A WHERE NOT EXISTS(SELECT * FROM SglPrt4BOMLog B WHERE "
'    StrSql = StrSql & " B.BOM = A.BOM"
'    StrSql = StrSql & " AND B.ParentID=A.ParentID"
'    StrSql = StrSql & " AND B.ChildID=A.ChildID"
'    StrSql = StrSql & " AND B.BOMVersion=A.BOMVersion"
'    StrSql = StrSql & " AND B.ChgStatus=A.ChgStatus"
'    StrSql = StrSql & ")"
    StrSql = "INSERT INTO SglPrt4BOMLog SELECT * FROM " & temp_tb_SglPrt4BOMLog
    Conn.Execute StrSql
    
    '############���������޸�#############
    StrSql = "SELECT ParentID,ChildID,ChgStatus,Quantity FROM " & temp_tb_SglPrt4BOMLog & " WHERE ChgStatus like '%Delete-Add-Upgrade%' OR ChgStatus ='Modify'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            '################ɾ���ظ��ļ�¼################
            If rs.Fields("chgStatus") = "Modify" Then
                StrSql = "DELETE FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion & " AND ParentID=" & rs.Fields("ParentID") & " AND ChildID=" & rs.Fields("ChildID") & " AND (ChgStatus like 'Add%' OR (ChgStatus='Modify' AND Quantity<>" & rs.Fields("Quantity") & "))"
            Else
                StrSql = "DELETE FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion & " AND ParentID=" & rs.Fields("ParentID") & " AND ChildID <= '" & rs.Fields("ChildID") & " AND ChgStatus not like  '%Delete-Add-Upgrade%'"
            End If
            Conn.Execute StrSql
        rs.MoveNext
        Loop
    End If
    rs.Close
    
    '###########5.ɾ����ʱ��###############
    Call DropTempTable
    Call buildInit4Version
    StrSql = "SELECT * FROM SglPrt4BOMLog  WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        StrSql = "INSERT INTO " & temp_tb_SglPrt4BOMLog & " SELECT * FROM SglPrt4BOMLog WHERE BOM=" & FinishGoodsNO & " AND BOMVersion=" & CurVersion
        Conn.Execute StrSql
        rs.Close
    Else
    '##############�ѱ������д����־��ʱ��##################
        With MSFlexGrid1
            For i = 2 To .Rows - 2
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    '�������µ��޸���־
                    StrSql = "IF NOT EXISTS(SELECT * FROM  " & temp_tb_SglPrt4BOMLog & "  Where BOM=" & FinishGoodsNO & " And BOMVersion=" & CStr(CurVersion) & " And ParentID='" & .TextMatrix(i, 2) & "' And ChildID='" & .TextMatrix(i, 3) & "' And Family='" & .TextMatrix(i, 11) & "' And ChgStatus='" & .TextMatrix(i, 9) & "') "
                    StrSql = StrSql & "Insert into  " & temp_tb_SglPrt4BOMLog & " (BOM,SeqIndex,ParentID,ChildID,BOMVersion,Quantity,PrtUnit,Description,ItemType,SERNmbr,CommtNote,Family) Values("
                    StrSql = StrSql & "" & FinishGoodsNO
                    StrSql = StrSql & "," & i + J
                    StrSql = StrSql & "," & .TextMatrix(i, 2)
                    StrSql = StrSql & "," & .TextMatrix(i, 3)
                    StrSql = StrSql & "," & CStr(CurVersion)
                    StrSql = StrSql & ",'" & .TextMatrix(i, 4)
                    StrSql = StrSql & "','" & .TextMatrix(i, 5)
                    StrSql = StrSql & "','" & Replace(.TextMatrix(i, 6), "'", "''")
                    StrSql = StrSql & "','" & .TextMatrix(i, 7)
                    StrSql = StrSql & "','" & .TextMatrix(i, 8)
                    StrSql = StrSql & "','" & .TextMatrix(i, 10)
                    StrSql = StrSql & "','" & .TextMatrix(i, 11) & "')"
                    Conn.Execute StrSql
                End If
            Next i
        End With
    End If
    '#############д�����##############
    '############�������ϼ���Ϣ#############
    StrSql = "SELECT * FROM PartVar WHERE BOM='" & FinishGoodsNO & "' AND CPCN='" & CPCN & "' ORDER BY [Index] ASC"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            If rs.Fields("TableName") = "SglPrt" Then
                StrSql = "UPDATE SglPrt SET " & rs.Fields("FieldName") & " = '" & rs.Fields("PartValue") & "' WHERE SglPrtIndex='" & left(rs.Fields("PartIndex"), 11) & "0'"
            Else
                StrSql = "UPDATE FinsGd SET " & rs.Fields("FieldName") & " = '" & rs.Fields("PartValue") & "' WHERE FinsGdIndex=" & rs.Fields("PartIndex")
            End If
            Conn.Execute StrSql
        rs.MoveNext
        Loop
    End If
    rs.Close
    
    '############��������BOM���ϰ汾##############
    StrSql = "SELECT ChildID FROM " & temp_tb_BOMOrigData
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            '���¸������
            Conn.Execute "UPDATE BOMOrigData SET ParentID = '" & rs.Fields("ChildID") & "' WHERE LEFT(ParentID,11)='" & left(rs.Fields("ChildID"), 11) & "'"
            '�����Ӽ����
            Conn.Execute "UPDATE BOMOrigData SET ChildID = '" & rs.Fields("ChildID") & "' WHERE LEFT(ChildID,11)='" & left(rs.Fields("ChildID"), 11) & "'"
        rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    If Err = 0 Then
        Conn.CommitTrans
        MsgBox "BOM save successfully.", vbInformation, "PDM Database"
        ChgMass = False
    Else
        Conn.RollbackTrans
        MsgBox "BOM save failure, please try again." & vbCrLf & "[" & Err.Number & "]:" & Err.Description, vbCritical, "PDM Database"
    End If
        
End Sub

Private Function getPartValue(ByVal tn As String, ByVal fn As String, ByVal pi As String, ByVal fv As String) As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn

    StrSql = "SELECT TOP 1 PartValue FROM PartVar  WHERE BOM='" & FinishGoodsNO & "' AND CPCN='" & CPCN & "' AND TableName='" & tn & "' AND [FieldName]='" & fn & "' AND PartIndex ='" & pi & "' ORDER BY [Index] DESC"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        getPartValue = IIf(IsNull(rs(0)), "", CStr(rs(0)))
    Else
        getPartValue = fv
    End If
    rs.Close
    Set rs = Nothing
    If Err <> 0 Then
        getPartValue = ""
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Function


Private Sub buildInit4Version()
    On Error Resume Next
    Dim userid As Integer
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    If CurVersion = 1 Then '#########��ʼ�汾���ô洢��ʱ��#########
        temp_tb_BOMOrigData = "BOMOrigData"
        temp_tb_SglPrt4BOMLog = "SglPrt4BOMLog"
    Else  '#########�ǳ�ʼ�汾�Ȱ��޸����ݴ洢����ʱ���ȷ���޸ĺ���д����ʽ��#########
        StrSql = "SELECT [UserId] FROM [Users] WHERE [Name]='" & PDMUserName & "'"
        rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            userid = rs.Fields("UserId")
        Else
            MsgBox "System Failed, pls contact admin.", vbCritical, "PDM"
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        temp_tb_BOMOrigData = "BOMOrigData_" & FinishGoodsNO & "_" & CStr(userid)
        temp_tb_SglPrt4BOMLog = "SglPrt4BOMLog_" & FinishGoodsNO & "_" & CStr(userid)
        '########################������ʱ��########################
        StrSql = "sp_create_temp_table '" & FinishGoodsNO & "_" & CStr(userid) & "'"
        Conn.Execute StrSql
        '#######################д����ʱ��#########################
        InsertBOMOrigData FinishGoodsNO
    End If
    If Err <> 0 Then
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Sub

Private Sub DropTempTable()
    On Error Resume Next
    '####��bomҪ�����ʱ��#########
    If FinishGoodsNO <> "" Then
        If temp_tb_BOMOrigData <> "" And temp_tb_BOMOrigData <> "BOMOrigData" Then
            StrSql = "DROP TABLE " & temp_tb_BOMOrigData
            Conn.Execute StrSql
        End If
        
        If temp_tb_SglPrt4BOMLog <> "" And temp_tb_SglPrt4BOMLog <> "SglPrt4BOMLog" Then
            StrSql = "DROP TABLE " & temp_tb_SglPrt4BOMLog
            Conn.Execute StrSql
        End If
    End If
    If Err <> 0 Then
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Sub

Private Function getBOMLocker() As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "SELECT TOP 1 ISNULL(IsLocked,0) AS IsLocked,ISNULL(Locker,'') AS Locker FROM BOMCPCN WHERE BOMID='" & FinishGoodsNO & "' ORDER BY BOMVersion DESC"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If rs.Fields("IsLocked") = True Then
            BOMLock = True
        Else
            BOMLock = False
        End If
        getBOMLocker = rs.Fields("Locker")
        rs.Close
        Set rs = Nothing
    Else
        BOMLock = False
        getBOMLocker = ""
    End If
    If Err <> 0 Then
        getBOMLocker = ""
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Function

Private Function IsBOMLocked() As Boolean
    On Error Resume Next
    Dim locker As String
    locker = getBOMLocker
    If BOMLock Then
        If locker = PDMUserName Or SystemAdmin = "Y" Then
            IsBOMLocked = False
        Else
            IsBOMLocked = True
        End If
    Else
        IsBOMLocked = False
    End If
    If Err <> 0 Then
        IsBOMLocked = False
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Function

Private Function getIsFG(ByVal fgno As String) As Boolean
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    StrSql = "SELECT * FROM FinsGd WHERE FinsGdIndex='" & fgno & "'"
    rs.Open StrSql, Conn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        getIsFG = True
    Else
        getIsFG = False
    End If
    rs.Close
    Set rs = Nothing
    If Err <> 0 Then
        getIsFG = False
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
End Function


Private Function checkModifyPermission(txtcode As String) As Boolean
    On Error Resume Next
    checkModifyPermission = True
    If txtcode <> FinishGoodsNO Then
        If Not IsNewAssembly(txtcode) Then
            If CheckIsApprovalForAll(arrBOM) Then
                MsgBox "The Assembly Part can't change, because it used to other formal BOMs.", vbCritical
                checkModifyPermission = False
            End If
        End If
    End If
    If Err <> 0 Then
        checkModifyPermission = False
        MsgBox Err.Number & ":" & Err.Description & vbCrLf & "Something Error, Please contact system admin."
    End If
        
End Function

