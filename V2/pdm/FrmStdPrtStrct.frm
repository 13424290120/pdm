VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmStdPrtLibStructr 
   Caption         =   "PDM-Standard Part Lib Structure Admin ���̹�����ϵͳ"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   -1605
   ClientWidth     =   13710
   Icon            =   "FrmStdPrtStrct.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   13710
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdRemoveCategory 
      Caption         =   "Remove Category / C1-5"
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
      Left            =   105
      TabIndex        =   37
      Top             =   10030
      Width           =   2220
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
      Left            =   6555
      TabIndex        =   30
      Top             =   10030
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
      Left            =   4575
      TabIndex        =   29
      Top             =   10030
      Width           =   1425
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
      Left            =   2820
      TabIndex        =   28
      Top             =   10030
      Width           =   1215
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
      Left            =   8310
      TabIndex        =   27
      Top             =   10030
      Width           =   1215
   End
   Begin VB.TextBox txtPage 
      Height          =   375
      Left            =   12600
      TabIndex        =   26
      Top             =   10030
      Width           =   975
   End
   Begin VB.TextBox txtPage_nd 
      Height          =   375
      Left            =   9990
      TabIndex        =   25
      Top             =   10030
      Width           =   735
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
      Left            =   10710
      TabIndex        =   24
      Top             =   11235
      Width           =   375
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
      Left            =   11340
      TabIndex        =   23
      Top             =   10030
      Width           =   1215
   End
   Begin VB.CommandButton CmdunLoadImage 
      Caption         =   "UnLoad Image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5490
      TabIndex        =   3
      Top             =   6300
      Width           =   1755
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7515
      TabIndex        =   2
      Top             =   6300
      Width           =   1755
   End
   Begin VB.CommandButton CmdLoadImage 
      Caption         =   "Load Image"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3450
      TabIndex        =   1
      Top             =   6300
      Width           =   1755
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6615
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   11668
      _Version        =   393217
      Indentation     =   88
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9060
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3045
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":08CA
            Key             =   "NEW"
            Object.Tag             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":09DC
            Key             =   "FILE"
            Object.Tag             =   "FILE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":2C8E
            Key             =   "CHILD"
            Object.Tag             =   "CHILD"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":2FA8
            Key             =   "FOLDER"
            Object.Tag             =   "FOLDER"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":32C2
            Key             =   "DELETE"
            Object.Tag             =   "DELETE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":33D4
            Key             =   "OPENFOLDER"
            Object.Tag             =   "OPENFOLDER"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":36EE
            Key             =   "SETTINGS"
            Object.Tag             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":3A08
            Key             =   "PREVIOUS"
            Object.Tag             =   "PREVIOUS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":3D5A
            Key             =   "NEXT"
            Object.Tag             =   "NEXT"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":40AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":44FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":4950
            Key             =   "BAS"
            Object.Tag             =   "BAS"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":4CA2
            Key             =   "CLS"
            Object.Tag             =   "CLS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":4F64
            Key             =   "VB"
            Object.Tag             =   "VB"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":52B6
            Key             =   "VIEWBOOKMARKS"
            Object.Tag             =   "VIEWBOOKMARKS"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":55D0
            Key             =   "ADDBOOKMARK"
            Object.Tag             =   "ADDBOOKMARK"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":58EA
            Key             =   "OPEN"
            Object.Tag             =   "OPEN"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":59FC
            Key             =   "PRINT"
            Object.Tag             =   "PRINT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStdPrtStrct.frx":5B0E
            Key             =   "FIND"
            Object.Tag             =   "FIND"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standard Part Lib Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   9720
      TabIndex        =   4
      Top             =   30
      Width           =   3900
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2550
         TabIndex        =   36
         Top             =   0
         Width           =   1245
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
         Height          =   255
         Left            =   90
         TabIndex        =   35
         Top             =   1890
         Width           =   1125
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
         Height          =   255
         Left            =   90
         TabIndex        =   34
         Top             =   1440
         Width           =   1125
      End
      Begin VB.TextBox TxtDrwlocate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   32
         Top             =   1620
         Width           =   2325
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2640
         TabIndex        =   22
         Top             =   6015
         Width           =   1125
      End
      Begin VB.CommandButton CmdRegister 
         Caption         =   "Register one Data into Standard Part Library"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   135
         TabIndex        =   21
         Top             =   6015
         Width           =   2340
      End
      Begin VB.TextBox TxtMainChrct5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   19
         Top             =   5580
         Width           =   2325
      End
      Begin VB.TextBox TxtMainChrct4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   17
         Top             =   4935
         Width           =   2325
      End
      Begin VB.TextBox TxtMainChrct3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   15
         Top             =   4275
         Width           =   2325
      End
      Begin VB.TextBox TxtMainChrct2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   13
         Top             =   3630
         Width           =   2325
      End
      Begin VB.TextBox TxtMainChrct1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   11
         Top             =   2970
         Width           =   2325
      End
      Begin VB.TextBox TxtStandrdPrtCateg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   9
         Top             =   2325
         Width           =   2325
      End
      Begin VB.TextBox TxtDescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         TabIndex        =   8
         Top             =   930
         Width           =   2325
      End
      Begin VB.TextBox TxtSglPrtIndex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1185
         TabIndex        =   6
         Top             =   315
         Width           =   2325
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000080FF&
         Height          =   255
         Left            =   120
         Top             =   0
         Width           =   2235
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   45
         Top             =   1425
         Width           =   3540
      End
      Begin VB.Label Lbl2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Drawing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   33
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label Lbl8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   20
         Top             =   5655
         Width           =   240
      End
      Begin VB.Label Lbl7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   18
         Top             =   5010
         Width           =   240
      End
      Begin VB.Label Lbl6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   16
         Top             =   4350
         Width           =   240
      End
      Begin VB.Label Lbl5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   14
         Top             =   3705
         Width           =   240
      End
      Begin VB.Label Lbl4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   12
         Top             =   3045
         Width           =   240
      End
      Begin VB.Label Lbl3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   10
         Top             =   2370
         Width           =   855
      End
      Begin VB.Label Lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrip."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   7
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Lbl0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item12NC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3095
      Left            =   150
      TabIndex        =   31
      Top             =   6780
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   5450
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
      ColumnCount     =   11
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
      BeginProperty Column04 
         DataField       =   "StandrdPrtCateg"
         Caption         =   "Category"
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
         DataField       =   "MainChrct1"
         Caption         =   "C1"
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
         DataField       =   "MainChrct2"
         Caption         =   "C2"
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
         DataField       =   "MainChrct3"
         Caption         =   "C3"
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
         DataField       =   "MainChrct4"
         Caption         =   "C4"
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
         DataField       =   "MainChrct5"
         Caption         =   "C5"
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
         DataField       =   "Drwlocate"
         Caption         =   "Drwlocate"
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
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   3120
      Top             =   120
      Width           =   6420
   End
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   3135
      Stretch         =   -1  'True
      Top             =   315
      Width           =   6420
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCode 
         Caption         =   "&Add New Item Here"
      End
      Begin VB.Menu mnuDeleteCode 
         Caption         =   "&Delete Selected Item"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename Item Here"
      End
   End
End
Attribute VB_Name = "FrmStdPrtLibStructr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ʹ��Stream���󣬿���ʵ�ֶ����ݿ��ͼ���ȡ��
'���ݿ��д��ͼ����ֶ�����image��AccessΪOLE���ͣ���
'���磬����á�CommonDialog���ؼ���ѡ����Ӳ���ϵ�ͼ���ļ���
'�á�Picture���ؼ�����ʾͼ����ô����Ĵ��빩�ο���
'������VB��ѡ�񡰹���\���á�������á�Microsoft ActiveX Date 2.5 Library�������������ݿ⣬������Ӧ�ļ�¼��rs��
' CommonDialog �ؼ��� Visual Basic �� Microsoft Windows ��̬���ӿ�Commdlg.dll ����֮���ṩ�˽ӿڡ�Ϊ���øÿؼ������Ի��򣬱���Ҫ��Commdlg.dll �� Microsoft Windows \System Ŀ¼�¡�
'��δ��� CommonDialog �ؼ�����Ӧ�ӡ����̡��˵���ѡ���������������ؼ�Microsoft Common Dialoge Control 6.0 ��ӵ��������С��ڱ�ǶԻ��ġ��ؼ������ҵ���ѡ���ؼ���Ȼ�󵥻���ȷ������ť��
'CommonDialog �ؼ�������ʾ���³��öԻ���showopen���򿪡�,  showsave�����Ϊ��,  showcolor����ɫ��, showfont �����塱,showprinter ����ӡ��,showhelp��windows������
'�������и�ʽ���� Filter ���ԣ�
'description1 | filter1 | description2 | filter2...
'Description ���б������ʾ���ַ����������磬"Text Files (*.txt)"��Filter ��ʵ�ʵ��ļ��������������磬"*.txt"��ÿ��description | filter ���ü�����ùܵ����ŷָ� (|)��
'Private Sub mnuFileOpen_Click()
     'CancelError Ϊ True��
'     On Error GoTo ErrHandler
     '���ù�������
'     CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
     'ָ��ȱʡ��������
'     CommonDialog1.FilterIndex = 2
     '��ʾ���򿪡��Ի���
'     CommonDialog1.ShowOpen
     '���ô��ļ��Ĺ��̡�
'     OpenFile (CommonDialog1.FileName)
'     Exit Sub
'ErrHandler:
     '�û�����ȡ������ť��
'     Exit Sub
'End Sub
'����Ϊʾ������
Option Explicit
Private SourceNode As Object  '����ڵ���ҷ��Դ�ڵ�
Private SourceNodeParent As String
Private DataGridSearchString As String
Private AddNodeOk As Boolean   '����ӽڵ�ɹ���ı��
Private lCurrentpage As Long           '���嵱ǰҳ����

'�жϽڵ�A�����ַ�����TreeView���Ƿ���ͬ�ڵ����Ľڵ����&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeKeyNameNodeExist(NodeAString As String) As Boolean
'����TreeView�����нڵ�ı�����
   Dim nodEachChild As Node
   For Each nodEachChild In TreeView1.Nodes
        If nodEachChild = NodeAString Then
            AddNodeKeyNameNodeExist = True
            Exit Function
       End If
   Next nodEachChild
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

Private Sub CmdClear_Click()
TxtSglPrtIndex = ""
TxtDescription = ""
TxtDrwlocate = ""
TxtStandrdPrtCateg = ""
TxtMainChrct1 = ""
TxtMainChrct2 = ""
TxtMainChrct3 = ""
TxtMainChrct4 = ""
TxtMainChrct5 = ""

End Sub

Private Sub CmdExit_Click()
    Unload Me
    FromForm.Show 0
End Sub

Private Sub CmdRemoveCategory_Click()
On Error GoTo vbErrorHandler
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    Dim SglPrtIndex2Clr As String
    
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
                  
With DataGrid1
.Col = 0
SglPrtIndex2Clr = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
End With

            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(SglPrtIndex2Clr), 1, Len(Trim(SglPrtIndex2Clr)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
                 MsgBox "The 12NC is not existing in Single Part Database", vbInformation, "System Info."
                 If rs.State = adStateOpen Then rs.Close
                 Exit Sub
               Else
                 rs("StandrdPrtCateg") = ""
                 rs("MainChrct1") = ""
                 rs("MainChrct2") = ""
                 rs("MainChrct3") = ""
                 rs("MainChrct4") = ""
                 rs("MainChrct5") = ""
                 rs.Update
               End If
            If rs.State = adStateOpen Then rs.Close
       
        Conn.Close
        CmdFresh_Click             'ˢ��һ��
Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:CmdRemoveCategory"
End Sub

Private Sub CmdSearch_Click()
Dim C12345 As String

If TxtStandrdPrtCateg.Text = "" Then
   MsgBox "At least, You need to set a Category to Search", vbInformation, "System Info"
   Exit Sub
End If

DataGridSearchString = "select * from SglPrt Where StandrdPrtCateg = '" & TxtStandrdPrtCateg & "'"

C12345 = ""
If Trim(TxtMainChrct1.Text) <> "" Then
C12345 = C12345 & " And MainChrct1 = '" & TxtMainChrct1 & "'"
End If
   If Trim(TxtMainChrct2.Text) <> "" Then
     C12345 = C12345 & " And MainChrct2 = '" & TxtMainChrct2 & "'"
   End If
        If Trim(TxtMainChrct3.Text) <> "" Then
          C12345 = C12345 & " And MainChrct3 = '" & TxtMainChrct3 & "'"
        End If
             If Trim(TxtMainChrct4.Text) <> "" Then
                C12345 = C12345 & " And MainChrct4 = '" & TxtMainChrct4 & "'"
             End If
                   If Trim(TxtMainChrct5.Text) <> "" Then
                     C12345 = C12345 & " And MainChrct5 = '" & TxtMainChrct5 & "'"
                   End If

DataGridSearchString = DataGridSearchString & C12345
lCurrentpage = 1
Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me

DataGrid1.AllowDelete = False
DataGrid1.AllowAddNew = False
DataGrid1.AllowUpdate = False
 
lCurrentpage = 1           '���ڴ�Ĭ���ǵ�1ҳ����
 
CommonDialog1.Filter = "All Files (*.*)|*.*|Jpg Files (*.jpg)|*.jpg|Bmp Files (*.bmp)|*.bmp"
'ָ��ȱʡ��������
CommonDialog1.FilterIndex = 2

FillTree
End Sub

Private Sub CmdLoadImage_Click()
' ���á�CancelError��Ϊ True
CommonDialog1.CancelError = True       '��CancelError��������Ϊ True ʱ�����ۺ�ʱѡȡ��ȡ������ť�������� 32755 (cdlCancel) �Ŵ���
On Error GoTo ErrHandler

    Dim StmPic  As ADODB.Stream

    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Set StmPic = New ADODB.Stream
        CommonDialog1.ShowOpen
        Dim tempNodekey As String
        tempNodekey = SourceNode.Key     'myNode.Key�Ǵ�TreeView1_MouseDown���͹����ڵ�key
        rs.Open "Select * from StdPrtLibStructr where ChildID ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
          StmPic.Type = adTypeBinary             'ָ�����Ƕ���������
          StmPic.Open                            '�����ݻ�ȡ��Stream������
          StmPic.LoadFromFile (CommonDialog1.Filename)     '��ѡ���ͼ����ص��򿪵�StmPic��
          rs.Fields("StdPrtImage") = StmPic.Read           '��StmPic�����ж�ȡ����
          rs.Update
          StmPic.Close
       End If
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
       Conn.Close
    CmdLoadImage.Enabled = False   '����Load Image��ť��λ�ɲ�����״̬
Exit Sub

ErrHandler:
' �û����ˡ�ȡ������ť
Exit Sub
End Sub

Private Sub CmdunLoadImage_Click()
On Error GoTo ErrHandler
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ж��ͼƬ
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Unload", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ж��ͼƬ
    
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
        
        Dim tempNodekey As String
        tempNodekey = SourceNode.Key     'myNode.Key�Ǵ�TreeView1_MouseDown���͹����ڵ�key
        rs.Open "Select * from StdPrtLibStructr where ChildID ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
          rs.Fields("StdPrtImage") = Null         '�������ΪNull
          rs.Update
        End If
        If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
        Conn.Close
    CmdLoadImage.Enabled = True   '����unLoad Image��ť���óɿ���״̬
Exit Sub

ErrHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:CmdunLoadImage"
End Sub

Private Sub Form_Resize()
        'ȷ������ı�ʱ�ؼ���֮�ı�
        Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    FromForm.Show 0
End Sub

Private Sub Frame1_DblClick()
' ���á�CancelError��Ϊ True
CommonDialog1.CancelError = True       '��CancelError��������Ϊ True ʱ�����ۺ�ʱѡȡ��ȡ������ť�������� 32755 (cdlCancel) �Ŵ���
On Error GoTo ErrHandler

    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    Dim rs2 As New ADODB.Recordset
    Set rs2.ActiveConnection = Conn
    
         If MsgBox("It's going to search all drawings in one specified Folder." & vbCrLf & "And add their PathName into Single Part Library " & vbCrLf & "Are You Sure to Continue?", vbYesNo + vbDefaultButton2, "Confirm to add File PathName") = vbNo Then
            Exit Sub
         End If
                
        Dim eachfilename As String
        Dim eachfilename_12NC As String
        Dim CurrentFilePath As String
        
        CommonDialog1.ShowOpen   'Dir֧�ֶ��ַ� (*) �͵��ַ� (?) ��ͨ�����ָ�������ļ�
        CurrentFilePath = left(CommonDialog1.Filename, InStrRev(CommonDialog1.Filename, "\") - 1)
        eachfilename = Dir(CurrentFilePath & "\" & "*.*")  '�ڵ�һ�ε���Dir(pathname)����ʱ,����ָ��pathname,������������Dir �᷵��ƥ�� pathname �ĵ�һ���ļ���.  vbDirectory ֵΪ16 ָ���������ļ�����·�����ļ��С�
        
        Do While Len(eachfilename) > 0
          
          eachfilename_12NC = (Mid(Replace(Trim(eachfilename), " ", ""), 1, 11) & "0")        '(Replace(Trim(eachfilename), " ", "") ȥ���ַ����м�Ŀո�
          rs.Open "Select * from SglPrt Where SglPrtIndex ='" & eachfilename_12NC & "'", Conn, adOpenKeyset, adLockOptimistic
          If rs.RecordCount > 0 Then
            If Not IsNull(rs("Drwlocate")) Then
              If MsgBox("The single part " & eachfilename_12NC & " already has pathname." & vbCrLf & "Are You Sure to Override existing pathname?", vbYesNo + vbDefaultButton2, "Confirm to override PathName") = vbNo Then
              GoTo Nextfilepathname
              End If
            End If
          Else
              If MsgBox("The single part " & eachfilename_12NC & " is NOT existing in Single Part Lib." & vbCrLf & "Are You going to add Single Part 12NC before adding pathname?", vbYesNo + vbDefaultButton2, "Confirm to add new Single Part 12NC") = vbNo Then
              GoTo Nextfilepathname
              Else
                rs2.Open "SglPrt", Conn, adOpenKeyset, adLockOptimistic
                rs2.AddNew
                  rs2!SglPrtIndex = eachfilename_12NC
                  rs2!SglPrtVer = 1                              '�̶�����ֵ
                  rs2!PrtUnit = "Piece"                          '�̶�����ֵ
                  rs2!Applicant = "NA"                              '�̶�����ֵ
                  rs2!Description = InputBox("Please Input Description", "Input Part Name ", "Frame TYoke Washer ...", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
                  rs2!IDSO = "Open"                              '�̶�����ֵ
                  rs2!NewOldStatus = "Old"                       '�̶�����ֵ
                  rs2!OpnDate = Date
                  rs2!ClosDate = Date
                  rs2!PJNOIndex = InputBox("Please Input Project Number" & vbCrLf & "Must be 6 Arabic numerals", "Input Project Number ", "120300 115360 ...", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
                  rs2!PjtName = "NA"                             '�̶�����ֵ
                  rs2!ProductLine = "5000"                       '�̶�����ֵ
                  rs2!ItemType = InputBox("Please Input Item Type" & vbCrLf & "Must be 3 Arabic numerals", "Input Item Type ", "050 070 100 ...", 10000, 1)    'InputBox (Message, Title, Default,����10000, 1�Ǵ��ڳ���λ��)
                  rs2!Location = "TR-AV"                         '�̶�����ֵ
                  rs2!CommtNote = "NA"                           '�̶�����ֵ
                 rs2.Update
                If rs2.State = adStateOpen Then rs2.Close   'ע����������State,����status  adStateOpenֵΪ1
                'ע��������Ҫ�ȹر�rs �����°��������new 12NC��rs
                If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
                rs.Open "Select * from SglPrt Where SglPrtIndex ='" & eachfilename_12NC & "'", Conn, adOpenKeyset, adLockOptimistic
              End If
          End If
          rs("Drwlocate") = CurrentFilePath & "\" & eachfilename
          rs.Update
Nextfilepathname:
       eachfilename = Dir  '��һ�ε���Dir,��Ҫʹ�ò����������û�кϺ��������ļ�,�� Dir �᷵��һ���㳤���ַ��� ("")��һ������ֵΪ�㳤���ַ�����Ҫ�ٴε���Dirʱ,�ͱ���ָ��pathname,������������
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
       Loop
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
       Conn.Close
Exit Sub

ErrHandler:
' �û����ˡ�ȡ������ť
Exit Sub
End Sub

'ɾ��һ���ڵ�Ĳ���(Check OK)
Private Sub mnuDeleteCode_Click()
' Delete the selected CodeItem and all it's children

On Error GoTo vbErrorHandler
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@�ж��Ƿ��ǹ���Ա�û���������ɾ��
    
    Dim sKey As String
    Dim oNode As Node
    Dim oParentNode As Node
    Dim sMessage As String
        
    Set oNode = TreeView1.SelectedItem
    sKey = oNode
    If sKey = "InputNew12NC" Then        '�����InputNew12NC,��ʾʵ��ֻ����ɾ��һ��treeview����ʱ�ڵ�,��Standard Part Library�и����޼�¼��
    FillTree
    Exit Sub
    End If
        
    If oNode.Key = "ROOT" Then Exit Sub     '����Ǹ��ڵ����˳�ɾ������
    
       If oNode.Parent.Key = "ROOT" And oNode.Parent.Children = 1 Then       '����ڵ㸸�ڵ���ROOT����ROOTֻ��һ��Child(��ѡ�е�����ڵ�)
           MsgBox "Delete Final Record in Library, Library will not Exist", vbInformation, "System Info"
           Exit Sub
       End If
       
    sMessage = "Delete selected Code "
    If oNode.Children > 0 Then    '����нڵ�ѡ�в������ӽڵ���Ҫ�в�ͬ��ʾ��Children�ǽڵ���ӽڵ�����
        sMessage = sMessage & "and all child records ?"
    Else
        sMessage = sMessage & "?"
    End If
    
      If MsgBox(sMessage, vbYesNo + vbExclamation, "Delete Category Record") = vbNo Then
         Exit Sub
      End If
      
    Set oParentNode = oNode.Parent  '��ǰѡ�е�Ҫɾ���Ľڵ�ĸ��ڵ㸳ֵ
    SourceNodeParent = oParentNode
    
    DeleteCodeItem SourceNodeParent, sKey
    FillTree
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:mnuDeleteCode"
End Sub

'ɾ��һ���ڵ�Ĳ���(Check OK)
Private Sub DeleteCodeItem(ParentNodeKey As String, ChildNodeKey As String)

On Error GoTo vbErrorHandler
 
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
            
    'ɾ��Դ�ڵ���ӽڵ�����
    rs.Open "Select * from StdPrtLibStructr Where ParentID ='" & ChildNodeKey & "'", Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
       DeleteCodeItem ChildNodeKey, rs("ChildID")      '�ݹ�����ҳ����в㼶���ӽڵ�����
       rs.MoveNext
    Loop
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    'ɾ��Դ�ڵ㱾�����һ������
    rs.Open "Delete from StdPrtLibStructr Where ChildID='" & ChildNodeKey & "'" & " and  ParentID ='" & ParentNodeKey & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.State = adStateOpen Then rs.Close
    Conn.Close
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:DeleteCodeItem"
End Sub

'����һ���ڵ�Ĳ���(Check OK)
Private Sub mnuAddCode_Click()
AddCode
End Sub

'����һ���ڵ�Ĳ���(Check OK)
Private Sub AddCode()

On Error GoTo vbErrorHandler
    Dim AddNodeParentImage As String
    Dim nNode As Node
    Dim nParentNode As Node
    'Dim sParentKey As String
    
    Set nNode = TreeView1.SelectedItem   'TreeView��ѡ�еĽڵ㸳ֵ��nNode
    
    If nNode.Key <> "ROOT" Then                        '�ж�������Ǹ��ڵ�Ļ�
        Set nParentNode = TreeView1.Nodes(nNode.Key)        'Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ㿪ʼ��һ�����ڵ�
        SourceNodeParent = nParentNode
        AddNodeParentImage = nParentNode.Image       '����Ҫ����һ���ڵ�Ĳ����б����ӵĽڵ��ԭͼ��
        nParentNode.Image = "FOLDER"
        nParentNode.ExpandedImage = "FOLDER"
   'ExpandedImage���Է��ػ������ڹ�����ImageList�ؼ��е�ListImage������������ֵ����Node����չ��ʱ��ʾ ListImage ����
    End If
    
    Set nNode = TreeView1.Nodes.Add(TreeView1.SelectedItem, tvwChild, "InputNewItem", "InputNewItem", "CHILD")
    Set TreeView1.SelectedItem = nNode       '�任ѡ�еĽڵ�(������ʾ)�ӱ����ӵĽڵ� ��Ϊ�ո���ӵĽڵ�
    nNode.EnsureVisible
    
    AddNodeOk = True                   '�ȼ����ǿ��Ա༭OK��
    TreeView1.StartLabelEdit    'StartLabelEdit���������û��༭��ǩ��
' �� LabelEdit ��������Ϊ 1���ֶ���ʱ�������� StartLabelEdit ����������һ��ǩ�༭������
' ��һ�����ϵ��� StartLabelEdit ����ʱ��BeforeLabelEdit �¼�Ҳͬʱ������
   
    If Not AddNodeOk Then
    TreeView1.Nodes.Remove nNode.Key
    nParentNode.Image = AddNodeParentImage
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmStdPrtLibStructr:AddCode"

End Sub

'����һ���ڵ�(��)�Ĳ���(Check OK)
Private Sub mnuRename_Click()
' Change the Label - remember, we only allow 12 Characters
    TreeView1.StartLabelEdit
End Sub

'����һ���ڵ��ǩ��ǰԤ������(Check OK)
Private Sub tvCodeItems_BeforeLabelEdit(Cancel As Integer)
'���û�в���Ҳ��Ҫɾ��,��������AfterLabelEdit����
End Sub

'����һ���ڵ��ǩ�ĺ�������(Check OK)
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo vbErrorHandler
    Dim sKey As String
    Dim sKeyb4Rename As String
    Dim oNode As Node
    Dim oParentNode As Node
    Dim sMessage As String
    
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    

        If Len(NewString) = 0 Then
          MsgBox "You must enter a new Name for the new item", vbInformation, "System Info."
          Cancel = True
          TreeView1.Nodes.Remove "InputNewItem"
          Exit Sub
        End If
        
    Set oNode = TreeView1.SelectedItem      'TreeView1.SelectedItem�����������һ����Ҫ�����Ľڵ㣬��һ���������ӵĽڵ�
    sKeyb4Rename = oNode
    
        If oNode.Key = "ROOT" Then Exit Sub     '����Ǹ��ڵ����˳���������
        
    If oNode Is Nothing Then      '�����û�нڵ�ѡ������ʾ���˳���������
        MsgBox "No Selected Record", vbInformation, "System Info"
        Exit Sub
    End If
    
    Set oParentNode = oNode.Parent  '��ǰѡ�е�Ҫ�����Ľڵ�ĸ��ڵ㸳ֵ
    SourceNodeParent = oParentNode
    
    
       If sKeyb4Rename = "InputNewItem" Then GoTo handleAddcode   'ֱ��ȥ�������ڵ�Ĳ���
    
rs.Open "Select * from StdPrtLibStructr Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'", Conn, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then   '������жϸ����Ľڵ���Add code�ո����ӵ�(ֻ��TreeView�пɼ�,��Standard Part Libû�м�¼��)����˵����Standard Part Lib(�м�¼��)
       sMessage = "Rename selected Code ?"
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
       If MsgBox(sMessage, vbYesNo + vbExclamation, "Rename Code Record") = vbNo Then
          Exit Sub
       End If
       
      sKey = NewString
      
       If AddNodeKeyNameNodeExist(sKey) Then  '�ж������NewString�ڴ�BOM(TreeView)���Ƿ�Ϊһ���Ѿ����ڵĽڵ���
         MsgBox "This New Name Item already Exist, Can NOT Rename.", vbInformation, "System Info"
         Exit Sub
       End If
      
       rs.Open "Select * from StdPrtLibStructr Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'", Conn, adOpenKeyset, adLockOptimistic
       rs("ChildID") = Trim(NewString)    '����(����)Ҫ�����Ľڵ���Ϊ�ӽڵ����һ����Standard Part Lib�еļ�¼
       rs.Update
    
         If rs.State = adStateOpen Then rs.Close
         Set rs = Nothing
         
             '���´�Category�������е�SinglePart��StandrdPrtCateg�ֶ�ȫ����Ϊ��ֵ
             rs.Open "Select * from SglPrt Where StandrdPrtCateg ='" & sKeyb4Rename & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount > 0 Then
                 rs.MoveFirst
                 Do While Not rs.EOF
                    rs("StandrdPrtCateg") = Trim(NewString)
                    rs.Update
                    rs.MoveNext
                 Loop
               End If
             If rs.State = adStateOpen Then rs.Close
    
      rs.Open "Select * from StdPrtLibStructr Where ParentID ='" & sKeyb4Rename & "'", Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then    '����(����)Ҫ�����Ľڵ����ӽڵ���Ҫ��������(��һ��)�ӽڵ�, ���¼������ӽڵ㲻�ø��¸Ķ�
                
           rs.MoveFirst
           Do While Not rs.EOF
                   rs("ParentID") = Trim(NewString)
                   rs.Update
                   rs.MoveNext
            Loop
            
           If rs.State = adStateOpen Then rs.Close
           FillTree
           Exit Sub
           
         End If
        
      If rs.State = adStateOpen Then rs.Close
      Conn.Close
      FillTree
      Exit Sub
    
Else
handleAddcode:        '�����Add code��TreeView�ո����ӵ�InputNewItem�ڵ�
    sKey = NewString
        
       If AddNodeKeyNameNodeExist(sKey) Then  '�ж������NewString�ڴ�BOM(TreeView)���Ƿ�Ϊһ���Ѿ����ڵĽڵ���
         MsgBox "This New Name Item already Exist, Can NOT Rename.", vbInformation, "System Info"
         Exit Sub
       End If
         
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
       rs.Open "INSERT INTO StdPrtLibStructr (ParentID, ChildID) VALUES ('" & SourceNodeParent & "','" & sKey & "')", Conn, adOpenKeyset, adLockOptimistic
       If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    
    Conn.Close
    FillTree
    Exit Sub
    
End If
vbErrorHandler:

    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmStdPrtLibStructr:TreeView1_AfterLabelEdit", , App.ProductName

End Sub

'�������ĳ�ڵ�ʱ����Picture������ʾ�ļ�¼��Image
Private Sub TreeView1_NodeClick(ByVal myNode As Node)     'myNode�Ǵ�TreeView1_MouseDown���͹����ڵ���(������ǰ��C��12NC)
On Error GoTo ErrHandler
    Dim StmPic  As ADODB.Stream
    Dim StrPicTemp  As String

    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Set StmPic = New ADODB.Stream
    
        If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"  '�ж�һ��Ŀ¼�Ƿ���ڲ���������
      StrPicTemp = "C:\Temp\temp.tmp"            '��ʱ�ļ�,�������������ͼƬ
      Image1.Picture = LoadPicture()             '�����ԭ��������
    If myNode.Index = 1 Then                   'myNode.Index = 1 ��ʾ��ȡ���Ǹ��ڵ�
        Image1.Picture = LoadPicture()         '����Ǹ��ڵ�,��û��ͼƬ
        DataGridSearchString = ""
    Else
           If SourceNode.Children > 0 Then
             CmdLoadImage.Enabled = False     '�������Ĳ�����ײ�ڵ���LoadImage��ť������,Ҳ����uploadͼƬ,ֱ���˳�
             Exit Sub
           End If
       
        Dim tempNodekey As String
        tempNodekey = myNode.Key     'myNode.Key�Ǵ�TreeView1_MouseDown���͹����ڵ�key
        TxtStandrdPrtCateg = tempNodekey
        DataGridSearchString = "select * from SglPrt Where StandrdPrtCateg = '" & TxtStandrdPrtCateg & "'"
        
        rs.Open "Select * from StdPrtLibStructr where ChildID ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
          If IsNull(rs.Fields("StdPrtImage")) Then
            CmdLoadImage.Enabled = True
            Exit Sub
          End If
        If Not rs.EOF Then
            With StmPic
              .Type = adTypeBinary
              .Open
              .Write rs.Fields("StdPrtImage")                       'д�����ݿ��е�������Stream��
              .SaveToFile StrPicTemp, adSaveCreateOverWrite         '��Stream������д����ʱ�ļ�C:\Temp\temp.tmp��
              .Close
           End With
           Image1.Picture = LoadPicture(StrPicTemp)         '��Picture�ؼ���ʾͼ��
           CmdLoadImage.Enabled = False                     '��ʾ��ͼƬLoad Image��ť��λ�ɲ�����״̬
        End If
            If rs.State = adStateOpen Then rs.Close   'ע����������State,����status  adStateOpenֵΪ1
    End If
Conn.Close
lCurrentpage = 1
Call Refresh_StdPrtLibSearch(lCurrentpage)
Exit Sub

ErrHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:TreeView1_NodeClick"
End Sub
 
'��������갴��ʱ��ȡ��Դ�ڵ� (Check OK)
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����ָ���µĶ���ֵ��Դ�ڵ�
    Set SourceNode = TreeView1.HitTest(x, y)  'HitTest����������������ض�λ��x��y�����Node���������
        
    If Not (SourceNode Is Nothing) Then
    If SourceNode.Key = "ROOT" Then CmdLoadImage.Enabled = False  '���������Ǹ��ڵ���LoadImage��ť������
    If SourceNode.Children > 0 Then CmdLoadImage.Enabled = False  '�������Ĳ�����ײ�ڵ���LoadImage��ť������
    Set TreeView1.DropHighlight = SourceNode
    Else
    Set TreeView1.SelectedItem = Nothing     'ȡ��ѡ�еĹ��
    Set TreeView1.DropHighlight = Nothing    'ȡ��ѡ�еĹ��
    CmdLoadImage.Enabled = False  '�������Ĳ��ǽڵ���ǿհ״���LoadImage��ť������
    Image1.Picture = LoadPicture()
    End If
    
End Sub

'���ɿ���갴��ʱ (Check OK)
Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim bIsRoot As Boolean
    
    If TreeView1.SelectedItem Is Nothing Then
    Set TreeView1.DropHighlight = Nothing
    Exit Sub
    End If

' Show Popup Menu
    If Button = vbRightButton Then
'�ж�����Ǹ��ڵ�Ļ���ɾ��������, TreeView1.SelectedItem�ǵ��еĽڵ���(������ǰ��C��12NC)
        bIsRoot = (StrComp(TreeView1.SelectedItem.Key, "ROOT", vbTextCompare) = 0) 'vbTextCompare ֵΪ1ִ��һ������ԭ�ĵıȽϡ�string1 ���� string2����ֵΪ0
        mnuDeleteCode.Enabled = Not (bIsRoot)
        mnuRename.Enabled = Not (bIsRoot)
        PopupMenu mnuEdit
    End If
    
End Sub


Private Sub FillTree()

On Error GoTo vbErrorHandler

'  ��TreeView����Ҫ����
' Populate our TreeView Control with the Data from our database
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Dim lCount As Long
    Dim sParent As String
    Dim sKey As String
    Dim sText As String
    Dim nNode As Node
    
    rs.Open "Select * from StdPrtLibStructr order by ChildID ", Conn, adOpenStatic, adLockOptimistic
    Set TreeView1.ImageList = Nothing   '����TreeView1�Ǵ�����TreeView�ؼ�
    Set TreeView1.ImageList = ImageList1

    If rs.BOF And rs.EOF Then
        TreeView1.Nodes.Add , , "ROOT", "Standard Part Library", "VIEWBOOKMARKS"
        'BoldTreeNode TreeView1.Nodes("ROOT")
        Exit Sub
    End If
        
    'TreeRedraw TreeView1.hWnd, False
    
    rs.MoveFirst
    Set TreeView1.ImageList = Nothing
    Set TreeView1.ImageList = ImageList1
'
' Populate the TreeView Nodes
'

    With TreeView1.Nodes
        .Clear
        .Add , , "ROOT", "Standard Part Library", "VIEWBOOKMARKS"
'
' Make our Root Item BOLD
'
        'BoldTreeNode TreeView1.Nodes("ROOT")
'
' Now add all nodes into TreeView, but under the root item.
' We reparent the nodes in the next step
'
        Do Until rs.EOF
            sParent = rs("ParentID").Value
            sKey = rs("ChildID").Value
            sText = rs("ChildID").Value
            Set nNode = .Add("ROOT", tvwChild, sKey, sText, "FOLDER")
'
' Record parent ID
'
            nNode.Tag = sParent
            rs.MoveNext
        Loop
    
    End With
'
' Here's where we rebuild the structure of the nodes
'
    For Each nNode In TreeView1.Nodes
        sParent = nNode.Tag
        If Len(sParent) > 0 Then        ' Don't try and reparent the ROOT !
            If sParent = "Standard Part Library" Then
                sParent = "ROOT"
            End If
            Set nNode.Parent = TreeView1.Nodes(sParent)
        End If
    Next
'
' Now setup the images for each node in the treeview & set each node to
' be sorted if it has children
'
    For Each nNode In TreeView1.Nodes
        If nNode.Children = 0 Then
            nNode.Image = "CHILD"
        Else
            nNode.Sorted = True
        End If
    Next
    
    Set rs = Nothing
'
' Expand the Root Node
'
    TreeView1.Nodes("ROOT").Sorted = True
    TreeView1.Nodes("ROOT").Expanded = True
    
    'TreeRedraw TreeView1.hWnd, True
    
    Exit Sub

vbErrorHandler:
    
    'TreeRedraw TreeView1.hWnd, True
    
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::FillTree", , App.ProductName

End Sub

Private Sub TxtStandrdPrtCateg_Click()
MsgBox "The Content can NOT be inputed manually in case of Error" & vbCrLf & "Please click the left Tree View to Choose one Category", vbInformation, "System Info"
TxtStandrdPrtCateg.Enabled = False
End Sub

Private Sub TxtSglPrtIndex_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vbErrorHandler
 
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
If KeyCode = vbKeyReturn Then
               
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '�ж�TxtSglPrtIndex(����SinglePart NO)���ݵĺϷ���
        MsgBox "You must enter a Single Part 12NC here", vbInformation, "System Info."
        Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
                 MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
                 Exit Sub
            Else        '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(TxtSglPrtIndex.Text), 1, Len(Trim(TxtSglPrtIndex.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
                 MsgBox "The 12NC is not existing in Single Part Database", vbInformation, "System Info."
                 If rs.State = adStateOpen Then rs.Close
                 Exit Sub
               Else
                  TxtDescription.Text = rs("Description")
                 
                  If IsNull(rs("Drwlocate")) Then
                    TxtDrwlocate.Text = ""
                  Else
                    TxtDrwlocate.Text = rs("Drwlocate")                                                   '##########��Ӧ���ֶθ�ֵ
                  End If
                 
               End If
            If rs.State = adStateOpen Then rs.Close
        End If
        
TxtMainChrct1.SetFocus
End If

Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:TxtSglPrtIndex_KeyDown"
End Sub

Private Sub CmdRegister_Click()
On Error GoTo vbErrorHandler
 
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
                  
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '�ж�TxtSglPrtIndex(����SinglePart NO)���ݵĺϷ���
        MsgBox "You must enter a Single Part 12NC here", vbInformation, "System Info."
        Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
                 MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
                 Exit Sub
            Else        '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
            rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(TxtSglPrtIndex.Text), 1, Len(Trim(TxtSglPrtIndex.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
                 MsgBox "The 12NC is not existing in Single Part Database", vbInformation, "System Info."
                 If rs.State = adStateOpen Then rs.Close
                 Exit Sub
               Else
                 rs("StandrdPrtCateg") = TxtStandrdPrtCateg
                 rs("MainChrct1") = Trim(TxtMainChrct1.Text)
                 rs("MainChrct2") = Trim(TxtMainChrct2.Text)
                 rs("MainChrct3") = Trim(TxtMainChrct3.Text)
                 rs("MainChrct4") = Trim(TxtMainChrct4.Text)
                 rs("MainChrct5") = Trim(TxtMainChrct5.Text)
                 rs.Update
               End If
            If rs.State = adStateOpen Then rs.Close
        End If
        Conn.Close
        CmdFresh_Click             'ˢ��һ��
Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:CmdRegister"
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub CmdDrwPathAdd_Click()
GeneralPathAdd TxtDrwlocate.Text, "Drwlocate"
End Sub
Private Sub Lbl2_Click()
ClearPathAdd "Drwlocate"
End Sub
Private Sub CmdDrwView_Click()
GeneralDocView (TxtDrwlocate.Text)
End Sub

Private Sub GeneralPathAdd(ByVal InputPathName As String, ByVal InputField As String)
On Error GoTo vbErrorHandler
Dim DocPathName As String
    
DocPathName = Trim(InputPathName)
   If DocPathName = "" Then
      MsgBox "The Document Path/name is Null", vbInformation, "System Info."
      Exit Sub
         ElseIf Mid(DocPathName, 1, 3) <> "X:\" Then
         MsgBox "The Document Path/name must be formal released (In X:\)", vbInformation, "System Info."
         Exit Sub
            ElseIf Not OpnFileExist(DocPathName) Then
            MsgBox "The Document Path/Name is NOT existing, Please Check Path/Name", vbInformation, "System Info."
            Exit Sub
   End If

    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '�ж�TxtSglPrtIndex(����SinglePart NO )���ݵĺϷ���
          MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
          Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
              MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
              Exit Sub
            Else
               
               '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
               rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(TxtSglPrtIndex.Text), 1, Len(Trim(TxtSglPrtIndex.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
               If rs.RecordCount = 0 Then
                 MsgBox "The Selected Item 12NC is not existing in Database", vbInformation, "System Info."
                 If rs.State = adStateOpen Then rs.Close
                 Exit Sub
               Else
                 rs(InputField) = DocPathName
                 rs.Update
               End If
               If rs.State = adStateOpen Then rs.Close
        End If
     MsgBox "The Item Drawing(Document) Path/Name has been added successfully ", vbInformation, "System Info."
    Conn.Close
Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:CmdDrwPathAdd"
End Sub

Private Sub ClearPathAdd(ByVal InputField As String)
On Error GoTo vbErrorHandler
'Dim DocPathName As String
    
    If MsgBox("Confirm to Clear the Path/Name?" + vbCrLf + "ȷ���Ƿ����·��?", vbYesNo + vbDefaultButton2, "Confirm to Clear ȷ�����") = vbNo Then
      Exit Sub
    End If

    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '�ж�TxtSglPrtIndex(����SinglePart NO )���ݵĺϷ���
          MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
          Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
              MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "������12λ���ֵı��,����ĸ", vbInformation, "System Info."
              Exit Sub
            Else
               
               '��ʼ�ж������Single Part NO �Ƿ������ݿ�������
               rs.Open "Select * from SglPrt Where SglPrtIndex ='" & (Mid(Trim(TxtSglPrtIndex.Text), 1, Len(Trim(TxtSglPrtIndex.Text)) - 1) & "0") & "'", Conn, adOpenKeyset, adLockOptimistic
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
     MsgBox "The Item Drawing(Document) Path/Name has been Cleared successfully ", vbInformation, "System Info."
     Conn.Close
Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:ClearPathAdd"
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

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub CmdFirst_Click()     '��1ҳ����
   lCurrentpage = 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdFresh_Click()        'ˢ�²���
 Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdLast_Click()          '��ĩҳ����
   lCurrentpage = 10000
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdNext_Click()           '��1ҳ����
   lCurrentpage = lCurrentpage + 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdPrevious_Click()       '��1ҳ����
 If lCurrentpage > 1 Then
   lCurrentpage = lCurrentpage - 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
 End If
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub DataGrid1_Click()

With DataGrid1
.Col = 0
TxtSglPrtIndex = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 3
TxtDescription = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 10
TxtDrwlocate = .Text               '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 4
TxtStandrdPrtCateg = .Text         '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 5
TxtMainChrct1 = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 6
TxtMainChrct2 = .Text               '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 7
TxtMainChrct3 = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 8
TxtMainChrct4 = .Text               '##########��Ӧ�༭���ڿؼ���ֵ
.Col = 9
TxtMainChrct5 = .Text              '##########��Ӧ�༭���ڿؼ���ֵ
End With
End Sub
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Private Sub txtPage_nd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then PageGO_Click
End Sub

Private Sub PageGO_Click()          'ȥ��ָ��ҳ
   If Not IsNumeric(txtPage_nd) Then
       MsgBox "Page No. must be Number, No letter " + vbCrLf + "������ҳ������ֱ��", vbInformation, "Error Info!"
       txtPage_nd.SetFocus
   End If
   
   If val(txtPage_nd.Text) < 1 Then
   lCurrentpage = 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
   Exit Sub
   End If
   
   lCurrentpage = val(txtPage_nd.Text)  'val�������ַ���ת������ֵ
   Call Refresh_StdPrtLibSearch(lCurrentpage)

End Sub

Private Sub Refresh_StdPrtLibSearch(lPage As Long)
On Error GoTo vbErrorHandler
  Dim Conn As New ADODB.Connection   '����һ��ADO����
  '�������ݿ�
  Conn.ConnectionString = connString
  Conn.Open

          Dim rcds As New ADODB.Recordset    '����һ��ADO��¼�����ڴ��ÿ��ȫ��ȡ���ļ�¼
          Dim objrs As New ADODB.Recordset    '������һ����¼�����ڴ��ÿһҳ�ļ�¼
          Set rcds.ActiveConnection = Conn
          Set objrs.ActiveConnection = Conn

          Dim lPageCount     As Long
          Dim nPageSize     As Integer
          Dim lCount     As Long
          
If DataGridSearchString = "" Then Exit Sub
Set rcds = Nothing  'ԭ��¼�е�������Ҫ����ղ���д
rcds.Open DataGridSearchString, Conn, adOpenKeyset, adOpenStatic  '����һ��Static���͵��α�,�����¼��RecordCount��Ϊ-1 '##########��Ӧ������SglPrt

 
   'ÿҳ��ʾ�ļ�¼��Ϊ15
   nPageSize = 15
   rcds.PageSize = nPageSize         'ÿҳ��ʾ�ļ�¼����ֵ����¼������. PageSize��ҳ��ʾʱÿһҳ�ļ�¼��
' ADO PageCount ����
'The PageCount property returns a long value that indicates the number of pages with data in a Recordset object.
'PageCount���Ե������ǣ�����һ����ֵ������ָ����¼������������ҳ���������

'Tip: To divide the Recordset into a series of pages, use the PageSize property.
'��ʾ: �����ʹ��PageSize���Խ���¼���ָ�Ϊһϵ�е�ҳ��

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
  
  If rcds.PageCount = 0 Then Exit Sub               '��һ������Ҳû�е�ʱ��PageCount = 0��Ҫ�˳�������������²���Ҫ����
  lPageCount = rcds.PageCount                       'ÿҳ�ļ�¼�����趨����ҳ��������
              If lCurrentpage > lPageCount Then     '�������ʵ����ʾ��ҳ��,����ʵ����ʾ��ҳ��
                  lCurrentpage = lPageCount
              End If
          rcds.AbsolutePage = lCurrentpage          'ÿҳ�ļ�¼��������ҳ�����к�Ϳ�������ȥ��ĳһҳ
          
Set objrs = Nothing  'ԭ��¼�е�������Ҫ����ղ���д

          '����ֶ�����
          For lCount = 0 To rcds.Fields.count - 1
            If lCount = 0 Or lCount = 1 Or lCount = 10 Then                            ' ############## ���ڴ����ֵ��ֶ���Ҫ����������ֶ����
              objrs.Fields.Append rcds.Fields(lCount).Name, adUnsignedBigInt, rcds.Fields(lCount).DefinedSize  'adUnsignedBigInt   8�ֽڲ�����������
              GoTo NextLine
            End If
            objrs.Fields.Append rcds.Fields(lCount).Name, adVarChar, rcds.Fields(lCount).DefinedSize 'adVarChar�����ֶ����ַ���
NextLine:
          Next
          
          '�������ֶκ�򿪼�¼��
          objrs.Open
          
          '��ָ����¼��ѭ����ӵ�objrs��
          For lCount = 1 To nPageSize   'nPageSizeÿҳ��ʾ�ļ�¼��Ϊ15
                  If rcds.EOF = True Then
                  Exit For
                  End If
                 
                  objrs.AddNew
                  objrs!SglPrtIndex = rcds!SglPrtIndex                                             '##########��Ӧ���ֶθ�ֵ
                  objrs!SglPrtVer = rcds!SglPrtVer                                                 '##########��Ӧ���ֶθ�ֵ
                  objrs!PrtUnit = rcds!PrtUnit                                                     '##########��Ӧ���ֶθ�ֵ
                  objrs!Description = rcds!Description                                             '##########��Ӧ���ֶθ�ֵ
                  
                  If IsNull(rcds!StandrdPrtCateg) Then
                  objrs!StandrdPrtCateg = ""
                  Else
                  objrs!StandrdPrtCateg = rcds!StandrdPrtCateg                                                          '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!MainChrct1) Then
                  objrs!MainChrct1 = ""
                  Else
                  objrs!MainChrct1 = rcds!MainChrct1                                                 '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!MainChrct2) Then
                  objrs!MainChrct2 = ""
                  Else
                  objrs!MainChrct2 = rcds!MainChrct2                                                 '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!MainChrct3) Then
                  objrs!MainChrct3 = ""
                  Else
                  objrs!MainChrct3 = rcds!MainChrct3                                                 '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!MainChrct4) Then
                  objrs!MainChrct4 = ""
                  Else
                  objrs!MainChrct4 = rcds!MainChrct4                                                  '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!MainChrct5) Then
                  objrs!MainChrct5 = ""
                  Else
                  objrs!MainChrct5 = rcds!MainChrct5                                                  '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  If IsNull(rcds!Drwlocate) Then
                  objrs!Drwlocate = ""
                  Else
                  objrs!Drwlocate = rcds!Drwlocate                                                    '##########��Ӧ���ֶθ�ֵ
                  End If
                  
                  rcds.MoveNext
          Next
          
          '��
          Set DataGrid1.DataSource = objrs
            
          '��ʾҳ��
          txtPage.Text = lPage & "/" & rcds.PageCount              'lPage��Refresh_StdPrtLibSearch(lPage As Long)��
          
'  ��¼�������Ӷ����ܹ�,�رպ�DataGrid1��û��������ʾ
'If objrs.State = adStateOpen Then objrs.Close
'If rcds.State = adStateOpen Then rcds.Close
'Conn.Close

Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:Refresh_StdPrtLibSearch"
End Sub

