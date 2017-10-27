VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmStdPrtLibStructr 
   Caption         =   "PDM-Standard Part Lib Structure Admin 工程管理子系统"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   -1605
   ClientWidth     =   13710
   Icon            =   "FrmStdPrtStrct.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   13710
   StartUpPosition =   2  '屏幕中心
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
'使用Stream对象，可以实现对数据库的图像存取。
'数据库中存放图像的字段类型image（Access为OLE类型）。
'比如，如果用“CommonDialog”控件来选择你硬盘上的图像文件；
'用“Picture”控件来显示图像，那么下面的代码供参考：
'（运行VB，选择“工程\引用”命令，引用“Microsoft ActiveX Date 2.5 Library”。已连接数据库，打开了相应的记录集rs）
' CommonDialog 控件在 Visual Basic 和 Microsoft Windows 动态连接库Commdlg.dll 例程之间提供了接口。为了用该控件创建对话框，必须要求Commdlg.dll 在 Microsoft Windows \System 目录下。
'若未添加 CommonDialog 控件，则应从“工程”菜单中选定“部件”，将控件Microsoft Common Dialoge Control 6.0 添加到工具箱中。在标记对话的“控件”中找到并选定控件，然后单击“确定”按钮。
'CommonDialog 控件可以显示如下常用对话框：showopen“打开”,  showsave“另存为”,  showcolor“颜色”, showfont “字体”,showprinter “打印”,showhelp“windows帮助”
'可用下列格式设置 Filter 属性：
'description1 | filter1 | description2 | filter2...
'Description 是列表框中显示的字符串——例如，"Text Files (*.txt)"。Filter 是实际的文件过滤器─—例如，"*.txt"。每个description | filter 设置间必须用管道符号分隔 (|)。
'Private Sub mnuFileOpen_Click()
     'CancelError 为 True。
'     On Error GoTo ErrHandler
     '设置过滤器。
'     CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
     '指定缺省过滤器。
'     CommonDialog1.FilterIndex = 2
     '显示“打开”对话框。
'     CommonDialog1.ShowOpen
     '调用打开文件的过程。
'     OpenFile (CommonDialog1.FileName)
'     Exit Sub
'ErrHandler:
     '用户按“取消”按钮。
'     Exit Sub
'End Sub
'以上为示例代码
Option Explicit
Private SourceNode As Object  '定义节点拖曳的源节点
Private SourceNodeParent As String
Private DataGridSearchString As String
Private AddNodeOk As Boolean   '定义加节点成功否的标记
Private lCurrentpage As Long           '定义当前页变量

'判断节点A名字字符串在TreeView中是否有同节点名的节点存在&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function AddNodeKeyNameNodeExist(NodeAString As String) As Boolean
'整个TreeView中所有节点的遍历法
   Dim nodEachChild As Node
   For Each nodEachChild In TreeView1.Nodes
        If nodEachChild = NodeAString Then
            AddNodeKeyNameNodeExist = True
            Exit Function
       End If
   Next nodEachChild
 End Function

'判断一个字符串中是否含有数字  用IsNumeric判断0000d031为真(当成double型数字)
Private Function Isnum(Str As String) As Boolean
  Isnum = True
  Dim i  As Integer
  For i = 1 To Len(Str)
      Select Case Mid(Str, i, 1)
          Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          'Isnum = True  这里写Isnum = True就出错,因为如果中间是字母false了后面有数字的话又成为true了
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
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    Dim SglPrtIndex2Clr As String
    
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
                  
With DataGrid1
.Col = 0
SglPrtIndex2Clr = .Text              '##########对应编辑窗口控件赋值
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
        CmdFresh_Click             '刷新一下
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
 
lCurrentpage = 1           '窗口打开默认是第1页操作
 
CommonDialog1.Filter = "All Files (*.*)|*.*|Jpg Files (*.jpg)|*.jpg|Bmp Files (*.bmp)|*.bmp"
'指定缺省过滤器。
CommonDialog1.FilterIndex = 2

FillTree
End Sub

Private Sub CmdLoadImage_Click()
' 设置“CancelError”为 True
CommonDialog1.CancelError = True       '当CancelError属性设置为 True 时，无论何时选取“取消”按钮，均产生 32755 (cdlCancel) 号错误。
On Error GoTo ErrHandler

    Dim StmPic  As ADODB.Stream

    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Set StmPic = New ADODB.Stream
        CommonDialog1.ShowOpen
        Dim tempNodekey As String
        tempNodekey = SourceNode.Key     'myNode.Key是从TreeView1_MouseDown传送过来节点key
        rs.Open "Select * from StdPrtLibStructr where ChildID ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
          StmPic.Type = adTypeBinary             '指定流是二进制类型
          StmPic.Open                            '将数据获取到Stream对象中
          StmPic.LoadFromFile (CommonDialog1.Filename)     '将选择的图像加载到打开的StmPic中
          rs.Fields("StdPrtImage") = StmPic.Read           '从StmPic对象中读取数据
          rs.Update
          StmPic.Close
       End If
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
       Conn.Close
    CmdLoadImage.Enabled = False   '做完Load Image后按钮复位成不可用状态
Exit Sub

ErrHandler:
' 用户按了“取消”按钮
Exit Sub
End Sub

Private Sub CmdunLoadImage_Click()
On Error GoTo ErrHandler
    '@@@@@@@@@@判断是否是管理员用户，否则不能卸载图片
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Unload", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能卸载图片
    
    Dim Conn As New ADODB.Connection
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
        
        Dim tempNodekey As String
        tempNodekey = SourceNode.Key     'myNode.Key是从TreeView1_MouseDown传送过来节点key
        rs.Open "Select * from StdPrtLibStructr where ChildID ='" & tempNodekey & "'", Conn, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
          rs.Fields("StdPrtImage") = Null         '清空设置为Null
          rs.Update
        End If
        If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
        Conn.Close
    CmdLoadImage.Enabled = True   '做完unLoad Image后按钮设置成可用状态
Exit Sub

ErrHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:CmdunLoadImage"
End Sub

Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    FromForm.Show 0
End Sub

Private Sub Frame1_DblClick()
' 设置“CancelError”为 True
CommonDialog1.CancelError = True       '当CancelError属性设置为 True 时，无论何时选取“取消”按钮，均产生 32755 (cdlCancel) 号错误。
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
        
        CommonDialog1.ShowOpen   'Dir支持多字符 (*) 和单字符 (?) 的通配符来指定多重文件
        CurrentFilePath = left(CommonDialog1.Filename, InStrRev(CommonDialog1.Filename, "\") - 1)
        eachfilename = Dir(CurrentFilePath & "\" & "*.*")  '在第一次调用Dir(pathname)函数时,必须指定pathname,否则会产生错误。Dir 会返回匹配 pathname 的第一个文件名.  vbDirectory 值为16 指定无属性文件及其路径和文件夹。
        
        Do While Len(eachfilename) > 0
          
          eachfilename_12NC = (Mid(Replace(Trim(eachfilename), " ", ""), 1, 11) & "0")        '(Replace(Trim(eachfilename), " ", "") 去掉字符串中间的空格
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
                  rs2!SglPrtVer = 1                              '固定输入值
                  rs2!PrtUnit = "Piece"                          '固定输入值
                  rs2!Applicant = "NA"                              '固定输入值
                  rs2!Description = InputBox("Please Input Description", "Input Part Name ", "Frame TYoke Washer ...", 10000, 1)    'InputBox (Message, Title, Default,其中10000, 1是窗口出现位置)
                  rs2!IDSO = "Open"                              '固定输入值
                  rs2!NewOldStatus = "Old"                       '固定输入值
                  rs2!OpnDate = Date
                  rs2!ClosDate = Date
                  rs2!PJNOIndex = InputBox("Please Input Project Number" & vbCrLf & "Must be 6 Arabic numerals", "Input Project Number ", "120300 115360 ...", 10000, 1)    'InputBox (Message, Title, Default,其中10000, 1是窗口出现位置)
                  rs2!PjtName = "NA"                             '固定输入值
                  rs2!ProductLine = "5000"                       '固定输入值
                  rs2!ItemType = InputBox("Please Input Item Type" & vbCrLf & "Must be 3 Arabic numerals", "Input Item Type ", "050 070 100 ...", 10000, 1)    'InputBox (Message, Title, Default,其中10000, 1是窗口出现位置)
                  rs2!Location = "TR-AV"                         '固定输入值
                  rs2!CommtNote = "NA"                           '固定输入值
                 rs2.Update
                If rs2.State = adStateOpen Then rs2.Close   '注意这里是用State,不是status  adStateOpen值为1
                '注意这里需要先关闭rs 再重新按刚输入的new 12NC打开rs
                If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
                rs.Open "Select * from SglPrt Where SglPrtIndex ='" & eachfilename_12NC & "'", Conn, adOpenKeyset, adLockOptimistic
              End If
          End If
          rs("Drwlocate") = CurrentFilePath & "\" & eachfilename
          rs.Update
Nextfilepathname:
       eachfilename = Dir  '再一次调用Dir,不要使用参数。如果已没有合乎条件的文件,则 Dir 会返回一个零长度字符串 ("")。一旦返回值为零长度字符串并要再次调用Dir时,就必须指定pathname,否则会产生错误。
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
       Loop
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
       Conn.Close
Exit Sub

ErrHandler:
' 用户按了“取消”按钮
Exit Sub
End Sub

'删除一个节点的操作(Check OK)
Private Sub mnuDeleteCode_Click()
' Delete the selected CodeItem and all it's children

On Error GoTo vbErrorHandler
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    If SystemAdmin <> "Y" Then
        MsgBox "You are not administrator, No right to Delete", vbInformation, "System Info."
        Exit Sub
    End If
    '@@@@@@@@@@判断是否是管理员用户，否则不能删除
    
    Dim sKey As String
    Dim oNode As Node
    Dim oParentNode As Node
    Dim sMessage As String
        
    Set oNode = TreeView1.SelectedItem
    sKey = oNode
    If sKey = "InputNew12NC" Then        '如果是InputNew12NC,表示实际只是想删除一个treeview中临时节点,在Standard Part Library中根本无记录的
    FillTree
    Exit Sub
    End If
        
    If oNode.Key = "ROOT" Then Exit Sub     '如果是根节点则退出删除操作
    
       If oNode.Parent.Key = "ROOT" And oNode.Parent.Children = 1 Then       '如果节点父节点是ROOT并且ROOT只有一个Child(即选中的这个节点)
           MsgBox "Delete Final Record in Library, Library will not Exist", vbInformation, "System Info"
           Exit Sub
       End If
       
    sMessage = "Delete selected Code "
    If oNode.Children > 0 Then    '如果有节点选中并且有子节点则要有不同提示，Children是节点的子节点数量
        sMessage = sMessage & "and all child records ?"
    Else
        sMessage = sMessage & "?"
    End If
    
      If MsgBox(sMessage, vbYesNo + vbExclamation, "Delete Category Record") = vbNo Then
         Exit Sub
      End If
      
    Set oParentNode = oNode.Parent  '当前选中的要删除的节点的父节点赋值
    SourceNodeParent = oParentNode
    
    DeleteCodeItem SourceNodeParent, sKey
    FillTree
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:mnuDeleteCode"
End Sub

'删除一个节点的操作(Check OK)
Private Sub DeleteCodeItem(ParentNodeKey As String, ChildNodeKey As String)

On Error GoTo vbErrorHandler
 
    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
            
    '删除源节点的子节点数据
    rs.Open "Select * from StdPrtLibStructr Where ParentID ='" & ChildNodeKey & "'", Conn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
       DeleteCodeItem ChildNodeKey, rs("ChildID")      '递归调用找出所有层级的子节点数据
       rs.MoveNext
    Loop
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    '删除源节点本身的这一条数据
    rs.Open "Delete from StdPrtLibStructr Where ChildID='" & ChildNodeKey & "'" & " and  ParentID ='" & ParentNodeKey & "'", Conn, adOpenKeyset, adLockOptimistic
    If rs.State = adStateOpen Then rs.Close
    Conn.Close
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:DeleteCodeItem"
End Sub

'增加一个节点的操作(Check OK)
Private Sub mnuAddCode_Click()
AddCode
End Sub

'增加一个节点的操作(Check OK)
Private Sub AddCode()

On Error GoTo vbErrorHandler
    Dim AddNodeParentImage As String
    Dim nNode As Node
    Dim nParentNode As Node
    'Dim sParentKey As String
    
    Set nNode = TreeView1.SelectedItem   'TreeView上选中的节点赋值给nNode
    
    If nNode.Key <> "ROOT" Then                        '判断如果不是根节点的话
        Set nParentNode = TreeView1.Nodes(nNode.Key)        '要增加一个节点的操作中被增加的节点开始做一个父节点
        SourceNodeParent = nParentNode
        AddNodeParentImage = nParentNode.Image       '保存要增加一个节点的操作中被增加的节点的原图标
        nParentNode.Image = "FOLDER"
        nParentNode.ExpandedImage = "FOLDER"
   'ExpandedImage属性返回或设置在关联的ImageList控件中的ListImage对象的索引或键值，当Node对象被展开时显示 ListImage 对象。
    End If
    
    Set nNode = TreeView1.Nodes.Add(TreeView1.SelectedItem, tvwChild, "InputNewItem", "InputNewItem", "CHILD")
    Set TreeView1.SelectedItem = nNode       '变换选中的节点(蓝底显示)从被增加的节点 成为刚刚添加的节点
    nNode.EnsureVisible
    
    AddNodeOk = True                   '先假设是可以编辑OK的
    TreeView1.StartLabelEdit    'StartLabelEdit方法允许用户编辑标签。
' 当 LabelEdit 属性设置为 1（手动）时，必须用 StartLabelEdit 方法来启动一标签编辑操作。
' 在一对象上调用 StartLabelEdit 方法时，BeforeLabelEdit 事件也同时发生。
   
    If Not AddNodeOk Then
    TreeView1.Nodes.Remove nNode.Key
    nParentNode.Image = AddNodeParentImage
    End If
    
    Exit Sub
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmStdPrtLibStructr:AddCode"

End Sub

'更改一个节点(名)的操作(Check OK)
Private Sub mnuRename_Click()
' Change the Label - remember, we only allow 12 Characters
    TreeView1.StartLabelEdit
End Sub

'更改一个节点标签的前预备操作(Check OK)
Private Sub tvCodeItems_BeforeLabelEdit(Cancel As Integer)
'如果没有操作也不要删除,可以引发AfterLabelEdit操作
End Sub

'更改一个节点标签的后续操作(Check OK)
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
        
    Set oNode = TreeView1.SelectedItem      'TreeView1.SelectedItem有两种情况，一个是要改名的节点，另一个是新增加的节点
    sKeyb4Rename = oNode
    
        If oNode.Key = "ROOT" Then Exit Sub     '如果是根节点则退出改名操作
        
    If oNode Is Nothing Then      '如果是没有节点选中则提示并退出改名操作
        MsgBox "No Selected Record", vbInformation, "System Info"
        Exit Sub
    End If
    
    Set oParentNode = oNode.Parent  '当前选中的要改名的节点的父节点赋值
    SourceNodeParent = oParentNode
    
    
       If sKeyb4Rename = "InputNewItem" Then GoTo handleAddcode   '直接去到新增节点的操作
    
rs.Open "Select * from StdPrtLibStructr Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'", Conn, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then   '该语句判断改名的节点是Add code刚刚增加的(只在TreeView中可见,在Standard Part Lib没有记录的)还是说现有Standard Part Lib(有记录的)
       sMessage = "Rename selected Code ?"
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
       If MsgBox(sMessage, vbYesNo + vbExclamation, "Rename Code Record") = vbNo Then
          Exit Sub
       End If
       
      sKey = NewString
      
       If AddNodeKeyNameNodeExist(sKey) Then  '判断输入的NewString在此BOM(TreeView)中是否为一个已经存在的节点名
         MsgBox "This New Name Item already Exist, Can NOT Rename.", vbInformation, "System Info"
         Exit Sub
       End If
      
       rs.Open "Select * from StdPrtLibStructr Where ChildID='" & sKeyb4Rename & "'" & " and  ParentID ='" & SourceNodeParent & "'", Conn, adOpenKeyset, adLockOptimistic
       rs("ChildID") = Trim(NewString)    '先做(更新)要改名的节点作为子节点的这一条在Standard Part Lib中的记录
       rs.Update
    
         If rs.State = adStateOpen Then rs.Close
         Set rs = Nothing
         
             '更新此Category名下所有的SinglePart中StandrdPrtCateg字段全部改为新值
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
        If rs.RecordCount > 0 Then    '再做(更新)要改名的节点有子节点则要更新所有(下一级)子节点, 下下级以上子节点不用更新改动
                
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
handleAddcode:        '如果是Add code在TreeView刚刚增加的InputNewItem节点
    sKey = NewString
        
       If AddNodeKeyNameNodeExist(sKey) Then  '判断输入的NewString在此BOM(TreeView)中是否为一个已经存在的节点名
         MsgBox "This New Name Item already Exist, Can NOT Rename.", vbInformation, "System Info"
         Exit Sub
       End If
         
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
       rs.Open "INSERT INTO StdPrtLibStructr (ParentID, ChildID) VALUES ('" & SourceNodeParent & "','" & sKey & "')", Conn, adOpenKeyset, adLockOptimistic
       If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
    
    Conn.Close
    FillTree
    Exit Sub
    
End If
vbErrorHandler:

    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "FrmStdPrtLibStructr:TreeView1_AfterLabelEdit", , App.ProductName

End Sub

'当鼠标点击某节点时，在Picture里面显示改记录的Image
Private Sub TreeView1_NodeClick(ByVal myNode As Node)     'myNode是从TreeView1_MouseDown传送过来节点名(不包含前面C的12NC)
On Error GoTo ErrHandler
    Dim StmPic  As ADODB.Stream
    Dim StrPicTemp  As String

    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
    Set StmPic = New ADODB.Stream
    
        If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"  '判断一个目录是否存在不存在则建立
      StrPicTemp = "C:\Temp\temp.tmp"            '临时文件,用来保存读出的图片
      Image1.Picture = LoadPicture()             '先清除原来的内容
    If myNode.Index = 1 Then                   'myNode.Index = 1 表示点取的是根节点
        Image1.Picture = LoadPicture()         '如果是根节点,则没有图片
        DataGridSearchString = ""
    Else
           If SourceNode.Children > 0 Then
             CmdLoadImage.Enabled = False     '如果点击的不是最底层节点则LoadImage按钮不可用,也不用upload图片,直接退出
             Exit Sub
           End If
       
        Dim tempNodekey As String
        tempNodekey = myNode.Key     'myNode.Key是从TreeView1_MouseDown传送过来节点key
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
              .Write rs.Fields("StdPrtImage")                       '写入数据库中的数据至Stream中
              .SaveToFile StrPicTemp, adSaveCreateOverWrite         '将Stream中数据写入临时文件C:\Temp\temp.tmp中
              .Close
           End With
           Image1.Picture = LoadPicture(StrPicTemp)         '用Picture控件显示图像
           CmdLoadImage.Enabled = False                     '显示完图片Load Image后按钮复位成不可用状态
        End If
            If rs.State = adStateOpen Then rs.Close   '注意这里是用State,不是status  adStateOpen值为1
    End If
Conn.Close
lCurrentpage = 1
Call Refresh_StdPrtLibSearch(lCurrentpage)
Exit Sub

ErrHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:TreeView1_NodeClick"
End Sub
 
'当按下鼠标按键时，取得源节点 (Check OK)
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '将鼠标指针下的对象赋值给源节点
    Set SourceNode = TreeView1.HitTest(x, y)  'HitTest方法，这个方法返回对位于x和y坐标的Node对象的引用
        
    If Not (SourceNode Is Nothing) Then
    If SourceNode.Key = "ROOT" Then CmdLoadImage.Enabled = False  '如果点击的是根节点则LoadImage按钮不可用
    If SourceNode.Children > 0 Then CmdLoadImage.Enabled = False  '如果点击的不是最底层节点则LoadImage按钮不可用
    Set TreeView1.DropHighlight = SourceNode
    Else
    Set TreeView1.SelectedItem = Nothing     '取消选中的光标
    Set TreeView1.DropHighlight = Nothing    '取消选中的光标
    CmdLoadImage.Enabled = False  '如果点击的不是节点而是空白处则LoadImage按钮不可用
    Image1.Picture = LoadPicture()
    End If
    
End Sub

'当松开鼠标按键时 (Check OK)
Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim bIsRoot As Boolean
    
    If TreeView1.SelectedItem Is Nothing Then
    Set TreeView1.DropHighlight = Nothing
    Exit Sub
    End If

' Show Popup Menu
    If Button = vbRightButton Then
'判断如果是根节点的话则删除不可用, TreeView1.SelectedItem是点中的节点名(不包含前面C的12NC)
        bIsRoot = (StrComp(TreeView1.SelectedItem.Key, "ROOT", vbTextCompare) = 0) 'vbTextCompare 值为1执行一个按照原文的比较。string1 等于 string2返回值为0
        mnuDeleteCode.Enabled = Not (bIsRoot)
        mnuRename.Enabled = Not (bIsRoot)
        PopupMenu mnuEdit
    End If
    
End Sub


Private Sub FillTree()

On Error GoTo vbErrorHandler

'  对TreeView的主要操作
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
    Set TreeView1.ImageList = Nothing   '其中TreeView1是窗口上TreeView控件
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
               
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '判断TxtSglPrtIndex(输入SinglePart NO)数据的合法性
        MsgBox "You must enter a Single Part 12NC here", vbInformation, "System Info."
        Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
                 MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
                 Exit Sub
            Else        '开始判断输入的Single Part NO 是否在数据库表里存在
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
                    TxtDrwlocate.Text = rs("Drwlocate")                                                   '##########对应表字段赋值
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
                  
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '判断TxtSglPrtIndex(输入SinglePart NO)数据的合法性
        MsgBox "You must enter a Single Part 12NC here", vbInformation, "System Info."
        Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
                 MsgBox "Single Part is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
                 Exit Sub
            Else        '开始判断输入的Single Part NO 是否在数据库表里存在
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
        CmdFresh_Click             '刷新一下
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
    
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '判断TxtSglPrtIndex(输入SinglePart NO )数据的合法性
          MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
          Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
              MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
              Exit Sub
            Else
               
               '开始判断输入的Single Part NO 是否在数据库表里存在
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
    
    If MsgBox("Confirm to Clear the Path/Name?" + vbCrLf + "确认是否清除路径?", vbYesNo + vbDefaultButton2, "Confirm to Clear 确认清除") = vbNo Then
      Exit Sub
    End If

    Dim Conn As New ADODB.Connection
    
    
    Conn.Open connString
    
    Dim rs As New ADODB.Recordset
    Set rs.ActiveConnection = Conn
    
        If Len(Trim(TxtSglPrtIndex.Text)) = 0 Then        '判断TxtSglPrtIndex(输入SinglePart NO )数据的合法性
          MsgBox "You must enter a 12NC for the Selected Item", vbInformation, "System Info."
          Exit Sub
        ElseIf Not (Len(Trim(TxtSglPrtIndex.Text)) = 12 And Isnum(Trim(TxtSglPrtIndex.Text))) Then
              MsgBox "Selected Item is 12 Number, no Letter" + vbCrLf + "必须是12位数字的编号,无字母", vbInformation, "System Info."
              Exit Sub
            Else
               
               '开始判断输入的Single Part NO 是否在数据库表里存在
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
Private Sub CmdFirst_Click()     '第1页操作
   lCurrentpage = 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdFresh_Click()        '刷新操作
 Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdLast_Click()          '第末页操作
   lCurrentpage = 10000
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdNext_Click()           '下1页操作
   lCurrentpage = lCurrentpage + 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
End Sub

Private Sub CmdPrevious_Click()       '上1页操作
 If lCurrentpage > 1 Then
   lCurrentpage = lCurrentpage - 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
 End If
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub DataGrid1_Click()

With DataGrid1
.Col = 0
TxtSglPrtIndex = .Text              '##########对应编辑窗口控件赋值
.Col = 3
TxtDescription = .Text              '##########对应编辑窗口控件赋值
.Col = 10
TxtDrwlocate = .Text               '##########对应编辑窗口控件赋值
.Col = 4
TxtStandrdPrtCateg = .Text         '##########对应编辑窗口控件赋值
.Col = 5
TxtMainChrct1 = .Text              '##########对应编辑窗口控件赋值
.Col = 6
TxtMainChrct2 = .Text               '##########对应编辑窗口控件赋值
.Col = 7
TxtMainChrct3 = .Text              '##########对应编辑窗口控件赋值
.Col = 8
TxtMainChrct4 = .Text               '##########对应编辑窗口控件赋值
.Col = 9
TxtMainChrct5 = .Text              '##########对应编辑窗口控件赋值
End With
End Sub
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Private Sub txtPage_nd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then PageGO_Click
End Sub

Private Sub PageGO_Click()          '去到指定页
   If Not IsNumeric(txtPage_nd) Then
       MsgBox "Page No. must be Number, No letter " + vbCrLf + "请输入页码的数字编号", vbInformation, "Error Info!"
       txtPage_nd.SetFocus
   End If
   
   If val(txtPage_nd.Text) < 1 Then
   lCurrentpage = 1
   Call Refresh_StdPrtLibSearch(lCurrentpage)
   Exit Sub
   End If
   
   lCurrentpage = val(txtPage_nd.Text)  'val函数是字符串转换成数值
   Call Refresh_StdPrtLibSearch(lCurrentpage)

End Sub

Private Sub Refresh_StdPrtLibSearch(lPage As Long)
On Error GoTo vbErrorHandler
  Dim Conn As New ADODB.Connection   '定义一个ADO连接
  '连接数据库
  Conn.ConnectionString = connString
  Conn.Open

          Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
          Dim objrs As New ADODB.Recordset    '定义另一个记录集用于存放每一页的记录
          Set rcds.ActiveConnection = Conn
          Set objrs.ActiveConnection = Conn

          Dim lPageCount     As Long
          Dim nPageSize     As Integer
          Dim lCount     As Long
          
If DataGridSearchString = "" Then Exit Sub
Set rcds = Nothing  '原记录中的内容需要先清空才能写
rcds.Open DataGridSearchString, Conn, adOpenKeyset, adOpenStatic  '启动一个Static类型的游标,否则记录数RecordCount总为-1 '##########对应表名字SglPrt

 
   '每页显示的记录数为15
   nPageSize = 15
   rcds.PageSize = nPageSize         '每页显示的记录数赋值给记录集属性. PageSize分页显示时每一页的记录数
' ADO PageCount 属性
'The PageCount property returns a long value that indicates the number of pages with data in a Recordset object.
'PageCount属性的作用是：返回一个长值，用于指定记录集对象中数据页面的数量。

'Tip: To divide the Recordset into a series of pages, use the PageSize property.
'提示: 你可以使用PageSize属性将记录集分割为一系列的页面

'Note: If the last page contains fewer records than specified in PageSize, it still counts as an additional page in the PageCount property.
'注意：如果最后一页的记录数量少于在PageSize属性中指定的数量，那么它仍然被视为一页。

'Note: If this method is not supported it returns -1.
'注意：如果不支持这个方法，那么将返回-1。

'IntFix 函数返回参数的整数部分?
'语法
'Int(number)
'Fix(number)
'必要的 number 参数是 Double 或任何有效的数值表达式。如果 number 包含 Null，则返回 Null。
'说明
'Int 和 Fix 都会删除 number 的小数部份而返回剩下的整数。
'Int 和 Fix 的不同之处在于，如果 number 为负数，则 Int 返回小于或等于 number 的第一个负整数，而 Fix 则会返回大于或等于 number 的第一个负整数。例如，Int 将 -8.4 转换成 -9，而 Fix 将 -8.4 转换成 -8。
  
  If rcds.PageCount = 0 Then Exit Sub               '当一条数据也没有的时候PageCount = 0需要退出，否则进行以下操作要出错
  lPageCount = rcds.PageCount                       '每页的记录数量设定后总页数就有了
              If lCurrentpage > lPageCount Then     '如果超出实际显示的页数,则按照实际显示的页数
                  lCurrentpage = lPageCount
              End If
          rcds.AbsolutePage = lCurrentpage          '每页的记录数量和总页数就有后就可以设置去到某一页
          
Set objrs = Nothing  '原记录中的内容需要先清空才能写

          '添加字段名称
          For lCount = 0 To rcds.Fields.count - 1
            If lCount = 0 Or lCount = 1 Or lCount = 10 Then                            ' ############## 对于纯数字的字段需要在这里调整字段序号
              objrs.Fields.Append rcds.Fields(lCount).Name, adUnsignedBigInt, rcds.Fields(lCount).DefinedSize  'adUnsignedBigInt   8字节不带符号整型
              GoTo NextLine
            End If
            objrs.Fields.Append rcds.Fields(lCount).Name, adVarChar, rcds.Fields(lCount).DefinedSize 'adVarChar其余字段用字符串
NextLine:
          Next
          
          '定义完字段后打开记录集
          objrs.Open
          
          '将指定记录数循环添加到objrs中
          For lCount = 1 To nPageSize   'nPageSize每页显示的记录数为15
                  If rcds.EOF = True Then
                  Exit For
                  End If
                 
                  objrs.AddNew
                  objrs!SglPrtIndex = rcds!SglPrtIndex                                             '##########对应表字段赋值
                  objrs!SglPrtVer = rcds!SglPrtVer                                                 '##########对应表字段赋值
                  objrs!PrtUnit = rcds!PrtUnit                                                     '##########对应表字段赋值
                  objrs!Description = rcds!Description                                             '##########对应表字段赋值
                  
                  If IsNull(rcds!StandrdPrtCateg) Then
                  objrs!StandrdPrtCateg = ""
                  Else
                  objrs!StandrdPrtCateg = rcds!StandrdPrtCateg                                                          '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!MainChrct1) Then
                  objrs!MainChrct1 = ""
                  Else
                  objrs!MainChrct1 = rcds!MainChrct1                                                 '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!MainChrct2) Then
                  objrs!MainChrct2 = ""
                  Else
                  objrs!MainChrct2 = rcds!MainChrct2                                                 '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!MainChrct3) Then
                  objrs!MainChrct3 = ""
                  Else
                  objrs!MainChrct3 = rcds!MainChrct3                                                 '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!MainChrct4) Then
                  objrs!MainChrct4 = ""
                  Else
                  objrs!MainChrct4 = rcds!MainChrct4                                                  '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!MainChrct5) Then
                  objrs!MainChrct5 = ""
                  Else
                  objrs!MainChrct5 = rcds!MainChrct5                                                  '##########对应表字段赋值
                  End If
                  
                  If IsNull(rcds!Drwlocate) Then
                  objrs!Drwlocate = ""
                  Else
                  objrs!Drwlocate = rcds!Drwlocate                                                    '##########对应表字段赋值
                  End If
                  
                  rcds.MoveNext
          Next
          
          '绑定
          Set DataGrid1.DataSource = objrs
            
          '显示页数
          txtPage.Text = lPage & "/" & rcds.PageCount              'lPage在Refresh_StdPrtLibSearch(lPage As Long)中
          
'  记录集和连接都不能关,关闭后DataGrid1将没有数据显示
'If objrs.State = adStateOpen Then objrs.Close
'If rcds.State = adStateOpen Then rcds.Close
'Conn.Close

Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " FrmStdPrtLibStructr:Refresh_StdPrtLibSearch"
End Sub

