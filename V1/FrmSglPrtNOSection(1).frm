VERSION 5.00
Begin VB.Form FrmSglPrtNOSection 
   Caption         =   "Single Part Number Section Overview"
   ClientHeight    =   10896
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   12612
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSglPrtNOSection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10896
   ScaleWidth      =   12612
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdSysDistrb37 
      Caption         =   "4341073-01000/04999 Non-Metal Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   76
      Top             =   3090
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb0 
      Caption         =   "8241063-00000/99999 Temporary Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   74
      Top             =   0
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb36 
      Caption         =   "3141-13800000/13999999 Electrical Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   72
      Top             =   10530
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb35 
      Caption         =   "3141-13760000/13799999 Mechanical Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   70
      Top             =   10245
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb34 
      Caption         =   "3141-13740000/13759999 PCB Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   68
      Top             =   9960
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb33 
      Caption         =   "3141-13725000/13739999 Plastic Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   66
      Top             =   9675
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb32 
      Caption         =   "3141-13640000/13724999 Packing components"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   64
      Top             =   9390
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb31 
      Caption         =   "3141-13540000/13639999 Instruction for Use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   62
      Top             =   9105
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb30 
      Caption         =   "3141-13430000/13539999 Plastic Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   60
      Top             =   8820
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb29 
      Caption         =   "3141-13350000/13429999 PCB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   58
      Top             =   8535
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb28 
      Caption         =   "3141-13130000/13349999 Mechanical Part, no plastic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   56
      Top             =   8250
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb27 
      Caption         =   "3141-13060000/13129999 Electronic Part"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   55
      Top             =   7965
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb26 
      Caption         =   "3141130-30000/59999 Electronic Part Label/Sticker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   53
      Top             =   7680
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb25 
      Caption         =   "4341078-90000/99999 Other Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   41
      Top             =   7365
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb24 
      Caption         =   "4341078-60000/89999 Speaker Assy, SPK Box Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   40
      Top             =   7080
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb23 
      Caption         =   "4341078-45000/49999 Damper Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   39
      Top             =   6795
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb22 
      Caption         =   "4341078-25000/29999 VoiceCoil Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   38
      Top             =   6510
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb21 
      Caption         =   "4341078-15000/19999 Cone/Membrane/Edge Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   37
      Top             =   6225
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb20 
      Caption         =   "4341078-05000/09999 Frame Assy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   36
      Top             =   5940
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb1 
      Caption         =   "4341070-25000/29999 Connection Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   35
      Top             =   255
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb19 
      Caption         =   "4341076-95000/99999 Packing Sets(Pallet)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   34
      Top             =   5655
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb18 
      Caption         =   "4341076-30000/94999 Packing Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   33
      Top             =   5370
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb17 
      Caption         =   "4341076-10000/29999 Packing Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   32
      Top             =   5085
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb15 
      Caption         =   "4341-07450000/07549999 Plastic Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   31
      Top             =   4515
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb16 
      Caption         =   "4341-07550000/07599999 Label Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   30
      Top             =   4800
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb14 
      Caption         =   "4341073-45000/49999 Non-Metal Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   29
      Top             =   4230
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb13 
      Caption         =   "4341073-25000/29999 Non-Metal Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   28
      Top             =   3945
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb12 
      Caption         =   "4341073-15000/19999 Non-Metal Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   27
      Top             =   3660
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb11 
      Caption         =   "4341073-05000/09999 Non-Metal Component"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   26
      Top             =   3375
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb10 
      Caption         =   "4341071-85000/89999 Ferrite Magnets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   25
      Top             =   2820
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb2 
      Caption         =   "4341070-45000/49999 Fixation Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   7
      Top             =   540
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb3 
      Caption         =   "1341-50000000/59999999 Glues.            "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   6
      Top             =   810
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb4 
      Caption         =   "4341071-15000/19999 Non Ferite Magnet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   5
      Top             =   1110
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb5 
      Caption         =   "4341071-25000/29999 Metal Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   4
      Top             =   1395
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb6 
      Caption         =   "4341071-35000/39999 stamped component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   1680
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb7 
      Caption         =   "4341071-55000/59999 Milled(lathed) component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   2
      Top             =   1965
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb8 
      Caption         =   "4341071-65000/69999 Non Ferro metal frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Top             =   2250
      Width           =   4755
   End
   Begin VB.CommandButton CmdSysDistrb9 
      Caption         =   "4341071-75000/79999 No ferro part except frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   0
      Top             =   2535
      Width           =   4755
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 090: Cloth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7830
      TabIndex        =   79
      Top             =   3090
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 080: Wood"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5940
      TabIndex        =   78
      Top             =   3090
      Width           =   1830
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 200: Other Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   9810
      TabIndex        =   77
      Top             =   3090
      Width           =   2550
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item Type see following"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   75
      Top             =   -15
      Width           =   2085
   End
   Begin VB.Label Label36 
      Caption         =   "Item Type 040: Electronics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   73
      Top             =   10530
      Width           =   2310
   End
   Begin VB.Label Label35 
      Caption         =   "Item Type 040: Electronics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   71
      Top             =   10245
      Width           =   2310
   End
   Begin VB.Label Label34 
      Caption         =   "Item Type 040: Electronics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   69
      Top             =   9960
      Width           =   2310
   End
   Begin VB.Label Label33 
      Caption         =   "Item Type 010: Plastic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   67
      Top             =   9675
      Width           =   2070
   End
   Begin VB.Label Label32 
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   65
      Top             =   9390
      Width           =   2070
   End
   Begin VB.Label Label31 
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   63
      Top             =   9105
      Width           =   2070
   End
   Begin VB.Label Label30 
      Caption         =   "Item Type 010: Plastic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   61
      Top             =   8820
      Width           =   2070
   End
   Begin VB.Label Label29 
      Caption         =   "Item Type 040: Electronics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   59
      Top             =   8535
      Width           =   2310
   End
   Begin VB.Label Label28 
      Caption         =   "Item Type 110: Metal Speaker Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   57
      Top             =   8250
      Width           =   2985
   End
   Begin VB.Label Label14 
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   54
      Top             =   7680
      Width           =   2070
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 300: Sub-Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   52
      Top             =   7380
      Width           =   2535
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5925
      TabIndex        =   51
      Top             =   6810
      Width           =   2220
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5925
      TabIndex        =   50
      Top             =   6525
      Width           =   2235
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5925
      TabIndex        =   49
      Top             =   6240
      Width           =   2220
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 020 : Metal Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   48
      Top             =   5955
      Width           =   2565
   End
   Begin VB.Label Label24 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   47
      Top             =   5670
      Width           =   2070
   End
   Begin VB.Label Label23 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   46
      Top             =   5385
      Width           =   2070
   End
   Begin VB.Label Label22 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   45
      Top             =   5100
      Width           =   2070
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   44
      Top             =   4230
      Width           =   2145
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   43
      Top             =   3945
      Width           =   2145
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 030:  Magnet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   42
      Top             =   2820
      Width           =   2025
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 110:  Metal Speaker Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   24
      Top             =   2550
      Width           =   2640
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 020 : Metal Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   23
      Top             =   2265
      Width           =   2565
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 020 : Metal Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   22
      Top             =   1980
      Width           =   2565
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 020 : Metal Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   21
      Top             =   1695
      Width           =   2565
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Item Type 300: Sub-Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   20
      Top             =   7095
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 200: Other Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9810
      TabIndex        =   19
      Top             =   3375
      Width           =   2550
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Item Type 110: Metal Speaker Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   18
      Top             =   555
      Width           =   2955
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Item Type 100: Cable,Connector, terminal,wire..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   17
      Top             =   270
      Width           =   4140
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 090: Cloth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   7845
      TabIndex        =   16
      Top             =   3375
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 080: Wood"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5925
      TabIndex        =   15
      Top             =   3375
      Width           =   1830
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Item Type 070: Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   14
      Top             =   4815
      Width           =   2070
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Item Type 060: Chemical"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   13
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item Type 050: Acoustics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   12
      Top             =   3660
      Width           =   2145
   End
   Begin VB.Label Label3 
      Caption         =   "Item Type 040: Electronics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   11
      Top             =   7965
      Width           =   2310
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 030: Magnet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   10
      Top             =   1125
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Type 020 : Metal Speaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   9
      Top             =   1410
      Width           =   2565
   End
   Begin VB.Label LblItmTyp 
      BackColor       =   &H000080FF&
      Caption         =   "Item Type 010 : Plastic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5925
      TabIndex        =   8
      Top             =   4515
      Width           =   2070
   End
End
Attribute VB_Name = "FrmSglPrtNOSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'需要一个继承上一个窗口的判断是否Modify状态的逻辑变量
Public ModifyFm As Boolean


'搜索步距为1的函数
Private Function DataSection_DB1(ByVal StartString As String, EndString As String) As String    '###########变量改成对应的表字段名
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim Rscntstring As String

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next           '############以下相关改成对应的表字段名字
Rscntstring = "select min(CAST(SglPrtIndex AS DECIMAL(12,0))+10) from SglPrt WHERE (((CAST(SglPrtIndex AS DECIMAL(12,0))+10) Not In (select SglPrtIndex from SglPrt where sglprtindex between " + StartString + " and " + EndString + ")) and (CAST(SglPrtIndex AS DECIMAL(12,0))+10) between " + StartString + " and " + EndString + ")"
rcds.Open Rscntstring, Conn, adOpenKeyset, adOpenStatic     'PJNOIndex+1表示每加1位申请一个号
'Debug.Print Rscntstring
   '如果不能查到记录
If Not (rcds.EOF Or rcds.BOF) Then
    If ModifyFm = False Then
        DataSection_DB1 = Trim(rcds.Fields(0).Value)
        'MsgBox "Succeed to Add" + vbCrLf + "增加成功"   这句可以不用，用了还要关窗口，麻烦
    Else
        DataSection_DB1 = Trim(FrmSglPrtEdit.TxtSglPrtIndex.Text)
    End If
Else
   '系统提示信息，没有推荐号，请自行选择
   MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
   DataSection_DB1 = Trim(FrmSglPrtEdit.TxtSglPrtIndex.Text)
End If
rcds.Close
Conn.Close
  
End Function

'搜索步距为10的函数
Private Function DataSection_DB(ByVal StartString As String, EndString As String) As String    '###########变量改成对应的表字段名
Dim Conn As New ADODB.Connection   '定义一个ADO连接
Dim rcds As New ADODB.Recordset    '定义一个ADO记录集用于存放每次全部取出的记录
Dim Rscntstring As String

Conn.ConnectionString = connString
Conn.Open

On Error Resume Next           '############以下相关改成对应的表字段名字
Rscntstring = "select top 1 min(SglPrtIndex+10) From (select SglPrtIndex from SglPrt union select " + StartString + " as SglPrtIndex from SglPrt) a where (SglPrtIndex+10) between " + StartString + " and " + EndString + "  and sglprtindex+10 not in (select sglprtindex from sglprt)"

rcds.Open Rscntstring, Conn, adOpenKeyset, adOpenStatic     'PJNOIndex+1表示每加1位申请一个号

    If rcds.EOF Or rcds.BOF Then
        MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
        DataSection_DB = Trim(FrmSglPrtEdit.TxtSglPrtIndex.Text)
        Exit Function
    End If

      '如果不能查到记录
    If rcds.RecordCount = 0 Then
      '系统提示信息，没有推荐号，请自行选择
      MsgBox "System has no recommended Number, Please choose manually", vbInformation, " System information"
      DataSection_DB = Trim(FrmSglPrtEdit.TxtSglPrtIndex.Text)
      Exit Function
    End If
        If ModifyFm = False Then
            DataSection_DB = Trim(rcds.Fields(0).Value)
            'MsgBox "Succeed to Add" + vbCrLf + "增加成功"   这句可以不用，用了还要关窗口，麻烦
        Else
            DataSection_DB = Trim(FrmSglPrtEdit.TxtSglPrtIndex.Text)
        End If
Conn.Close
  
End Function

Private Sub CmdSysDistrb0_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB1("824106320000", "824106399999")   '使用搜索步距为1的函数
FrmSglPrtEdit.CombItemType.Text = "???"
sMsgBox "All temporary parts, ItemType can be refered to following"
End Sub

Private Sub CmdSysDistrb1_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107025000", "434107029999")
FrmSglPrtEdit.CombItemType.Text = "100"
sMsgBox "Connection part (Cable,terminals,Leadwire)"
End Sub

Private Sub CmdSysDistrb2_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107045000", "434107049999")
FrmSglPrtEdit.CombItemType.Text = "110"
sMsgBox "Fixation part(Screw,Nut,Bolt,Insert Nuts,Staple,etc)"
End Sub


Private Sub CmdSysDistrb3_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("134150000000", "134159999999")
FrmSglPrtEdit.CombItemType.Text = "060"
sMsgBox "Glues(All kinds of Glues including Ferroe Fluid)"
End Sub

Private Sub CmdSysDistrb37_Click()
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107301000", "434107304999")
FrmSglPrtEdit.CombItemType.Text = "???"
sMsgBox ("Wood, damping material, gaskets, foam & felt, rubber feet, textile,cable tie, solder wires, etc.")
End Sub

Private Sub CmdSysDistrb4_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107115000", "434107119999")
FrmSglPrtEdit.CombItemType.Text = "030"
sMsgBox "Non Ferite Magnet(Neodym,Alnico etc)"
End Sub

Private Sub CmdSysDistrb5_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107125000", "434107129999")
FrmSglPrtEdit.CombItemType.Text = "020"
sMsgBox "Metal Frame"
End Sub

Private Sub CmdSysDistrb6_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107135000", "434107139999")
FrmSglPrtEdit.CombItemType.Text = "020"
sMsgBox "All Stamped components except frames (T-yoke, U-yoke, Washer, shielding pot etc.), soldering tags"
End Sub

Private Sub CmdSysDistrb7_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107155000", "434107159999")
FrmSglPrtEdit.CombItemType.Text = "020"
sMsgBox "Milled(lathed) component(phase plug, milled U-Yoke, etc)"
End Sub

Private Sub CmdSysDistrb8_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107165000", "434107169999")
FrmSglPrtEdit.CombItemType.Text = "020"
sMsgBox "Non Ferro metal frame(aluminium etc.)"
End Sub

Private Sub CmdSysDistrb9_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107175000", "434107179999")
FrmSglPrtEdit.CombItemType.Text = "110"
sMsgBox "No ferro part except frame(Metal cabinet,Aluminium stands, Cupper Bush etc.)"
End Sub


Private Sub CmdSysDistrb10_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107185000", "434107189999")
FrmSglPrtEdit.CombItemType.Text = "030"
sMsgBox "Ferite Magnet"
End Sub

Private Sub CmdSysDistrb11_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107305000", "434107309999")
FrmSglPrtEdit.CombItemType.Text = "???"
sMsgBox "Wood, damping material, gaskets, foam & felt, rubber feet, textile,cable tie, solder wires, etc."
End Sub

Private Sub CmdSysDistrb12_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107315000", "434107319999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "Dust caps,Edges loose parts"
End Sub

Private Sub CmdSysDistrb13_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107325000", "434107329999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "Conebody,Membranes loose parts etc."
End Sub
Private Sub CmdSysDistrb14_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107345000", "434107349999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "VoiceCoil loose part(Voice coil loose parts: Coilformer,insulating sleeve,voice coil papers, ...)"
End Sub

Private Sub CmdSysDistrb15_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107450000", "434107549999")
FrmSglPrtEdit.CombItemType.Text = "010"
sMsgBox "AAll plastic parts for loudspeakers, cabinets,boxes,connection plates and plug connectors, grommets"
End Sub

Private Sub CmdSysDistrb16_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107550000", "434107599999")
FrmSglPrtEdit.CombItemType.Text = "070"
sMsgBox "Brand labels,POS label,Product Label,Wordmarks,All Labels"
End Sub

Private Sub CmdSysDistrb17_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107610000", "434107629999")
FrmSglPrtEdit.CombItemType.Text = "070"
sMsgBox "Packing related IFU ,Booklet,warranty card,safety card.etc"
End Sub

Private Sub CmdSysDistrb18_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107630000", "434107694999")
FrmSglPrtEdit.CombItemType.Text = "070"
sMsgBox "Packing components including transparent Tapes,PE Bags,Films,Nesting assy,Polyfoam,Paper tray,Carton,vacuum forming parts etc"
End Sub

Private Sub CmdSysDistrb19_Click()
sMsgBox "Standard transportation packing sets (Pallet)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107695000", "434107699999")
FrmSglPrtEdit.CombItemType.Text = "070"
End Sub


Private Sub CmdSysDistrb20_Click()
sMsgBox "Frame assy (Frame with washer, terminals)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107805000", "434107809999")
FrmSglPrtEdit.CombItemType.Text = "020"
End Sub

Private Sub CmdSysDistrb21_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107815000", "434107819999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "Cone assy, Membtane assy, Edge assy, Dome+VoiceCoil assy"
End Sub

Private Sub CmdSysDistrb22_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107825000", "434107829999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "Voic Coil assy"
End Sub

Private Sub CmdSysDistrb23_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107845000", "434107849999")
FrmSglPrtEdit.CombItemType.Text = "050"
sMsgBox "Damper assy"
End Sub

Private Sub CmdSysDistrb24_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107860000", "434107889999")
FrmSglPrtEdit.CombItemType.Text = "300"
sMsgBox "LoudSpeaker Driver assy, Box assy(internal use only)"
End Sub

Private Sub CmdSysDistrb25_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("434107890000", "434107899999")
FrmSglPrtEdit.CombItemType.Text = "300"
sMsgBox "Various other assy (clothframe assy, Label assy, etc."
End Sub

Private Sub CmdSysDistrb26_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113030000", "314113059999")
FrmSglPrtEdit.CombItemType.Text = "070"
sMsgBox "Sticker, Labels,WordMarks etc for electronic items."
End Sub

Private Sub CmdSysDistrb27_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113060000", "314113129999")
FrmSglPrtEdit.CombItemType.Text = "040"
sMsgBox "Electrical component (Capacitor,Resistor,IC,Software,Coil,Transformer,Connector,Transistor,Semiconductor,Switch)"
End Sub

Private Sub CmdSysDistrb28_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113130000", "314113349999")
FrmSglPrtEdit.CombItemType.Text = "110"
sMsgBox "Single mechanical component,Non plastic (Heatsink,Cooling fin,Antenna,Screw,Spacer,Washer, etc)"
End Sub

Private Sub CmdSysDistrb29_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113350000", "314113429999")
FrmSglPrtEdit.CombItemType.Text = "040"
sMsgBox "PCB (Transmitter PCB, Receiver PCB, Amplifer)"
End Sub

Private Sub CmdSysDistrb30_Click()

FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113430000", "314113539999")
FrmSglPrtEdit.CombItemType.Text = "010"
sMsgBox "Single Plastic components (Support, Pinion,Front,Back,Foot,Cover)"
End Sub

Private Sub CmdSysDistrb31_Click()
sMsgBox "Instruction for use"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113540000", "314113639999")
FrmSglPrtEdit.CombItemType.Text = "070"
End Sub

Private Sub CmdSysDistrb32_Click()
sMsgBox "Packaging components"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113640000", "314113724999")
FrmSglPrtEdit.CombItemType.Text = "070"
End Sub

Private Sub CmdSysDistrb33_Click()
sMsgBox "Plastic assembly, Cabinet assy(Top,Bottom)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113725000", "314113739999")
FrmSglPrtEdit.CombItemType.Text = "010"
End Sub

Private Sub CmdSysDistrb34_Click()
sMsgBox "PCB assembly (Transmitter,Receiver,Amplifer,Power unit,Family board,etc.)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113740000", "314113759999")
FrmSglPrtEdit.CombItemType.Text = "040"
End Sub

Private Sub CmdSysDistrb35_Click()
sMsgBox "Mechanical assembly(PCB assy + mains lead, Cable assy, Flat cable, Interface,etc.)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113760000", "314113799999")
FrmSglPrtEdit.CombItemType.Text = "040"
End Sub

Private Sub CmdSysDistrb36_Click()
sMsgBox "Electrical assembly (Transmitter assy, Receiver assy, Amplifer assy(with cabinets), Power unit supply, Electrical apparatus, etc.)"
FrmSglPrtEdit.TxtSglPrtIndex.Text = DataSection_DB("314113800000", "314113999999")
FrmSglPrtEdit.CombItemType.Text = "040"
End Sub


Private Sub sMsgBox(Msg As String)
    If MsgBox("The single part description: " & vbCrLf & Msg & vbCrLf & vbCrLf & "Would you like to create it?", vbYesNo, "System Alert") = vbYes Then
        Unload Me
    Else
    
    End If
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
'LoadSkin Me
'ResizeInit Me

End Sub
Private Sub Form_Resize()
        '确保窗体改变时控件随之改变
        Resize_ALL Me
End Sub
