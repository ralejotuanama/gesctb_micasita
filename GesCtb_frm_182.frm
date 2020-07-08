VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_AsiCtb_03_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16845
   Icon            =   "GesCtb_frm_182.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   16845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16845
      _Version        =   65536
      _ExtentX        =   29713
      _ExtentY        =   14473
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel4 
         Height          =   5055
         Left            =   30
         TabIndex        =   1
         Top             =   3090
         Width           =   16740
         _Version        =   65536
         _ExtentX        =   29527
         _ExtentY        =   8916
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_Tit_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Contable"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   1500
            TabIndex        =   3
            Top             =   60
            Width           =   3835
            _Version        =   65536
            _ExtentX        =   6765
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Cuenta Contable"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   5280
            TabIndex        =   4
            Top             =   60
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Glosa"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   11430
            TabIndex        =   5
            Top             =   60
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D/H"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   12030
            TabIndex        =   6
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe (MN)"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   13110
            TabIndex        =   7
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber (MN)"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   14190
            TabIndex        =   8
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe (ME)"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   15270
            TabIndex        =   9
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber (ME)"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_TotDeb_MN 
            Height          =   285
            Left            =   12060
            TabIndex        =   10
            Top             =   4410
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotHab_MN 
            Height          =   285
            Left            =   13140
            TabIndex        =   11
            Top             =   4410
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotDeb_ME 
            Height          =   285
            Left            =   14220
            TabIndex        =   12
            Top             =   4410
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotHab_ME 
            Height          =   285
            Left            =   15300
            TabIndex        =   13
            Top             =   4410
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_DifDeb_MN 
            Height          =   285
            Left            =   12060
            TabIndex        =   14
            Top             =   4710
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_DifHab_MN 
            Height          =   285
            Left            =   13140
            TabIndex        =   15
            Top             =   4710
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_DifDeb_ME 
            Height          =   285
            Left            =   14220
            TabIndex        =   16
            Top             =   4710
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_DifHab_ME 
            Height          =   285
            Left            =   15300
            TabIndex        =   17
            Top             =   4710
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_DetAsi 
            Height          =   3945
            Left            =   30
            TabIndex        =   18
            Top             =   360
            Width           =   16665
            _ExtentX        =   29395
            _ExtentY        =   6959
            _Version        =   393216
            Rows            =   6
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   10080
            TabIndex        =   41
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Contable"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin VB.Label Label8 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   10770
            TabIndex        =   20
            Top             =   4410
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Diferencia ==>"
            Height          =   285
            Left            =   10770
            TabIndex        =   19
            Top             =   4710
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   21
         Top             =   2280
         Width           =   16740
         _Version        =   65536
         _ExtentX        =   29527
         _ExtentY        =   1349
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_GloCab 
            Height          =   315
            Left            =   1410
            TabIndex        =   22
            Top             =   60
            Width           =   15135
            _Version        =   65536
            _ExtentX        =   26696
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   6600
            TabIndex        =   44
            Top             =   390
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MonCtb 
            Height          =   315
            Left            =   1410
            TabIndex        =   46
            Top             =   390
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   47
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Cambio:"
            Height          =   255
            Left            =   5190
            TabIndex        =   45
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Glosa Cabecera:"
            Height          =   285
            Left            =   60
            TabIndex        =   23
            Top             =   75
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   60
         Width           =   16740
         _Version        =   65536
         _ExtentX        =   29527
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   25
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Asientos Contables"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   14640
            Top             =   150
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_182.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   26
         Top             =   780
         Width           =   16740
         _Version        =   65536
         _ExtentX        =   29527
         _ExtentY        =   1138
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_182.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16140
            Picture         =   "GesCtb_frm_182.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   28
         Top             =   1470
         Width           =   16740
         _Version        =   65536
         _ExtentX        =   29527
         _ExtentY        =   1349
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_Empres 
            Height          =   315
            Left            =   1410
            TabIndex        =   29
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Sucurs 
            Height          =   315
            Left            =   1410
            TabIndex        =   30
            Top             =   390
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumAsi 
            Height          =   315
            Left            =   12030
            TabIndex        =   31
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   6600
            TabIndex        =   32
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   315
            Left            =   12030
            TabIndex        =   33
            Top             =   390
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_LibCon 
            Height          =   315
            Left            =   6600
            TabIndex        =   34
            Top             =   390
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecReg 
            Height          =   315
            Left            =   15180
            TabIndex        =   42
            Top             =   75
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Registro:"
            Height          =   285
            Left            =   13740
            TabIndex        =   43
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label33 
            Caption         =   "Período:"
            Height          =   255
            Left            =   5190
            TabIndex        =   40
            Top             =   90
            Width           =   1425
         End
         Begin VB.Label lbl_NumAsi 
            Caption         =   "Nro. Asiento:"
            Height          =   255
            Left            =   10470
            TabIndex        =   39
            Top             =   90
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   38
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   60
            TabIndex        =   37
            Top             =   420
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Libro Contable:"
            Height          =   255
            Index           =   1
            Left            =   5190
            TabIndex        =   36
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Comprob.:"
            Height          =   285
            Left            =   10470
            TabIndex        =   35
            Top             =   405
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_03_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ParEmp()      As moddat_tpo_Genera

Private Sub cmd_Imprim_Click()
   'Exportación al Crystal Report
   Screen.MousePointer = 11
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CNTBL_ASIENTO"
   crp_Imprim.DataFiles(1) = "CNTBL_ASIENTO_DET"

   'Se selecciona la formula
   crp_Imprim.SelectionFormula = ""

   'Se realiza la validación para codigo de instancia y fechas
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.ORIGEN} = {CNTBL_ASIENTO_DET.ORIGEN} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.ANO} = {CNTBL_ASIENTO_DET.ANO} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.MES} = {CNTBL_ASIENTO_DET.MES} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.NRO_LIBRO} = {CNTBL_ASIENTO_DET.NRO_LIBRO} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.NRO_ASIENTO} = {CNTBL_ASIENTO_DET.NRO_ASIENTO} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.ANO} = " & modctb_int_PerAno & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.MES} = " & modctb_int_PerMes & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.NRO_LIBRO} = " & modctb_int_CodLib & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.NRO_ASIENTO} = " & modctb_lng_NumAsi & "  "

   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTAMA_01.RPT"

   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   Screen.MousePointer = 0
End Sub

Private Sub grd_DetAsi_SelChange()
   If grd_DetAsi.Rows > 2 Then
      grd_DetAsi.RowSel = grd_DetAsi.Row
   End If
End Sub

Private Sub Form_Load()
Dim r_str_Origen  As String
Dim r_int_NumCla  As Integer

   Screen.MousePointer = 11
   r_str_Origen = "LM"
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_LimpiaCab
   pnl_NumAsi.Caption = CStr(modctb_lng_NumAsi)
      
   'Leyendo Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT ORIGEN, ANO, MES, NRO_LIBRO, NRO_ASIENTO, DESC_GLOSA, FECHA_CNTBL, FEC_REGISTRO,TASA_CAMBIO, COD_MONEDA  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_ASIENTO  "
   g_str_Parame = g_str_Parame & "   WHERE ORIGEN = '" & r_str_Origen & "'"
   g_str_Parame = g_str_Parame & "     AND ANO = " & CStr(modctb_int_PerAno) & "  "
   g_str_Parame = g_str_Parame & "     AND MES = " & CStr(modctb_int_PerMes) & "  "
   g_str_Parame = g_str_Parame & "     AND NRO_LIBRO = " & CStr(modctb_int_CodLib) & "  "
   g_str_Parame = g_str_Parame & "     AND NRO_ASIENTO = " & CStr(modctb_lng_NumAsi) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      pnl_FecCtb.Caption = g_rst_Princi!FECHA_CNTBL
      pnl_FecReg.Caption = Format(CDate(g_rst_Princi!FEC_REGISTRO), "DD/MM/YYYY")
      pnl_GloCab.Caption = Trim(g_rst_Princi!DESC_GLOSA & "")
      
      If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, modctb_str_CodEmp, "102", CStr(g_rst_Princi!COD_MONEDA)) Then
         pnl_MonCtb.Caption = l_arr_ParEmp(1).Genera_Nombre
      End If
         
      pnl_TipCam.Caption = Format(g_rst_Princi!TASA_CAMBIO, "###,###,##0.000000") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Leyendo Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT CNTA_CTBL, DET_GLOSA, FECHA_CNTBL, FLAG_DEBHAB, IMP_MOVSOL, IMP_MOVDOL"
   g_str_Parame = g_str_Parame & "     FROM CNTBL_ASIENTO_DET  "
   g_str_Parame = g_str_Parame & "    WHERE ORIGEN = '" & r_str_Origen & "'"
   g_str_Parame = g_str_Parame & "      AND ANO = " & CStr(modctb_int_PerAno) & "  "
   g_str_Parame = g_str_Parame & "      AND MES = " & CStr(modctb_int_PerMes) & "  "
   g_str_Parame = g_str_Parame & "      AND NRO_LIBRO = " & CStr(modctb_int_CodLib) & "  "
   g_str_Parame = g_str_Parame & "      AND NRO_ASIENTO = " & CStr(modctb_lng_NumAsi) & " "
   g_str_Parame = g_str_Parame & "    ORDER BY ITEM ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      grd_DetAsi.Redraw = False
      
      Do While Not g_rst_Princi.EOF
         grd_DetAsi.Rows = grd_DetAsi.Rows + 1
         grd_DetAsi.Row = grd_DetAsi.Rows - 1
         grd_DetAsi.Col = 0:  grd_DetAsi.Text = Trim(g_rst_Princi!CNTA_CTBL)
         grd_DetAsi.Col = 1:  grd_DetAsi.Text = moddat_gf_Consulta_CtaCtb(Trim(g_rst_Princi!CNTA_CTBL))
         grd_DetAsi.Col = 2:  grd_DetAsi.Text = Trim(g_rst_Princi!DET_GLOSA & "")
         grd_DetAsi.Col = 3:  grd_DetAsi.Text = Trim(g_rst_Princi!FECHA_CNTBL & "")
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = g_rst_Princi!FLAG_DEBHAB
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"
   
         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            grd_DetAsi.Col = 5:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVSOL), 0, g_rst_Princi!IMP_MOVSOL), "###,###,##0.00")                           'Debe MN
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVDOL), 0, g_rst_Princi!IMP_MOVDOL), "###,###,##0.00") 'Format(g_rst_Princi!IMP_MOVSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")  'Debe ME
         Else
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVSOL), 0, g_rst_Princi!IMP_MOVSOL), "###,###,##0.00")                             'Haber MN
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVDOL), 0, g_rst_Princi!IMP_MOVDOL), "###,###,##0.00") 'Format(g_rst_Princi!IMP_MOVSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")  'Haber ME
         End If
         g_rst_Princi.MoveNext
      Loop
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      grd_DetAsi.Redraw = True
      Call gs_UbiIniGrid(grd_DetAsi)
   End If
   Call fs_TotDebHab
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Function moddat_gf_Consulta_CtaCtb(ByVal p_CtaCtb As String) As String
   moddat_gf_Consulta_CtaCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DESC_CNTA, FLAG_PERMITE_MOV FROM CNTBL_CNTA  "
   g_str_Parame = g_str_Parame & " WHERE CNTA_CTBL = " & CStr(p_CtaCtb) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_CtaCtb = Trim(g_rst_Listas!DESC_CNTA)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_Inicia()
   grd_DetAsi.ColWidth(0) = 1447
   grd_DetAsi.ColWidth(1) = 3835
   grd_DetAsi.ColWidth(2) = 4785
   grd_DetAsi.ColWidth(3) = 1315
   grd_DetAsi.ColWidth(4) = 600
   grd_DetAsi.ColWidth(5) = 1075
   grd_DetAsi.ColWidth(6) = 1075
   grd_DetAsi.ColWidth(7) = 1075
   grd_DetAsi.ColWidth(8) = 1075
   grd_DetAsi.ColAlignment(0) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(1) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(2) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(3) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(4) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(5) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(6) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(7) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(8) = flexAlignRightCenter
End Sub

Private Sub fs_LimpiaCab()
   pnl_Empres.Caption = modctb_str_NomEmp
   pnl_Period.Caption = moddat_gf_Consulta_ParDes("033", CStr(modctb_int_PerMes)) & " " & Format(modctb_int_PerAno, "0000")
   pnl_Sucurs.Caption = modctb_str_NomSuc
   pnl_LibCon.Caption = modctb_str_NomLib

   Call gs_LimpiaGrid(grd_DetAsi)
   
   pnl_TotDeb_MN.Caption = "0.00 "
   pnl_TotHab_MN.Caption = "0.00 "
   pnl_DifDeb_MN.Caption = "0.00 "
   pnl_DifHab_MN.Caption = "0.00 "
   pnl_TotDeb_ME.Caption = "0.00 "
   pnl_TotHab_ME.Caption = "0.00 "
   pnl_DifDeb_ME.Caption = "0.00 "
   pnl_DifHab_ME.Caption = "0.00 "
End Sub

Private Sub fs_TotDebHab()
   Dim r_int_FilAct     As Integer
   Dim r_int_Contad     As Integer
   Dim r_dbl_TDebMN     As Double
   Dim r_dbl_TDebME     As Double
   Dim r_dbl_THabMN     As Double
   Dim r_dbl_THabME     As Double
   
   pnl_TotDeb_MN.Caption = Format(r_dbl_TDebMN, "###,###,###,##0.00") & " "
   pnl_TotHab_MN.Caption = Format(r_dbl_THabMN, "###,###,###,##0.00") & " "
   pnl_TotDeb_ME.Caption = Format(r_dbl_TDebME, "###,###,###,##0.00") & " "
   pnl_TotHab_ME.Caption = Format(r_dbl_THabME, "###,###,###,##0.00") & " "
   pnl_DifDeb_MN.Caption = "0.00 "
   pnl_DifDeb_ME.Caption = "0.00 "
   pnl_DifHab_MN.Caption = "0.00 "
   pnl_DifHab_ME.Caption = "0.00 "
   
   If grd_DetAsi.Rows = 0 Then
      Exit Sub
   End If
   
   grd_DetAsi.Redraw = False
   r_int_FilAct = grd_DetAsi.Row
   r_dbl_TDebMN = 0
   r_dbl_TDebME = 0
   r_dbl_THabMN = 0
   r_dbl_THabME = 0
   
   For r_int_Contad = 0 To grd_DetAsi.Rows - 1
      grd_DetAsi.Row = r_int_Contad
      
      grd_DetAsi.Col = 5:  r_dbl_TDebMN = r_dbl_TDebMN + CDbl(IIf(grd_DetAsi.Text = "", 0, grd_DetAsi.Text))
      grd_DetAsi.Col = 6:  r_dbl_THabMN = r_dbl_THabMN + CDbl(IIf(grd_DetAsi.Text = "", 0, grd_DetAsi.Text))
      grd_DetAsi.Col = 7:  r_dbl_TDebME = r_dbl_TDebME + CDbl(IIf(grd_DetAsi.Text = "", 0, grd_DetAsi.Text))
      grd_DetAsi.Col = 8:  r_dbl_THabME = r_dbl_THabME + CDbl(IIf(grd_DetAsi.Text = "", 0, grd_DetAsi.Text))
   Next r_int_Contad
   
   grd_DetAsi.Row = r_int_FilAct
   grd_DetAsi.Redraw = True
   
   Call gs_RefrescaGrid(grd_DetAsi)
   
   pnl_TotDeb_MN.Caption = Format(r_dbl_TDebMN, "###,###,###,##0.00") & " "
   pnl_TotHab_MN.Caption = Format(r_dbl_THabMN, "###,###,###,##0.00") & " "
   pnl_TotDeb_ME.Caption = Format(r_dbl_TDebME, "###,###,###,##0.00") & " "
   pnl_TotHab_ME.Caption = Format(r_dbl_THabME, "###,###,###,##0.00") & " "
   
   If r_dbl_TDebMN - r_dbl_THabMN > 0 Then
      pnl_DifDeb_MN.Caption = Format(r_dbl_TDebMN - r_dbl_THabMN, "###,###,###,##0.00") & " "
      pnl_DifHab_MN.Caption = "0.00 "
      pnl_DifDeb_ME.Caption = Format(r_dbl_TDebME - r_dbl_THabME, "###,###,###,##0.00") & " "
      pnl_DifHab_ME.Caption = "0.00 "
   Else
      pnl_DifDeb_MN.Caption = "0.00 "
      pnl_DifHab_MN.Caption = Format(Abs(r_dbl_TDebMN - r_dbl_THabMN), "###,###,###,##0.00") & " "
   
      pnl_DifDeb_ME.Caption = "0.00 "
      pnl_DifHab_ME.Caption = Format(Abs(r_dbl_TDebME - r_dbl_THabME), "###,###,###,##0.00") & " "
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub
