VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_Asient_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10830
   ClientLeft      =   1785
   ClientTop       =   1500
   ClientWidth     =   17940
   Icon            =   "GesCtb_frm_153.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10815
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   17925
      _Version        =   65536
      _ExtentX        =   31618
      _ExtentY        =   19076
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
      Begin Threed.SSPanel SSPanel24 
         Height          =   1395
         Left            =   30
         TabIndex        =   75
         Top             =   6210
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   2461
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
         Begin Threed.SSPanel SSPanel26 
            Height          =   285
            Left            =   60
            TabIndex        =   90
            Top             =   60
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Comprobante de Pago"
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
         Begin Threed.SSPanel SSPanel28 
            Height          =   285
            Left            =   3240
            TabIndex        =   91
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Persona"
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
         Begin Threed.SSPanel SSPanel30 
            Height          =   285
            Left            =   6150
            TabIndex        =   92
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Referencia"
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
         Begin Threed.SSPanel SSPanel32 
            Height          =   285
            Left            =   15360
            TabIndex        =   93
            Top             =   60
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
         Begin Threed.SSPanel SSPanel33 
            Height          =   285
            Left            =   16170
            TabIndex        =   94
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto"
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
         Begin MSFlexGridLib.MSFlexGrid grd_DocAsi 
            Height          =   1005
            Left            =   30
            TabIndex        =   5
            Top             =   360
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   1773
            _Version        =   393216
            Rows            =   6
            Cols            =   26
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   9060
            TabIndex        =   99
            Top             =   60
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Movimiento Bancario"
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
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1935
         Left            =   30
         TabIndex        =   56
         Top             =   3090
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   3413
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
            TabIndex        =   57
            Top             =   60
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "C�digo Cuenta Contable"
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
            Left            =   2760
            TabIndex        =   58
            Top             =   60
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripci�n Cuenta Contable"
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
            Left            =   6660
            TabIndex        =   59
            Top             =   60
            Width           =   5055
            _Version        =   65536
            _ExtentX        =   8916
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
            Left            =   11700
            TabIndex        =   60
            Top             =   60
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
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
            Left            =   12630
            TabIndex        =   61
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   13830
            TabIndex        =   62
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   15030
            TabIndex        =   63
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   16230
            TabIndex        =   64
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   12630
            TabIndex        =   65
            Top             =   1290
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   13830
            TabIndex        =   66
            Top             =   1290
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   15030
            TabIndex        =   67
            Top             =   1290
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   16230
            TabIndex        =   68
            Top             =   1290
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   12630
            TabIndex        =   70
            Top             =   1590
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   13830
            TabIndex        =   71
            Top             =   1590
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   15030
            TabIndex        =   72
            Top             =   1590
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Left            =   16230
            TabIndex        =   73
            Top             =   1590
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Height          =   915
            Left            =   30
            TabIndex        =   23
            Top             =   360
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   1614
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
         Begin VB.Label Label10 
            Caption         =   "Diferencia ==>"
            Height          =   285
            Left            =   11340
            TabIndex        =   74
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   11340
            TabIndex        =   69
            Top             =   1290
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   52
         Top             =   2280
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
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
         Begin VB.TextBox txt_GloCab 
            Height          =   315
            Left            =   1530
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "txt_GloCab"
            Top             =   60
            Width           =   16275
         End
         Begin VB.ComboBox cmb_MonCtb 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   3465
         End
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   8040
            TabIndex        =   4
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.000000"
            DecimalPlaces   =   6
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label7 
            Caption         =   "Glosa Cabecera:"
            Height          =   285
            Left            =   60
            TabIndex        =   55
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Cambio:"
            Height          =   255
            Left            =   6420
            TabIndex        =   54
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   43
         Top             =   60
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
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
            TabIndex        =   44
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Comprobantes Contables"
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   555
            Left            =   15300
            TabIndex        =   105
            Top             =   60
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4419
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Per�odo Actual:  NOVIEMBRE 2009"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_153.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   45
         Top             =   780
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
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
         Begin VB.CommandButton cmd_DetNue 
            Height          =   585
            Left            =   1590
            Picture         =   "GesCtb_frm_153.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Agregar registro de Detalle Contable"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_DetMod 
            Height          =   585
            Left            =   2220
            Picture         =   "GesCtb_frm_153.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Editar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetBor 
            Height          =   585
            Left            =   2820
            Picture         =   "GesCtb_frm_153.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Borrar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetAce 
            Height          =   585
            Left            =   3420
            Picture         =   "GesCtb_frm_153.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Grabar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetCan 
            Height          =   585
            Left            =   4020
            Picture         =   "GesCtb_frm_153.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cancelar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DocAce 
            Height          =   585
            Left            =   6780
            Picture         =   "GesCtb_frm_153.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Grabar registro de Detalle Operativo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DocBor 
            Height          =   585
            Left            =   6180
            Picture         =   "GesCtb_frm_153.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Borrar registro de Detalle Operativo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DocMod 
            Height          =   585
            Left            =   5580
            Picture         =   "GesCtb_frm_153.frx":185C
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Editar registro de Detalle Operativo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DocNue 
            Height          =   585
            Left            =   4980
            Picture         =   "GesCtb_frm_153.frx":1B66
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agregar registro de Detalle Operativo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DocCan 
            Height          =   585
            Left            =   7380
            Picture         =   "GesCtb_frm_153.frx":1E70
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Cancelar registro de Detalle Operativo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ComNue 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_153.frx":217A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Nuevo Comprobante Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ComGra 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_153.frx":2484
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Grabar Comprobante Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   17220
            Picture         =   "GesCtb_frm_153.frx":28C6
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   47
         Top             =   1470
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
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
         Begin VB.ComboBox cmb_LibCon 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   60
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   60
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecCtb 
            Height          =   315
            Left            =   8040
            TabIndex        =   1
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Comprob.:"
            Height          =   285
            Left            =   6420
            TabIndex        =   51
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Libro Contable:"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   50
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   6420
            TabIndex        =   49
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel25 
         Height          =   3105
         Left            =   30
         TabIndex        =   76
         Top             =   7650
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   5477
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
         Begin VB.TextBox txt_BcoNum 
            Height          =   315
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   20
            Text            =   "txt_BcoNum"
            ToolTipText     =   "Nro. de Cheque, Nro. de Transferencia, Nro. Dep�sito"
            Top             =   2400
            Width           =   1365
         End
         Begin VB.ComboBox cmb_BcoCod 
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2070
            Width           =   3465
         End
         Begin VB.ComboBox cmb_BcoCta 
            Height          =   315
            Left            =   14370
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2070
            Width           =   3405
         End
         Begin VB.ComboBox cmb_BcoTip 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2070
            Width           =   3465
         End
         Begin VB.ComboBox cmb_OrgMon 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2730
            Width           =   3465
         End
         Begin VB.ComboBox cmb_RefTip 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1740
            Width           =   3465
         End
         Begin VB.TextBox txt_IdeNDo 
            Height          =   315
            Left            =   14370
            MaxLength       =   250
            TabIndex        =   13
            Text            =   "txt_IdeNDo"
            Top             =   1410
            Width           =   3405
         End
         Begin VB.ComboBox cmb_IdeTDo 
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1410
            Width           =   3465
         End
         Begin VB.ComboBox cmb_IdeTip 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1410
            Width           =   3465
         End
         Begin VB.TextBox txt_DocNum 
            Height          =   315
            Left            =   14370
            MaxLength       =   12
            TabIndex        =   8
            Text            =   "txt_DocNum"
            Top             =   60
            Width           =   3405
         End
         Begin VB.ComboBox cmb_MovSuc 
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   750
            Width           =   3465
         End
         Begin VB.TextBox txt_DocSer 
            Height          =   315
            Left            =   8040
            MaxLength       =   5
            TabIndex        =   7
            Text            =   "txt_DocSer"
            Top             =   60
            Width           =   1335
         End
         Begin VB.ComboBox cmb_DocTip 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   3465
         End
         Begin EditLib.fpDoubleSingle ipp_OrgMto 
            Height          =   315
            Left            =   8040
            TabIndex        =   22
            Top             =   2730
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin MSMask.MaskEdBox msk_MovNum 
            Height          =   315
            Left            =   14370
            TabIndex        =   10
            Top             =   750
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "##-#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox msk_RefOpe 
            Height          =   315
            Left            =   8040
            TabIndex        =   15
            Top             =   1740
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox msk_RefSol 
            Height          =   315
            Left            =   14370
            TabIndex        =   16
            Top             =   1740
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin EditLib.fpDateTime ipp_BcoFec 
            Height          =   315
            Left            =   14370
            TabIndex        =   106
            Top             =   2400
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_DocFec 
            Height          =   315
            Left            =   8040
            TabIndex        =   108
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_MovFec 
            Height          =   315
            Left            =   8040
            TabIndex        =   110
            Top             =   1080
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label28 
            Caption         =   "F. Emisi�n Compr.:"
            Height          =   285
            Left            =   6420
            TabIndex        =   111
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha Emisi�n Doc.:"
            Height          =   285
            Left            =   6420
            TabIndex        =   109
            Top             =   390
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Movim."
            Height          =   285
            Left            =   12810
            TabIndex        =   107
            Top             =   2400
            Width           =   1245
         End
         Begin VB.Label Label32 
            Caption         =   "Nro. Referencia:"
            Height          =   285
            Left            =   6420
            TabIndex        =   98
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Banco:"
            Height          =   255
            Left            =   6420
            TabIndex        =   97
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   12810
            TabIndex        =   96
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Movim. Bancario:"
            Height          =   255
            Left            =   60
            TabIndex        =   95
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Tipo Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   89
            Top             =   2730
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Nro. Solicitud:"
            Height          =   255
            Left            =   12810
            TabIndex        =   88
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Nro. Operaci�n:"
            Height          =   255
            Left            =   6420
            TabIndex        =   87
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Tipo Referencia:"
            Height          =   255
            Left            =   60
            TabIndex        =   86
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   12810
            TabIndex        =   85
            Top             =   1410
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   255
            Left            =   6420
            TabIndex        =   84
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Tipo Persona:"
            Height          =   255
            Left            =   60
            TabIndex        =   83
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Nro. Movimiento:"
            Height          =   285
            Left            =   12810
            TabIndex        =   82
            Top             =   750
            Width           =   1725
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Documento:"
            Height          =   285
            Left            =   12810
            TabIndex        =   81
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   6420
            TabIndex        =   80
            Top             =   2730
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   6420
            TabIndex        =   79
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Nro. Serie Doc.:"
            Height          =   285
            Left            =   6420
            TabIndex        =   78
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Tipo Compr. Pago:"
            Height          =   255
            Left            =   60
            TabIndex        =   77
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel23 
         Height          =   1095
         Left            =   30
         TabIndex        =   100
         Top             =   5070
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_TipMov 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   3465
         End
         Begin VB.TextBox txt_GloDet 
            Height          =   315
            Left            =   1530
            MaxLength       =   250
            TabIndex        =   25
            Text            =   "txt_GloDet"
            Top             =   390
            Width           =   16275
         End
         Begin VB.ComboBox cmb_CtaCtb 
            Height          =   315
            Left            =   1530
            TabIndex        =   24
            Text            =   "cmb_CtaCtb"
            Top             =   60
            Width           =   16275
         End
         Begin EditLib.fpDoubleSingle ipp_MtoCta 
            Height          =   315
            Left            =   8040
            TabIndex        =   27
            Top             =   720
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Contable:"
            Height          =   255
            Left            =   60
            TabIndex        =   104
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   6420
            TabIndex        =   103
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Movimiento:"
            Height          =   255
            Left            =   60
            TabIndex        =   102
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Glosa Detalle:"
            Height          =   285
            Left            =   60
            TabIndex        =   101
            Top             =   390
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_Asient_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()   As moddat_tpo_Genera
Dim l_arr_Sucurs()   As moddat_tpo_Genera
Dim l_arr_MovSuc()   As moddat_tpo_Genera
Dim l_arr_MonCtb()   As moddat_tpo_Genera
Dim l_arr_BcoCod()   As moddat_tpo_Genera
Dim l_arr_ParEmp()   As moddat_tpo_Genera
Dim l_arr_CtaCtb()   As moddat_tpo_Genera
Dim l_arr_BcoCta()   As moddat_tpo_Genera
Dim l_int_TopNiv     As Integer
Dim l_str_CtaCtb     As String
Dim l_int_FlgCmb     As Integer
Dim l_str_FecIni     As String
Dim l_str_FecFin     As String
Dim l_int_GrbDet     As Integer
Dim l_int_GrbDoc     As Integer
Dim l_var_ColAnt     As Variant
Dim l_int_PerMes     As Integer
Dim l_int_PerAno     As Integer

Private Sub cmb_BcoCod_Click()
   If cmb_BcoCod.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_BcoCod(cmb_BcoCod.ListIndex + 1).Genera_Codigo, cmb_BcoCta, l_arr_BcoCta)
      Screen.MousePointer = 0
      
      Call gs_SetFocus(cmb_BcoCta)
   End If
End Sub

Private Sub cmb_BcoCod_GotFocus()
   l_var_ColAnt = cmb_BcoCod.BackColor
   cmb_BcoCod.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_BcoCod_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BcoCod_Click
   End If
End Sub

Private Sub cmb_BcoCod_LostFocus()
   cmb_BcoCod.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_BcoCta_Click()
   Call gs_SetFocus(txt_BcoNum)
End Sub

Private Sub cmb_BcoCta_GotFocus()
   l_var_ColAnt = cmb_BcoCta.BackColor
   cmb_BcoCta.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_BcoCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BcoCta_Click
   End If
End Sub

Private Sub cmb_BcoCta_LostFocus()
   cmb_BcoCta.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_BcoTip_Click()
   If cmb_BcoTip.ListIndex > -1 Then
      If cmb_BcoTip.ItemData(cmb_BcoTip.ListIndex) = 1 Then
         cmb_BcoCod.ListIndex = -1
         cmb_BcoCta.Clear
         txt_BcoNum.Text = ""
         ipp_BcoFec.Text = Format(Date, "dd/mm/yyyy")
         
         cmb_BcoCod.Enabled = False
         cmb_BcoCta.Enabled = False
         txt_BcoNum.Enabled = False
         ipp_BcoFec.Enabled = False
         
         Call gs_SetFocus(cmb_OrgMon)
      Else
         cmb_BcoCod.Enabled = True
         cmb_BcoCta.Enabled = True
         txt_BcoNum.Enabled = True
         ipp_BcoFec.Enabled = True
         
         Call gs_SetFocus(cmb_BcoCod)
      End If
   End If
End Sub

Private Sub cmb_BcoTip_GotFocus()
   l_var_ColAnt = cmb_BcoTip.BackColor
   cmb_BcoTip.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_BcoTip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BcoTip_Click
   End If
End Sub

Private Sub cmb_BcoTip_LostFocus()
   cmb_BcoTip.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_CtaCtb_Change()
   l_str_CtaCtb = cmb_CtaCtb.Text
   
   cmb_CtaCtb.SelLength = Len(l_str_CtaCtb)
End Sub

Private Sub cmb_CtaCtb_Click()
   If cmb_CtaCtb.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_GloDet)
      End If
   End If
End Sub

Private Sub cmb_CtaCtb_GotFocus()
   l_var_ColAnt = cmb_CtaCtb.BackColor
   cmb_CtaCtb.BackColor = modgen_g_con_ColAma

   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
End Sub

Private Sub cmb_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CtaCtb, l_str_CtaCtb)
      l_int_FlgCmb = True
      
      If cmb_CtaCtb.ListIndex > -1 Then
         l_str_CtaCtb = ""
      End If
      
      Call gs_SetFocus(txt_GloDet)
   End If
End Sub

Private Sub cmb_CtaCtb_LostFocus()
   cmb_CtaCtb.BackColor = l_var_ColAnt
   
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DocTip_Click()
   If cmb_DocTip.ListIndex > -1 Then
      If cmb_DocTip.ItemData(cmb_DocTip.ListIndex) = 1 Then
         txt_DocSer.Text = ""
         txt_DocNum.Text = ""
         ipp_DocFec.Text = Format(Date, "dd/mm/yyyy")
         
         txt_DocSer.Enabled = False
         txt_DocNum.Enabled = False
         ipp_DocFec.Enabled = False
         
         cmb_MovSuc.ListIndex = -1
         msk_MovNum.Text = ""
         ipp_MovFec.Text = Format(Date, "dd/mm/yyyy")
         
         cmb_MovSuc.Enabled = False
         msk_MovNum.Enabled = False
         ipp_MovFec.Enabled = False
         
         Call gs_SetFocus(cmb_IdeTip)
      ElseIf cmb_DocTip.ItemData(cmb_DocTip.ListIndex) = 5 Then
         txt_DocSer.Text = ""
         txt_DocNum.Text = ""
         ipp_DocFec.Text = Format(Date, "dd/mm/yyyy")
         
         txt_DocSer.Enabled = False
         txt_DocNum.Enabled = False
         ipp_DocFec.Enabled = False
         
         cmb_MovSuc.Enabled = True
         msk_MovNum.Enabled = True
         ipp_MovFec.Enabled = True
         
         Call gs_SetFocus(cmb_MovSuc)
      Else
         txt_DocSer.Enabled = True
         txt_DocNum.Enabled = True
         ipp_DocFec.Enabled = True
         
         cmb_MovSuc.ListIndex = -1
         msk_MovNum.Text = ""
         ipp_MovFec.Text = Format(Date, "dd/mm/yyyy")
         
         cmb_MovSuc.Enabled = False
         msk_MovNum.Enabled = False
         ipp_MovFec.Enabled = False
         
         Call gs_SetFocus(txt_DocSer)
      End If
   End If
End Sub

Private Sub cmb_DocTip_GotFocus()
   l_var_ColAnt = cmb_DocTip.BackColor
   cmb_DocTip.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_DocTip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DocTip_Click
   End If
End Sub

Private Sub cmb_DocTip_LostFocus()
   cmb_DocTip.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
      
      Call moddat_gs_Carga_ParEmp(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, "102", cmb_MonCtb, l_arr_MonCtb)
      
      pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, l_str_FecIni, l_str_FecFin, l_int_PerMes, l_int_PerAno)
      
      If Date > CDate(l_str_FecFin) Then
         ipp_FecCtb.Text = Format(l_str_FecFin, "dd/mm/yyyy")
      ElseIf Date < CDate(l_str_FecIni) Then
         ipp_FecCtb.Text = Format(l_str_FecIni, "dd/mm/yyyy")
      End If
      
      ipp_FecCtb.DateMin = Format(CDate(l_str_FecIni), "yyyymmdd")
      ipp_FecCtb.DateMax = Format(CDate(l_str_FecFin), "yyyymmdd")
      
      cmb_Sucurs.ListIndex = 0
      cmb_MovSuc.ListIndex = 0
      
      Call moddat_gs_Carga_CtaCtb(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
      
      cmb_MonCtb.ListIndex = gf_Busca_Arregl(l_arr_MonCtb, Format("2", "000000")) - 1
      
      l_int_TopNiv = -1
      If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, "100", "001") Then
         l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
      End If
      
      Call gs_SetFocus(cmb_Sucurs)
      
      Screen.MousePointer = 0
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Empres_GotFocus()
   l_var_ColAnt = cmb_Empres.BackColor
   cmb_Empres.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Empres_LostFocus()
   cmb_Empres.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_IdeTDo_Click()
   If cmb_IdeTDo.ListIndex > -1 Then
      Select Case cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex)
         Case 1:     txt_IdeNDo.MaxLength = 8
         Case 7:     txt_IdeNDo.MaxLength = 11
         Case Else:  txt_IdeNDo.MaxLength = 12
      End Select
      
      Call gs_SetFocus(txt_IdeNDo)
   Else
      txt_IdeNDo.MaxLength = 0
   End If
End Sub

Private Sub cmb_IdeTDo_GotFocus()
   l_var_ColAnt = cmb_IdeTDo.BackColor
   cmb_IdeTDo.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_IdeTDo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_IdeTDo_Click
   End If
End Sub

Private Sub cmb_IdeTDo_LostFocus()
   cmb_IdeTDo.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_IdeTip_Click()
   If cmb_IdeTip.ListIndex > -1 Then
      If cmb_IdeTip.ItemData(cmb_IdeTip.ListIndex) = 1 Then
         cmb_IdeTDo.ListIndex = -1
         txt_IdeNDo.Text = ""
         
         cmb_IdeTDo.Enabled = False
         txt_IdeNDo.Enabled = False
         
         Call gs_SetFocus(cmb_RefTip)
      Else
         cmb_IdeTDo.Enabled = True
         txt_IdeNDo.Enabled = True
         
         Call gs_SetFocus(cmb_IdeTDo)
      End If
   End If
End Sub

Private Sub cmb_IdeTip_GotFocus()
   l_var_ColAnt = cmb_IdeTip.BackColor
   cmb_IdeTip.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_IdeTip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_IdeTip_Click
   End If
End Sub

Private Sub cmb_IdeTip_LostFocus()
   cmb_IdeTip.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_LibCon_Click()
   Call ipp_FecCtb_LostFocus
   Call gs_SetFocus(ipp_FecCtb)
End Sub

Private Sub cmb_LibCon_GotFocus()
   l_var_ColAnt = cmb_LibCon.BackColor
   cmb_LibCon.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_LibCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_LibCon_Click
   End If
End Sub

Private Sub cmb_LibCon_LostFocus()
   cmb_LibCon.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_MonCtb_Click()
   Call gs_SetFocus(ipp_TipCam)
End Sub

Private Sub cmb_MonCtb_GotFocus()
   l_var_ColAnt = cmb_MonCtb.BackColor
   cmb_MonCtb.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_MonCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonCtb_Click
   End If
End Sub

Private Sub cmb_MonCtb_LostFocus()
   cmb_MonCtb.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_MovSuc_Click()
   Call gs_SetFocus(msk_MovNum)
End Sub

Private Sub cmb_MovSuc_GotFocus()
   l_var_ColAnt = cmb_MovSuc.BackColor
   cmb_MovSuc.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_MovSuc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MovSuc_Click
   End If
End Sub

Private Sub cmb_MovSuc_LostFocus()
   cmb_MovSuc.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_OrgMon_Click()
   Call gs_SetFocus(ipp_OrgMto)
End Sub

Private Sub cmb_OrgMon_GotFocus()
   l_var_ColAnt = cmb_OrgMon.BackColor
   cmb_OrgMon.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_OrgMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_OrgMon_Click
   End If
End Sub

Private Sub cmb_OrgMon_LostFocus()
   cmb_OrgMon.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_RefTip_Click()
   If cmb_RefTip.ListIndex > -1 Then
      If cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 1 Then
         msk_RefOpe.Text = ""
         msk_RefSol.Text = ""
         
         msk_RefOpe.Enabled = False
         msk_RefSol.Enabled = False
         
         Call gs_SetFocus(cmb_BcoTip)
      ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 2 Then
         msk_RefOpe.Enabled = True
         
         msk_RefSol.Text = ""
         msk_RefSol.Enabled = False
         
         Call gs_SetFocus(msk_RefOpe)
      ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 3 Then
         msk_RefOpe.Text = ""
         msk_RefOpe.Enabled = False
         
         msk_RefSol.Enabled = True
         
         Call gs_SetFocus(msk_RefSol)
      End If
   End If
End Sub

Private Sub cmb_RefTip_GotFocus()
   l_var_ColAnt = cmb_RefTip.BackColor
   cmb_RefTip.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_RefTip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RefTip_Click
   End If
End Sub

Private Sub cmb_RefTip_LostFocus()
   cmb_RefTip.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(cmb_LibCon)
End Sub

Private Sub cmb_Sucurs_GotFocus()
   l_var_ColAnt = cmb_Sucurs.BackColor
   cmb_Sucurs.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmb_Sucurs_LostFocus()
   cmb_Sucurs.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_TipMov_Click()
   Call gs_SetFocus(ipp_MtoCta)
End Sub

Private Sub cmb_TipMov_GotFocus()
   l_var_ColAnt = cmb_TipMov.BackColor
   cmb_TipMov.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_TipMov_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMov_Click
   End If
End Sub

Private Sub cmb_TipMov_LostFocus()
   cmb_TipMov.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_BcoFec_GotFocus()
   l_var_ColAnt = ipp_BcoFec.BackColor
   ipp_BcoFec.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_BcoFec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_OrgMon)
   End If
End Sub

Private Sub ipp_BcoFec_LostFocus()
   ipp_BcoFec.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_DocFec_GotFocus()
   l_var_ColAnt = ipp_DocFec.BackColor
   ipp_DocFec.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_DocFec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_IdeTip)
   End If
End Sub

Private Sub ipp_DocFec_LostFocus()
   ipp_DocFec.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_FecCtb_GotFocus()
   l_var_ColAnt = ipp_FecCtb.BackColor
   ipp_FecCtb.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_FecCtb_InvalidData(NextWnd As Long)
   If CDate(ipp_FecCtb.Text) < CDate(l_str_FecIni) Then
      ipp_FecCtb.Text = l_str_FecIni
   ElseIf CDate(ipp_FecCtb.Text) > CDate(l_str_FecIni) Then
      ipp_FecCtb.Text = l_str_FecFin
   End If
End Sub

Private Sub ipp_FecCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_GloCab)
   End If
End Sub

Private Sub ipp_FecCtb_LostFocus()
   ipp_FecCtb.BackColor = l_var_ColAnt
   
   If cmb_LibCon.ListIndex > -1 Then
      Select Case cmb_LibCon.ItemData(cmb_LibCon.ListIndex)
         Case 8:     ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(3, CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo), Format(CDate(ipp_FecCtb.Text), "yyyymmdd"), 2)
         Case 9:     ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(3, CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo), Format(CDate(ipp_FecCtb.Text), "yyyymmdd"), 1)
         Case Else:: ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(2, CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo), Format(CDate(ipp_FecCtb.Text), "yyyymmdd"), 2)
      End Select
   End If
End Sub

Private Sub ipp_MovFec_GotFocus()
   l_var_ColAnt = ipp_MovFec.BackColor
   ipp_MovFec.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_MovFec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_IdeTip)
   End If
End Sub

Private Sub ipp_MovFec_LostFocus()
   ipp_MovFec.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_MtoCta_GotFocus()
   l_var_ColAnt = ipp_MtoCta.BackColor
   ipp_MtoCta.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_MtoCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_DetAce)
   End If
End Sub

Private Sub ipp_MtoCta_LostFocus()
   ipp_MtoCta.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_OrgMto_GotFocus()
   l_var_ColAnt = ipp_OrgMto.BackColor
   ipp_OrgMto.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_OrgMto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_DocAce)
   End If
End Sub

Private Sub ipp_OrgMto_LostFocus()
   ipp_OrgMto.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_TipCam_GotFocus()
   l_var_ColAnt = ipp_TipCam.BackColor
   ipp_TipCam.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_DetNue)
   End If
End Sub

Private Sub ipp_TipCam_LostFocus()
   ipp_TipCam.BackColor = l_var_ColAnt
End Sub

Private Sub msk_MovNum_GotFocus()
   l_var_ColAnt = msk_MovNum.BackColor
   msk_MovNum.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(msk_MovNum)
End Sub

Private Sub msk_MovNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MovFec)
   End If
End Sub

Private Sub msk_MovNum_LostFocus()
   msk_MovNum.BackColor = l_var_ColAnt
End Sub

Private Sub msk_RefOpe_GotFocus()
   l_var_ColAnt = msk_RefOpe.BackColor
   msk_RefOpe.BackColor = modgen_g_con_ColAma
End Sub

Private Sub msk_RefOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BcoTip)
   End If
End Sub

Private Sub msk_RefOpe_LostFocus()
   msk_RefOpe.BackColor = l_var_ColAnt
End Sub

Private Sub msk_RefSol_GotFocus()
   l_var_ColAnt = msk_RefSol.BackColor
   msk_RefSol.BackColor = modgen_g_con_ColAma
End Sub

Private Sub msk_RefSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BcoTip)
   End If
End Sub

Private Sub msk_RefSol_LostFocus()
   msk_RefSol.BackColor = l_var_ColAnt
End Sub

Private Sub txt_BcoNum_GotFocus()
   l_var_ColAnt = txt_BcoNum.BackColor
   txt_BcoNum.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_BcoNum)
End Sub

Private Sub txt_BcoNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_BcoFec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_")
   End If
End Sub

Private Sub txt_BcoNum_LostFocus()
   txt_BcoNum.BackColor = l_var_ColAnt
End Sub

Private Sub txt_DocNum_GotFocus()
   l_var_ColAnt = txt_DocNum.BackColor
   txt_DocNum.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_DocNum)
End Sub

Private Sub txt_DocNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DocFec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DocNum_LostFocus()
   txt_DocNum.Text = Format(txt_DocNum.Text, "000000000000")
   txt_DocNum.BackColor = l_var_ColAnt
End Sub

Private Sub txt_DocSer_GotFocus()
   l_var_ColAnt = txt_DocSer.BackColor
   txt_DocSer.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_DocSer)
End Sub

Private Sub txt_DocSer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DocNum)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DocSer_LostFocus()
   txt_DocSer.Text = Format(txt_DocSer.Text, "00000")
   txt_DocSer.BackColor = l_var_ColAnt
End Sub

Private Sub txt_GloCab_GotFocus()
   l_var_ColAnt = txt_GloCab.BackColor
   txt_GloCab.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_GloCab)
End Sub

Private Sub txt_GloCab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonCtb)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\")
   End If
End Sub

Private Sub txt_GloCab_LostFocus()
   txt_GloCab.BackColor = l_var_ColAnt
End Sub

Private Sub txt_GloDet_GotFocus()
   l_var_ColAnt = txt_GloDet.BackColor
   txt_GloDet.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_GloDet)
End Sub

Private Sub txt_GloDet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMov)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\")
   End If
End Sub

Private Sub txt_GloDet_LostFocus()
   txt_GloDet.BackColor = l_var_ColAnt
End Sub

Private Sub txt_IdeNDo_GotFocus()
   l_var_ColAnt = txt_IdeNDo.BackColor
   txt_IdeNDo.BackColor = modgen_g_con_ColAma
   
   Call gs_SelecTodo(txt_IdeNDo)
End Sub

Private Sub txt_IdeNDo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_RefTip)
   Else
      If cmb_IdeTDo.ListIndex > -1 Then
         Select Case cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex)
            Case 1, 7:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_IdeNDo_LostFocus()
   txt_IdeNDo.BackColor = l_var_ColAnt
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   Call fs_Inicia
   Call fs_LimpiaCab
   Call fs_LimpiaDoc
   Call fs_LimpiaDet
   
   Call fs_Activa(False)
   Call fs_ActivaDoc(True)
   Call fs_ActivaDet(True)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LibCtb(cmb_LibCon)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_DocTip, 1, "257")
   Call moddat_gs_Carga_SucAge(cmb_MovSuc, l_arr_MovSuc, "000001")
   
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_IdeTip, 1, "259")
   Call moddat_gs_Carga_LisIte_Combo(cmb_IdeTDo, 1, "203")
   
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_RefTip, 1, "258")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_BcoTip, 1, "260")
   Call moddat_gs_Carga_LisIte(cmb_BcoCod, l_arr_BcoCod, 1, "516")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_OrgMon, 1, "204")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMov, 1, "255")

   grd_DocAsi.ColWidth(0) = 3185
   grd_DocAsi.ColWidth(1) = 2915
   grd_DocAsi.ColWidth(2) = 2915
   grd_DocAsi.ColWidth(3) = 6305
   grd_DocAsi.ColWidth(4) = 815
   grd_DocAsi.ColWidth(5) = 1295
   grd_DocAsi.ColWidth(6) = 0
   grd_DocAsi.ColWidth(7) = 0
   grd_DocAsi.ColWidth(8) = 0
   grd_DocAsi.ColWidth(9) = 0
   grd_DocAsi.ColWidth(10) = 0
   grd_DocAsi.ColWidth(11) = 0
   grd_DocAsi.ColWidth(12) = 0
   grd_DocAsi.ColWidth(13) = 0
   grd_DocAsi.ColWidth(14) = 0
   grd_DocAsi.ColWidth(15) = 0
   grd_DocAsi.ColWidth(16) = 0
   grd_DocAsi.ColWidth(17) = 0
   grd_DocAsi.ColWidth(18) = 0
   grd_DocAsi.ColWidth(19) = 0
   grd_DocAsi.ColWidth(20) = 0
   grd_DocAsi.ColWidth(21) = 0
   grd_DocAsi.ColWidth(22) = 0
   grd_DocAsi.ColWidth(23) = 0
   grd_DocAsi.ColWidth(24) = 0
   grd_DocAsi.ColWidth(25) = 0

   grd_DocAsi.ColAlignment(0) = flexAlignLeftCenter
   grd_DocAsi.ColAlignment(1) = flexAlignLeftCenter
   grd_DocAsi.ColAlignment(2) = flexAlignLeftCenter
   grd_DocAsi.ColAlignment(3) = flexAlignLeftCenter
   grd_DocAsi.ColAlignment(4) = flexAlignCenterCenter
   grd_DocAsi.ColAlignment(5) = flexAlignRightCenter


   grd_DetAsi.ColWidth(0) = 2705
   grd_DetAsi.ColWidth(1) = 3905
   grd_DetAsi.ColWidth(2) = 5055
   grd_DetAsi.ColWidth(3) = 935
   grd_DetAsi.ColWidth(4) = 1205
   grd_DetAsi.ColWidth(5) = 1205
   grd_DetAsi.ColWidth(6) = 1205
   grd_DetAsi.ColWidth(7) = 1205
   grd_DetAsi.ColWidth(8) = 0
   
   grd_DetAsi.ColAlignment(0) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(1) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(2) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(3) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(4) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(5) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(6) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_LimpiaCab()
   l_str_FecIni = ""
   l_str_FecFin = ""
   l_int_PerAno = 0
   l_int_PerMes = 0

   cmb_Empres.ListIndex = 0
   
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   Call moddat_gs_Carga_ParEmp(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, "102", cmb_MonCtb, l_arr_MonCtb)
   
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, l_str_FecIni, l_str_FecFin, l_int_PerMes, l_int_PerAno)
   
   If Date > CDate(l_str_FecFin) Then
      ipp_FecCtb.Text = Format(l_str_FecFin, "dd/mm/yyyy")
   ElseIf Date < CDate(l_str_FecIni) Then
      ipp_FecCtb.Text = Format(l_str_FecIni, "dd/mm/yyyy")
   End If
   
   ipp_FecCtb.DateMin = Format(CDate(l_str_FecIni), "yyyymmdd")
   ipp_FecCtb.DateMax = Format(CDate(l_str_FecFin), "yyyymmdd")
   
   cmb_Sucurs.ListIndex = 0
   cmb_LibCon.ListIndex = -1
   txt_GloCab.Text = ""
   
   cmb_MonCtb.ListIndex = gf_Busca_Arregl(l_arr_MonCtb, Format("2", "000000")) - 1
   
   ipp_TipCam.Value = 0
   
   Call moddat_gs_Carga_CtaCtb(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)

   Call gs_LimpiaGrid(grd_DocAsi)
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

Private Sub fs_LimpiaDoc()
   'Tipo Comprobante de Pago
   cmb_DocTip.ListIndex = -1
   txt_DocSer.Text = ""
   txt_DocNum.Text = ""
   ipp_DocFec.Text = Format(Date, "dd/mm/yyyy")
   
   txt_DocSer.Enabled = False
   txt_DocNum.Enabled = False
   ipp_DocFec.Enabled = False
   
   cmb_MovSuc.ListIndex = -1
   msk_MovNum.Text = ""
   ipp_MovFec.Enabled = False
   
   cmb_MovSuc.Enabled = False
   msk_MovNum.Enabled = False
   ipp_MovFec.Enabled = False
   
   'Tipo de Persona
   cmb_IdeTip.ListIndex = -1
   cmb_IdeTDo.ListIndex = -1
   txt_IdeNDo.Text = ""
   
   cmb_IdeTDo.Enabled = False
   txt_IdeNDo.Enabled = False
   
   'Tipo Referencia
   cmb_RefTip.ListIndex = -1
   msk_RefOpe.Text = ""
   msk_RefSol.Text = ""
   
   msk_RefOpe.Enabled = False
   msk_RefSol.Enabled = False
   
   'Movimiento Bancario
   cmb_BcoTip.ListIndex = -1
   txt_BcoNum.Text = ""
   cmb_BcoCod.ListIndex = -1
   cmb_BcoCta.Clear
   ipp_BcoFec.Text = Format(Date, "dd/mm/yyyy")
   
   txt_BcoNum.Enabled = False
   cmb_BcoCod.Enabled = False
   cmb_BcoCta.Enabled = False
   ipp_BcoFec.Enabled = False
   
   'Monto Original
   cmb_OrgMon.ListIndex = -1
   ipp_OrgMto.Value = 0
End Sub

Private Sub fs_LimpiaDet()
   cmb_CtaCtb.ListIndex = -1
   txt_GloDet.Text = ""
   cmb_TipMov.ListIndex = -1
   ipp_MtoCta.Value = 0
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_ComGra.Enabled = Not p_Activa

   cmb_Empres.Enabled = Not p_Activa
   cmb_Sucurs.Enabled = Not p_Activa
   cmb_LibCon.Enabled = Not p_Activa
   ipp_FecCtb.Enabled = Not p_Activa
   
   txt_GloCab.Enabled = Not p_Activa
   cmb_MonCtb.Enabled = Not p_Activa
   ipp_TipCam.Enabled = Not p_Activa
End Sub

Private Sub fs_ActivaDoc(ByVal p_Activa As Integer)
   cmd_DocNue.Enabled = p_Activa
   cmd_DocMod.Enabled = p_Activa
   cmd_DocBor.Enabled = p_Activa
   
   grd_DocAsi.Enabled = p_Activa
   
   cmd_DocAce.Enabled = Not p_Activa
   cmd_DocCan.Enabled = Not p_Activa
   
   cmb_DocTip.Enabled = Not p_Activa
   txt_DocSer.Enabled = Not p_Activa
   txt_DocNum.Enabled = Not p_Activa
   ipp_DocFec.Enabled = Not p_Activa
   cmb_MovSuc.Enabled = Not p_Activa
   msk_MovNum.Enabled = Not p_Activa
   ipp_MovFec.Enabled = Not p_Activa
   
   cmb_IdeTip.Enabled = Not p_Activa
   cmb_IdeTDo.Enabled = Not p_Activa
   txt_IdeNDo.Enabled = Not p_Activa
   
   cmb_RefTip.Enabled = Not p_Activa
   msk_RefOpe.Enabled = Not p_Activa
   msk_RefSol.Enabled = Not p_Activa
   
   cmb_BcoTip.Enabled = Not p_Activa
   cmb_BcoCod.Enabled = Not p_Activa
   cmb_BcoCta.Enabled = Not p_Activa
   txt_BcoNum.Enabled = Not p_Activa
   ipp_BcoFec.Enabled = Not p_Activa
   
   cmb_OrgMon.Enabled = Not p_Activa
   ipp_OrgMto.Enabled = Not p_Activa
End Sub

Private Sub fs_ActivaDet(ByVal p_Activa As Integer)
   cmd_DetNue.Enabled = p_Activa
   cmd_DetMod.Enabled = p_Activa
   cmd_DetBor.Enabled = p_Activa
   
   grd_DetAsi.Enabled = p_Activa
   
   cmd_DetAce.Enabled = Not p_Activa
   cmd_DetCan.Enabled = Not p_Activa
   
   cmb_CtaCtb.Enabled = Not p_Activa
   txt_GloDet.Enabled = Not p_Activa
   cmb_TipMov.Enabled = Not p_Activa
   ipp_MtoCta.Enabled = Not p_Activa
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
      
      grd_DetAsi.Col = 4:  r_dbl_TDebMN = r_dbl_TDebMN + CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 5:  r_dbl_THabMN = r_dbl_THabMN + CDbl(grd_DetAsi.Text)
      
      grd_DetAsi.Col = 6:  r_dbl_TDebME = r_dbl_TDebME + CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 7:  r_dbl_THabME = r_dbl_THabME + CDbl(grd_DetAsi.Text)
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

Private Sub cmd_DetAce_Click()
   If cmb_CtaCtb.ListIndex = -1 Then
      MsgBox "Debe ingresar la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCtb)
      
      Exit Sub
   End If
   
   If Len(Trim(txt_GloDet.Text)) = 0 Then
      MsgBox "Debe ingresar la Glosa de Detalle.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GloDet)
      
      Exit Sub
   End If
   
   If cmb_TipMov.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMov)
      
      Exit Sub
   End If
   
   If CDbl(ipp_MtoCta.Value) = 0 Then
      MsgBox "Debe ingresar el Monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoCta)
      
      Exit Sub
   End If
   
   If MsgBox("�Esta seguro de agregar la Cuenta?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_DetAsi.Redraw = False
   
   If l_int_GrbDet = 1 Then
      grd_DetAsi.Rows = grd_DetAsi.Rows + 1
      grd_DetAsi.Row = grd_DetAsi.Rows - 1
   End If
   
   grd_DetAsi.Col = 0
   grd_DetAsi.Text = l_arr_CtaCtb(cmb_CtaCtb.ListIndex + 1).Genera_Codigo
   
   grd_DetAsi.Col = 1
   grd_DetAsi.Text = l_arr_CtaCtb(cmb_CtaCtb.ListIndex + 1).Genera_Nombre
   
   grd_DetAsi.Col = 2
   grd_DetAsi.Text = txt_GloDet.Text
   
   grd_DetAsi.Col = 3
   grd_DetAsi.Text = cmb_TipMov.Text
   
   grd_DetAsi.Col = 4:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 5:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
   
   If CInt(Mid(l_arr_CtaCtb(cmb_CtaCtb.ListIndex + 1).Genera_Codigo, 3, 1)) = 1 Then
      If cmb_TipMov.ItemData(cmb_TipMov.ListIndex) = 1 Then
         'Debe MN
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")
         
         'Debe ME
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) / CDbl(ipp_TipCam.Value), "###,###,##0.00")
         
         'Haber MN
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = "0.00"
      
         'Haber ME
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
      Else
         'Debe MN
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = "0.00"
         
         'Debe ME
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
         
         'Haber MN
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")
         
         'Haber ME
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) / CDbl(ipp_TipCam.Value), "###,###,##0.00")
      End If
   Else
      If cmb_TipMov.ItemData(cmb_TipMov.ListIndex) = 1 Then
         'Debe MN
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) * CDbl(ipp_TipCam.Value), "###,###,##0.00")
         
         'Debe ME
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")
         
         'Haber MN
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = "0.00"
         
         'Haber ME
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
      Else
         'Debe MN
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = "0.00"
         
         'Debe ME
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
         
         'Haber MN
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) * CDbl(ipp_TipCam.Value), "###,###,##0.00")
         
         'Haber ME
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")
      End If
   End If
   
   grd_DetAsi.Col = 8
   grd_DetAsi.Text = cmb_TipMov.ItemData(cmb_TipMov.ListIndex)
   
   grd_DetAsi.Redraw = True
   
   Call gs_RefrescaGrid(grd_DetAsi)
   
   Call fs_TotDebHab
   
   If l_int_GrbDet = 1 Then
      Call cmd_DetNue_Click
   Else
      Call cmd_DetCan_Click
   End If
End Sub

Private Sub cmd_DetCan_Click()
   Call fs_LimpiaDet
   Call fs_ActivaDet(True)
   
   Call gs_SetFocus(grd_DetAsi)
End Sub

Private Sub cmd_DetNue_Click()
   If CDbl(ipp_TipCam.Value) = 0# Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   
   l_int_GrbDet = 1
   
   Call fs_LimpiaDet
   Call fs_ActivaDet(False)
   Call gs_SetFocus(cmb_CtaCtb)
End Sub

Private Sub cmd_DetMod_Click()
   Dim r_str_CtaCtb     As String
   Dim r_str_GloDet     As String
   Dim r_int_DebHab     As Integer
   Dim r_dbl_MtoCta     As Double
   
   If grd_DetAsi.Rows = 0 Then
      Exit Sub
   End If
   
   
   'Leyendo de Grid
   grd_DetAsi.Redraw = False
   
   grd_DetAsi.Col = 0:  r_str_CtaCtb = grd_DetAsi.Text
   grd_DetAsi.Col = 2:  r_str_GloDet = grd_DetAsi.Text
   grd_DetAsi.Col = 8:  r_int_DebHab = CInt(grd_DetAsi.Text)
   
   If Mid(r_str_CtaCtb, 3, 1) = 1 Then
      If r_int_DebHab = 1 Then
         grd_DetAsi.Col = 4:  r_dbl_MtoCta = CDbl(grd_DetAsi.Text)
      Else
         grd_DetAsi.Col = 5:  r_dbl_MtoCta = CDbl(grd_DetAsi.Text)
      End If
   Else
      If r_int_DebHab = 1 Then
         grd_DetAsi.Col = 6:  r_dbl_MtoCta = CDbl(grd_DetAsi.Text)
      Else
         grd_DetAsi.Col = 7:  r_dbl_MtoCta = CDbl(grd_DetAsi.Text)
      End If
   End If

   Call gs_RefrescaGrid(grd_DetAsi)

   grd_DetAsi.Redraw = True
   
   Call fs_ActivaDet(False)
   
   'Igualando a Controles
'   Call gs_BuscarCombo_Item(cmb_TipCta, CInt(Left(r_str_CtaCtb, 1)))
'   Call moddat_gs_Carga_CtaCtb(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, cmb_CtaCtb, l_arr_CtaCtb, cmb_TipCta.ItemData(cmb_TipCta.ListIndex), l_int_TopNiv, -1)

   cmb_CtaCtb.ListIndex = gf_Busca_Arregl(l_arr_CtaCtb, r_str_CtaCtb) - 1
   
   txt_GloDet.Text = r_str_GloDet
   
   Call gs_BuscarCombo_Item(cmb_TipMov, r_int_DebHab)

   ipp_MtoCta.Value = r_dbl_MtoCta

   
   Call gs_SetFocus(cmb_CtaCtb)

   l_int_GrbDet = 2
End Sub

Private Sub cmd_DocCan_Click()
   Call fs_LimpiaDoc
   Call fs_ActivaDoc(True)
   
   Call gs_SetFocus(grd_DocAsi)
End Sub

Private Sub cmd_DocNue_Click()
   l_int_GrbDoc = 1
   
   Call fs_ActivaDoc(False)
   Call fs_LimpiaDoc
   Call gs_SetFocus(cmb_DocTip)
End Sub

Private Sub cmd_DetBor_Click()
   If grd_DetAsi.Rows = 0 Then
      Exit Sub
   End If
   
   If grd_DetAsi.Rows = 1 Then
      Call gs_LimpiaGrid(grd_DetAsi)
   Else
      grd_DetAsi.RemoveItem grd_DetAsi.Row
      grd_DetAsi.Row = 0
   End If
   
   Call fs_TotDebHab
End Sub

Private Sub cmd_DocAce_Click()
   Dim r_int_Contad     As Integer

   If cmb_DocTip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Comprobante de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DocTip)
      Exit Sub
   ElseIf cmb_DocTip.ItemData(cmb_DocTip.ListIndex) = 5 Then
      If cmb_MovSuc.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Sucursal del Comprobante de Pago.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_MovSuc)
         Exit Sub
      End If
      
      If Len(Trim(msk_MovNum.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de Comprobante de Pago.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_MovNum)
         Exit Sub
      End If
   ElseIf cmb_DocTip.ItemData(cmb_DocTip.ListIndex) <> 1 Then
      If Len(Trim(txt_DocSer.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de Serie de " & Left(cmb_DocTip.Text, 1) & LCase(Mid(cmb_DocTip.Text, 2)) & ".", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DocSer)
         Exit Sub
      End If
      
      If Len(Trim(txt_DocNum.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de " & Left(cmb_DocTip.Text, 1) & LCase(Mid(cmb_DocTip.Text, 2)) & ".", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DocNum)
         Exit Sub
      End If
      
      txt_DocSer.Text = Format(txt_DocSer.Text, "00000")
      txt_DocNum.Text = Format(txt_DocNum.Text, "000000000000")
   End If
   
   If cmb_IdeTip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Persona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_IdeTip)
      Exit Sub
   ElseIf cmb_IdeTip.ItemData(cmb_IdeTip.ListIndex) <> 1 Then
      If cmb_IdeTDo.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_IdeTDo)
         Exit Sub
      End If
      
      If cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex) = 1 Then
         If Len(Trim(txt_IdeNDo.Text)) <> 8 Then
            MsgBox "Debe ingresar correctamente el N�mero de DNI.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_IdeNDo)
            Exit Sub
         End If
      ElseIf cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex) = 7 Then
         If Len(Trim(txt_IdeNDo.Text)) <> 11 Then
            MsgBox "Debe ingresar correctamente el N�mero de RUC.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_IdeNDo)
            Exit Sub
         End If
      End If
   End If
   
   If cmb_RefTip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Referencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_RefTip)
      Exit Sub
   ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 2 Then
      If Len(Trim(msk_RefOpe.Text)) <> 10 Then
         MsgBox "Debe ingresar el N�mero de Operaci�n Crediticia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_RefOpe)
         Exit Sub
      End If
   ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 3 Then
      If Len(Trim(msk_RefSol.Text)) <> 12 Then
         MsgBox "Debe ingresar el N�mero de Solicitud de Cr�dito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_RefSol)
         Exit Sub
      End If
   End If
   
   If cmb_BcoTip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Movimiento Bancario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BcoTip)
      Exit Sub
   ElseIf cmb_BcoTip.ItemData(cmb_BcoTip.ListIndex) <> 1 Then
      If cmb_BcoCod.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_BcoCod)
         Exit Sub
      End If
      
      If cmb_BcoCta.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_BcoCta)
         Exit Sub
      End If
      
      If Len(Trim(txt_BcoNum.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de Referencia Bancaria (Cheque, Transferencia, etc.).", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_BcoNum)
         Exit Sub
      End If
   End If
   
   If cmb_OrgMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OrgMon)
      Exit Sub
   End If
   
   If CDbl(ipp_OrgMto.Value) = 0 Then
      MsgBox "Debe ingresar el Monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_OrgMto)
      Exit Sub
   End If

   If MsgBox("�Esta seguro de agregar el Documento de Referencia?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   grd_DocAsi.Redraw = False
   
   If l_int_GrbDoc = 1 Then
      grd_DocAsi.Rows = grd_DocAsi.Rows + 1
      grd_DocAsi.Row = grd_DocAsi.Rows - 1
   End If
   
   'Inicializando Fila
   For r_int_Contad = 0 To grd_DocAsi.Cols - 1
      grd_DocAsi.Col = r_int_Contad
      grd_DocAsi.Text = ""
   Next r_int_Contad
   
   grd_DocAsi.Col = 6:        grd_DocAsi.Text = cmb_DocTip.ItemData(cmb_DocTip.ListIndex)
   grd_DocAsi.Col = 0
   If cmb_DocTip.ItemData(cmb_DocTip.ListIndex) = 1 Then
      grd_DocAsi.Text = "---"
      grd_DocAsi.CellAlignment = flexAlignCenterCenter
   ElseIf cmb_DocTip.ItemData(cmb_DocTip.ListIndex) = 5 Then
      grd_DocAsi.Text = Left(cmb_DocTip.Text, 3) & " (" & l_arr_MovSuc(cmb_MovSuc.ListIndex + 1).Genera_Codigo & "-" & msk_MovNum.Text & " - " & ipp_MovFec.Text & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 7:     grd_DocAsi.Text = l_arr_MovSuc(cmb_MovSuc.ListIndex + 1).Genera_Codigo
      grd_DocAsi.Col = 8:     grd_DocAsi.Text = msk_MovNum.Text
      grd_DocAsi.Col = 9:     grd_DocAsi.Text = ipp_MovFec.Text
   Else
      grd_DocAsi.Text = Left(cmb_DocTip.Text, 3) & " (" & txt_DocSer.Text & "-" & txt_DocNum.Text & " - " & ipp_DocFec.Text & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 10:    grd_DocAsi.Text = txt_DocSer.Text
      grd_DocAsi.Col = 11:    grd_DocAsi.Text = txt_DocNum.Text
      grd_DocAsi.Col = 12:    grd_DocAsi.Text = ipp_DocFec.Text
   End If
   
   
   grd_DocAsi.Col = 13:       grd_DocAsi.Text = cmb_IdeTip.ItemData(cmb_IdeTip.ListIndex)
   grd_DocAsi.Col = 1
   If cmb_IdeTip.ItemData(cmb_IdeTip.ListIndex) = 1 Then
      grd_DocAsi.Text = "---"
      grd_DocAsi.CellAlignment = flexAlignCenterCenter
   Else
      grd_DocAsi.Text = Left(cmb_IdeTip.Text, 3) & " (" & CStr(cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex)) & "-" & Trim(txt_IdeNDo.Text) & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 14:    grd_DocAsi.Text = CStr(cmb_IdeTDo.ItemData(cmb_IdeTDo.ListIndex))
      grd_DocAsi.Col = 15:    grd_DocAsi.Text = Trim(txt_IdeNDo.Text)
   End If
   
   grd_DocAsi.Col = 16:       grd_DocAsi.Text = cmb_RefTip.ItemData(cmb_RefTip.ListIndex)
   grd_DocAsi.Col = 2
   If cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 1 Then
      grd_DocAsi.Text = "---"
      grd_DocAsi.CellAlignment = flexAlignCenterCenter
   ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 2 Then
      grd_DocAsi.Text = Left(cmb_RefTip.Text, 3) & " (" & gf_Formato_NumOpe(msk_RefOpe.Text) & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 17:    grd_DocAsi.Text = msk_RefOpe.Text
   ElseIf cmb_RefTip.ItemData(cmb_RefTip.ListIndex) = 3 Then
      grd_DocAsi.Text = Left(cmb_RefTip.Text, 3) & " (" & gf_Formato_NumSol(msk_RefSol.Text) & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 18:    grd_DocAsi.Text = msk_RefSol.Text
   End If
   
   grd_DocAsi.Col = 19:       grd_DocAsi.Text = cmb_BcoTip.ItemData(cmb_BcoTip.ListIndex)
   grd_DocAsi.Col = 3
   If cmb_BcoTip.ItemData(cmb_BcoTip.ListIndex) = 1 Then
      grd_DocAsi.Text = "---"
      grd_DocAsi.CellAlignment = flexAlignCenterCenter
   Else
      grd_DocAsi.Text = Left(cmb_BcoTip.Text, 3) & " (" & cmb_BcoCod.Text & " - " & l_arr_BcoCta(cmb_BcoCta.ListIndex + 1).Genera_Codigo & " - " & txt_BcoNum.Text & " - " & ipp_BcoFec.Text & ")"
      grd_DocAsi.CellAlignment = flexAlignLeftCenter
      
      grd_DocAsi.Col = 20:    grd_DocAsi.Text = l_arr_BcoCod(cmb_BcoCod.ListIndex + 1).Genera_Codigo
      grd_DocAsi.Col = 21:    grd_DocAsi.Text = l_arr_BcoCta(cmb_BcoCta.ListIndex + 1).Genera_Codigo
      grd_DocAsi.Col = 22:    grd_DocAsi.Text = txt_BcoNum.Text
      grd_DocAsi.Col = 23:    grd_DocAsi.Text = ipp_BcoFec.Text
   End If
   
   grd_DocAsi.Col = 4:        grd_DocAsi.Text = moddat_gf_Consulta_ParDes("229", CStr(cmb_OrgMon.ItemData(cmb_OrgMon.ListIndex)))
   grd_DocAsi.Col = 24:       grd_DocAsi.Text = CStr(cmb_OrgMon.ItemData(cmb_OrgMon.ListIndex))
   
   grd_DocAsi.Col = 5:        grd_DocAsi.Text = Format(CDbl(ipp_OrgMto.Text), "###,###,##0.00")
   grd_DocAsi.Col = 25:       grd_DocAsi.Text = CStr(CDbl(ipp_OrgMto.Text))
   
   grd_DocAsi.Redraw = True
   
   Call gs_RefrescaGrid(grd_DocAsi)
   
   If l_int_GrbDoc = 1 Then
      Call cmd_DocNue_Click
   Else
      Call cmd_DocCan_Click
   End If
End Sub

Private Sub cmd_DocBor_Click()
   If grd_DocAsi.Rows = 0 Then
      Exit Sub
   End If
   
   If grd_DocAsi.Rows = 1 Then
      Call gs_LimpiaGrid(grd_DocAsi)
   Else
      grd_DocAsi.RemoveItem grd_DocAsi.Row
      grd_DocAsi.Row = 0
   End If
End Sub

Private Sub cmd_DocMod_Click()
   Dim r_int_DocTip     As Integer
   Dim r_str_MovSuc     As String
   Dim r_str_MovNum     As String
   Dim r_str_MovFec     As String
   Dim r_str_DocSer     As String
   Dim r_str_DocNum     As String
   Dim r_str_DocFec     As String
   
   Dim r_int_IdeTip     As Integer
   Dim r_int_IdeTDo     As Integer
   Dim r_str_IdeNDo     As String
   
   Dim r_int_RefTip     As Integer
   Dim r_str_RefOpe     As String
   Dim r_str_RefSol     As String

   Dim r_int_BcoTip     As Integer
   Dim r_str_BcoCod     As String
   Dim r_str_BcoCta     As String
   Dim r_str_BcoNum     As String
   Dim r_str_BcoFec     As String

   Dim r_int_OrgMon     As Integer
   Dim r_dbl_OrgMto     As Double

   If grd_DocAsi.Rows = 0 Then
      Exit Sub
   End If
   
   grd_DocAsi.Redraw = False

   'Documento de Referencia
   grd_DocAsi.Col = 6:        r_int_DocTip = CInt(grd_DocAsi.Text)
   grd_DocAsi.Col = 7:        r_str_MovSuc = grd_DocAsi.Text
   grd_DocAsi.Col = 8:        r_str_MovNum = grd_DocAsi.Text
   grd_DocAsi.Col = 9:        r_str_MovFec = grd_DocAsi.Text
   grd_DocAsi.Col = 10:       r_str_DocSer = grd_DocAsi.Text
   grd_DocAsi.Col = 11:       r_str_DocNum = grd_DocAsi.Text
   grd_DocAsi.Col = 12:       r_str_DocFec = grd_DocAsi.Text

   'Persona de Referencia
   grd_DocAsi.Col = 13:       r_int_IdeTip = CInt(grd_DocAsi.Text)
   
   If r_int_IdeTip <> 1 Then
      grd_DocAsi.Col = 14:    r_int_IdeTDo = CInt(grd_DocAsi.Text)
   End If
   
   grd_DocAsi.Col = 15:       r_str_IdeNDo = grd_DocAsi.Text
   
   'Operaci�n Financiera de Referencia
   grd_DocAsi.Col = 16:       r_int_RefTip = CInt(grd_DocAsi.Text)
   grd_DocAsi.Col = 17:       r_str_RefOpe = grd_DocAsi.Text
   grd_DocAsi.Col = 18:       r_str_RefSol = grd_DocAsi.Text
   
   'Operaci�n Bancaria de Referencia
   grd_DocAsi.Col = 19:       r_int_BcoTip = CInt(grd_DocAsi.Text)
   grd_DocAsi.Col = 20:       r_str_BcoCod = grd_DocAsi.Text
   grd_DocAsi.Col = 21:       r_str_BcoCta = grd_DocAsi.Text
   grd_DocAsi.Col = 22:       r_str_BcoNum = grd_DocAsi.Text
   grd_DocAsi.Col = 23:       r_str_BcoFec = grd_DocAsi.Text
   
   'Otros Datos
   grd_DocAsi.Col = 24:       r_int_OrgMon = CInt(grd_DocAsi.Text)
   grd_DocAsi.Col = 25:       r_dbl_OrgMto = CDbl(grd_DocAsi.Text)
   
   Call gs_RefrescaGrid(grd_DocAsi)

   grd_DocAsi.Redraw = True

   Call fs_ActivaDoc(False)

   'Documento de Referencia
   Call gs_BuscarCombo_Item(cmb_DocTip, r_int_DocTip)

   If r_int_DocTip <> 1 Then
      If r_int_DocTip = 5 Then
         cmb_MovSuc.ListIndex = gf_Busca_Arregl(l_arr_MovSuc, r_str_MovSuc) - 1
         msk_MovNum.Text = r_str_MovNum
         ipp_MovFec.Text = r_str_MovFec
      Else
         txt_DocSer.Text = r_str_DocSer
         txt_DocNum.Text = r_str_DocNum
         ipp_DocFec.Text = r_str_DocFec
      End If
   End If
   
   Call cmb_DocTip_Click

   'Persona de Referencia
   Call gs_BuscarCombo_Item(cmb_IdeTip, r_int_IdeTip)
   
   If r_int_IdeTip <> 1 Then
      Call gs_BuscarCombo_Item(cmb_IdeTDo, r_int_IdeTDo)
      txt_IdeNDo.Text = r_str_IdeNDo
   End If
   
   Call cmb_IdeTip_Click

   'Operaci�n de Referencia
   Call gs_BuscarCombo_Item(cmb_RefTip, r_int_RefTip)
   If r_int_RefTip <> 1 Then
      If r_int_RefTip = 2 Then
         msk_RefOpe.Text = r_str_RefOpe
      ElseIf r_int_RefTip = 3 Then
         msk_RefSol.Text = r_str_RefSol
      End If
   End If
   
   Call cmb_RefTip_Click
   
   'Operaci�n Bancaria de Referencia
   Call gs_BuscarCombo_Item(cmb_BcoTip, r_int_BcoTip)
   If r_int_BcoTip <> 1 Then
      cmb_BcoCod.ListIndex = gf_Busca_Arregl(l_arr_BcoCod, r_str_BcoCod) - 1
      
      Call moddat_gs_Carga_CtaBan(l_arr_BcoCod(cmb_BcoCod.ListIndex + 1).Genera_Codigo, cmb_BcoCta, l_arr_BcoCta)
      cmb_BcoCta.ListIndex = gf_Busca_Arregl(l_arr_BcoCta, r_str_BcoCta) - 1
      
      txt_BcoNum.Text = r_str_BcoNum
      ipp_BcoFec.Text = r_str_BcoFec
   End If
   
   Call cmb_BcoTip_Click
   
   'Otros Datos
   Call gs_BuscarCombo_Item(cmb_OrgMon, r_int_OrgMon)
   ipp_OrgMto.Value = r_dbl_OrgMto
   
   Call gs_SetFocus(cmb_DocTip)

   l_int_GrbDoc = 2
End Sub

Private Sub cmd_ComGra_Click()
   Dim r_int_CuaAsi     As Integer
   Dim r_lng_NumAsi     As Long
   
   Dim r_int_DocTip     As Integer
   Dim r_str_MovSuc     As String
   Dim r_str_MovNum     As String
   Dim r_str_MovFec     As String
   Dim r_str_DocSer     As String
   Dim r_str_DocNum     As String
   Dim r_str_DocFec     As String
   
   Dim r_int_IdeTip     As Integer
   Dim r_int_IdeTDo     As Integer
   Dim r_str_IdeNDo     As String
   
   Dim r_int_RefTip     As Integer
   Dim r_str_RefOpe     As String
   Dim r_str_RefSol     As String

   Dim r_int_BcoTip     As Integer
   Dim r_str_BcoCod     As String
   Dim r_str_BcoCta     As String
   Dim r_str_BcoNum     As String
   Dim r_str_BcoFec     As String

   Dim r_int_OrgMon     As Integer
   Dim r_dbl_OrgMto     As Double
   
   Dim r_int_Contad     As Integer
   
   Dim r_str_CtaCtb     As String
   Dim r_str_GloDet     As String
   Dim r_int_DebHab     As Integer
   Dim r_dbl_MtoCta_DMN As Double
   Dim r_dbl_MtoCta_DME As Double
   Dim r_dbl_MtoCta_HMN As Double
   Dim r_dbl_MtoCta_HME As Double
   
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   
   If cmb_Sucurs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Sucurs)
      Exit Sub
   End If
   
   If cmb_LibCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Libro Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_LibCon)
      Exit Sub
   End If
   
   If Len(Trim(txt_GloCab.Text)) = 0 Then
      MsgBox "Debe ingresar la Glosa de Cabecera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GloCab)
      Exit Sub
   End If
   
   If cmb_MonCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda del Asiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonCtb)
      Exit Sub
   End If
   
   If CDbl(ipp_TipCam.Value) = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   
   If grd_DetAsi.Rows = 0 Then
      MsgBox "Debe ingresar Detalle de Asiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_DetAsi)
      Exit Sub
   End If
   
   r_int_CuaAsi = 1
   If CDbl(pnl_DifDeb_MN.Caption) > 0 Or CDbl(pnl_DifDeb_ME.Caption) > 0 Or CDbl(pnl_DifHab_MN.Caption) > 0 Or CDbl(pnl_DifHab_ME.Caption) > 0 Then
      If MsgBox("El Asiento se encuentra descuadrado. �Desea continuar sin cuadrar el asiento?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Call gs_SetFocus(cmd_DetNue)
         Exit Sub
      End If
      
      r_int_CuaAsi = 2
   End If
   
   If grd_DocAsi.Rows = 0 Then
      If MsgBox("No ha ingresado Documentos de Referencia del Asiento. �Desea continuar sin ingresar esta informaci�n?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Call gs_SetFocus(cmd_DocNue)
         Exit Sub
      End If
   End If
   
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Obtener N�mero de Asiento
   r_lng_NumAsi = modctb_gf_Genera_NumAsi(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo, l_int_PerAno, l_int_PerMes, cmb_LibCon.ItemData(cmb_LibCon.ListIndex))
   
   'Grabando Documentos de Referencia
   grd_DocAsi.Redraw = False
   
   For r_int_Contad = 0 To grd_DocAsi.Rows - 1
      grd_DocAsi.Row = r_int_Contad
      
      'Documento de Referencia
      grd_DocAsi.Col = 6:        r_int_DocTip = CInt(grd_DocAsi.Text)
      grd_DocAsi.Col = 7:        r_str_MovSuc = grd_DocAsi.Text
      grd_DocAsi.Col = 8:        r_str_MovNum = grd_DocAsi.Text
      grd_DocAsi.Col = 9:        r_str_MovFec = grd_DocAsi.Text
      grd_DocAsi.Col = 10:       r_str_DocSer = grd_DocAsi.Text
      grd_DocAsi.Col = 11:       r_str_DocNum = grd_DocAsi.Text
      grd_DocAsi.Col = 12:       r_str_DocFec = grd_DocAsi.Text
   
      'Persona de Referencia
      grd_DocAsi.Col = 13:       r_int_IdeTip = CInt(grd_DocAsi.Text)
      
      r_int_IdeTDo = 0
      If r_int_IdeTip <> 1 Then
         grd_DocAsi.Col = 14:    r_int_IdeTDo = CInt(grd_DocAsi.Text)
      End If
      
      grd_DocAsi.Col = 15:       r_str_IdeNDo = grd_DocAsi.Text
      
      'Operaci�n Financiera de Referencia
      grd_DocAsi.Col = 16:       r_int_RefTip = CInt(grd_DocAsi.Text)
      grd_DocAsi.Col = 17:       r_str_RefOpe = grd_DocAsi.Text
      grd_DocAsi.Col = 18:       r_str_RefSol = grd_DocAsi.Text
      
      'Operaci�n Bancaria de Referencia
      grd_DocAsi.Col = 19:       r_int_BcoTip = CInt(grd_DocAsi.Text)
      grd_DocAsi.Col = 20:       r_str_BcoCod = grd_DocAsi.Text
      grd_DocAsi.Col = 21:       r_str_BcoCta = grd_DocAsi.Text
      grd_DocAsi.Col = 22:       r_str_BcoNum = grd_DocAsi.Text
      grd_DocAsi.Col = 23:       r_str_BcoFec = grd_DocAsi.Text
      
      'Otros Datos
      grd_DocAsi.Col = 24:       r_int_OrgMon = CInt(grd_DocAsi.Text)
      grd_DocAsi.Col = 25:       r_dbl_OrgMto = CDbl(grd_DocAsi.Text)
      
      
      'Grabando en BD
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_ASIDOC ("
         
         'Datos Principales
         g_str_Parame = g_str_Parame & "'" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_PerAno) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_PerMes) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(r_lng_NumAsi) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_Contad + 1) & ", "
         
         'Documento de Referencia
         g_str_Parame = g_str_Parame & CStr(r_int_DocTip) & ", "
         
         If r_int_DocTip = 1 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         ElseIf r_int_DocTip = 5 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'" & r_str_MovSuc & "', "
            g_str_Parame = g_str_Parame & r_str_MovNum & ", "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_MovFec), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & r_str_DocSer & ", "
            g_str_Parame = g_str_Parame & r_str_DocNum & ", "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_DocFec), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         'Tipo de Operaci�n
         g_str_Parame = g_str_Parame & CStr(r_int_RefTip) & ", "
         
         If r_int_RefTip = 1 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         ElseIf r_int_RefTip = 2 Then
            g_str_Parame = g_str_Parame & "'" & r_str_RefOpe & "', "
            g_str_Parame = g_str_Parame & "'', "
         ElseIf r_int_RefTip = 3 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'" & r_str_RefSol & "', "
         End If
         
         'Tipo de Persona
         g_str_Parame = g_str_Parame & CStr(r_int_IdeTip) & ", "
         
         If r_int_IdeTip = 1 Then
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
         Else
            g_str_Parame = g_str_Parame & CStr(r_int_IdeTDo) & ", "
            g_str_Parame = g_str_Parame & "'" & r_str_IdeNDo & "', "
         End If
         
         'Movimiento Bancario
         g_str_Parame = g_str_Parame & CStr(r_int_BcoTip) & ", "
         
         If r_int_BcoTip = 1 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & "'" & r_str_BcoCod & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_BcoCta & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_BcoNum & "', "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_BcoFec), "yyyymmdd") & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(r_int_OrgMon) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_OrgMto) & ", "
         
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_CTB_ASIDOC. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Next r_int_Contad
   
   grd_DocAsi.Redraw = True
   
   Call gs_RefrescaGrid(grd_DocAsi)
   
   'Grabando Detalle de Asiento
   grd_DetAsi.Redraw = False
   
   For r_int_Contad = 0 To grd_DetAsi.Rows - 1
      grd_DetAsi.Row = r_int_Contad
      
      grd_DetAsi.Col = 0:  r_str_CtaCtb = grd_DetAsi.Text
      grd_DetAsi.Col = 2:  r_str_GloDet = grd_DetAsi.Text
      grd_DetAsi.Col = 8:  r_int_DebHab = CInt(grd_DetAsi.Text)
      
      grd_DetAsi.Col = 4:  r_dbl_MtoCta_DMN = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 5:  r_dbl_MtoCta_HMN = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 6:  r_dbl_MtoCta_DME = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 7:  r_dbl_MtoCta_HME = CDbl(grd_DetAsi.Text)
      
      'Grabando en BD
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_ASIDET ("
         
         'Datos Principales
         g_str_Parame = g_str_Parame & "'" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_PerAno) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_PerMes) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(r_lng_NumAsi) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_Contad + 1) & ", "
         
         'Datos de Linea
         g_str_Parame = g_str_Parame & "'" & r_str_CtaCtb & "',"
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCtb.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_DebHab) & ","
         
         If r_int_DebHab = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_DMN) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_DME) & ", "
         Else
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_HMN) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_HME) & ", "
         End If
         
         g_str_Parame = g_str_Parame & "'" & r_str_GloDet & "',"
         
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_CTB_ASIDET. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Next r_int_Contad
   
   grd_DetAsi.Redraw = True
   
   Call gs_RefrescaGrid(grd_DetAsi)
   
   'Grabando Cabecera de Asiento
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_ASICAB ("
      
      'Datos Principales
      g_str_Parame = g_str_Parame & "'" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_PerAno) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumAsi) & ", "
      
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCtb.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_TipCam.Value)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_GloCab.Text & "', "
      
      If r_int_CuaAsi = 2 Then
         g_str_Parame = g_str_Parame & "1, "
      Else
         g_str_Parame = g_str_Parame & "2, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotDeb_MN.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotHab_MN.Caption)) & ", "
      
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotDeb_ME.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotHab_ME.Caption)) & ", "

      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "

      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1) "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CTB_ASICAB. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   MsgBox "El N�mero de Asiento generado es el " & CStr(r_lng_NumAsi), vbInformation, modgen_g_str_NomPlt
   
   Call cmd_ComNue_Click
End Sub

Private Sub cmd_ComNue_Click()
   Call fs_LimpiaCab
   Call fs_LimpiaDoc
   Call fs_LimpiaDet
   
   Call fs_Activa(False)
   Call fs_ActivaDoc(True)
   Call fs_ActivaDet(True)
   
   Call gs_SetFocus(cmb_LibCon)
End Sub


