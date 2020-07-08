VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_AsiCtb_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8910
   ClientLeft      =   450
   ClientTop       =   1500
   ClientWidth     =   17940
   Icon            =   "GesCtb_frm_161.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17925
      _Version        =   65536
      _ExtentX        =   31618
      _ExtentY        =   15690
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
         Height          =   2745
         Left            =   30
         TabIndex        =   1
         Top             =   6090
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   4842
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
            TabIndex        =   2
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
            TabIndex        =   3
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
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   6
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
            Height          =   2355
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   4154
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
            TabIndex        =   8
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
         Height          =   2955
         Left            =   30
         TabIndex        =   9
         Top             =   3090
         Width           =   17835
         _Version        =   65536
         _ExtentX        =   31459
         _ExtentY        =   5212
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
            TabIndex        =   10
            Top             =   60
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Cuenta Contable"
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
            TabIndex        =   11
            Top             =   60
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
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
            Left            =   6660
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
            Top             =   2280
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
            TabIndex        =   19
            Top             =   2280
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
            TabIndex        =   20
            Top             =   2280
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
            TabIndex        =   21
            Top             =   2280
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
            TabIndex        =   22
            Top             =   2580
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
            TabIndex        =   23
            Top             =   2580
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
            TabIndex        =   24
            Top             =   2580
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
            TabIndex        =   25
            Top             =   2580
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
            Height          =   1875
            Left            =   30
            TabIndex        =   26
            Top             =   360
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   3307
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
            TabIndex        =   28
            Top             =   2580
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   11340
            TabIndex        =   27
            Top             =   2280
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   29
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
         Begin Threed.SSPanel pnl_MonCtb 
            Height          =   315
            Left            =   1530
            TabIndex        =   51
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
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   8070
            TabIndex        =   52
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
         Begin Threed.SSPanel pnl_GloCab 
            Height          =   315
            Left            =   1530
            TabIndex        =   53
            Top             =   60
            Width           =   16275
            _Version        =   65536
            _ExtentX        =   28707
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
         Begin VB.Label Label7 
            Caption         =   "Glosa Cabecera:"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Cambio:"
            Height          =   255
            Left            =   6420
            TabIndex        =   31
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   33
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
            TabIndex        =   34
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_161.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   35
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
         Begin VB.CommandButton cmd_ComGra 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_161.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Grabar Comprobante Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   17220
            Picture         =   "GesCtb_frm_161.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   38
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
         Begin Threed.SSPanel pnl_Empres 
            Height          =   315
            Left            =   1530
            TabIndex        =   39
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
            Left            =   1530
            TabIndex        =   40
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
            Left            =   14370
            TabIndex        =   41
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
            Left            =   8070
            TabIndex        =   42
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
            Left            =   14370
            TabIndex        =   49
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
            Left            =   8070
            TabIndex        =   50
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   15780
            TabIndex        =   54
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
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
         Begin VB.Label Label6 
            Caption         =   "Fecha Comprob.:"
            Height          =   285
            Left            =   12810
            TabIndex        =   48
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Libro Contable:"
            Height          =   255
            Index           =   1
            Left            =   6420
            TabIndex        =   47
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   60
            TabIndex        =   46
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NumAsi 
            Caption         =   "Nro. Asiento:"
            Height          =   255
            Left            =   12810
            TabIndex        =   44
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label Label33 
            Caption         =   "Período:"
            Height          =   255
            Left            =   6420
            TabIndex        =   43
            Top             =   60
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ParEmp()      As moddat_tpo_Genera

Private Sub grd_DetAsi_SelChange()
   If grd_DetAsi.Rows > 2 Then
      grd_DetAsi.RowSel = grd_DetAsi.Row
   End If
End Sub

Private Sub grd_DocAsi_SelChange()
   If grd_DocAsi.Rows > 2 Then
      grd_DocAsi.RowSel = grd_DocAsi.Row
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   Call fs_Inicia
   Call fs_LimpiaCab
   
      
   pnl_NumAsi.Caption = CStr(modctb_lng_NumAsi)
      
   'Leyendo Cabecera
   g_str_Parame = "SELECT * FROM CTB_ASICAB WHERE "
   g_str_Parame = g_str_Parame & "ASICAB_CODEMP = '" & modctb_str_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "ASICAB_CODSUC = '" & modctb_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "ASICAB_PERANO = " & CStr(modctb_int_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "ASICAB_PERMES = " & CStr(modctb_int_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "ASICAB_CODLIB = " & CStr(modctb_int_CodLib) & " AND "
   g_str_Parame = g_str_Parame & "ASICAB_NUMASI = " & CStr(modctb_lng_NumAsi) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   g_rst_Princi.MoveFirst
      
   pnl_FecCtb.Caption = gf_FormatoFecha(CStr(g_rst_Princi!ASICAB_FECCTB))
   pnl_GloCab.Caption = Trim(g_rst_Princi!ASICAB_DESCRI & "")
   
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, modctb_str_CodEmp, "102", CStr(g_rst_Princi!ASICAB_TIPMON)) Then
      pnl_MonCtb.Caption = l_arr_ParEmp(1).Genera_Nombre
   End If
      
   pnl_TipCam.Caption = Format(g_rst_Princi!ASICAB_TIPCAM, "###,###,##0.000000") & " "
   
   pnl_Situac.Caption = moddat_gf_Consulta_ParDes("261", CStr(g_rst_Princi!ASICAB_SITUAC))
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Leyendo Detalle
   g_str_Parame = "SELECT * FROM CTB_ASIDET WHERE "
   g_str_Parame = g_str_Parame & "ASIDET_CODEMP = '" & modctb_str_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "ASIDET_CODSUC = '" & modctb_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "ASIDET_PERANO = " & CStr(modctb_int_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "ASIDET_PERMES = " & CStr(modctb_int_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "ASIDET_CODLIB = " & CStr(modctb_int_CodLib) & " AND "
   g_str_Parame = g_str_Parame & "ASIDET_NUMASI = " & CStr(modctb_lng_NumAsi) & " "
   g_str_Parame = g_str_Parame & "ORDER BY ASIDET_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_DetAsi.Redraw = False
   
   Do While Not g_rst_Princi.EOF
      grd_DetAsi.Rows = grd_DetAsi.Rows + 1
      grd_DetAsi.Row = grd_DetAsi.Rows - 1
   
      grd_DetAsi.Col = 0:  grd_DetAsi.Text = Trim(g_rst_Princi!ASIDET_CODCTA)
      grd_DetAsi.Col = 1:  grd_DetAsi.Text = moddat_gf_Consulta_NomCtaCtb(modctb_str_CodEmp, Trim(g_rst_Princi!ASIDET_CODCTA))
      grd_DetAsi.Col = 2:  grd_DetAsi.Text = Trim(g_rst_Princi!ASIDET_DETGLO & "")
      grd_DetAsi.Col = 3:  grd_DetAsi.Text = moddat_gf_Consulta_ParDes("255", CStr(g_rst_Princi!ASIDET_FLAGDH))
      grd_DetAsi.Col = 4:  grd_DetAsi.Text = "0.00"
      grd_DetAsi.Col = 5:  grd_DetAsi.Text = "0.00"
      grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
      grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
      
      If Mid(Trim(g_rst_Princi!ASIDET_CODCTA), 3, 1) = 1 Then
         If g_rst_Princi!ASIDET_FLAGDH = 1 Then
            'Debe MN
            grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPSOL, "###,###,##0.00")
         
            'Debe ME
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")
         Else
            'Haber MN
            grd_DetAsi.Col = 5:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPSOL, "###,###,##0.00")
            
            'Haber ME
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")
         End If
      Else
         If g_rst_Princi!ASIDET_FLAGDH = 1 Then
            'Debe MN
            grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPDOL * CDbl(pnl_TipCam.Caption), "###,###,##0.00")
            
            'Debe ME
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPDOL, "###,###,##0.00")
         Else
            'Haber MN
            grd_DetAsi.Col = 5:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPDOL * CDbl(pnl_TipCam.Caption), "###,###,##0.00")
            
            'Haber ME
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(g_rst_Princi!ASIDET_IMPDOL, "###,###,##0.00")
         End If
      End If
   
      grd_DetAsi.Col = 8:  grd_DetAsi.Text = CStr(g_rst_Princi!ASIDET_FLAGDH)
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   grd_DetAsi.Redraw = True
   
   Call gs_UbiIniGrid(grd_DetAsi)
   Call fs_TotDebHab

   'Leyendo Documentos
   g_str_Parame = "SELECT * FROM CTB_ASIDOC WHERE "
   g_str_Parame = g_str_Parame & "ASIDOC_CODEMP = '" & modctb_str_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "ASIDOC_CODSUC = '" & modctb_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "ASIDOC_PERANO = " & CStr(modctb_int_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "ASIDOC_PERMES = " & CStr(modctb_int_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "ASIDOC_CODLIB = " & CStr(modctb_int_CodLib) & " AND "
   g_str_Parame = g_str_Parame & "ASIDOC_NUMASI = " & CStr(modctb_lng_NumAsi) & " "
   g_str_Parame = g_str_Parame & "ORDER BY ASIDOC_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      grd_DocAsi.Redraw = False
   
      Do While Not g_rst_Princi.EOF
         grd_DocAsi.Rows = grd_DocAsi.Rows + 1
         grd_DocAsi.Row = grd_DocAsi.Rows - 1
         
         grd_DocAsi.Col = 6:        grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_DOCTIP)
         grd_DocAsi.Col = 0
         If g_rst_Princi!ASIDOC_DOCTIP = 1 Then
            grd_DocAsi.Text = "---"
            grd_DocAsi.CellAlignment = flexAlignCenterCenter
         ElseIf g_rst_Princi!ASIDOC_DOCTIP = 5 Then
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("257", CStr(g_rst_Princi!ASIDOC_DOCTIP)), 3) & " (" & Trim(g_rst_Princi!ASIDOC_MOVSUC) & "-" & Format(g_rst_Princi!ASIDOC_MOVNUM, "0000000") & " - " & gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_MOVFEC)) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 7:     grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_MOVSUC)
            grd_DocAsi.Col = 8:     grd_DocAsi.Text = Format(g_rst_Princi!ASIDOC_MOVNUM, "0000000")
            grd_DocAsi.Col = 9:     grd_DocAsi.Text = gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_MOVFEC))
         Else
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("257", CStr(g_rst_Princi!ASIDOC_DOCTIP)), 3) & " (" & Trim(g_rst_Princi!ASIDOC_DOCSER) & "-" & Trim(g_rst_Princi!ASIDOC_DOCNUM) & " - " & gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_DOCFEC)) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 10:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_DOCSER)
            grd_DocAsi.Col = 11:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_DOCNUM)
            grd_DocAsi.Col = 12:    grd_DocAsi.Text = gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_DOCFEC))
         End If
         
         grd_DocAsi.Col = 13:       grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_IDETIP)
         grd_DocAsi.Col = 1
         If g_rst_Princi!ASIDOC_IDETIP = 1 Then
            grd_DocAsi.Text = "---"
            grd_DocAsi.CellAlignment = flexAlignCenterCenter
         Else
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("259", CStr(g_rst_Princi!ASIDOC_IDETIP)), 3) & " (" & CStr(g_rst_Princi!ASIDOC_IDETDO) & "-" & Trim(g_rst_Princi!ASIDOC_IDENDO) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 14:    grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_IDETDO)
            grd_DocAsi.Col = 15:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_IDENDO)
         End If
         
         grd_DocAsi.Col = 16:       grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_REFTIP)
         grd_DocAsi.Col = 2
         If g_rst_Princi!ASIDOC_REFTIP = 1 Then
            grd_DocAsi.Text = "---"
            grd_DocAsi.CellAlignment = flexAlignCenterCenter
         ElseIf g_rst_Princi!ASIDOC_REFTIP = 2 Then
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("258", CStr(g_rst_Princi!ASIDOC_REFTIP)), 3) & " (" & gf_Formato_NumOpe(Trim(g_rst_Princi!ASIDOC_REFOPE)) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 17:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_REFOPE)
         ElseIf g_rst_Princi!ASIDOC_REFTIP = 3 Then
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("258", CStr(g_rst_Princi!ASIDOC_REFTIP)), 3) & " (" & gf_Formato_NumSol(Trim(g_rst_Princi!ASIDOC_REFSOL)) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 18:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_REFSOL)
         End If
         
         grd_DocAsi.Col = 19:       grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_BCOTIP)
         grd_DocAsi.Col = 3
         If g_rst_Princi!ASIDOC_BCOTIP = 1 Then
            grd_DocAsi.Text = "---"
            grd_DocAsi.CellAlignment = flexAlignCenterCenter
         Else
            grd_DocAsi.Text = Left(moddat_gf_Consulta_ParDes("260", CStr(g_rst_Princi!ASIDOC_BCOTIP)), 3) & " (" & moddat_gf_Consulta_ParDes("516", g_rst_Princi!ASIDOC_BCOCOD) & " - " & Trim(g_rst_Princi!ASIDOC_BCOCTA) & " - " & Trim(g_rst_Princi!ASIDOC_BCONUM) & " - " & gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_BCOFEC)) & ")"
            grd_DocAsi.CellAlignment = flexAlignLeftCenter
            
            grd_DocAsi.Col = 20:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_BCOCOD)
            grd_DocAsi.Col = 21:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_BCOCTA)
            grd_DocAsi.Col = 22:    grd_DocAsi.Text = Trim(g_rst_Princi!ASIDOC_BCONUM)
            grd_DocAsi.Col = 23:    grd_DocAsi.Text = gf_FormatoFecha(CStr(g_rst_Princi!ASIDOC_BCOFEC))
         End If
         
         grd_DocAsi.Col = 4:        grd_DocAsi.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!ASIDOC_ORGMON))
         grd_DocAsi.Col = 24:       grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_ORGMON)
         
         grd_DocAsi.Col = 5:        grd_DocAsi.Text = Format(g_rst_Princi!ASIDOC_ORGMTO, "###,###,##0.00")
         grd_DocAsi.Col = 25:       grd_DocAsi.Text = CStr(g_rst_Princi!ASIDOC_ORGMTO)
         
         g_rst_Princi.MoveNext
      Loop
   
      grd_DocAsi.Redraw = True
      
      Call gs_UbiIniGrid(grd_DocAsi)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
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
   pnl_Empres.Caption = modctb_str_NomEmp
   pnl_Period.Caption = moddat_gf_Consulta_ParDes("033", CStr(modctb_int_PerMes)) & " " & Format(modctb_int_PerAno, "0000")
   pnl_Sucurs.Caption = modctb_str_NomSuc
   pnl_LibCon.Caption = modctb_str_NomLib

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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

