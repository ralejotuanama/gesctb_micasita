VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegCom_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18045
   Icon            =   "GesCtb_frm_185.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   18045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   18060
      _Version        =   65536
      _ExtentX        =   31856
      _ExtentY        =   16325
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   6135
         Left            =   60
         TabIndex        =   18
         Top             =   2340
         Width           =   17940
         _Version        =   65536
         _ExtentX        =   31644
         _ExtentY        =   10821
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
         Begin Threed.SSPanel pnl_Origen 
            Height          =   285
            Left            =   10980
            TabIndex        =   46
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2311
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Origen"
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
         Begin Threed.SSPanel pnl_FecComp 
            Height          =   285
            Left            =   12270
            TabIndex        =   45
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Compen"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5715
            Left            =   30
            TabIndex        =   19
            Top             =   360
            Width           =   17870
            _ExtentX        =   31512
            _ExtentY        =   10081
            _Version        =   393216
            Rows            =   24
            Cols            =   25
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   285
            Left            =   5190
            TabIndex        =   29
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Contable"
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
         Begin Threed.SSPanel pnl_TipComp 
            Height          =   285
            Left            =   6210
            TabIndex        =   30
            Top             =   60
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Comprobante"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   9210
            TabIndex        =   34
            Top             =   60
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1199
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
         Begin Threed.SSPanel pnl_TotComp 
            Height          =   285
            Left            =   9870
            TabIndex        =   35
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1976
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Comp."
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
         Begin Threed.SSPanel pnl_NumRef 
            Height          =   285
            Left            =   14400
            TabIndex        =   36
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "N° Referencia"
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
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   285
            Left            =   1200
            TabIndex        =   37
            Top             =   60
            Width           =   1340
            _Version        =   65536
            _ExtentX        =   2364
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Documento"
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   285
            Left            =   2520
            TabIndex        =   38
            Top             =   60
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         Begin Threed.SSPanel pnl_Select 
            Height          =   285
            Left            =   16710
            TabIndex        =   41
            Top             =   60
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1464
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   " Select."
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   600
               TabIndex        =   42
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Contab 
            Height          =   285
            Left            =   15540
            TabIndex        =   40
            Top             =   60
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1041
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Contab"
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
         Begin Threed.SSPanel pnl_Asiento 
            Height          =   285
            Left            =   16110
            TabIndex        =   43
            Top             =   60
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1094
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Asiento"
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
         Begin Threed.SSPanel pnl_TipCom 
            Height          =   285
            Left            =   13305
            TabIndex        =   44
            Top             =   60
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Compen."
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   7620
            TabIndex        =   47
            Top             =   60
            Width           =   700
            _Version        =   65536
            _ExtentX        =   1235
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Serie"
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
            Left            =   8310
            TabIndex        =   48
            Top             =   60
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Numero"
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   20
         Top             =   60
         Width           =   17940
         _Version        =   65536
         _ExtentX        =   31644
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   21
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Compras"
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
            Left            =   30
            Picture         =   "GesCtb_frm_185.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   22
         Top             =   780
         Width           =   17940
         _Version        =   65536
         _ExtentX        =   31644
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
         Begin VB.CommandButton cmd_Reversa 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_185.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   5460
            Picture         =   "GesCtb_frm_185.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Generar Asientos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   4860
            Picture         =   "GesCtb_frm_185.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Generar Archivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_185.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_185.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_185.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   17340
            Picture         =   "GesCtb_frm_185.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_185.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_185.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_185.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   4260
            Picture         =   "GesCtb_frm_185.frx":22B2
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   23
         Top             =   1470
         Width           =   17940
         _Version        =   65536
         _ExtentX        =   31644
         _ExtentY        =   1455
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
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6780
            TabIndex        =   2
            Top             =   420
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8160
            TabIndex        =   3
            Top             =   420
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   6780
            TabIndex        =   24
            Top             =   90
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Contable:"
            Height          =   195
            Left            =   5310
            TabIndex        =   28
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   480
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   120
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   25
            Top             =   120
            Width           =   1200
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   630
         Left            =   60
         TabIndex        =   31
         Top             =   8520
         Width           =   17010
         _Version        =   65536
         _ExtentX        =   30004
         _ExtentY        =   1111
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
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   17
            Top             =   180
            Width           =   4425
         End
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   32
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegCom_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_RegCom
   regcom_CodCom        As String
   regcom_FecCtb        As String
   regcom_FecEmi        As String
   regcom_FecVto        As String
   regcom_TipCpb_Lrg    As String
   regcom_TipCpb        As String
   regcom_Nserie        As String
   regcom_NroCom        As String
   regcom_TipDoc_Lrg    As String
   regcom_TipDoc        As String
   regcom_NumDoc        As String
   MaePrv_RazSoc        As String
   regcom_Descrp        As String
   
   regcom_Grv           As String
   regcom_Igv           As String
   regcom_Ngrv          As String
   regcom_Total         As String
   
   regcom_Deb_Grv1      As Double
   regcom_Deb_Grv2      As Double
   regcom_Deb_Ngv1      As Double
   regcom_Deb_Ngv2      As Double
   regcom_Deb_Igv1      As Double
   regcom_Deb_Ret1      As Double
   regcom_Deb_Det1      As Double
   regcom_Deb_Ppg1      As Double
   regcom_Hab_Grv1      As Double
   regcom_Hab_Grv2      As Double
   regcom_Hab_Ngv1      As Double
   regcom_Hab_Ngv2      As Double
   regcom_Hab_Igv1      As Double
   regcom_Hab_Ret1      As Double
   regcom_Hab_Det1      As Double
   regcom_Hab_Ppg1      As Double
   
   regcom_CodMon        As String
   regcom_Moneda        As String
   regcom_TipCam        As String
   regcom_Ref_FecEmi    As String
   regcom_Ref_TipCpb    As String
   regcom_Ref_Nserie    As String
   regcom_Ref_NroCom    As String
   regcom_apptrb        As Integer
   regcom_FecDet        As String
   regcom_CodDet        As String
   regcom_Numdet        As String
   regcom_CatCtb        As String
   regcom_FlgCnt        As Integer
   
   regcom_Cnt_Grv1      As String
   regcom_Cnt_Grv2      As String
   regcom_Cnt_Ngv1      As String
   regcom_Cnt_Ngv2      As String
   regcom_Cnt_Igv1      As String
   regcom_Cnt_Ret1      As String
   regcom_Cnt_Det1      As String
   regcom_Cnt_Ppg1      As String
   regcom_TipTab        As Integer
   regcom_CodCaj_Chc    As String
   cajchc_TipPag        As Integer
   
   regcom_CodBnc        As String
   regcom_CtaCrr        As String
   MaePrv_CtaDet        As String
End Type
   
Dim l_arr_GenArc()      As arr_RegCom
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera
Dim r_str_Origen        As String
Dim l_var_ColAnt        As Variant
Dim l_int_Contar        As Long
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer

Private Sub chkSeleccionar_Click()
 Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 13)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 15) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 13)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 15) = "X"
             End If
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmb_Buscar_Click()
    If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
        txt_Buscar.Enabled = False
        Call gs_SetFocus(cmd_Buscar)
    Else
        txt_Buscar.Enabled = True
    End If
    txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      moddat_g_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
      moddat_g_str_RazSoc = cmb_Empres.Text
      
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
      cmb_Sucurs.ListIndex = 0
      Call gs_SetFocus(cmb_Sucurs)
      Screen.MousePointer = 0
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_FlgAct = 0
   moddat_g_int_FlgGrb = 1
   moddat_g_int_InsAct = 0 'Registro de compras
   frm_Ctb_RegCom_04.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   moddat_g_str_CodGen = "" 'CODIGO DE CAJA
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 13
   If UCase(Trim(grd_Listad.Text)) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar el registro por que esta contabilizado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'validacion de caja chica
   grd_Listad.Col = 17 'codigo de caja
   moddat_g_str_CodGen = Trim(grd_Listad.Text & "")
   grd_Listad.Col = 16
   If grd_Listad.Text = 1 And Len(Trim(moddat_g_str_CodGen)) > 0 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar, corresponde al registro a otro modulo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_REGCOM_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'REGCOM_CODCOM
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El proveedor se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarComp
   Call gs_SetFocus(grd_Listad)
End Sub

Private Function fs_ValMod_Aut(p_Codigo As String) As Boolean

   fs_ValMod_Aut = True
   '---------------------------------
   'procesado por Compensasion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL((SELECT DISTINCT COMAUT_CODEST FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "              WHERE A.COMAUT_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODEST IN (1,2,4,5)  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODOPE = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "                AND ROWNUM = 1)  "
   g_str_Parame = g_str_Parame & "           ,0) AS CODEST  "
   g_str_Parame = g_str_Parame & "   FROM DUAL  "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   If g_rst_Princi!CODEST <> 0 Then
      fs_ValMod_Aut = False
      Exit Function
   End If
End Function

Private Sub fs_Buscar()
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_BuscarComp
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
End Sub

Public Sub fs_BuscarComp()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String
Dim r_str_FecPag  As String
Dim r_str_CodPag  As String
Dim r_str_TipPag  As String

   ReDim l_arr_GenArc(0)
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.REGCOM_CODCOM, A.REGCOM_TIPDOC || '-' || A.REGCOM_NUMDOC ID_CLIENTE, TRIM(B.MAEPRV_RAZSOC) MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_FECCTB, TRIM(C.PARDES_DESCRI) TIP_COMPROBANTE, TRIM(D.PARDES_DESCRI) MONEDA,  "
   g_str_Parame = g_str_Parame & "       regcom_CodBnc , regcom_CtaCrr, MAEPRV_CTADET,  "
                                         
   g_str_Parame = g_str_Parame & "       REGCOM_DEB_GRV1, REGCOM_DEB_GRV2, REGCOM_DEB_NGV1, REGCOM_DEB_NGV2, REGCOM_DEB_IGV1, REGCOM_DEB_RET1, "
   g_str_Parame = g_str_Parame & "       REGCOM_DEB_DET1, REGCOM_DEB_PPG1, " 'TOTAL DEBE
   
   g_str_Parame = g_str_Parame & "       REGCOM_HAB_GRV1, REGCOM_HAB_GRV2, REGCOM_HAB_NGV1, REGCOM_HAB_NGV2, REGCOM_HAB_IGV1, REGCOM_HAB_RET1, "
   g_str_Parame = g_str_Parame & "       REGCOM_HAB_DET1, REGCOM_HAB_PPG1, " 'TOT_HABER
   '----------
   g_str_Parame = g_str_Parame & "       A.REGCOM_FECCTB, A.REGCOM_FECEMI, A.REGCOM_FECVTO, A.REGCOM_TIPCPB, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_NSERIE, A.REGCOM_NROCOM, TRIM(B.MAEPRV_RAZSOC) MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_CODMON, A.REGCOM_TIPCAM, A.REGCOM_REF_FECEMI, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_REF_TIPCPB, REGCOM_REF_NSERIE, REGCOM_REF_NROCOM, REGCOM_APPTRB, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_FECDET, REGCOM_CODDET, REGCOM_NUMDET, REGCOM_CATCTB, A.REGCOM_TIPDOC, A.REGCOM_NUMDOC, "
   '----------
   g_str_Parame = g_str_Parame & "       A.REGCOM_CNT_GRV1, A.REGCOM_CNT_GRV2, A.REGCOM_CNT_NGV1, A.REGCOM_CNT_NGV2, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_CNT_IGV1, A.REGCOM_CNT_RET1, A.REGCOM_CNT_DET1, A.REGCOM_CNT_PPG1, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_FLGCNT, A.REGCOM_DESCRP, TRIM(E.PARDES_DESCRI) TIPO_DOCUMENTO, TRIM(SUBSTR(REGCOM_DATCNT,15,20)) NRO_ASIENTO, "
   g_str_Parame = g_str_Parame & "       A.REGCOM_CODDET_CHC, A.REGCOM_CODCAJ_CHC, A.REGCOM_TIPREG, REGCOM_TIPTAB,  "
   
   g_str_Parame = g_str_Parame & "       (CASE A.REGCOM_TIPTAB  "
   g_str_Parame = g_str_Parame & "            WHEN 1 THEN (SELECT AB.CAJCHC_FECCAJ FROM CNTBL_CAJCHC AB WHERE AB.CAJCHC_TIPTAB = 1 AND AB.CAJCHC_CODCAJ = A.REGCOM_CODCAJ_CHC)  "
   g_str_Parame = g_str_Parame & "            WHEN 6 THEN (SELECT TO_NUMBER(AB.CAJCHC_PERANO || LPAD(AB.CAJCHC_PERMES,2,'0')||'01') FROM CNTBL_CAJCHC AB WHERE AB.CAJCHC_TIPTAB = 6 AND AB.CAJCHC_CODCAJ = A.REGCOM_CODCAJ_CHC)  "
   g_str_Parame = g_str_Parame & "            WHEN 2 THEN (SELECT AB.CAJCHC_FECCAJ FROM CNTBL_CAJCHC AB WHERE AB.CAJCHC_TIPTAB = 2 AND AB.CAJCHC_CODCAJ = A.REGCOM_CODCAJ_CHC)  "
   g_str_Parame = g_str_Parame & "            WHEN 3 THEN  J.COMPAG_FECPAG  "
   g_str_Parame = g_str_Parame & "        END) AS FECHA_COMPENSA,  "
   g_str_Parame = g_str_Parame & "       (CASE A.REGCOM_TIPTAB  "
   g_str_Parame = g_str_Parame & "           WHEN 1 THEN 'EFECTIVO'  "
   g_str_Parame = g_str_Parame & "           WHEN 6 THEN 'EFECTIVO'  "
   g_str_Parame = g_str_Parame & "           WHEN 2 THEN (SELECT Y.PARDES_DESCRI  "
   g_str_Parame = g_str_Parame & "                          FROM CNTBL_CAJCHC X  "
   g_str_Parame = g_str_Parame & "                          LEFT JOIN MNT_PARDES Y ON Y.PARDES_CODGRP = 138 AND Y.PARDES_CODITE = X.CAJCHC_TIPPAG  "
   g_str_Parame = g_str_Parame & "                         WHERE X.CAJCHC_TIPTAB = 2 AND X.CAJCHC_CODCAJ = A.REGCOM_CODCAJ_CHC)  "
   g_str_Parame = g_str_Parame & "           WHEN 3 THEN TRIM(K.PARDES_DESCRI)  "
   g_str_Parame = g_str_Parame & "        END) AS TIPO_PAGO,  "
   g_str_Parame = g_str_Parame & "        J.COMPAG_CODCOM,  "
   g_str_Parame = g_str_Parame & "       (CASE A.REGCOM_TIPTAB  "
   g_str_Parame = g_str_Parame & "           WHEN 2 THEN (SELECT NVL(X.CAJCHC_TIPPAG,0)  "
   g_str_Parame = g_str_Parame & "                          FROM CNTBL_CAJCHC X  "
   g_str_Parame = g_str_Parame & "                         WHERE X.CAJCHC_TIPTAB = 2 AND X.CAJCHC_CODCAJ = A.REGCOM_CODCAJ_CHC)  "
   g_str_Parame = g_str_Parame & "           ELSE 0  "
   g_str_Parame = g_str_Parame & "        END) AS CAJCHC_TIPPAG  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_REGCOM A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_MAEPRV B ON A.REGCOM_TIPDOC = B.MAEPRV_TIPDOC AND A.REGCOM_NUMDOC = B.MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 123 AND A.REGCOM_TIPCPB = C.PARDES_CODITE " 'comprobante
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND A.REGCOM_CODMON = D.PARDES_CODITE " 'moneda
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 118 AND A.REGCOM_TIPDOC = E.PARDES_CODITE " 'documento
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.REGCOM_CODCOM) AND H.COMAUT_TIPOPE = 1 AND H.COMAUT_CODEST NOT IN (3)  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                               AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES K ON K.PARDES_CODGRP = 135 AND K.PARDES_CODITE = J.COMPAG_TIPPAG  "
   g_str_Parame = g_str_Parame & " WHERE A.REGCOM_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND A.REGCOM_FECCTB BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   
   If (cmb_Buscar.ListIndex = 1) Then 'numero de documento
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND REGCOM_NUMDOC = '" & Trim(txt_Buscar.Text) & "' "
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'razon social
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND MAEPRV_RAZSOC LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 3) Then 'contabilizado
       r_str_Cadena = ""
       Select Case UCase(Trim(txt_Buscar.Text))
              Case "S", "SI", "I": r_str_Cadena = "1"
              Case "N", "NO", "O": r_str_Cadena = "0"
       End Select
       If (Len(Trim(r_str_Cadena)) > 0) Then
           g_str_Parame = g_str_Parame & "   AND REGCOM_FLGCNT = " & r_str_Cadena
       End If
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.REGCOM_CODCOM ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   ReDim l_arr_GenArc(0)
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!regcom_CodCom)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!ID_CLIENTE & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(Trim(g_rst_Princi!regcom_FecCtb & ""))
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!TIP_COMPROBANTE & "")
      '--------------------
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!regcom_Nserie & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!regcom_NroCom & "")
            
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
                        
      grd_Listad.Col = 8 'TOTAL COMPROBANTE
      grd_Listad.Text = g_rst_Princi!regcom_Deb_Ppg1 + g_rst_Princi!regcom_Hab_Ppg1
      grd_Listad.Text = Format(grd_Listad.Text, "###,###,###,##0.00") & " "
            
      If g_rst_Princi!regcom_TipTab = 1 Then
         grd_Listad.Col = 9: grd_Listad.Text = "CAJA CHICA"
      ElseIf g_rst_Princi!regcom_TipTab = 2 Then
          grd_Listad.Col = 9: grd_Listad.Text = "ENT.RENDIR"
      ElseIf g_rst_Princi!regcom_TipTab = 3 Then
          grd_Listad.Col = 9: grd_Listad.Text = "REG.COMPRAS"
      ElseIf g_rst_Princi!regcom_TipTab = 6 Then
          grd_Listad.Col = 9: grd_Listad.Text = "TARJ.CREDITO"
      End If
      
      If Trim(g_rst_Princi!FECHA_COMPENSA & "") <> "" Then
         grd_Listad.Col = 10
         'If g_rst_Princi!regcom_TipTab = 6 Then
         '   grd_Listad.Text = Trim(g_rst_Princi!FECHA_COMPENSA)
         'Else
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!FECHA_COMPENSA)
         'End If
      End If
      
      grd_Listad.Col = 11
      grd_Listad.Text = Trim(g_rst_Princi!TIPO_PAGO & "")
               
      grd_Listad.Col = 12
      grd_Listad.Text = Format(Trim(g_rst_Princi!regcom_CodCaj_Chc & ""), "0000000000")
      
      If g_rst_Princi!regcom_TipTab <> 3 Then
         'CAJA CHICA / ENTREGAS A RENDIR
         grd_Listad.Col = 12
         grd_Listad.Text = Format(Trim(g_rst_Princi!regcom_CodCaj_Chc & ""), "0000000000")
      Else
         grd_Listad.Col = 12
         grd_Listad.Text = Format(Trim(g_rst_Princi!COMPAG_CODCOM & ""), "0000000000")
      End If
                        
      grd_Listad.Col = 13
      grd_Listad.Text = IIf(g_rst_Princi!regcom_FlgCnt = 1, "SI", "NO")
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!NRO_ASIENTO & "")
      
      'grd_Listad.Col = 13
      'grd_Listad.Text = 'SELECCION
      
      grd_Listad.Col = 16
      grd_Listad.Text = CInt(g_rst_Princi!REGCOM_TIPREG)
      
      grd_Listad.Col = 17
      grd_Listad.Text = Trim(g_rst_Princi!regcom_CodCaj_Chc & "")
      
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(g_rst_Princi!regcom_TipDoc & "")
      
      grd_Listad.Col = 19
      grd_Listad.Text = Trim(g_rst_Princi!regcom_NumDoc & "")
      
      grd_Listad.Col = 20
      grd_Listad.Text = Trim(g_rst_Princi!regcom_TipDoc & "") & Trim(g_rst_Princi!regcom_NumDoc & "")

      grd_Listad.Col = 21
      grd_Listad.Text = g_rst_Princi!regcom_FecCtb
      
      If Trim(g_rst_Princi!FECHA_COMPENSA & "") <> "" Then
         grd_Listad.Col = 22
         grd_Listad.Text = g_rst_Princi!FECHA_COMPENSA
      End If
      
      grd_Listad.Col = 23
      grd_Listad.Text = g_rst_Princi!cajchc_TipPag
      
      grd_Listad.Col = 24
      grd_Listad.Text = g_rst_Princi!regcom_TipTab
      
      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_GenArc(UBound(l_arr_GenArc) + 1)
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CodCom = Trim(g_rst_Princi!regcom_CodCom & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Moneda = Trim(g_rst_Princi!Moneda & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_FecCtb = Trim(g_rst_Princi!regcom_FecCtb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_FecEmi = Trim(g_rst_Princi!regcom_FecEmi & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_FecVto = Trim(g_rst_Princi!regcom_FecVto & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipCpb_Lrg = Trim(g_rst_Princi!TIP_COMPROBANTE & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipCpb = Trim(g_rst_Princi!regcom_TipCpb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Nserie = Trim(g_rst_Princi!regcom_Nserie & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_NroCom = Trim(g_rst_Princi!regcom_NroCom & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipDoc_Lrg = Trim(g_rst_Princi!TIPO_DOCUMENTO & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipDoc = Trim(g_rst_Princi!regcom_TipDoc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_NumDoc = Trim(g_rst_Princi!regcom_NumDoc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).MaePrv_RazSoc = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Grv = g_rst_Princi!regcom_Deb_Grv1 + g_rst_Princi!regcom_Hab_Grv1 + _
                                                      g_rst_Princi!regcom_Deb_Grv2 + g_rst_Princi!regcom_Hab_Grv2 'Grv
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Igv = g_rst_Princi!regcom_Deb_Igv1 + g_rst_Princi!regcom_Hab_Igv1 'Igv
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Ngrv = g_rst_Princi!regcom_Deb_Ngv1 + g_rst_Princi!regcom_Hab_Ngv1 + _
                                                       g_rst_Princi!regcom_Deb_Ngv2 + g_rst_Princi!regcom_Hab_Ngv2 'Ngrv
                                                       
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Total = g_rst_Princi!regcom_Deb_Grv1 + g_rst_Princi!regcom_Hab_Grv1 + _
                                                        g_rst_Princi!regcom_Deb_Grv2 + g_rst_Princi!regcom_Hab_Grv2 + _
                                                        g_rst_Princi!regcom_Deb_Igv1 + g_rst_Princi!regcom_Hab_Igv1 + _
                                                        g_rst_Princi!regcom_Deb_Ngv1 + g_rst_Princi!regcom_Hab_Ngv1 + _
                                                        g_rst_Princi!regcom_Deb_Ngv2 + g_rst_Princi!regcom_Hab_Ngv2  'Total
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Grv1 = g_rst_Princi!regcom_Deb_Grv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Grv2 = g_rst_Princi!regcom_Deb_Grv2
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Ngv1 = g_rst_Princi!regcom_Deb_Ngv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Ngv2 = g_rst_Princi!regcom_Deb_Ngv2
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Igv1 = g_rst_Princi!regcom_Deb_Igv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Ret1 = g_rst_Princi!regcom_Deb_Ret1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Det1 = g_rst_Princi!regcom_Deb_Det1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Deb_Ppg1 = g_rst_Princi!regcom_Deb_Ppg1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Grv1 = g_rst_Princi!regcom_Hab_Grv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Grv2 = g_rst_Princi!regcom_Hab_Grv2
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Ngv1 = g_rst_Princi!regcom_Hab_Ngv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Ngv2 = g_rst_Princi!regcom_Hab_Ngv2
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Igv1 = g_rst_Princi!regcom_Hab_Igv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Ret1 = g_rst_Princi!regcom_Hab_Ret1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Det1 = g_rst_Princi!regcom_Hab_Det1
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Hab_Ppg1 = g_rst_Princi!regcom_Hab_Ppg1
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Grv1 = Trim(g_rst_Princi!regcom_Cnt_Grv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Grv2 = Trim(g_rst_Princi!regcom_Cnt_Grv2 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Ngv1 = Trim(g_rst_Princi!regcom_Cnt_Ngv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Ngv2 = Trim(g_rst_Princi!regcom_Cnt_Ngv2 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Igv1 = Trim(g_rst_Princi!regcom_Cnt_Igv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Ret1 = Trim(g_rst_Princi!regcom_Cnt_Ret1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Det1 = Trim(g_rst_Princi!regcom_Cnt_Det1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Cnt_Ppg1 = Trim(g_rst_Princi!regcom_Cnt_Ppg1 & "")
   
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CodMon = Trim(g_rst_Princi!regcom_CodMon & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipCam = Trim(g_rst_Princi!regcom_TipCam & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Ref_FecEmi = Trim(g_rst_Princi!regcom_Ref_FecEmi & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Ref_TipCpb = Trim(g_rst_Princi!regcom_Ref_TipCpb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Ref_Nserie = Trim(g_rst_Princi!regcom_Ref_Nserie & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Ref_NroCom = Trim(g_rst_Princi!regcom_Ref_NroCom & "")
            
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_apptrb = Trim(g_rst_Princi!regcom_apptrb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_FecDet = Trim(g_rst_Princi!regcom_FecDet & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CodDet = Trim(g_rst_Princi!regcom_CodDet & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Numdet = Trim(g_rst_Princi!regcom_Numdet & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CatCtb = Trim(g_rst_Princi!regcom_CatCtb & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_FlgCnt = Trim(g_rst_Princi!regcom_FlgCnt & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_Descrp = Trim(g_rst_Princi!regcom_Descrp & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_TipTab = g_rst_Princi!regcom_TipTab
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CodCaj_Chc = Trim(g_rst_Princi!regcom_CodCaj_Chc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).cajchc_TipPag = g_rst_Princi!cajchc_TipPag
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CodBnc = Trim(g_rst_Princi!regcom_CodBnc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regcom_CtaCrr = Trim(g_rst_Princi!regcom_CtaCrr & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).MaePrv_CtaDet = Trim(g_rst_Princi!MaePrv_CtaDet & "") 'del banco de la nacion
      '***
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   'Call gs_SetFocus(grd_Listad)
   
   Call grd_Listad_SelChange
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 18
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 19
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgAct = 0
   moddat_g_int_FlgGrb = 0
   moddat_g_int_InsAct = 0 'Registro de compras
   frm_Ctb_RegCom_04.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   moddat_g_str_CodGen = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 18))
   moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 19))
   moddat_g_str_CodGen = CStr(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) 'codigo de caja chica
   
   'Estado de la edicion
   moddat_g_int_FlgAct = 0 'edicion normal
   
   If Len(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) > 0 Then '15
'      'moddat_g_int_FlgAct = 2 'reversa caja chica
'      MsgBox "Solo se puede editar registros que son de origen: registro de compras.", vbExclamation, modgen_g_str_NomPlt
'      Exit Sub
       moddat_g_int_FlgAct = 4 'modificar - contabilizado
   End If
   
   If UCase(grd_Listad.TextMatrix(grd_Listad.Row, 13)) = "SI" Then '11
      moddat_g_int_FlgAct = 3 'modificar - contabilizado
      'MsgBox "Solo se puede editar registros que no hayan sido contabilizados.", vbExclamation, modgen_g_str_NomPlt
      'Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_InsAct = 0 'Registro de compras
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_RegCom_04.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_ExpArc_Click()
Dim r_int_NroCor     As Integer
Dim r_str_NomRes1    As String
Dim r_str_NomRes2    As String
Dim r_str_NumRuc     As String
Dim r_str_DetGlo     As String
Dim r_dbl_TipCam     As Double
Dim r_int_NumRes     As Integer
Dim r_rst_Total      As ADODB.Recordset
Dim r_str_Nombre     As String
Dim R_STR_CONSTT     As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el archivo PLE? ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Verifica que exista ruta
   If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
      MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   Screen.MousePointer = 11

   '----------Creando Archivo - Registro de Compras----------
   r_str_NomRes1 = moddat_g_str_RutLoc & "\LE20511904162" & Format(modctb_str_FecIni, "yyyymm") & "00080100001111.TXT"
   r_int_NumRes = FreeFile
   Open r_str_NomRes1 For Output As r_int_NumRes
   r_int_NroCor = 1
   R_STR_CONSTT = "|"
   Dim r_str_Fecvto      As String
   Dim r_str_Fecemi      As String
   Dim r_str_Ref_fecemi  As String
   Dim r_str_Fecdet      As String
   Dim r_str_TipDoc      As String
   Dim r_str_tipcpb      As String
   Dim r_str_Ref_tipcpb  As String
   Dim r_str_NumDet      As String
   Dim r_str_Col_41      As String
   Dim r_str_Col_14, r_str_Col_15, r_str_Col_16 As String
   Dim r_str_Col_17, r_str_Col_18, r_str_Col_19 As String
   Dim r_str_Col_20, r_str_Col_21, r_str_Col_22 As String
   Dim r_str_Col_23      As String
   Dim r_str_CadAux      As String
   
   For l_int_Contar = 1 To UBound(l_arr_GenArc)
       r_str_Fecemi = ""
       r_str_Fecvto = ""
       r_str_Ref_fecemi = ""
       r_str_Fecdet = ""
       r_str_TipDoc = ""
       r_str_tipcpb = ""
       r_str_Ref_tipcpb = ""
       r_str_NumDet = "0"
       r_str_Col_41 = "6"
       r_dbl_TipCam = 0
       r_str_CadAux = 0
       
       'GENERA ARCHIVO MENOS (00)OTROS, (02)RECIBO HONORARIOS
       If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) <> 9999 Then 'OTROS
          If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) <> 2 Then 'RECIBO HONORARIOS
             If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) <> 91 Then '91 - INVOCE  -- 99 ACEEDORES
                If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) <> 88 Then '88 - DEVOLUCIONES
                   r_str_Fecemi = "01/01/0001"
                   If Len(Trim(l_arr_GenArc(l_int_Contar).regcom_FecEmi)) > 0 Then
                      r_str_Fecemi = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecEmi)
                   End If
                   r_str_Fecvto = "01/01/0001"
                   If Len(Trim(l_arr_GenArc(l_int_Contar).regcom_FecVto)) > 0 Then
                      r_str_Fecvto = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecVto)
                   End If
                   r_str_Ref_fecemi = "01/01/0001"
                   If Len(Trim(l_arr_GenArc(l_int_Contar).regcom_Ref_FecEmi)) > 0 Then
                      r_str_Ref_fecemi = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_Ref_FecEmi)
                   End If
                   
                   r_str_Fecdet = "01/01/0001"
                   r_str_NumDet = "0"
                   If l_arr_GenArc(l_int_Contar).regcom_apptrb = 2 Then 'Detraccion
                      If Len(Trim(l_arr_GenArc(l_int_Contar).regcom_FecDet)) > 0 Then
                         r_str_Fecdet = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecDet)
                      End If
                      r_str_NumDet = Trim(l_arr_GenArc(l_int_Contar).regcom_Numdet)
                   End If
                   'tipo documento
                   r_str_TipDoc = Trim(l_arr_GenArc(l_int_Contar).regcom_TipDoc)
                   If (Trim(l_arr_GenArc(l_int_Contar).regcom_TipDoc) = "9999") Then
                       r_str_TipDoc = "0"
                   End If
                   'tipo comprobante
                   r_str_tipcpb = l_arr_GenArc(l_int_Contar).regcom_TipCpb
                   If (Trim(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = "9999") Then
                       r_str_tipcpb = "00"
                   End If
                   'tipo comprobante referencia
                   r_str_Ref_tipcpb = l_arr_GenArc(l_int_Contar).regcom_Ref_TipCpb
                   If (Trim(l_arr_GenArc(l_int_Contar).regcom_Ref_TipCpb) = "9999") Then
                       r_str_Ref_tipcpb = "00"
                   End If
                   r_str_Col_41 = "6"
                   'ultima columna
                   If CDbl(Format(r_str_Fecemi, "yyyymm")) >= CDbl(Left(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 6)) And _
                      CDbl(Format(r_str_Fecemi, "yyyymm")) <= CDbl(Left(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 6)) Then
                      r_str_Col_41 = 1
                   End If
                   'columna 14 a la 23
                   r_str_CadAux = ""
                   If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = CInt("07") Then
                      r_str_CadAux = "-"
                   End If
                   r_str_Col_14 = r_str_CadAux & "0"
                   r_str_Col_15 = r_str_CadAux & "0"
                   r_str_Col_18 = r_str_CadAux & "0"
                   r_str_Col_19 = r_str_CadAux & "0"
                   r_str_Col_21 = r_str_CadAux & "0"
                   r_str_Col_22 = r_str_CadAux & "0"
                   If CInt(l_arr_GenArc(l_int_Contar).regcom_CodMon) = 2 Then
                      r_dbl_TipCam = CDbl(l_arr_GenArc(l_int_Contar).regcom_TipCam)
                      r_str_Col_20 = r_str_CadAux & CDbl(Format(CDbl(l_arr_GenArc(l_int_Contar).regcom_Ngrv * r_dbl_TipCam), "########0.00"))
                      r_str_Col_16 = r_str_CadAux & CDbl(Format(CDbl(l_arr_GenArc(l_int_Contar).regcom_Grv * r_dbl_TipCam), "########0.00"))
                      r_str_Col_17 = r_str_CadAux & CDbl(Format(CDbl(l_arr_GenArc(l_int_Contar).regcom_Igv * r_dbl_TipCam), "########0.00"))
                      r_str_Col_23 = r_str_CadAux & CDbl(Format(CDbl(l_arr_GenArc(l_int_Contar).regcom_Total * r_dbl_TipCam), "########0.0"))
                   Else
                      r_str_Col_20 = r_str_CadAux & l_arr_GenArc(l_int_Contar).regcom_Ngrv
                      r_str_Col_16 = r_str_CadAux & l_arr_GenArc(l_int_Contar).regcom_Grv
                      r_str_Col_17 = r_str_CadAux & l_arr_GenArc(l_int_Contar).regcom_Igv
                      r_str_Col_23 = r_str_CadAux & l_arr_GenArc(l_int_Contar).regcom_Total
                   End If
                   
                   Print #1, Mid(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1, 6) & "00"; R_STR_CONSTT; _
                             Mid(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1, 6) & Format(CStr(r_int_NroCor), "0000"); R_STR_CONSTT; _
                             "M" & Format(r_int_NroCor, "000000000"); R_STR_CONSTT; _
                             r_str_Fecemi; R_STR_CONSTT; r_str_Fecvto; R_STR_CONSTT; _
                             Format(r_str_tipcpb, "00"); R_STR_CONSTT; _
                             l_arr_GenArc(l_int_Contar).regcom_Nserie; R_STR_CONSTT; "0"; R_STR_CONSTT; _
                             l_arr_GenArc(l_int_Contar).regcom_NroCom; R_STR_CONSTT; ""; R_STR_CONSTT; _
                             r_str_TipDoc; R_STR_CONSTT; l_arr_GenArc(l_int_Contar).regcom_NumDoc; R_STR_CONSTT; _
                             l_arr_GenArc(l_int_Contar).MaePrv_RazSoc; R_STR_CONSTT; _
                             r_str_Col_14; R_STR_CONSTT; r_str_Col_15; R_STR_CONSTT; _
                             r_str_Col_16; R_STR_CONSTT; r_str_Col_17; R_STR_CONSTT; _
                             r_str_Col_18; R_STR_CONSTT; r_str_Col_19; R_STR_CONSTT; _
                             r_str_Col_20; R_STR_CONSTT; r_str_Col_21; R_STR_CONSTT; _
                             r_str_Col_22; R_STR_CONSTT; r_str_Col_23; R_STR_CONSTT; _
                             IIf(CInt(l_arr_GenArc(l_int_Contar).regcom_CodMon) = 1, "PEN", "USD"); R_STR_CONSTT; _
                             Format(l_arr_GenArc(l_int_Contar).regcom_TipCam, "#,##0.000"); R_STR_CONSTT; _
                             r_str_Ref_fecemi; R_STR_CONSTT; Format(r_str_Ref_tipcpb, "00"); R_STR_CONSTT; _
                             IIf(Trim(l_arr_GenArc(l_int_Contar).regcom_Ref_Nserie) = "", "-", l_arr_GenArc(l_int_Contar).regcom_Ref_Nserie); R_STR_CONSTT; _
                             "244"; R_STR_CONSTT; _
                             IIf(Trim(l_arr_GenArc(l_int_Contar).regcom_Ref_NroCom) = "", "-", l_arr_GenArc(l_int_Contar).regcom_Ref_NroCom); R_STR_CONSTT; _
                             r_str_Fecdet; R_STR_CONSTT; r_str_NumDet; R_STR_CONSTT; ""; R_STR_CONSTT; _
                             l_arr_GenArc(l_int_Contar).regcom_CatCtb; R_STR_CONSTT; ""; R_STR_CONSTT; ""; R_STR_CONSTT; _
                             "1"; R_STR_CONSTT; "1"; R_STR_CONSTT; IIf(CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = 3, "0", "1"); R_STR_CONSTT; _
                             "1"; R_STR_CONSTT; r_str_Col_41; R_STR_CONSTT
                         
                   r_int_NroCor = r_int_NroCor + 1
                End If
             End If
          End If
       End If
   Next
   Close #1
   
   '----------Creando Archivo - Registro de no domiciliados----------
   r_str_NomRes2 = moddat_g_str_RutLoc & "\LE20511904162" & Format(modctb_str_FecIni, "yyyymm") & "00080200001011.TXT"
   r_int_NumRes = FreeFile
   Open r_str_NomRes2 For Output As r_int_NumRes
   r_int_NroCor = 1
   R_STR_CONSTT = "|"
   For l_int_Contar = 1 To UBound(l_arr_GenArc)
       r_str_Fecemi = ""
       r_str_tipcpb = ""

       'GENERA ARCHIVO (91) INVOCE
       If CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = 91 Then
          r_str_Fecemi = "01/01/0001"
          If Len(Trim(l_arr_GenArc(l_int_Contar).regcom_FecEmi)) > 0 Then
             r_str_Fecemi = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecEmi)
          End If
          'tipo comprobante
          r_str_tipcpb = l_arr_GenArc(l_int_Contar).regcom_TipCpb
          If (Trim(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = "9999") Then
              r_str_tipcpb = "00"
          End If

          Print #1, Mid(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1, 6) & "00"; R_STR_CONSTT; _
                    Mid(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1, 6) & Format(CStr(r_int_NroCor), "0000"); R_STR_CONSTT; _
                    "M" & Format(r_int_NroCor, "000000000"); R_STR_CONSTT; _
                    r_str_Fecemi; R_STR_CONSTT; Format(r_str_tipcpb, "00"); R_STR_CONSTT; _
                    l_arr_GenArc(l_int_Contar).regcom_Nserie; R_STR_CONSTT; _
                    l_arr_GenArc(l_int_Contar).regcom_NroCom; R_STR_CONSTT; _
                    l_arr_GenArc(l_int_Contar).regcom_Ngrv; R_STR_CONSTT; "0"; R_STR_CONSTT; _
                    l_arr_GenArc(l_int_Contar).regcom_Ngrv; R_STR_CONSTT; _
                    "00"; R_STR_CONSTT; "0"; R_STR_CONSTT; Mid(l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1, 4); R_STR_CONSTT; _
                    "000000000"; R_STR_CONSTT; ""; R_STR_CONSTT; _
                    IIf(CInt(l_arr_GenArc(l_int_Contar).regcom_CodMon) = 1, "PEN", "USD"); R_STR_CONSTT; _
                    Format(l_arr_GenArc(l_int_Contar).regcom_TipCam, "#,##0.000"); R_STR_CONSTT; _
                    "9249"; R_STR_CONSTT; l_arr_GenArc(l_int_Contar).MaePrv_RazSoc; R_STR_CONSTT; ""; R_STR_CONSTT; _
                    "-"; R_STR_CONSTT; "20511904162"; R_STR_CONSTT; "MI CASITA SA"; R_STR_CONSTT; _
                    "9589"; R_STR_CONSTT; "00"; R_STR_CONSTT; "0"; R_STR_CONSTT; ""; R_STR_CONSTT; _
                    "0"; R_STR_CONSTT; "0"; R_STR_CONSTT; "0"; R_STR_CONSTT; "0"; R_STR_CONSTT; _
                    ""; R_STR_CONSTT; "18"; R_STR_CONSTT; "1"; R_STR_CONSTT; ""; R_STR_CONSTT; "0"; R_STR_CONSTT;
          r_int_NroCor = r_int_NroCor + 1
       End If
   Next
   Close #1

   Screen.MousePointer = 0 'vbCrLf r_str_NomRes
   MsgBox "El archivo ha sido creado. " & vbCrLf & _
          "Registro de compras:         " & Trim(r_str_NomRes1) & vbCrLf & _
          "Registro de no domiciliados: " & Trim(r_str_NomRes2), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Generar_Click()
Dim r_int_Contad        As Integer
Dim r_bol_Estado        As Boolean

   r_bol_Estado = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contad, 13) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contad, 15) = "X" Then
             r_bol_Estado = True
             Exit For
          End If
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No se han seleccionados registros para generar asientos automáticos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'confirma
   If MsgBox("¿Está seguro de generar los asientos contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiento
   Call cmd_Buscar_Click
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Empres.Enabled = True
   cmb_Sucurs.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmd_Reversa_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   moddat_g_str_CodGen = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 18))
   moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 19))
   moddat_g_str_CodGen = CStr(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) 'codigo de caja chica
             
   If (UCase(grd_Listad.TextMatrix(grd_Listad.Row, 13)) = "NO" And Len(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) = 0) Then
      MsgBox "Solo se da reversa a registros que son contabilizados o que vengan de caja chica, entregas a rendir.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   '---------------------------------
   'procesado por Compensasion
    If fs_ValMod_Aut(moddat_g_str_Codigo) = False Then
       MsgBox "El registro se encuentra en el módulo de compensación.", vbExclamation, modgen_g_str_NomPlt
       Screen.MousePointer = 0
       Exit Sub
    End If
      
   'Estado de la edicion
   moddat_g_int_FlgAct = 0 'edicion normal
   If (Len(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) > 0) Then
       moddat_g_int_FlgAct = 2 'reversa caja chica
       If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 24) & "") = "2" Then
          'SOLO ENTREGAS A RENDIR
          If Mid(Format(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17)), "0000000000"), 1, 2) = "05" Then
             If fs_ValMod_Aut(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))) = False Then
                MsgBox "El registro de origen de entrega a rendir se encuentra en el modulo de compensación.", vbExclamation, modgen_g_str_NomPlt
                Screen.MousePointer = 0
                Exit Sub
             End If
          End If
          
       End If
   End If
   If (UCase(grd_Listad.TextMatrix(grd_Listad.Row, 13)) = "SI") Then
       moddat_g_int_FlgAct = 1 'reversa contabilidad
   End If
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_InsAct = 0 'Registro de compras
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_RegCom_04.Show 1
   
   'Call fs_BuscarComp
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_GeneraAsiento()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_Import        As Double
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_CadAux        As String
Dim r_int_Fila          As Integer
Dim r_bol_Estado        As Boolean
Dim r_dbl_ImpRnd        As Double
Dim r_int_NumIte        As Integer
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
Dim r_str_CtaHab        As String
             
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "K"
   r_int_NumLib = 15
   r_str_AsiGen = ""
   r_dbl_ImpRnd = 0
   
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Fila, 15) = "X" Then
       
   For l_int_Contar = 1 To UBound(l_arr_GenArc)
       If Trim(grd_Listad.TextMatrix(r_int_Fila, 0)) = Trim(l_arr_GenArc(l_int_Contar).regcom_CodCom) Then
          If l_arr_GenArc(l_int_Contar).regcom_FlgCnt = 0 Then
             
             'Inicializa variables
             r_dbl_ImpRnd = 0
             r_int_NumAsi = 0
             r_int_NumIte = 0
             r_str_FecPrPgoC = l_arr_GenArc(l_int_Contar).regcom_FecCtb
             r_str_FecPrPgoL = gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecCtb)
             
             'r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, 2, l_arr_GenArc(l_int_Contar).regcom_FecCtb, 1)
             r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(3, 2, l_arr_GenArc(l_int_Contar).regcom_FecEmi, 1)
             
             r_str_Glosa = ""
             If l_arr_GenArc(l_int_Contar).regcom_TipTab = "2" Then 'Entregas a Rendir
                r_str_Glosa = l_arr_GenArc(l_int_Contar).regcom_NumDoc & "/" & _
                              IIf(CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = CInt("9999"), "00", _
                              Format(CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb), "00")) & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_Nserie) & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_NroCom) & "/" & _
                              Format(Trim(l_arr_GenArc(l_int_Contar).regcom_CodCaj_Chc), "0000000000") & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_Descrp)
             Else
                r_str_Glosa = l_arr_GenArc(l_int_Contar).regcom_NumDoc & "/" & _
                              IIf(CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = CInt("9999"), "00", _
                              Format(CInt(l_arr_GenArc(l_int_Contar).regcom_TipCpb), "00")) & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_Nserie) & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_NroCom) & "/" & _
                              Trim(l_arr_GenArc(l_int_Contar).regcom_Descrp)
             End If
             r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
             
             'l_int_PerMes = modctb_int_PerMes 'Month(gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecCtb))
             'l_int_PerAno = modctb_int_PerAno 'Year(gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecCtb))
             l_int_PerMes = Month(gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecCtb))
             l_int_PerAno = Year(gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecCtb))
              
             'Obteniendo Nro. de Asiento
             r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
             r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
                
             'Insertar en CABECERA
             Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                  r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
             '--------------GRAVADO_1-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Grv1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Grv1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Grv1 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Grv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------GRAVADO_2-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Grv2 + l_arr_GenArc(l_int_Contar).regcom_Hab_Grv2
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Grv2 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Grv2, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------No_GRAVADO_1-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Ngv1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv1 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
         
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Ngv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------No_GRAVADO_2-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv2 + l_arr_GenArc(l_int_Contar).regcom_Hab_Ngv2
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv2 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Ngv2, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------IGV-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Igv1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Igv1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Igv1 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Igv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------RETENCION-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Ret1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Ret1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Ret1 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                 
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Ret1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '--------------DETRACCION-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Det1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Det1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Det1 > 0, "D", "H")
             r_dbl_ImpRnd = 0
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                 '*************************
                 'PREGUNTAR A JULIO CABRERA
                 If l_arr_GenArc(l_int_Contar).regcom_CodMon = 2 Then
                    r_dbl_ImpRnd = Round(r_dbl_MtoSol, 0) - r_dbl_MtoSol 'Format(Round(importe, 0), "###,###,##0.00") & " "
                    r_dbl_ImpRnd = Math.Abs(Round(r_dbl_ImpRnd, 2))
                    r_dbl_MtoSol = Round(r_dbl_MtoSol, 0)
                 End If
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Det1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             '/**/,
             '--------------POR PAGAR-------------------------
             r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Ppg1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Ppg1
             r_str_DebHab = IIf(l_arr_GenArc(l_int_Contar).regcom_Deb_Ppg1 > 0, "D", "H")
             If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                 '*************************
                 'PREGUNTAR A JULIO CABRERA - SE RESTA LA DIFERENCIA
                 If l_arr_GenArc(l_int_Contar).regcom_CodMon = 2 Then
                    r_dbl_MtoSol = r_dbl_MtoSol - r_dbl_ImpRnd
                 End If
                 
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(l_int_Contar).regcom_Cnt_Ppg1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
             End If
             'Actualiza flag de contabilizacion
             r_str_CadAux = ""
             r_str_CadAux = r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & "UPDATE CNTBL_REGCOM "
             g_str_Parame = g_str_Parame & "   SET REGCOM_FLGCNT = 1, "
             g_str_Parame = g_str_Parame & "       REGCOM_FECCNT = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ", "
             g_str_Parame = g_str_Parame & "       REGCOM_DATCNT = '" & r_str_CadAux & "' "
             g_str_Parame = g_str_Parame & " WHERE REGCOM_CODCOM  = '" & l_arr_GenArc(l_int_Contar).regcom_CodCom & "' "
              
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                Exit Sub
             End If
             
             
             r_bol_Estado = False
             If l_arr_GenArc(l_int_Contar).regcom_TipTab = 3 Then
                'REG. COMPRAS
                r_bol_Estado = True
             End If
             
              If r_bol_Estado = True Then
                'Enviar a la tabla de autorizaciones - POR PAGAR
                r_str_CtaHab = ""
                r_str_CtaHab = l_arr_GenArc(l_int_Contar).regcom_Cnt_Ppg1
                r_dbl_Import = 0
                r_dbl_Import = CDbl(l_arr_GenArc(l_int_Contar).regcom_Deb_Ppg1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Ppg1)
                If r_dbl_Import > 0 Then
                   If CLng(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = 7 Or CLng(l_arr_GenArc(l_int_Contar).regcom_TipCpb) = 88 Then
                      r_dbl_Import = -r_dbl_Import
                   End If
                   g_str_Parame = ""
                   g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
                   g_str_Parame = g_str_Parame & " " & CLng(l_arr_GenArc(l_int_Contar).regcom_CodCom) & ", " 'COMAUT_CODOPE
                   g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_FecCtb & ", " 'COMAUT_FECOPE
                   g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_TipDoc & ", "      'COMAUT_TIPDOC
                   g_str_Parame = g_str_Parame & " '" & l_arr_GenArc(l_int_Contar).regcom_NumDoc & "', "    'COMAUT_NUMDOC
                   g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_CodMon & ", "      'COMAUT_CODMON
                   g_str_Parame = g_str_Parame & " " & r_dbl_Import & ", " 'COMAUT_IMPPAG
                   If l_arr_GenArc(l_int_Contar).regcom_CodBnc = "" Then
                      g_str_Parame = g_str_Parame & " NULL, "          'COMAUT_CODBNC
                   Else
                      g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_CodBnc & ", " 'COMAUT_CODBNC
                   End If
                   g_str_Parame = g_str_Parame & " '" & l_arr_GenArc(l_int_Contar).regcom_CtaCrr & "', "  'COMAUT_CTACRR
                   g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB ??????
                   g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
                   g_str_Parame = g_str_Parame & " '" & "POR PAGAR',  " 'COMAUT_DESCRIPCION
                   g_str_Parame = g_str_Parame & " 1,  " 'COMAUT_TIPOPE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'SEGUSUCRE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'SEGPLTCRE
                   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'SEGSUCCRE
                  
                   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                      Exit Sub
                   End If
                End If
             End If
             
             r_bol_Estado = False
             If l_arr_GenArc(l_int_Contar).regcom_TipTab = 3 Then
                'REG. COMPRAS
                r_bol_Estado = True
             ElseIf l_arr_GenArc(l_int_Contar).regcom_TipTab = 2 Then
                'ENT. RENDIR (SOLO REG. ANTIGUOS 0 no pasa)
                 'Select Case CInt(grd_Listad.TextMatrix(r_int_fila, 21)) 'CAJCHC_TIPPAG
                 Select Case CInt(l_arr_GenArc(l_int_Contar).cajchc_TipPag)  'CAJCHC_TIPPAG
                        Case 0: r_bol_Estado = False
                        Case Else
                             r_bol_Estado = True
                 End Select
             End If
             
             If r_bol_Estado = True Then
                'Enviar a la tabla de autorizaciones  - DETRACCION
                r_str_CtaHab = ""
                r_str_CtaHab = l_arr_GenArc(l_int_Contar).regcom_Cnt_Det1
                r_dbl_Import = 0
                r_dbl_Import = l_arr_GenArc(l_int_Contar).regcom_Deb_Det1 + l_arr_GenArc(l_int_Contar).regcom_Hab_Det1
                If r_dbl_Import > 0 Then
                   Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(l_int_Contar).regcom_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                   '*************************
                   'PREGUNTAR A JULIO CABRERA
                   If l_arr_GenArc(l_int_Contar).regcom_CodMon = 2 Then
                      r_dbl_MtoSol = Round(r_dbl_MtoSol, 0)
                   End If
                   
                   If l_arr_GenArc(l_int_Contar).regcom_TipCpb = 7 Or l_arr_GenArc(l_int_Contar).regcom_TipCpb = 88 Then
                      r_dbl_MtoSol = -r_dbl_MtoSol
                   End If
                   g_str_Parame = ""
                   g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
                   g_str_Parame = g_str_Parame & " " & CLng(l_arr_GenArc(l_int_Contar).regcom_CodCom) & ", " 'COMAUT_CODOPE
                   g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_FecCtb & ", " 'COMAUT_FECOPE
                   g_str_Parame = g_str_Parame & " " & l_arr_GenArc(l_int_Contar).regcom_TipDoc & ", "   'COMAUT_TIPDOC
                   g_str_Parame = g_str_Parame & " '" & l_arr_GenArc(l_int_Contar).regcom_NumDoc & "', " 'COMAUT_NUMDOC
                   g_str_Parame = g_str_Parame & " " & 1 & ", "      'COMAUT_CODMON - SIEMPRE SOLES
                   g_str_Parame = g_str_Parame & " " & CDbl(r_dbl_MtoSol) & ", " 'COMAUT_IMPPAG
                   If l_arr_GenArc(l_int_Contar).MaePrv_CtaDet = "" Then
                      g_str_Parame = g_str_Parame & " NULL, "    'COMAUT_CODBNC
                   Else
                      g_str_Parame = g_str_Parame & " 18, "    'COMAUT_CODBNC - Banco nacion
                   End If
                   g_str_Parame = g_str_Parame & " '" & l_arr_GenArc(l_int_Contar).MaePrv_CtaDet & "', "  'COMAUT_CTACRR - cuenta detractora
                   g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB ??????
                   g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
                   g_str_Parame = g_str_Parame & " '" & "DETRACCION',  " 'COMAUT_DESCRIPCION
                   g_str_Parame = g_str_Parame & " 2,  " 'COMAUT_TIPOPE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'SEGUSUCRE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'SEGPLTCRE
                   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'SEGSUCCRE
                    
                   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                      Exit Sub
                   End If
                End If
             End If
             
             Exit For
          End If
       End If
   Next
   
   End If
   Next
   MsgBox "Se culminó el proceso de generación de asientos contables." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_Calc_Imp(p_Import As Double, p_Moneda As String, p_TipSbs As Double, ByRef p_ImpDol As Double, ByRef p_ImpSol As Double)
   If (CInt(p_Moneda) = 1) Then
      'SOLES
      p_ImpSol = Format(p_Import, "###,###,##0.00") 'Importe soles
      p_ImpDol = Format(CDbl(p_ImpSol / p_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
   Else
      'DOLARES
      p_ImpDol = Format(p_Import, "###,###,##0.00") 'Importe dolares
      p_ImpSol = Format(CDbl(p_ImpDol * p_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
      
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "NRO DOCUMENTO"
   cmb_Buscar.AddItem "RAZÓN SOCIAL"
   cmb_Buscar.AddItem "CONTABILIZADO"
   
   grd_Listad.ColWidth(0) = 1140  'codigo
   grd_Listad.ColWidth(1) = 1320  'Num documento
   grd_Listad.ColWidth(2) = 2680  'Razon Social
   grd_Listad.ColWidth(3) = 1040  'Fecha Contable
   grd_Listad.ColWidth(4) = 1400  'Tipo Comprobante
   grd_Listad.ColWidth(5) = 680  'serie
   grd_Listad.ColWidth(6) = 915  'numero
   
   grd_Listad.ColWidth(7) = 650   'Moneda
   grd_Listad.ColWidth(8) = 1100  'Total comprobante
   grd_Listad.ColWidth(9) = 1310  'origen
   grd_Listad.ColWidth(10) = 1040  'fecha comprobante
   grd_Listad.ColWidth(11) = 1090  'tipo comprobante
   grd_Listad.ColWidth(12) = 1160 'referencia
   grd_Listad.ColWidth(13) = 560  'contabilizado
   grd_Listad.ColWidth(14) = 600 'Asiento
   grd_Listad.ColWidth(15) = 800 'seleccion
   grd_Listad.ColWidth(16) = 0    'REGCOM_TIPREG
   grd_Listad.ColWidth(17) = 0    'REGCOM_CODCAJ_CHC
   grd_Listad.ColWidth(18) = 0    'regcom_TipDoc
   grd_Listad.ColWidth(19) = 0    'regcom_NumDoc
   grd_Listad.ColWidth(20) = 0    'regcom_NumDoc - ORDEN
   grd_Listad.ColWidth(21) = 0    'regcom_FECCTB - ORDEN
   grd_Listad.ColWidth(22) = 0    'FECHA COMPENSACION - ORDEN
   grd_Listad.ColWidth(23) = 0    'CAJCHC_TIPPAG
   grd_Listad.ColWidth(24) = 0    'REGCOM_TIPTAB
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter  'codigo
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter  'Num documento
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter    'Razon Social
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter  'Fecha Contable
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter    'Tipo Comprobante
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter  'serie
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter  'numero
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter    'Moneda
   grd_Listad.ColAlignment(8) = flexAlignRightCenter   'Total comprobante
   grd_Listad.ColAlignment(9) = flexAlignLeftCenter    'origen
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter  'fecha comprobante
   grd_Listad.ColAlignment(11) = flexAlignLeftCenter    'tipo comprobante
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter 'referencia
   grd_Listad.ColAlignment(13) = flexAlignCenterCenter 'contabilizado
   grd_Listad.ColAlignment(14) = flexAlignCenterCenter 'Asiento
   grd_Listad.ColAlignment(15) = flexAlignCenterCenter 'seleccion
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Buscar.ListIndex = 0
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE REGISTRO DE COMPRAS"
      .Range(.Cells(2, 2), .Cells(2, 18)).Merge
      .Range(.Cells(2, 2), .Cells(2, 18)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 18)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "FECHA DE EMISIÓN"
      .Cells(3, 4) = "TIPO COMPROBANTE"
      .Cells(3, 5) = "SERIE"
      .Cells(3, 6) = "NÚMERO"
      .Cells(3, 7) = "TIPO DOCUMENTO"
      .Cells(3, 8) = "DOCUMENTO"
      .Cells(3, 9) = "PROVEEDOR"
      .Cells(3, 10) = "MONEDA"
      .Cells(3, 11) = "GRAVADO 1"
      .Cells(3, 12) = "GRAVADO 2"
      .Cells(3, 13) = "NO GRAVADO 1"
      .Cells(3, 14) = "NO GRAVADO 2"
      .Cells(3, 15) = "IGV"
      .Cells(3, 16) = "RETENCIÓN"
      .Cells(3, 17) = "DETRACCIÓN"
      .Cells(3, 18) = "TOTAL"
         
      .Range(.Cells(3, 2), .Cells(3, 18)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 18)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 18 'fecha de emision
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 30 'tipo de comprobante
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 8 'serie
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 10 'numero
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 40 'tipo de documento
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 15 'documento
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 50 'proveedor
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      
      .Columns("J").ColumnWidth = 21 'moneda
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      
      .Columns("K").ColumnWidth = 13 'gravado 1
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      
      .Columns("L").ColumnWidth = 13 'gravado 2
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 15 'no gravado 1
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 15 'no gravado 2
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13 'igv
      .Columns("O").NumberFormat = "###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13 'retencion
      .Columns("P").NumberFormat = "###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13 'detraccion
      .Columns("Q").NumberFormat = "###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      
      .Columns("R").ColumnWidth = 13 'total
      .Columns("R").NumberFormat = "###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 18)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 18)).Font.Size = 11
      
      r_int_NumFil = 2
      For l_int_Contar = 1 To UBound(l_arr_GenArc)
          .Cells(r_int_NumFil + 2, 2) = "'" & l_arr_GenArc(l_int_Contar).regcom_CodCom     'codigo
          .Cells(r_int_NumFil + 2, 3) = "'" & gf_FormatoFecha(l_arr_GenArc(l_int_Contar).regcom_FecEmi) 'fecha de emision
          .Cells(r_int_NumFil + 2, 4) = "'" & l_arr_GenArc(l_int_Contar).regcom_TipCpb_Lrg 'tipo de comprobante
          .Cells(r_int_NumFil + 2, 5) = "'" & l_arr_GenArc(l_int_Contar).regcom_Nserie     'serie
          .Cells(r_int_NumFil + 2, 6) = "'" & l_arr_GenArc(l_int_Contar).regcom_NroCom     'numero
          .Cells(r_int_NumFil + 2, 7) = "'" & l_arr_GenArc(l_int_Contar).regcom_TipDoc_Lrg 'tipo de documento
          .Cells(r_int_NumFil + 2, 8) = "'" & l_arr_GenArc(l_int_Contar).regcom_NumDoc     'documento
          .Cells(r_int_NumFil + 2, 9) = "'" & l_arr_GenArc(l_int_Contar).MaePrv_RazSoc     'proveedor
          
          .Cells(r_int_NumFil + 2, 10) = "'" & l_arr_GenArc(l_int_Contar).regcom_Moneda    'moneda
          
          .Cells(r_int_NumFil + 2, 11) = l_arr_GenArc(l_int_Contar).regcom_Deb_Grv1 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Grv1 'gravado 1
                                         
          .Cells(r_int_NumFil + 2, 12) = l_arr_GenArc(l_int_Contar).regcom_Deb_Grv2 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Grv2 'gravado 2
                                         
          .Cells(r_int_NumFil + 2, 13) = l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv1 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Ngv1 'no gravado 1
                                         
          .Cells(r_int_NumFil + 2, 14) = l_arr_GenArc(l_int_Contar).regcom_Deb_Ngv2 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Ngv2 'no gravado 2
                                         
          .Cells(r_int_NumFil + 2, 15) = l_arr_GenArc(l_int_Contar).regcom_Deb_Igv1 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Igv1 'igv
                                         
          .Cells(r_int_NumFil + 2, 16) = l_arr_GenArc(l_int_Contar).regcom_Deb_Ret1 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Ret1 'retencion
                                         
          .Cells(r_int_NumFil + 2, 17) = l_arr_GenArc(l_int_Contar).regcom_Deb_Det1 + _
                                         l_arr_GenArc(l_int_Contar).regcom_Hab_Det1  'detraccion
                                         
          .Cells(r_int_NumFil + 2, 18) = .Cells(r_int_NumFil + 2, 11) + .Cells(r_int_NumFil + 2, 12) + _
                                         .Cells(r_int_NumFil + 2, 13) + .Cells(r_int_NumFil + 2, 14) + _
                                         .Cells(r_int_NumFil + 2, 15) + .Cells(r_int_NumFil + 2, 16) + _
                                         .Cells(r_int_NumFil + 2, 17) + .Cells(r_int_NumFil + 2, 18) 'total
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 18)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 13
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 15
         If grd_Listad.Text = "X" Then
             grd_Listad.Text = ""
         Else
              grd_Listad.Text = "X"
         End If
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_SelChange()
Dim r_str_Fecha As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
      
   r_str_Fecha = grd_Listad.TextMatrix(grd_Listad.Row, 3)
   
   cmd_Editar.Enabled = True
   cmd_Borrar.Enabled = True
   If CDate(r_str_Fecha) < CDate(modctb_str_FecIni) Then
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
   ElseIf CDate(r_str_Fecha) > CDate(modctb_str_FecFin) Then
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarComp
   Else
      If (cmb_Buscar.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub

Private Sub pnl_Codigo_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_NumDoc_Click()
   If pnl_NumDoc.Tag = "" Then
      pnl_NumDoc.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 18, "N")
   Else
      pnl_NumDoc.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 18, "N-")
   End If
End Sub

Private Sub pnl_RazSoc_Click()
   If pnl_RazSoc.Tag = "" Then
      pnl_RazSoc.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_RazSoc.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_FecCtb_Click()
   If pnl_FecCtb.Tag = "" Then
      pnl_FecCtb.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 19, "N")
   Else
      pnl_FecCtb.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 19, "N-")
   End If
End Sub

Private Sub pnl_TipComp_Click()
   If pnl_TipComp.Tag = "" Then
      pnl_TipComp.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_TipComp.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If pnl_Moneda.Tag = "" Then
      pnl_Moneda.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Moneda.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_TotComp_Click()
   If pnl_TotComp.Tag = "" Then
      pnl_TotComp.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_TotComp.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_Origen_Click()
   If pnl_Origen.Tag = "" Then
      pnl_Origen.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 7, "C")
   Else
      pnl_Origen.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 7, "C-")
   End If
End Sub

Private Sub pnl_FecComp_Click()
   If pnl_FecComp.Tag = "" Then
      pnl_FecComp.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 20, "N")
   Else
      pnl_FecComp.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 20, "N-")
   End If
End Sub

Private Sub pnl_TipCom_Click()
   If pnl_TipCom.Tag = "" Then
      pnl_TipCom.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 9, "C")
   Else
      pnl_TipCom.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 9, "C-")
   End If
End Sub

Private Sub pnl_NumRef_Click()
   If pnl_NumRef.Tag = "" Then
      pnl_NumRef.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 10, "N")
   Else
      pnl_NumRef.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 10, "N-")
   End If
End Sub

Private Sub pnl_Contab_Click()
   If pnl_Contab.Tag = "" Then
      pnl_Contab.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 11, "C")
   Else
      pnl_Contab.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 11, "C-")
   End If
End Sub

Private Sub pnl_Asiento_Click()
   If pnl_Asiento.Tag = "" Then
      pnl_Asiento.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 12, "N")
   Else
      pnl_Asiento.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 12, "N-")
   End If
End Sub

Private Sub pnl_Select_Click()
   If pnl_Select.Tag = "" Then
      pnl_Select.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 13, "C")
   Else
      pnl_Select.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 13, "C-")
   End If
End Sub

