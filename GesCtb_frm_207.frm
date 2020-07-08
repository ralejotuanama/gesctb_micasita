VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13500
   Icon            =   "GesCtb_frm_207.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8595
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13520
      _Version        =   65536
      _ExtentX        =   23848
      _ExtentY        =   15161
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
         Height          =   6045
         Left            =   60
         TabIndex        =   17
         Top             =   2340
         Width           =   13390
         _Version        =   65536
         _ExtentX        =   23618
         _ExtentY        =   10663
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5685
            Left            =   30
            TabIndex        =   18
            Top             =   360
            Width           =   13335
            _ExtentX        =   23521
            _ExtentY        =   10028
            _Version        =   393216
            Rows            =   24
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   285
            Left            =   2220
            TabIndex        =   19
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Cambio"
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
            Left            =   5050
            TabIndex        =   20
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
         Begin Threed.SSPanel pnl_CtaCtb 
            Height          =   285
            Left            =   5940
            TabIndex        =   21
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2505
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
         Begin Threed.SSPanel pnl_Referencia 
            Height          =   285
            Left            =   7350
            TabIndex        =   22
            Top             =   60
            Width           =   2050
            _Version        =   65536
            _ExtentX        =   3616
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
         Begin Threed.SSPanel pnl_Fecha 
            Height          =   285
            Left            =   1110
            TabIndex        =   23
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1976
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha"
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
            TabIndex        =   24
            Top             =   60
            Width           =   1070
            _Version        =   65536
            _ExtentX        =   1887
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
         Begin Threed.SSPanel pnl_PagNeto 
            Height          =   285
            Left            =   9390
            TabIndex        =   32
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Pago Neto"
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
            Left            =   11820
            TabIndex        =   37
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   " Seleccionar"
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
               Left            =   960
               TabIndex        =   38
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Contab 
            Height          =   285
            Left            =   10770
            TabIndex        =   39
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1870
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Contabilizado"
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
         Begin Threed.SSPanel pnl_TipPago 
            Height          =   285
            Left            =   3255
            TabIndex        =   25
            Top             =   60
            Width           =   1810
            _Version        =   65536
            _ExtentX        =   3193
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Pago"
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
         TabIndex        =   26
         Top             =   60
         Width           =   13390
         _Version        =   65536
         _ExtentX        =   23618
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
            TabIndex        =   27
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Compensaciones"
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
            Picture         =   "GesCtb_frm_207.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   28
         Top             =   780
         Width           =   13390
         _Version        =   65536
         _ExtentX        =   23618
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
         Begin VB.CommandButton cmd_Seguim 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_207.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Aprobados"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Reversa 
            Height          =   585
            Left            =   4260
            Picture         =   "GesCtb_frm_207.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_207.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   6060
            Picture         =   "GesCtb_frm_207.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Generar Asientos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   4860
            Picture         =   "GesCtb_frm_207.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_207.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_207.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_207.frx":185C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12780
            Picture         =   "GesCtb_frm_207.frx":1B66
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_207.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_207.frx":22B2
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   5460
            Picture         =   "GesCtb_frm_207.frx":25BC
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Generar Archivo"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   30
         Top             =   1470
         Width           =   13390
         _Version        =   65536
         _ExtentX        =   23618
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
         Begin VB.CheckBox chk_Estado 
            Caption         =   "Incluir Contabilizados"
            Height          =   195
            Left            =   10410
            TabIndex        =   4
            Top             =   450
            Width           =   2200
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6960
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
            Left            =   8370
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
            Left            =   6960
            TabIndex        =   33
            Top             =   90
            Width           =   2755
            _Version        =   65536
            _ExtentX        =   4860
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   360
            TabIndex        =   36
            Top             =   450
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   35
            Top             =   120
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5490
            TabIndex        =   34
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operación:"
            Height          =   195
            Left            =   5490
            TabIndex        =   31
            Top             =   450
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera
Dim l_int_Contar        As Integer

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 8)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 9) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 8)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 9) = "X"
             End If
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 8) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo borrar el registro, el pago esta contabilizado.", vbExclamation, modgen_g_str_NomPlt
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
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMPAG_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'COMPAG_CODCOM
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_Buscar
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_Buscar
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
   chk_Estado.Enabled = False
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 0
   frm_Ctb_PagCom_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
      
   If grd_Listad.TextMatrix(grd_Listad.Row, 8) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo editar el registro, el pago esta contabilizado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
    
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   frm_Ctb_PagCom_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_ExpArc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el archito de texto?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If grd_Listad.TextMatrix(grd_Listad.Row, 11) = 1 Then
      'TRANSFERENCIA
      Call fs_GenTxt_Tra
   ElseIf grd_Listad.TextMatrix(grd_Listad.Row, 11) = 6 Then
      'PAGO PROVEEDORES
      Call fs_GenTxt_Tra
   ElseIf grd_Listad.TextMatrix(grd_Listad.Row, 11) = 7 Then
      'ORDENES DE PAGO
      Call fs_GenTxt_Tra
   ElseIf grd_Listad.TextMatrix(grd_Listad.Row, 11) = 4 Then
      'DETRACCIONES
      Call fs_GenTxt_Det
   ElseIf grd_Listad.TextMatrix(grd_Listad.Row, 11) = 8 Then
   'HABERES
      Call fs_GenText_Hab
   Else
      'CHEQUE
      Call fs_GenTxt_Chq
   End If
   Screen.MousePointer = 0
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
       If grd_Listad.TextMatrix(r_int_Contad, 8) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
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
   chk_Estado.Enabled = True
   chk_Estado.Value = 0
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmd_Reversa_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
 
   If grd_Listad.TextMatrix(grd_Listad.Row, 8) = "NO" Then
      MsgBox "Solo se pueden revertir los registros contabilizados.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   moddat_g_int_FlgGrb = 3
   frm_Ctb_PagCom_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Seguim_Click()
   frm_Ctb_PagCom_06.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ctb_PagCom_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   grd_Listad.ColWidth(0) = 1050 'CODIGO
   grd_Listad.ColWidth(1) = 1110 'FECHA
   grd_Listad.ColWidth(2) = 1040 'TIPO CAMBIO
   grd_Listad.ColWidth(3) = 1800 'TIPO PAGO
   grd_Listad.ColWidth(4) = 890  'MONEDA
   grd_Listad.ColWidth(5) = 1410 'CUENTA CONTABLE
   grd_Listad.ColWidth(6) = 2040 'REFERENCIA
   grd_Listad.ColWidth(7) = 1380 'IMPORTE PAGAR
   grd_Listad.ColWidth(8) = 1050 'CONTABILIZADO
   grd_Listad.ColWidth(9) = 1190 'SELECCIONAR
   grd_Listad.ColWidth(10) = 0   'NUM_REG
   grd_Listad.ColWidth(11) = 0   'COMPAG_TIPPAG
   grd_Listad.ColWidth(12) = 0   'COMPAG_CODMON
   grd_Listad.ColWidth(13) = 0   'COMPAG_NUMLOT - DETRACCION
   grd_Listad.ColWidth(14) = 0   'COMPAG_FECPAG
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter  'CODIGO
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter  'FECHA
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter  'TIPO CAMBIO
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter    'TIPO PAGO
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter    'MONEDA
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter  'CUENTA CONTABLE
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter    'REFERENCIA
   grd_Listad.ColAlignment(7) = flexAlignRightCenter   'IMPORTE PAGAR
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter  'CONTABILIZADO
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter  'SELECCIONAR
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
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

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Public Sub fs_Buscar()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.COMPAG_CODCOM, A.COMPAG_FECPAG, A.COMPAG_TIPCAM, A.COMPAG_TIPPAG, A.COMPAG_CODMON,  "
   g_str_Parame = g_str_Parame & "         A.COMPAG_CTACTB, A.COMPAG_REFERE, TRIM(B.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "         TRIM(C.PARDES_DESCRI) As TIPPAGO, DECODE(COMPAG_FLGCTB, 1, 'SI','NO') AS CONTABILIZADO,  "
   g_str_Parame = g_str_Parame & "         (SELECT COUNT(*) FROM CNTBL_COMDET W WHERE W.COMDET_CODCOM = A.COMPAG_CODCOM) AS NUM_REG,  "
   g_str_Parame = g_str_Parame & "         (CASE A.COMPAG_CODMON  "
   g_str_Parame = g_str_Parame & "                 WHEN 1 THEN (SELECT SUM(DECODE(X.COMDET_CODMON,1,SUM(X.COMDET_IMPPAG - X.COMDET_IMPDST), SUM((X.COMDET_IMPPAG - X.COMDET_IMPDST) * A.COMPAG_TIPCAM)))  "
   g_str_Parame = g_str_Parame & "                                FROM CNTBL_COMDET X  "
   g_str_Parame = g_str_Parame & "                               Where x.COMDET_CODCOM = A.COMPAG_CODCOM  "
   g_str_Parame = g_str_Parame & "                               GROUP BY X.COMDET_CODMON)  "
   g_str_Parame = g_str_Parame & "                 WHEN 2 THEN (SELECT SUM(DECODE(X.COMDET_CODMON,2,SUM(X.COMDET_IMPPAG - X.COMDET_IMPDST), SUM((X.COMDET_IMPPAG - X.COMDET_IMPDST) / A.COMPAG_TIPCAM)))  "
   g_str_Parame = g_str_Parame & "                                FROM CNTBL_COMDET X  "
   g_str_Parame = g_str_Parame & "                               Where x.COMDET_CODCOM = A.COMPAG_CODCOM  "
   g_str_Parame = g_str_Parame & "                               GROUP BY X.COMDET_CODMON)  "
   g_str_Parame = g_str_Parame & "         END) AS PAGO_NETO, A.COMPAG_NUMLOT  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_COMPAG A  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 204 AND B.PARDES_CODITE = A.COMPAG_CODMON  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 135 AND C.PARDES_CODITE = A.COMPAG_TIPPAG  "
   g_str_Parame = g_str_Parame & "   WHERE A.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "     AND A.COMPAG_FECPAG BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   
   If chk_Estado.Value = 0 Then
      'solo procesados
      g_str_Parame = g_str_Parame & "    AND A.COMPAG_FLGCTB  = 0  "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY COMPAG_CODCOM, COMPAG_FECPAG ASC  "
   
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

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      grd_Listad.Col = 0
      grd_Listad.Text = Format(CStr(g_rst_Princi!COMPAG_CODCOM), "00000000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!COMPAG_TIPCAM, "###,###,##0.000000")
               
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!TIPPAGO & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!Moneda & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!COMPAG_CTACTB & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!COMPAG_REFERE & "")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!PAGO_NETO, "###,###,##0.00")
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!CONTABILIZADO & "")
      
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(g_rst_Princi!NUM_REG & "")
      
      grd_Listad.Col = 11
      grd_Listad.Text = g_rst_Princi!COMPAG_TIPPAG
      
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!COMPAG_CODMON
      
      If Trim(g_rst_Princi!COMPAG_NUMLOT & "") <> "" Then
         grd_Listad.Col = 13
         grd_Listad.Text = g_rst_Princi!COMPAG_NUMLOT
      End If
      
      grd_Listad.Col = 14
      grd_Listad.Text = g_rst_Princi!COMPAG_FECPAG
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE COMPENSACIONES"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "FECHA"
      .Cells(3, 4) = "TIPO CAMBIO"
      .Cells(3, 5) = "TIPO PAGO"
      .Cells(3, 6) = "MONEDA"
      .Cells(3, 7) = "CUENTA CONTABLE"
      .Cells(3, 8) = "REFERENCIA"
      .Cells(3, 9) = "PAGO NETO"
      .Cells(3, 10) = "CONTABILIZADO"
         
      .Range(.Cells(3, 2), .Cells(3, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 11 'fecha
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 13 'tipo cambio
      .Columns("D").NumberFormat = "###,###,##0.000000"
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17 'tipo pago
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 21 'moneda
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 16 'cuenta cargo
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 20 'referencia
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 16 'importe NETO
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15 'contabilizado
      .Columns("J").HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 4
      For l_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(l_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = "'" & grd_Listad.TextMatrix(l_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(l_int_Contar, 2)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(l_int_Contar, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(l_int_Contar, 4)
         .Cells(r_int_NumFil, 7) = "'" & grd_Listad.TextMatrix(l_int_Contar, 5)
         .Cells(r_int_NumFil, 8) = grd_Listad.TextMatrix(l_int_Contar, 6)
         .Cells(r_int_NumFil, 9) = grd_Listad.TextMatrix(l_int_Contar, 7)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(l_int_Contar, 8)
         
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(3, 3), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GeneraAsiento()
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_int_NumAsi    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_str_CtaHab_1  As String
Dim r_str_CtaHab_2  As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
Dim r_str_AsiGen    As String
Dim r_int_NumAux    As Integer
Dim r_str_CadAux    As String
Dim r_rst_Record    As ADODB.Recordset
Dim r_dbl_TotSol    As Double
Dim r_dbl_TotDol    As Double
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   
   For l_int_Contar = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(l_int_Contar, 8) = "NO" Then
          If grd_Listad.TextMatrix(l_int_Contar, 9) = "X" Then
         
             r_int_NumAsi = 0
             r_int_NumIte = 0
             r_str_FechaC = Format(grd_Listad.TextMatrix(l_int_Contar, 1), "yyyymmdd")
             r_str_FechaL = grd_Listad.TextMatrix(l_int_Contar, 1)        'FECHA
             r_int_PerAno = Year(grd_Listad.TextMatrix(l_int_Contar, 1))  'FECHA
             r_int_PerMes = Month(grd_Listad.TextMatrix(l_int_Contar, 1)) 'FECHA
             
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " SELECT COUNT(*) AS CONTEO  "
             g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
             g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM =  " & CLng(grd_Listad.TextMatrix(l_int_Contar, 0))
             g_str_Parame = g_str_Parame & "    AND A.COMDET_IMPPAG < 0  "
             
             If Not gf_EjecutaSQL(g_str_Parame, r_rst_Record, 3) Then
                Exit Sub
             End If
             
             If r_rst_Record.BOF And r_rst_Record.EOF Then 'No se ha encontrado ningún registro
                r_rst_Record.Close
                Set r_rst_Record = Nothing
                Screen.MousePointer = 0
                Exit Sub
             End If
            
             If r_rst_Record!CONTEO = 0 Then
                'SIN AGRUPACION
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & "  SELECT C.COMDET_TIPDOC, C.COMDET_NUMDOC, C.COMDET_CODMON, C.COMDET_CTACTB, C.COMDET_TIPDST,  "
                g_str_Parame = g_str_Parame & "         C.COMDET_IMPPAG, C.COMDET_IMPDST, C.COMDET_TIPCAM,  "
                g_str_Parame = g_str_Parame & "         DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
                g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS PROVEEDOR  "
                g_str_Parame = g_str_Parame & "    FROM CNTBL_COMDET C  "
                g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = C.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
                g_str_Parame = g_str_Parame & "    LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = C.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
                g_str_Parame = g_str_Parame & "   WHERE C.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(l_int_Contar, 0))
                g_str_Parame = g_str_Parame & "   ORDER BY PROVEEDOR ASC  "
             Else
                'AGRUPACION POR EMPRESA
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & " SELECT C.*, DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
                g_str_Parame = g_str_Parame & "             ,B.MaePrv_RazSoc) AS PROVEEDOR  "
                g_str_Parame = g_str_Parame & "   FROM (SELECT A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, A.COMDET_CTACTB, A.COMDET_TIPDST,  "
                g_str_Parame = g_str_Parame & "                SUM(A.COMDET_IMPPAG) AS COMDET_IMPPAG, SUM(A.COMDET_IMPDST) AS COMDET_IMPDST, A.COMDET_TIPCAM  "
                g_str_Parame = g_str_Parame & "           FROM CNTBL_COMDET A  "
                g_str_Parame = g_str_Parame & "          WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(l_int_Contar, 0))
                g_str_Parame = g_str_Parame & "          GROUP BY A.COMDET_TIPDST, A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, A.COMDET_CTACTB, A.COMDET_TIPCAM) C  "
                g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = C.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
                g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = C.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
                g_str_Parame = g_str_Parame & "  ORDER BY PROVEEDOR ASC  "
             End If
            
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                Exit Sub
             End If
             
             If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ningún registro
                g_rst_Princi.Close
                Set g_rst_Princi = Nothing
                Screen.MousePointer = 0
                Exit Sub
             End If
               
             'Obteniendo Nro. de Asiento
             r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
             r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
             ''TIPO CAMBIO SBS(2) - VENTA(1) - SOLO PARA LA CABECERA -  'moddat_gf_ObtieneTipCamDia(2, 2, r_str_FechaC, 1)
             r_dbl_TipSbs = CDbl(grd_Listad.TextMatrix(l_int_Contar, 2))
             'Insertar en CABECERA
             r_str_Glosa = "PAGO/" & Trim(grd_Listad.TextMatrix(l_int_Contar, 3)) & "/" & Trim(grd_Listad.TextMatrix(l_int_Contar, 6)) & _
                           "/" & CStr(grd_Listad.TextMatrix(l_int_Contar, 0))
             r_str_Glosa = Mid(r_str_Glosa, 1, 60)
             
             Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                           r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, r_str_Glosa, r_str_FechaL, "1")
            'Insertar en DETALLE
             r_dbl_TipSbs = 0
             r_int_NumIte = 0
             r_dbl_TotSol = 0
             r_dbl_TotDol = 0
             r_str_CtaHab_2 = Trim(grd_Listad.TextMatrix(l_int_Contar, 5))
             
             g_rst_Princi.MoveFirst
             Do While Not g_rst_Princi.EOF
                r_dbl_TipSbs = g_rst_Princi!COMDET_TIPCAM
                
                r_str_Glosa = "PAGO/" & Trim(g_rst_Princi!COMDET_TIPDOC) & "-" & Trim(g_rst_Princi!COMDET_NUMDOC) & "/" & Trim(g_rst_Princi!PROVEEDOR) & "/" & Trim(CStr(grd_Listad.TextMatrix(l_int_Contar, 0)))
                r_str_Glosa = Mid(r_str_Glosa, 1, 60)
                
                r_str_CtaDeb = Trim(g_rst_Princi!COMDET_CTACTB & "")
                'r_str_CtaHab_2 = Trim(grd_Listad.TextMatrix(l_int_Contar, 5))
                
                r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
                Call fs_Convertir(g_rst_Princi!COMDET_CODMON, r_dbl_TipSbs, CDbl(g_rst_Princi!COMDET_IMPPAG), r_dbl_MtoSol, r_dbl_MtoDol)
                'REGISTRO DEBE
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, _
                                                     r_str_CtaDeb, CDate(r_str_FechaL), r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
                r_str_CtaHab_1 = ""
                If g_rst_Princi!COMDET_TIPDST > 1 And g_rst_Princi!COMDET_IMPDST > 0 Then
                   'REGISTRO HABER 1
                   r_int_NumIte = r_int_NumIte + 1
                   If g_rst_Princi!COMDET_TIPDST = 2 Then
                      r_str_CtaHab_1 = "251705010107" 'ITF
                   Else
                      r_str_CtaHab_1 = "251705010103" '4TA
                   End If
                   r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
                   Call fs_Convertir(g_rst_Princi!COMDET_CODMON, r_dbl_TipSbs, CDbl(g_rst_Princi!COMDET_IMPDST), r_dbl_MtoSol, r_dbl_MtoDol)
                   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, _
                                                        r_str_CtaHab_1, CDate(r_str_FechaL), r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
                End If
                'REGISTRO HABER 2
                r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
                Call fs_Convertir(g_rst_Princi!COMDET_CODMON, r_dbl_TipSbs, (CDbl(g_rst_Princi!COMDET_IMPPAG) - CDbl(g_rst_Princi!COMDET_IMPDST)), r_dbl_MtoSol, r_dbl_MtoDol)
                r_dbl_TotSol = r_dbl_TotSol + r_dbl_MtoSol
                r_dbl_TotDol = r_dbl_TotDol + r_dbl_MtoDol
             
                'r_int_NumIte = r_int_NumIte + 1
                'Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaHab_2, CDate(r_str_FechaL), r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
                                                     
'                'Enviar a la tabla de autorizaciones
'                g_str_Parame = ""
'                g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT_ESTADO ( "
'                g_str_Parame = g_str_Parame & " " & g_rst_Princi!COMDET_CODAUT & ", " 'COMDET_CODAUT
'                g_str_Parame = g_str_Parame & " 5, " 'COMAUT_CODEST -- ESTADO PAGADO
'                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "  'SEGUSUCRE
'                g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "  'SEGPLTCRE
'                g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
'                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "  'SEGSUCCRE
'
'                If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'                   Exit Sub
'                End If
                                                     
                g_rst_Princi.MoveNext
             Loop
             
             r_str_Glosa = ""
             'r_str_Glosa = Mid("PAGO MASIVO/" & CStr(grd_Listad.TextMatrix(l_int_Contar, 0)), 1, 60)
             r_str_Glosa = Mid("PAGO MASIVO/" & CStr(grd_Listad.TextMatrix(l_int_Contar, 0)) & "/" & Trim(grd_Listad.TextMatrix(l_int_Contar, 6)), 1, 60)
             
             r_int_NumIte = r_int_NumIte + 1
             Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaHab_2, CDate(r_str_FechaL), r_str_Glosa, "H", r_dbl_TotSol, r_dbl_TotDol, 1, CDate(r_str_FechaL))
             
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " SELECT A.COMDET_CODCOM, A.COMDET_CODAUT, A.COMDET_CODOPE  "
             g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
             g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(l_int_Contar, 0))

             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
                Exit Sub
             End If

             g_rst_Listas.MoveFirst
             Do While Not g_rst_Listas.EOF
                'Enviar a la tabla de autorizaciones
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT_ESTADO ( "
                g_str_Parame = g_str_Parame & " " & g_rst_Listas!COMDET_CODAUT & ", " 'COMDET_CODAUT
                g_str_Parame = g_str_Parame & " 5, " 'COMAUT_CODEST -- ESTADO PAGADO
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'SEGUSUCRE
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'SEGPLTCRE
                g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'SEGSUCCRE

                If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                   Exit Sub
                End If

                g_rst_Listas.MoveNext
             Loop
                                                                    
             'Actualiza flag de contabilizacion
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " UPDATE CNTBL_COMPAG  "
             g_str_Parame = g_str_Parame & "    SET COMPAG_DATCTB = '" & r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "',  "
             g_str_Parame = g_str_Parame & "        COMPAG_FLGCTB = 1 ,  "
             g_str_Parame = g_str_Parame & "        COMPAG_FECCTB = " & Format(moddat_g_str_FecSis, "yyyymmdd")
             g_str_Parame = g_str_Parame & "  WHERE COMPAG_CODCOM = " & CLng(grd_Listad.TextMatrix(l_int_Contar, 0))
             
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
                Exit Sub
             End If
                          
          End If
       End If
   Next
   
   MsgBox "Se procesaron los registros seleccionados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_Convertir(ByVal p_CodMon As Integer, ByVal p_TipCam As Double, ByVal p_Importe As Double, ByRef p_ImpSol As Double, ByRef p_ImpDol As Double)
   If p_CodMon = 1 Then
      p_ImpSol = p_Importe
      p_ImpDol = Format(p_Importe / p_TipCam, "###,###,##0.00")
   Else
      p_ImpSol = Format(p_Importe * p_TipCam, "###,###,##0.00")
      p_ImpDol = p_Importe
   End If
End Sub

Private Sub fs_ConvertTXT(ByVal p_CabMon As Integer, ByVal p_DetMon As Integer, ByVal p_TipCam As Double, ByVal p_Importe As Double, ByRef p_ImpCnv As Double)
   p_ImpCnv = 0

   If p_CabMon = p_DetMon Then
      'MISMA MONEDA
      p_ImpCnv = p_Importe
   Else
      If p_CabMon = 1 Then
         'CABECERA SOLES
         p_ImpCnv = Format(p_Importe * p_TipCam, "###,###,##0.00")
      Else
         'CABECERA DOLARES
         p_ImpCnv = Format(p_Importe / p_TipCam, "###,###,##0.00")
      End If
   End If
End Sub

Private Sub fs_GenTxt_Tra()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_str_CadAux    As String
Dim r_str_FmtAux    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_RegTot    As Integer
Dim r_int_Contad    As Integer
Dim r_int_PosIni    As Integer
Dim r_str_AuxRuc    As String
Dim r_str_TipDoc    As String
Dim r_str_TipCta    As String
Dim r_str_NumCta    As String
Dim r_dbl_ImpCnv    As Double
Dim r_rst_Record    As ADODB.Recordset

   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_PAGO_" & Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & ".TXT"
   
   '===============================================
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(*) AS CONTEO  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM =  " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   g_str_Parame = g_str_Parame & "    AND A.COMDET_IMPPAG < 0  "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Record, 3) Then
      Exit Sub
   End If
    
   If r_rst_Record.BOF And r_rst_Record.EOF Then 'No se ha encontrado ningún registro
      r_rst_Record.Close
      Set r_rst_Record = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If r_rst_Record!CONTEO = 0 Then
      'SIN AGRUPACION
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT A.COMDET_TIPDOC , A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR, A.COMDET_IMPPAG, A.COMDET_IMPDST,  "
      g_str_Parame = g_str_Parame & "        DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
      g_str_Parame = g_str_Parame & "              ,B.MAEPRV_RAZSOC) AS PROVEEDOR  "
      g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
      g_str_Parame = g_str_Parame & "  ORDER BY PROVEEDOR ASC  "
   Else
      'AGRUPACION POR EMPRESA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT C.*, DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
      g_str_Parame = g_str_Parame & "                   ,B.MAEPRV_RAZSOC) AS PROVEEDOR  "
      g_str_Parame = g_str_Parame & "   FROM (SELECT A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR,  "
      g_str_Parame = g_str_Parame & "                SUM(A.COMDET_IMPPAG) AS COMDET_IMPPAG, SUM(A.COMDET_IMPDST) AS COMDET_IMPDST  " '-- A.COMDET_TIPCAM,
      g_str_Parame = g_str_Parame & "           FROM CNTBL_COMDET A  "
      g_str_Parame = g_str_Parame & "          WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
      g_str_Parame = g_str_Parame & "          GROUP BY A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR) C  "
      g_str_Parame = g_str_Parame & "           LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = C.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "           LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = C.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "  ORDER BY PROVEEDOR ASC  "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
                      
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   r_str_Cadena = ""
   r_dbl_PlaTot = 0
       
   r_str_CadAux = ""
   For r_int_Contad = 1 To 68
       r_str_CadAux = r_str_CadAux & " "
   Next
      
   'Calcular totales
   r_dbl_PlaTot = 0
   r_int_RegTot = 0
   'r_int_RegTot = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 10))
   
   'detalle del archivo
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_int_RegTot = r_int_RegTot + 1
      g_rst_Princi.MoveNext
   Loop
   r_dbl_PlaTot = CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 7))
   
   '750
   r_str_FmtAux = ""
   If grd_Listad.TextMatrix(grd_Listad.Row, 11) = 6 Or grd_Listad.TextMatrix(grd_Listad.Row, 11) = 7 Then
      r_str_FmtAux = "750"
   Else
      r_str_FmtAux = "500"
   End If
      
   If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 12)) = 1 Then
      'SOLES
      r_str_Cadena = r_str_Cadena & r_str_FmtAux & "00110661000100040896PEN" & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                    "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "TRANSFERENCIAS           " & _
                                    Format(r_int_RegTot, "000000") & "N" & r_str_CadAux
   Else
      'DOLARES
      r_str_Cadena = r_str_Cadena & r_str_FmtAux & "00110661000100040918USD" & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                    "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "TRANSFERENCIAS           " & _
                                    Format(r_int_RegTot, "000000") & "N" & r_str_CadAux
   End If
                                 
   Print #r_int_NumRes, r_str_Cadena
      
   'detalle del archivo
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_AuxRuc = ""
      r_str_TipDoc = ""
      r_str_AuxRuc = Trim(g_rst_Princi!COMDET_NUMDOC)
      Select Case CLng(g_rst_Princi!COMDET_TIPDOC)
             Case 1
                  r_str_TipDoc = "L"
             Case 4
                  r_str_TipDoc = "E"
             Case 6
                 r_str_TipDoc = "R"
             Case 7
                 r_str_TipDoc = "P"
      End Select
      r_str_TipCta = ""
      r_str_NumCta = ""
      If Trim(g_rst_Princi!COMDET_CODBNC & "") <> "" Then
         If CLng(g_rst_Princi!COMDET_CODBNC) = 11 Then
            r_str_TipCta = "P"
            r_str_NumCta = Mid(Trim(g_rst_Princi!COMDET_CTACRR), 1, 8) & "00" & Mid(Trim(g_rst_Princi!COMDET_CTACRR), 9, 10)
         Else
            r_str_TipCta = "I"
            r_str_NumCta = Left(Trim(g_rst_Princi!COMDET_CTACRR) & "                    ", 20)
         End If
      End If
      
      r_str_CadAux = ""
      r_str_FmtAux = ""
      If grd_Listad.TextMatrix(grd_Listad.Row, 11) = 7 Then
         'orden de pago(7-cheque gerencia)
         r_str_TipCta = "O"
         r_str_NumCta = Left("                    ", 20)
         r_str_FmtAux = "F000000000000N" & Left(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 3)) & "                                        ", 40)
         
         For r_int_Contad = 1 To 131
             r_str_CadAux = r_str_CadAux & " "
         Next
         r_str_CadAux = r_str_FmtAux & r_str_CadAux
      ElseIf grd_Listad.TextMatrix(grd_Listad.Row, 11) = 6 Then
         'pago proveedores(6)
         r_str_FmtAux = "F000000000000N" & Left(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 3)) & "                                        ", 40)
         For r_int_Contad = 1 To 131
             r_str_CadAux = r_str_CadAux & " "
         Next
         r_str_CadAux = r_str_FmtAux & r_str_CadAux
      Else
         'otros
         For r_int_Contad = 1 To 101
             r_str_CadAux = r_str_CadAux & " "
         Next
         r_str_CadAux = Left("000000000000000" & "                                        ", 40) & r_str_CadAux
      End If
      
      r_dbl_ImpCnv = 0
      Call fs_ConvertTXT(grd_Listad.TextMatrix(grd_Listad.Row, 12), g_rst_Princi!COMDET_CODMON, CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 2)), _
                        (g_rst_Princi!COMDET_IMPPAG - g_rst_Princi!COMDET_IMPDST), r_dbl_ImpCnv)
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "002" & r_str_TipDoc & _
                                    Left(Trim(r_str_AuxRuc) & "            ", 12) & _
                                    r_str_TipCta & r_str_NumCta & _
                                    Left(Trim(g_rst_Princi!PROVEEDOR) & "                                        ", 40) & _
                                    Format(r_dbl_ImpCnv * 100, "000000000000000") & r_str_CadAux
      Print #r_int_NumRes, r_str_Cadena
      
      g_rst_Princi.MoveNext
   Loop
               
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Public Sub fs_GenTxt_Det()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_str_CadAux    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_RegTot    As Integer
Dim r_int_Contad    As Integer
Dim r_int_PosIni    As Integer
Dim r_str_AuxRuc    As String
Dim r_str_TipDoc    As String
Dim r_str_TipCta    As String
Dim r_str_NumCta    As String
Dim r_dbl_ImpCnv    As Double
Dim r_str_NumLot    As String
Dim r_rst_Record    As ADODB.Recordset

   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NumLot = Mid(Left(grd_Listad.TextMatrix(grd_Listad.Row, 14), 4), 3, 2) & Right("0000" & grd_Listad.TextMatrix(grd_Listad.Row, 13), 4)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & "D20511904162" & r_str_NumLot & ".TXT"
   
   '===============================================
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(*) AS CONTEO  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM =  " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   g_str_Parame = g_str_Parame & "    AND A.COMDET_IMPPAG < 0  "
    
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Record, 3) Then
      Exit Sub
   End If
    
   If r_rst_Record.BOF And r_rst_Record.EOF Then 'No se ha encontrado ningún registro
      r_rst_Record.Close
      Set r_rst_Record = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If r_rst_Record!CONTEO = 0 Then
      'SIN AGRUPACION
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
      g_str_Parame = g_str_Parame & "              ,B.MAEPRV_RAZSOC) AS PROVEEDOR,  "
      g_str_Parame = g_str_Parame & "        A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR,  "
      g_str_Parame = g_str_Parame & "        A.COMDET_IMPPAG, A.COMDET_IMPDST, G.regcom_CodDet, regcom_FecEmi, regcom_TipCpb, regcom_Nserie, regcom_NroCom  "
      g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_REGCOM G ON TO_NUMBER(G.REGCOM_CODCOM) = A.COMDET_CODOPE  "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM =  " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
      g_str_Parame = g_str_Parame & "  ORDER BY PROVEEDOR ASC  "
   Else
      'AGRUPACION POR EMPRESA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT C.*, DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
      g_str_Parame = g_str_Parame & "                   ,B.MAEPRV_RAZSOC) AS PROVEEDOR  "
      g_str_Parame = g_str_Parame & "   FROM (SELECT A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR,  "
      g_str_Parame = g_str_Parame & "                SUM(A.COMDET_IMPPAG) AS COMDET_IMPPAG, SUM(A.COMDET_IMPDST) AS COMDET_IMPDST, G.REGCOM_CODDET, regcom_FecEmi, regcom_TipCpb, regcom_Nserie, regcom_NroCom  " '-- A.COMDET_TIPCAM,
      g_str_Parame = g_str_Parame & "           FROM CNTBL_COMDET A  "
      g_str_Parame = g_str_Parame & "           LEFT JOIN CNTBL_REGCOM G ON TO_NUMBER(G.REGCOM_CODCOM) = A.COMDET_CODOPE  "
      g_str_Parame = g_str_Parame & "          WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
      g_str_Parame = g_str_Parame & "          GROUP BY A.COMDET_TIPDOC, A.COMDET_NUMDOC, A.COMDET_CODMON, COMDET_CODBNC, COMDET_CTACRR, G.REGCOM_CODDET, regcom_FecEmi, regcom_TipCpb, regcom_Nserie, regcom_NroCom) C  "
      g_str_Parame = g_str_Parame & "           LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = C.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
      g_str_Parame = g_str_Parame & "           LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = C.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(C.COMDET_NUMDOC)  "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
                      
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   r_str_Cadena = ""
   r_dbl_PlaTot = 0
       
   r_str_CadAux = ""
   For r_int_Contad = 1 To 68
       r_str_CadAux = r_str_CadAux & " "
   Next
      
   'Calcular totales
   r_dbl_PlaTot = 0
   r_int_RegTot = 0
   r_int_RegTot = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 10))
   r_dbl_PlaTot = CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 7))

   'SOLES CABECERA
   r_str_Cadena = r_str_Cadena & "*20511904162" & Mid("EDPYMEMICASITASA" & "                                   ", 1, 35) & _
                                 r_str_NumLot & Right("000000000000000" & (r_dbl_PlaTot * 100), 15)
                                 
   Print #r_int_NumRes, r_str_Cadena
   
   r_str_CadAux = ""
   For r_int_Contad = 1 To 101
       r_str_CadAux = r_str_CadAux & " "
   Next
   
   'detalle del archivo
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_AuxRuc = ""
      r_str_AuxRuc = Trim(g_rst_Princi!COMDET_NUMDOC)
            
      r_dbl_ImpCnv = 0
      r_dbl_ImpCnv = (g_rst_Princi!COMDET_IMPPAG - g_rst_Princi!COMDET_IMPDST) * 100
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & g_rst_Princi!COMDET_TIPDOC & Left(Trim(r_str_AuxRuc) & "                                   ", 47) & _
                                    "000000000" & Right("00" & Trim(g_rst_Princi!regcom_CodDet), 3) & _
                                    Trim(g_rst_Princi!COMDET_CTACRR) & Right(("000000000000000" & CStr(r_dbl_ImpCnv)), 15) & _
                                    "01" & Left(g_rst_Princi!regcom_FecEmi, 6) & Right("00" & Trim(g_rst_Princi!regcom_TipCpb), 2) & _
                                    Trim(g_rst_Princi!regcom_Nserie) & Right("000" & Trim(g_rst_Princi!regcom_NroCom), 8)
      Print #r_int_NumRes, r_str_Cadena
      
      g_rst_Princi.MoveNext
   Loop
               
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Public Sub fs_GenTxt_Chq()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_PerDia    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_Contad    As Integer
Dim r_str_NomPrv    As String
Dim r_str_FecChq    As String
Dim r_str_ImpSol    As String
Dim r_str_CadAux    As String
Dim r_str_CadAux2   As String
Dim r_dbl_ImpSol    As Double
Dim r_dbl_ImpDol    As Double

   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_int_PerDia = Day(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_PAGO_" & Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & ".TXT"
   r_str_NomPrv = ""
   r_str_ImpSol = ""
   r_str_FecChq = ""
   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
   Print #r_int_NumRes, r_str_Cadena
   
   '===============================================
   'Extraer registros de la compensacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.COMDET_CODCOM,  "
   g_str_Parame = g_str_Parame & "      DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
   g_str_Parame = g_str_Parame & "             ,B.MAEPRV_RAZSOC) AS PROVEEDOR,  "
   g_str_Parame = g_str_Parame & "      A.COMDET_TIPDOC , A.COMDET_NUMDOC  "
   g_str_Parame = g_str_Parame & " FROM CNTBL_COMDET A  "
   g_str_Parame = g_str_Parame & " LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
   g_str_Parame = g_str_Parame & " LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   g_str_Parame = g_str_Parame & " GROUP BY COMDET_CODCOM, COMDET_TIPDOC, COMDET_NUMDOC,  "
   g_str_Parame = g_str_Parame & "         DECODE(B.MAEPRV_RAZSOC,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
   g_str_Parame = g_str_Parame & "                ,B.MAEPRV_RAZSOC)  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   r_str_NomPrv = Left(Replace(Trim(g_rst_Princi!PROVEEDOR), " ", "*") & "************************************************************", 60)
   r_str_ImpSol = Left(Format(CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 7)), "###,###,##0.00") & "****************", 16)
   r_str_FecChq = Format(moddat_g_str_FecSis, "dd / mm / yyyy")
   r_str_CadAux2 = Format(CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 7)), "###,###,##0.00")
     
   r_str_Cadena = "                                 " & r_str_FecChq & "    " & r_str_ImpSol
   Print #r_int_NumRes, r_str_Cadena
   r_str_Cadena = ""
   Print #r_int_NumRes, r_str_Cadena
   r_str_Cadena = ""
   Print #r_int_NumRes, r_str_Cadena
   r_str_Cadena = "           " & r_str_NomPrv
   Print #r_int_NumRes, r_str_Cadena
   r_str_Cadena = ""
   Print #r_int_NumRes, r_str_Cadena
   r_str_CadAux = fs_NumLetra(Left(r_str_CadAux2, Len(r_str_CadAux2) - 3))
   r_str_CadAux = r_str_CadAux & " CON " & Right(Format(CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 7)), "###,###,##0.00"), 2) & "/100"
   r_str_CadAux = Left(Replace(r_str_CadAux, " ", "*") & "****************************************************************", 64)
   r_str_Cadena = " " & r_str_CadAux
   Print #r_int_NumRes, r_str_Cadena
               
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_GenText_Hab()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_str_CadAux    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_RegTot    As Integer
   
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_int_NumAsi    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_int_Contar    As Integer
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
Dim r_dbl_ImpCnv    As Double
Dim r_str_moneda    As String
Dim r_str_NumCta_Cab    As String
Dim r_str_NumCta_Cli    As String
Dim r_str_TipCta_Cli    As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_CODSIC, A.MAEPRV_NUMDOC, CASE A.MAEPRV_TIPDOC "
   g_str_Parame = g_str_Parame & "                                             WHEN 1 THEN 'L' "
   g_str_Parame = g_str_Parame & "                                             WHEN 4 THEN 'E' "
   g_str_Parame = g_str_Parame & "                                             WHEN 7 THEN 'P' "
   g_str_Parame = g_str_Parame & "                                             END AS TIPDOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_RAZSOC, "
   
   'g_str_Parame = g_str_Parame & "        DECODE(A.MAEPRV_CODBNC_MN1,11,'P','I') AS TIPCTA_SOLES, "
   'g_str_Parame = g_str_Parame & "        DECODE(A.MAEPRV_CODBNC_DL1,11,'P','I') AS TIPCTA_DOLARES, "
   'g_str_Parame = g_str_Parame & "        DECODE(A.MAEPRV_CODBNC_MN1,11,SUBSTR(TRIM(A.MAEPRV_CTACRR_MN1),1,8) || '00'|| SUBSTR(TRIM(A.MAEPRV_CTACRR_MN1),9,10),"
   'g_str_Parame = g_str_Parame & "               A.MAEPRV_NROCCI_MN1) AS NUM_CUENTA_SOLES,"
   'g_str_Parame = g_str_Parame & "        DECODE(A.MAEPRV_CODBNC_DL1,11,SUBSTR(TRIM(A.MAEPRV_CTACRR_DL1),1,8) || '00'|| SUBSTR(TRIM(A.MAEPRV_CTACRR_DL1),9,10),"
   'g_str_Parame = g_str_Parame & "               A.MAEPRV_NROCCI_DL1) AS NUM_CUENTA_DOLARES,"
                  
   g_str_Parame = g_str_Parame & "        DECODE(B.COMDET_CODBNC,11,'P','I') AS TIPCTA, "
   g_str_Parame = g_str_Parame & "        DECODE(B.COMDET_CODBNC,11,SUBSTR(TRIM(B.COMDET_CTACRR),1,8) || '00'|| SUBSTR(TRIM(B.COMDET_CTACRR),9,10), "
   g_str_Parame = g_str_Parame & "               B.COMDET_CTACRR) AS NUM_CUENTA, "
              
   g_str_Parame = g_str_Parame & "        B.COMDET_IMPPAG, B.COMDET_IMPDST, COMDET_CODMON "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET B "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV A ON A.MAEPRV_TIPDOC = B.COMDET_TIPDOC AND TRIM(A.MAEPRV_NUMDOC) = TRIM(B.COMDET_NUMDOC) "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "    AND B.COMDET_CODCOM = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   
   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_PAGO_HABERES_" & Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & ".TXT"
                      
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   r_str_Cadena = ""
   r_dbl_PlaTot = 0
   r_int_RegTot = 0
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
       r_dbl_ImpCnv = 0
       Call fs_ConvertTXT(grd_Listad.TextMatrix(grd_Listad.Row, 12), g_rst_Princi!COMDET_CODMON, CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 2)), _
                         (g_rst_Princi!COMDET_IMPPAG - g_rst_Princi!COMDET_IMPDST), r_dbl_ImpCnv)
                        
       r_dbl_PlaTot = r_dbl_PlaTot + r_dbl_ImpCnv
       r_int_RegTot = r_int_RegTot + 1
       
       g_rst_Princi.MoveNext
   Loop
   
   r_str_CadAux = ""
   For r_int_Contar = 1 To 68
       r_str_CadAux = r_str_CadAux & " "
   Next
   
   r_str_TipCta_Cli = ""
   r_str_NumCta_Cab = ""
   r_str_moneda = "PEN"
   r_str_NumCta_Cab = "00110661000100040896"
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 12) = 2 Then
      r_str_moneda = "USD"
      r_str_NumCta_Cab = "00110661000100040918"
   End If
   r_str_Cadena = r_str_Cadena & "700" & r_str_NumCta_Cab & r_str_moneda & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                 "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "HABERES 5TA CATEGORIA    " & _
                                 Format(r_int_RegTot, "000000") & "S" & r_str_CadAux
   Print #r_int_NumRes, r_str_Cadena
      
   r_str_CadAux = ""
   For r_int_Contar = 1 To 101
       r_str_CadAux = r_str_CadAux & " "
   Next
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_dbl_ImpCnv = 0
      Call fs_ConvertTXT(grd_Listad.TextMatrix(grd_Listad.Row, 12), g_rst_Princi!COMDET_CODMON, CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 2)), _
                        (g_rst_Princi!COMDET_IMPPAG - g_rst_Princi!COMDET_IMPDST), r_dbl_ImpCnv)
                         
       r_str_NumCta_Cli = ""
       r_str_TipCta_Cli = ""
       
       r_str_NumCta_Cli = g_rst_Princi!NUM_CUENTA
       r_str_TipCta_Cli = g_rst_Princi!TIPCTA
       
       'r_str_NumCta_Cli = g_rst_Princi!NUM_CUENTA_SOLES
       'r_str_TipCta_Cli = g_rst_Princi!TIPCTA_SOLES
       'If grd_Listad.TextMatrix(grd_Listad.Row, 12) = 2 Then
       '   r_str_NumCta_Cli = g_rst_Princi!NUM_CUENTA_DOLARES
       '   r_str_TipCta_Cli = g_rst_Princi!TIPCTA_DOLARES
       'End If
       
       If Trim(r_str_NumCta_Cli) = "" Then
          MsgBox "El proveedor no tiene registrado una cuenta.", vbExclamation, modgen_g_str_NomPlt
          Exit Sub
       End If
       
       r_str_Cadena = ""
       r_str_Cadena = r_str_Cadena & "002" & g_rst_Princi!TIPDOC & _
                                     Left(Trim(g_rst_Princi!MAEPRV_NUMDOC) & "            ", 12) & _
                                     Trim(r_str_TipCta_Cli) & _
                                     Trim(r_str_NumCta_Cli) & _
                                     Left("HABERES" & Trim(g_rst_Princi!MAEPRV_CODSIC) & r_int_PerAno & Format(r_int_PerMes, "00") & "                                        ", 40) & _
                                     Format(r_dbl_ImpCnv * 100, "000000000000000") & _
                                     Left("HABERES " & fs_nombresMes(r_int_PerMes) & "                                        ", 40) & _
                                     r_str_CadAux
       Print #r_int_NumRes, r_str_Cadena
       
      g_rst_Princi.MoveNext
   Loop
   
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Function fs_nombresMes(p_Mes As Integer) As String
   Select Case p_Mes
          Case 1: fs_nombresMes = "ENERO"
          Case 2: fs_nombresMes = "FEBRERO"
          Case 3: fs_nombresMes = "MARZO"
          Case 4: fs_nombresMes = "ABRIL"
          Case 5: fs_nombresMes = "MAYO"
          Case 6: fs_nombresMes = "JUNIO"
          Case 7: fs_nombresMes = "JULIO"
          Case 8: fs_nombresMes = "AGOSTO"
          Case 9: fs_nombresMes = "SETIEMBRE"
          Case 10: fs_nombresMes = "OCTUBRE"
          Case 11: fs_nombresMes = "NOVIEMBRE"
          Case 12: fs_nombresMes = "DICIEMBRE"
   End Select
End Function

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 8
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 9
         If grd_Listad.Text = "X" Then
             grd_Listad.Text = ""
         Else
              grd_Listad.Text = "X"
         End If
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Estado)
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub chk_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

'============================================
 Public Function fs_NumLetra(ByVal p_Numero As Double) As String
        Select Case p_Numero
            Case 0: fs_NumLetra = "CERO"
            Case 1: fs_NumLetra = "UNO"
            Case 2: fs_NumLetra = "DOS"
            Case 3: fs_NumLetra = "TRES"
            Case 4: fs_NumLetra = "CUATRO"
            Case 5: fs_NumLetra = "CINCO"
            Case 6: fs_NumLetra = "SEIS"
            Case 7: fs_NumLetra = "SIETE"
            Case 8: fs_NumLetra = "OCHO"
            Case 9: fs_NumLetra = "NUEVE"
            Case 10: fs_NumLetra = "DIEZ"
            Case 11: fs_NumLetra = "ONCE"
            Case 12: fs_NumLetra = "DOCE"
            Case 13: fs_NumLetra = "TRECE"
            Case 14: fs_NumLetra = "CATORCE"
            Case 15: fs_NumLetra = "QUINCE"
            Case Is < 20: fs_NumLetra = "DIECI" & fs_NumLetra(p_Numero - 10)
            Case 20: fs_NumLetra = "VEINTE"
            Case Is < 30: fs_NumLetra = "VEINTI" & fs_NumLetra(p_Numero - 20)
            Case 30: fs_NumLetra = "TREINTA"
            Case 40: fs_NumLetra = "CUARENTA"
            Case 50: fs_NumLetra = "CINCUENTA"
            Case 60: fs_NumLetra = "SESENTA"
            Case 70: fs_NumLetra = "SETENTA"
            Case 80: fs_NumLetra = "OCHENTA"
            Case 90: fs_NumLetra = "NOVENTA"
            Case Is < 100: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 10) * 10) & " Y " & fs_NumLetra(p_Numero Mod 10)
            Case 100: fs_NumLetra = "CIEN"
            Case Is < 200: fs_NumLetra = "CIENTO " & fs_NumLetra(p_Numero - 100)
            Case 200, 300, 400, 600, 800: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 100)) & "CIENTOS"
            Case 500: fs_NumLetra = "QUINIENTOS"
            Case 700: fs_NumLetra = "SETECIENTOS"
            Case 900: fs_NumLetra = "NOVECIENTOS"
            Case Is < 1000: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 100) * 100) & " " & fs_NumLetra(p_Numero Mod 100)
            Case 1000: fs_NumLetra = "MIL"
            Case Is < 2000: fs_NumLetra = "MIL " & fs_NumLetra(p_Numero Mod 1000)
            Case Is < 1000000: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 1000)) & " MIL"
                If p_Numero Mod 1000 Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero Mod 1000)
            Case 1000000: fs_NumLetra = "UN MILLON"
            Case Is < 2000000: fs_NumLetra = "UN MILLON " & fs_NumLetra(p_Numero Mod 1000000)
            Case Is < 1000000000000#: fs_NumLetra = fs_NumLetra(Int(p_Numero / 1000000)) & " MILLONES "
                If (p_Numero - Int(p_Numero / 1000000) * 1000000) Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000) * 1000000)
            Case 1000000000000#: fs_NumLetra = "UN BILLON"
            Case Is < 2000000000000#: fs_NumLetra = "UN BILLON " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#)
            Case Else: fs_NumLetra = fs_NumLetra(Int(p_Numero / 1000000000000#)) & " BILLONES"
                If (p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#) Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#)
        End Select
    End Function

Private Sub pnl_Codigo_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Contab_Click()
   If pnl_Contab.Tag = "" Then
      pnl_Contab.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_Contab.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 8, "C-")
   End If
End Sub

Private Sub pnl_CtaCtb_Click()
   If pnl_CtaCtb.Tag = "" Then
      pnl_CtaCtb.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 5, "N")
   Else
      pnl_CtaCtb.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "N-")
   End If
End Sub

Private Sub pnl_Fecha_Click()
   If pnl_Fecha.Tag = "" Then
      pnl_Fecha.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 14, "N")
   Else
      pnl_Fecha.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 14, "N-")
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If pnl_Moneda.Tag = "" Then
      pnl_Moneda.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Moneda.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_PagNeto_Click()
   If pnl_PagNeto.Tag = "" Then
      pnl_PagNeto.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_PagNeto.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_Referencia_Click()
   If pnl_Referencia.Tag = "" Then
      pnl_Referencia.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Referencia.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Select_Click()
   If pnl_Select.Tag = "" Then
      pnl_Select.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 9, "C")
   Else
      pnl_Select.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 9, "C-")
   End If
End Sub

Private Sub pnl_TipCambio_Click()
   If pnl_TipCambio.Tag = "" Then
      pnl_TipCambio.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 2, "N")
   Else
      pnl_TipCambio.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "N-")
   End If
End Sub

Private Sub pnl_TipPago_Click()
   If pnl_TipPago.Tag = "" Then
      pnl_TipPago.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_TipPago.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub
