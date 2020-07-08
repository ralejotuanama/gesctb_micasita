VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegVen_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16500
   Icon            =   "GesCtb_frm_203.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   16500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9225
      Left            =   -150
      TabIndex        =   16
      Top             =   -30
      Width           =   16710
      _Version        =   65536
      _ExtentX        =   29475
      _ExtentY        =   16272
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
         Height          =   6165
         Left            =   210
         TabIndex        =   17
         Top             =   2370
         Width           =   16410
         _Version        =   65536
         _ExtentX        =   28945
         _ExtentY        =   10874
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
            Height          =   5775
            Left            =   30
            TabIndex        =   18
            Top             =   360
            Width           =   16380
            _ExtentX        =   28893
            _ExtentY        =   10186
            _Version        =   393216
            Rows            =   24
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   285
            Left            =   6450
            TabIndex        =   19
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
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
         Begin Threed.SSPanel pnl_Glosa 
            Height          =   285
            Left            =   7500
            TabIndex        =   20
            Top             =   60
            Width           =   2550
            _Version        =   65536
            _ExtentX        =   4498
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
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   10020
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_DebME 
            Height          =   285
            Left            =   10890
            TabIndex        =   22
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Debe"
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
         Begin Threed.SSPanel pnl_HabME 
            Height          =   285
            Left            =   12030
            TabIndex        =   23
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2081
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Haber"
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
            Left            =   1110
            TabIndex        =   24
            Top             =   60
            Width           =   1420
            _Version        =   65536
            _ExtentX        =   2505
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2490
            TabIndex        =   25
            Top             =   60
            Width           =   3980
            _Version        =   65536
            _ExtentX        =   7020
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   60
            TabIndex        =   26
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   15180
            TabIndex        =   27
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seleccionar"
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
               Left            =   930
               TabIndex        =   28
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   13200
            TabIndex        =   29
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   14250
            TabIndex        =   30
            Top             =   60
            Width           =   950
            _Version        =   65536
            _ExtentX        =   1676
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "N° Asiento"
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
         Left            =   210
         TabIndex        =   31
         Top             =   90
         Width           =   16410
         _Version        =   65536
         _ExtentX        =   28945
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
            TabIndex        =   32
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Ventas"
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
            Picture         =   "GesCtb_frm_203.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   210
         TabIndex        =   33
         Top             =   810
         Width           =   16410
         _Version        =   65536
         _ExtentX        =   28945
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_203.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_203.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_203.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_203.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15810
            Picture         =   "GesCtb_frm_203.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_203.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_203.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_203.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   4260
            Picture         =   "GesCtb_frm_203.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Generar Archivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   4860
            Picture         =   "GesCtb_frm_203.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Generar Asientos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   210
         TabIndex        =   34
         Top             =   1500
         Width           =   16410
         _Version        =   65536
         _ExtentX        =   28945
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
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
            TabIndex        =   35
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
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   39
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   38
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   37
            Top             =   450
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Asiento:"
            Height          =   195
            Left            =   5310
            TabIndex        =   36
            Top             =   450
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   585
         Left            =   210
         TabIndex        =   40
         Top             =   8580
         Width           =   16410
         _Version        =   65536
         _ExtentX        =   28945
         _ExtentY        =   1041
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
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   2595
         End
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   15
            Top             =   180
            Width           =   4425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   42
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegVen_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_RegVen
   regven_CodVen        As String
   regven_FecCtb        As String
   regven_FecEmi        As String
   regven_FecVto        As String
   regven_TipCpb_Lrg    As String
   regven_TipCpb        As String
   regven_Nserie        As String
   regven_NroCom        As String
   regven_TipDoc_Lrg    As String
   regven_TipDoc        As String
   regven_NumDoc        As String
   MaePrv_RazSoc        As String
   regven_Descrp        As String
   regven_Grv           As String
   regven_Igv           As String
   regven_Ngrv          As String
   regven_Total         As String
   regven_Deb_Grv1      As Double
   regven_Deb_Ngv1      As Double
   regven_Deb_Igv1      As Double
   regven_Deb_Ppg1      As Double
   regven_Hab_Grv1      As Double
   regven_Hab_Ngv1      As Double
   regven_Hab_Igv1      As Double
   regven_Hab_Ppg1      As Double
   regven_CodMon        As String
   regven_Moneda        As String
   regven_TipCam        As String
   regven_Ref_FecEmi    As String
   regven_Ref_TipCpb    As String
   regven_Ref_Nserie    As String
   regven_Ref_NroCom    As String
   regven_FlgCnt        As Integer
   regven_Cnt_Grv1      As String
   regven_Cnt_Ngv1      As String
   regven_Cnt_Igv1      As String
   regven_Cnt_Ppg1      As String
   regven_retbie        As Integer
End Type
   
Dim l_arr_GenArc()      As arr_RegVen
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera

Private Sub chkSeleccionar_Click()
 Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 12) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 12) = "X"
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
   moddat_g_int_FlgGrb = 1
   frm_Ctb_RegVen_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 10
   If UCase(Trim(grd_Listad.Text)) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar el registro por que esta contabilizado.", vbExclamation, modgen_g_str_NomPlt
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
   g_str_Parame = g_str_Parame & " USP_CNTBL_REGVEN_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'REGVEN_CODVEN
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

Private Sub cmd_Buscar_Click()
   Call fs_BuscarComp
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
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
   grd_Listad.Col = 8
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 9
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 0
   frm_Ctb_RegVen_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 8
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 9
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
      
   'Estado de la edicion
   moddat_g_int_FlgAct = 0 'edicion normal
   
   grd_Listad.Col = 10
   If (UCase(grd_Listad.Text) = "SI") Then
       moddat_g_int_FlgAct = 1 'reversa contabilidad
   End If
   
   moddat_g_int_FlgGrb = 2
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_RegVen_02.Show 1
   
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
Dim r_str_FecUlt     As String
Dim r_str_FecAno     As String
Dim r_str_FecMes     As String
Dim r_str_FecDia     As String
Dim r_int_Contad     As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el archivo de texto? ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Verifica que exista ruta
   If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
      MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   Screen.MousePointer = 11
   r_str_FecUlt = ""
   For r_int_Contad = 1 To grd_Listad.Rows - 1
       r_str_FecDia = Format(ff_Ultimo_Dia_Mes(Month(grd_Listad.TextMatrix(r_int_Contad, 3)), Year(grd_Listad.TextMatrix(r_int_Contad, 3))), "00")
       r_str_FecMes = Format(Format(grd_Listad.TextMatrix(r_int_Contad, 3), "mm"), "00")
       r_str_FecAno = Format(grd_Listad.TextMatrix(r_int_Contad, 3), "yyyy")
       r_str_FecUlt = r_str_FecDia & "/" & r_str_FecMes & "/" & r_str_FecAno
       Exit For
   Next
   'Dice:               LE20100118760201702001401000011
   'Debe Decir:         LE2051190416220170200140100001111
   '----------Creando Archivo - Registro de Ventas----------
   r_str_NomRes1 = moddat_g_str_RutLoc & "\LE20511904162" & r_str_FecAno & r_str_FecMes & "00140100001111.TXT"
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
   Dim r_str_Col_13, r_str_Col_24 As String
   Dim r_str_Col_23      As String
   Dim r_str_CadAux      As String
   
   For r_int_Contad = 1 To UBound(l_arr_GenArc)
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
       If CInt(l_arr_GenArc(r_int_Contad).regven_TipCpb) <> "9999" Then 'OTROS
          r_str_Fecemi = "01/01/0001"
          If Len(Trim(l_arr_GenArc(r_int_Contad).regven_FecEmi)) > 0 Then
             r_str_Fecemi = gf_FormatoFecha(l_arr_GenArc(r_int_Contad).regven_FecEmi)
             r_str_FecDia = Format(ff_Ultimo_Dia_Mes(Month(r_str_Fecemi), Year(r_str_Fecemi)), "00")
             r_str_FecMes = Format(Format(r_str_Fecemi, "mm"), "00")
             r_str_FecAno = Format(r_str_Fecemi, "yyyy")
             r_str_Fecemi = r_str_FecDia & "/" & r_str_FecMes & "/" & r_str_FecAno
          End If
          
          r_str_Fecvto = "" '"01/01/0001"
          If Len(Trim(l_arr_GenArc(r_int_Contad).regven_FecVto)) > 0 Then
             r_str_Fecvto = gf_FormatoFecha(l_arr_GenArc(r_int_Contad).regven_FecVto)
          End If
          'tipo comprobante
          r_str_tipcpb = l_arr_GenArc(r_int_Contad).regven_TipCpb
          If (Trim(l_arr_GenArc(r_int_Contad).regven_TipCpb) = "9999") Then
              r_str_tipcpb = "00"
          End If
         'fecha emision referencia
          r_str_Ref_fecemi = "01/01/0001"
          If Len(Trim(l_arr_GenArc(r_int_Contad).regven_Ref_FecEmi)) > 0 Then
             r_str_Ref_fecemi = gf_FormatoFecha(l_arr_GenArc(r_int_Contad).regven_Ref_FecEmi)
          End If
          'tipo documento
          r_str_TipDoc = Trim(l_arr_GenArc(r_int_Contad).regven_TipDoc)
          If (Trim(l_arr_GenArc(r_int_Contad).regven_TipDoc) = "9999") Then
              r_str_TipDoc = "0"
          End If
          'tipo comprobante referencia
          r_str_Ref_tipcpb = l_arr_GenArc(r_int_Contad).regven_Ref_TipCpb
          If (Trim(l_arr_GenArc(r_int_Contad).regven_Ref_TipCpb) = "9999") Then
              r_str_Ref_tipcpb = "00"
          End If
          r_str_Col_41 = "6"
          'ultima columna
          If CDbl(Format(r_str_Fecemi, "yyyymm")) >= CDbl(Left(l_arr_GenArc(r_int_Contad).regven_FecCtb, 6)) And _
             CDbl(Format(r_str_Fecemi, "yyyymm")) <= CDbl(Left(l_arr_GenArc(r_int_Contad).regven_FecCtb, 6)) Then
             r_str_Col_41 = 1
          End If
          r_str_Col_13 = "0"
          r_str_Col_15 = "0"
          r_str_Col_17 = "0"
          r_str_Col_18 = "0"
          r_str_Col_19 = "0"
          r_str_Col_20 = "0"
          r_str_Col_21 = "0"
          r_str_Col_22 = "0"
          r_str_Col_23 = "0"
          If CInt(l_arr_GenArc(r_int_Contad).regven_CodMon) = 2 Then
             r_dbl_TipCam = Trim(CDbl(l_arr_GenArc(r_int_Contad).regven_TipCam))
             r_str_Col_14 = Trim(CDbl(Format(CDbl(l_arr_GenArc(r_int_Contad).regven_Grv * r_dbl_TipCam), "########0.00")))
             r_str_Col_16 = Trim(CDbl(Format(CDbl(l_arr_GenArc(r_int_Contad).regven_Igv * r_dbl_TipCam), "########0.00")))
             r_str_Col_24 = Trim(CDbl(Format(CDbl(l_arr_GenArc(r_int_Contad).regven_Total * r_dbl_TipCam), "########0.0")))
          Else
             r_str_Col_14 = Trim(l_arr_GenArc(r_int_Contad).regven_Grv)
             r_str_Col_16 = Trim(l_arr_GenArc(r_int_Contad).regven_Igv)
             r_str_Col_24 = Trim(l_arr_GenArc(r_int_Contad).regven_Total)
          End If
                     
          Print #1, Mid(l_arr_GenArc(r_int_Contad).regven_FecCtb, 1, 6) & "00"; R_STR_CONSTT; _
                    Format(r_int_NroCor, "000000000000"); R_STR_CONSTT; _
                    "M" & Format(r_int_NroCor, "000000000"); R_STR_CONSTT; _
                    r_str_Fecemi; R_STR_CONSTT; ""; R_STR_CONSTT; _
                    Format(r_str_tipcpb, "00"); R_STR_CONSTT; _
                    l_arr_GenArc(r_int_Contad).regven_Nserie; R_STR_CONSTT; _
                    l_arr_GenArc(r_int_Contad).regven_NroCom; R_STR_CONSTT; ""; R_STR_CONSTT; _
                    r_str_TipDoc; R_STR_CONSTT; l_arr_GenArc(r_int_Contad).regven_NumDoc; R_STR_CONSTT; _
                    l_arr_GenArc(r_int_Contad).MaePrv_RazSoc; R_STR_CONSTT; _
                    r_str_Col_13; R_STR_CONSTT; r_str_Col_14; R_STR_CONSTT; _
                    r_str_Col_15; R_STR_CONSTT; r_str_Col_16; R_STR_CONSTT; _
                    r_str_Col_17; R_STR_CONSTT; r_str_Col_18; R_STR_CONSTT; _
                    r_str_Col_19; R_STR_CONSTT; r_str_Col_20; R_STR_CONSTT; _
                    r_str_Col_21; R_STR_CONSTT; r_str_Col_22; R_STR_CONSTT; _
                    r_str_Col_23; R_STR_CONSTT; r_str_Col_24; R_STR_CONSTT; _
                    IIf(CInt(l_arr_GenArc(r_int_Contad).regven_CodMon) = 1, "PEN", "USD"); R_STR_CONSTT; _
                    Format(l_arr_GenArc(r_int_Contad).regven_TipCam, "#,##0.000"); R_STR_CONSTT; _
                    r_str_Ref_fecemi; R_STR_CONSTT; Format(r_str_Ref_tipcpb, "00"); R_STR_CONSTT; _
                    IIf(Trim(l_arr_GenArc(r_int_Contad).regven_Ref_Nserie) = "", "-", l_arr_GenArc(r_int_Contad).regven_Ref_Nserie); R_STR_CONSTT; _
                    IIf(Trim(l_arr_GenArc(r_int_Contad).regven_Ref_NroCom) = "", "-", l_arr_GenArc(r_int_Contad).regven_Ref_NroCom); R_STR_CONSTT; _
                    ""; R_STR_CONSTT; ""; R_STR_CONSTT; "1"; R_STR_CONSTT; "1"; R_STR_CONSTT
                                        
           r_int_NroCor = r_int_NroCor + 1
     End If
   Next
   Print #1, r_str_FecAno & r_str_FecMes & "00|" & Format(r_int_NroCor, "000000000000") & "|M" & Format(r_int_NroCor, "000000000") & "|" & r_str_FecUlt & "||13|0013|13|||||0|0|0|0|0|0|0|0|0|0|0|0|PEN|1.000|01/01/0001|00|-|-|||1|1|"
   r_int_NroCor = r_int_NroCor + 1
   Print #1, r_str_FecAno & r_str_FecMes & "00|" & Format(r_int_NroCor, "000000000000") & "|M" & Format(r_int_NroCor, "000000000") & "|" & r_str_FecUlt & "||13|0013|13|||||0|0|0|0|0|0|0|0|0|0|0|0|PEN|1.000|01/01/0001|00|-|-|||1|1|"

   Close #1

   Screen.MousePointer = 0 'vbCrLf r_str_NomRes
   MsgBox "El archivo ha sido creado. " & vbCrLf & _
          "Registro de ventas: " & Trim(r_str_NomRes1), vbInformation, modgen_g_str_NomPlt
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
       If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contad, 12) = "X" Then
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

Private Sub cmd_Salida_Click()
   Unload Me
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
   
   grd_Listad.ColWidth(0) = 1070 'Codigo
   grd_Listad.ColWidth(1) = 1400 'Nro Documento
   grd_Listad.ColWidth(2) = 3950 'Razon Social
   grd_Listad.ColWidth(3) = 1060 'Fecha Contable
   grd_Listad.ColWidth(4) = 2510 'Tipo Comprobante
   grd_Listad.ColWidth(5) = 870  'Moneda
   grd_Listad.ColWidth(6) = 1140 'Total Debe
   grd_Listad.ColWidth(7) = 1150 'Total Haber
   grd_Listad.ColWidth(8) = 0    '
   grd_Listad.ColWidth(9) = 0    '
   grd_Listad.ColWidth(10) = 1040 'Contabilizado
   grd_Listad.ColWidth(11) = 930 'Nro Asiento
   grd_Listad.ColWidth(12) = 920 'Seleccionar
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
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
   
   cmb_Buscar.ListIndex = 0
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 10
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 12
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

Public Sub fs_BuscarComp()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String

   ReDim l_arr_GenArc(0)
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT  LPAD(A.REGVEN_CODVEN,10,'0') AS REGVEN_CODVEN, A.REGVEN_TIPDOC || '-' || A.REGVEN_NUMDOC ID_CLIENTE, TRIM(B.MAEPRV_RAZSOC) MAEPRV_RAZSOC,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_FECCTB, TRIM(C.PARDES_DESCRI) TIP_COMPROBANTE, TRIM(D.PARDES_DESCRI) MONEDA,  "
   g_str_Parame = g_str_Parame & "        REGVEN_DEB_GRV1, REGVEN_DEB_NGV1, REGVEN_DEB_IGV1, REGVEN_DEB_PPG1,  " 'TOTAL DEBE
   g_str_Parame = g_str_Parame & "        REGVEN_HAB_GRV1, REGVEN_HAB_NGV1, REGVEN_HAB_IGV1, REGVEN_HAB_PPG1,  " 'TOT_HABER
   g_str_Parame = g_str_Parame & "        A.REGVEN_FECCTB, A.REGVEN_FECEMI, A.REGVEN_FECVTO, A.REGVEN_TIPCPB,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_NSERIE, A.REGVEN_NROCOM, TRIM(B.MAEPRV_RAZSOC) MAEPRV_RAZSOC,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_CODMON, A.REGVEN_TIPCAM, A.REGVEN_REF_FECEMI,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_REF_TIPCPB, REGVEN_REF_NSERIE, REGVEN_REF_NROCOM,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_CNT_GRV1, A.REGVEN_CNT_NGV1, A.REGVEN_CNT_IGV1, A.REGVEN_CNT_PPG1,  "
   g_str_Parame = g_str_Parame & "        A.REGVEN_FLGCNT, A.REGVEN_DESCRP, TRIM(E.PARDES_DESCRI) TIPO_DOCUMENTO,  "
   g_str_Parame = g_str_Parame & "        TRIM(SUBSTR(REGVEN_DATCNT_1,15,20)) NRO_ASIENTO_1,  "
   g_str_Parame = g_str_Parame & "        TRIM(SUBSTR(REGVEN_DATCNT_2,15,20)) NRO_ASIENTO_2, A.REGVEN_TIPDOC, A.REGVEN_NUMDOC, REGVEN_RETBIE  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_REGVEN A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON A.REGVEN_TIPDOC = B.MAEPRV_TIPDOC AND A.REGVEN_NUMDOC = B.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 123 AND A.REGVEN_TIPCPB = C.PARDES_CODITE  "  'comprobante"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND A.REGVEN_CODMON = D.PARDES_CODITE  "  'moneda"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 118 AND A.REGVEN_TIPDOC = E.PARDES_CODITE  "  'documento"
   g_str_Parame = g_str_Parame & "  WHERE A.REGVEN_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND A.REGVEN_FECCTB BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   
   If (cmb_Buscar.ListIndex = 1) Then 'numero de documento
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND REGVEN_NUMDOC = '" & Trim(txt_Buscar.Text) & "' "
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
           g_str_Parame = g_str_Parame & "   AND REGVEN_FLGCNT = " & r_str_Cadena
       End If
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.REGVEN_CODVEN ASC "

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
      grd_Listad.Text = CStr(g_rst_Princi!regven_CodVen)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!ID_CLIENTE & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(Trim(g_rst_Princi!regven_FecCtb & ""))
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!TIP_COMPROBANTE & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 6 'TOTAL DEBE
      grd_Listad.Text = g_rst_Princi!regven_Deb_Grv1 + g_rst_Princi!regven_Deb_Ngv1 + _
                        g_rst_Princi!regven_Deb_Igv1 + g_rst_Princi!regven_Deb_Ppg1
      grd_Listad.Text = Format(grd_Listad.Text, "###,###,###,##0.00")
                        
      grd_Listad.Col = 7 'TOT_HABER
      grd_Listad.Text = g_rst_Princi!regven_Hab_Grv1 + g_rst_Princi!regven_Hab_Ngv1 + _
                        g_rst_Princi!regven_Hab_Igv1 + g_rst_Princi!regven_Hab_Ppg1
      grd_Listad.Text = Format(grd_Listad.Text, "###,###,###,##0.00")

      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!regven_TipDoc & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!regven_NumDoc & "")
            
      grd_Listad.Col = 10
      grd_Listad.Text = IIf(g_rst_Princi!regven_FlgCnt = 1, "SI", "NO")
      
      If Trim(g_rst_Princi!NRO_ASIENTO_2 & "") = "" Then
         grd_Listad.Col = 11
         grd_Listad.Text = Trim(g_rst_Princi!NRO_ASIENTO_1 & "")
      Else
         grd_Listad.Col = 11
         grd_Listad.Text = Trim(g_rst_Princi!NRO_ASIENTO_1 & "") & " - " & Trim(g_rst_Princi!NRO_ASIENTO_2 & "")
      End If

      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_GenArc(UBound(l_arr_GenArc) + 1)
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_CodVen = Trim(g_rst_Princi!regven_CodVen & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Moneda = Trim(g_rst_Princi!Moneda & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_FecCtb = Trim(g_rst_Princi!regven_FecCtb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_FecEmi = Trim(g_rst_Princi!regven_FecEmi & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_FecVto = Trim(g_rst_Princi!regven_FecVto & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_TipCpb_Lrg = Trim(g_rst_Princi!TIP_COMPROBANTE & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_TipCpb = Trim(g_rst_Princi!regven_TipCpb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Nserie = Trim(g_rst_Princi!regven_Nserie & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_NroCom = Trim(g_rst_Princi!regven_NroCom & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_TipDoc_Lrg = Trim(g_rst_Princi!TIPO_DOCUMENTO & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_TipDoc = Trim(g_rst_Princi!regven_TipDoc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_NumDoc = Trim(g_rst_Princi!regven_NumDoc & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).MaePrv_RazSoc = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Grv = g_rst_Princi!regven_Deb_Grv1 + g_rst_Princi!regven_Hab_Grv1 'Grv
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Igv = g_rst_Princi!regven_Deb_Igv1 + g_rst_Princi!regven_Hab_Igv1 'Igv
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Ngrv = g_rst_Princi!regven_Deb_Ngv1 + g_rst_Princi!regven_Hab_Ngv1 'Ngrv
                                                       
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Total = g_rst_Princi!regven_Deb_Grv1 + g_rst_Princi!regven_Hab_Grv1 + _
                                                        g_rst_Princi!regven_Deb_Igv1 + g_rst_Princi!regven_Hab_Igv1 + _
                                                        g_rst_Princi!regven_Deb_Ngv1 + g_rst_Princi!regven_Hab_Ngv1 'Total
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Deb_Grv1 = g_rst_Princi!regven_Deb_Grv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Deb_Ngv1 = g_rst_Princi!regven_Deb_Ngv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Deb_Igv1 = g_rst_Princi!regven_Deb_Igv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Deb_Ppg1 = g_rst_Princi!regven_Deb_Ppg1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Hab_Grv1 = g_rst_Princi!regven_Hab_Grv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Hab_Ngv1 = g_rst_Princi!regven_Hab_Ngv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Hab_Igv1 = g_rst_Princi!regven_Hab_Igv1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Hab_Ppg1 = g_rst_Princi!regven_Hab_Ppg1
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Cnt_Grv1 = Trim(g_rst_Princi!regven_Cnt_Grv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Cnt_Ngv1 = Trim(g_rst_Princi!regven_Cnt_Ngv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Cnt_Igv1 = Trim(g_rst_Princi!regven_Cnt_Igv1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Cnt_Ppg1 = Trim(g_rst_Princi!regven_Cnt_Ppg1 & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_CodMon = Trim(g_rst_Princi!regven_CodMon & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_TipCam = Trim(g_rst_Princi!regven_TipCam & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Ref_FecEmi = Trim(g_rst_Princi!regven_Ref_FecEmi & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Ref_TipCpb = Trim(g_rst_Princi!regven_Ref_TipCpb & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Ref_Nserie = Trim(g_rst_Princi!regven_Ref_Nserie & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Ref_NroCom = Trim(g_rst_Princi!regven_Ref_NroCom & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_FlgCnt = Trim(g_rst_Princi!regven_FlgCnt & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_Descrp = Trim(g_rst_Princi!regven_Descrp & "")
      
      l_arr_GenArc(UBound(l_arr_GenArc)).regven_retbie = g_rst_Princi!regven_retbie
      '***
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   Call grd_Listad_SelChange
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contar        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE REGISTRO DE VENTAS"
      .Range(.Cells(2, 2), .Cells(2, 14)).Merge
      .Range(.Cells(2, 2), .Cells(2, 14)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 14)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "FECHA DE EMISIÓN"
      .Cells(3, 4) = "TIPO COMPROBANTE"
      .Cells(3, 5) = "SERIE"
      .Cells(3, 6) = "NÚMERO"
      .Cells(3, 7) = "TIPO DOCUMENTO"
      .Cells(3, 8) = "DOCUMENTO"
      .Cells(3, 9) = "PROVEEDOR"
      .Cells(3, 10) = "MONEDA"
      .Cells(3, 11) = "GRAVADO"
      .Cells(3, 12) = "NO GRAVADO"
      .Cells(3, 13) = "IGV"
      .Cells(3, 14) = "TOTAL"
         
      .Range(.Cells(3, 2), .Cells(3, 14)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 14)).Font.Bold = True
       
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
      .Columns("L").ColumnWidth = 15 'no gravado 1
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 13 'igv
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 13 'total
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 14)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 14)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contar = 1 To UBound(l_arr_GenArc)
          .Cells(r_int_NumFil + 2, 2) = "'" & l_arr_GenArc(r_int_Contar).regven_CodVen     'codigo
          .Cells(r_int_NumFil + 2, 3) = "'" & gf_FormatoFecha(l_arr_GenArc(r_int_Contar).regven_FecEmi) 'fecha de emision
          .Cells(r_int_NumFil + 2, 4) = "'" & l_arr_GenArc(r_int_Contar).regven_TipCpb_Lrg 'tipo de comprobante
          .Cells(r_int_NumFil + 2, 5) = "'" & l_arr_GenArc(r_int_Contar).regven_Nserie     'serie
          .Cells(r_int_NumFil + 2, 6) = "'" & l_arr_GenArc(r_int_Contar).regven_NroCom     'numero
          .Cells(r_int_NumFil + 2, 7) = "'" & l_arr_GenArc(r_int_Contar).regven_TipDoc_Lrg 'tipo de documento
          .Cells(r_int_NumFil + 2, 8) = "'" & l_arr_GenArc(r_int_Contar).regven_NumDoc     'documento
          .Cells(r_int_NumFil + 2, 9) = "'" & l_arr_GenArc(r_int_Contar).MaePrv_RazSoc     'proveedor
          
          .Cells(r_int_NumFil + 2, 10) = "'" & l_arr_GenArc(r_int_Contar).regven_Moneda    'moneda
          
          .Cells(r_int_NumFil + 2, 11) = l_arr_GenArc(r_int_Contar).regven_Deb_Grv1 + _
                                         l_arr_GenArc(r_int_Contar).regven_Hab_Grv1 'gravado 1
                                         
          .Cells(r_int_NumFil + 2, 12) = l_arr_GenArc(r_int_Contar).regven_Deb_Ngv1 + _
                                         l_arr_GenArc(r_int_Contar).regven_Hab_Ngv1 'no gravado 1
                                         
          .Cells(r_int_NumFil + 2, 13) = l_arr_GenArc(r_int_Contar).regven_Deb_Igv1 + _
                                         l_arr_GenArc(r_int_Contar).regven_Hab_Igv1 'igv
                                         
          .Cells(r_int_NumFil + 2, 14) = .Cells(r_int_NumFil + 2, 11) + .Cells(r_int_NumFil + 2, 12) + _
                                         .Cells(r_int_NumFil + 2, 13) 'total
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 18)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
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
Dim r_int_NumAsi_2      As Integer
Dim r_int_Contar        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "K"
   r_int_NumLib = 16
   r_str_AsiGen = ""
   
   For r_int_Contar = 1 To UBound(l_arr_GenArc)
       If l_arr_GenArc(r_int_Contar).regven_FlgCnt = 0 Then
          If grd_Listad.TextMatrix(r_int_Contar - 1, 12) = "X" And _
             CLng(grd_Listad.TextMatrix(r_int_Contar - 1, 0)) = CLng(l_arr_GenArc(r_int_Contar).regven_CodVen) Then
             'Inicializa variables
             r_int_NumAsi = 0
             r_int_NumIte = 0
             r_str_FecPrPgoC = l_arr_GenArc(r_int_Contar).regven_FecCtb
             r_str_FecPrPgoL = gf_FormatoFecha(l_arr_GenArc(r_int_Contar).regven_FecCtb)
             
             r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(3, 2, l_arr_GenArc(r_int_Contar).regven_FecEmi, 1)
             
             r_str_Glosa = l_arr_GenArc(r_int_Contar).regven_NumDoc & "/" & _
                           IIf(CInt(l_arr_GenArc(r_int_Contar).regven_TipCpb) = CInt("9999"), "00", _
                           Format(CInt(l_arr_GenArc(r_int_Contar).regven_TipCpb), "00")) & "/" & _
                           Trim(l_arr_GenArc(r_int_Contar).regven_Nserie) & "/" & _
                           Trim(l_arr_GenArc(r_int_Contar).regven_NroCom) & "/" & _
                           Trim(l_arr_GenArc(r_int_Contar).regven_Descrp)
             r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
             
             r_int_PerMes = Month(gf_FormatoFecha(l_arr_GenArc(r_int_Contar).regven_FecCtb))
             r_int_PerAno = Year(gf_FormatoFecha(l_arr_GenArc(r_int_Contar).regven_FecCtb))
             
             'Obteniendo Nro. de Asiento
             r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
             r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
                
             'Insertar en CABECERA
              Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                   r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
              '--------------GRAVADO_1-------------------------
              r_dbl_Import = l_arr_GenArc(r_int_Contar).regven_Deb_Grv1 + l_arr_GenArc(r_int_Contar).regven_Hab_Grv1
              r_str_DebHab = IIf(l_arr_GenArc(r_int_Contar).regven_Deb_Grv1 > 0, "D", "H")
              If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(r_int_Contar).regven_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(r_int_Contar).regven_Cnt_Grv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
              End If
              '--------------No_GRAVADO_1-------------------------
              r_dbl_Import = l_arr_GenArc(r_int_Contar).regven_Deb_Ngv1 + l_arr_GenArc(r_int_Contar).regven_Hab_Ngv1
              r_str_DebHab = IIf(l_arr_GenArc(r_int_Contar).regven_Deb_Ngv1 > 0, "D", "H")
              If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(r_int_Contar).regven_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
         
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(r_int_Contar).regven_Cnt_Ngv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
              End If
              '--------------IGV-------------------------
              r_dbl_Import = l_arr_GenArc(r_int_Contar).regven_Deb_Igv1 + l_arr_GenArc(r_int_Contar).regven_Hab_Igv1
              r_str_DebHab = IIf(l_arr_GenArc(r_int_Contar).regven_Deb_Igv1 > 0, "D", "H")
              If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(r_int_Contar).regven_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(r_int_Contar).regven_Cnt_Igv1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
              End If
              '--------------POR PAGAR-------------------------
              r_dbl_Import = l_arr_GenArc(r_int_Contar).regven_Deb_Ppg1 + l_arr_GenArc(r_int_Contar).regven_Hab_Ppg1
              r_str_DebHab = IIf(l_arr_GenArc(r_int_Contar).regven_Deb_Ppg1 > 0, "D", "H")
              If (r_dbl_Import > 0) Then
                 r_int_NumIte = r_int_NumIte + 1
                 Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(r_int_Contar).regven_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                
                 Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, _
                      r_int_NumLib, r_int_NumAsi, r_int_NumIte, l_arr_GenArc(r_int_Contar).regven_Cnt_Ppg1, _
                      CDate(r_str_FecPrPgoL), r_str_Glosa, Trim(r_str_DebHab), r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
              End If
              
              '**********************************************************************
              r_int_NumAsi_2 = 0
              r_dbl_Import = 0
              If l_arr_GenArc(r_int_Contar).regven_retbie = 1 Then
                 r_dbl_Import = l_arr_GenArc(r_int_Contar).regven_Deb_Igv1 + l_arr_GenArc(r_int_Contar).regven_Hab_Igv1
                 If (r_dbl_Import > 0) Then
                     Call fs_Calc_Imp(r_dbl_Import, l_arr_GenArc(r_int_Contar).regven_CodMon, r_dbl_TipSbs, r_dbl_MtoDol, r_dbl_MtoSol)
                     'Obteniendo Nro. de Asiento
                     r_int_NumAsi_2 = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
                     r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi_2)
                      
                     'Insertar en CABECERA
                     Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                          r_int_NumAsi_2, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                     
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                          r_int_NumAsi_2, 1, "451301290110", CDate(r_str_FecPrPgoL), _
                          r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                                            
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                          r_int_NumAsi_2, 2, "251703020101", CDate(r_str_FecPrPgoL), _
                          r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                 End If
              End If
              
              'Actualiza flag de contabilizacion
              g_str_Parame = ""
              g_str_Parame = g_str_Parame & "UPDATE CNTBL_REGVEN "
              g_str_Parame = g_str_Parame & "   SET REGVEN_FLGCNT = 1, "
              g_str_Parame = g_str_Parame & "       REGVEN_FECCNT = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ", "
              g_str_Parame = g_str_Parame & "       REGVEN_DATCNT_1 = '" & r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
              If r_int_NumAsi_2 > 0 Then
                 g_str_Parame = g_str_Parame & "       ,REGVEN_DATCNT_2 = '" & r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi_2 & "' "
              End If
              g_str_Parame = g_str_Parame & " WHERE REGVEN_CODVEN  = '" & CLng(l_arr_GenArc(r_int_Contar).regven_CodVen) & "' "
              
              If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                 Exit Sub
              End If
          End If
       End If
   Next
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros no generados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
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
