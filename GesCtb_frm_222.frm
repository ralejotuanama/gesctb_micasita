VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CtaPag_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   Icon            =   "GesCtb_frm_222.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   16485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   40
      TabIndex        =   14
      Top             =   770
      Width           =   16390
      _Version        =   65536
      _ExtentX        =   28910
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
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   15780
         Picture         =   "GesCtb_frm_222.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Generar 
         Height          =   585
         Left            =   4830
         Picture         =   "GesCtb_frm_222.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Procesar Registro"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Agrega 
         Height          =   585
         Left            =   1230
         Picture         =   "GesCtb_frm_222.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Adicionar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Editar 
         Height          =   585
         Left            =   1830
         Picture         =   "GesCtb_frm_222.frx":0A62
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Modificar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Borrar 
         Height          =   585
         Left            =   2430
         Picture         =   "GesCtb_frm_222.frx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Eliminar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Buscar 
         Height          =   585
         Left            =   30
         Picture         =   "GesCtb_frm_222.frx":1076
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Limpia 
         Height          =   585
         Left            =   630
         Picture         =   "GesCtb_frm_222.frx":1380
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpiar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Consul 
         Height          =   585
         Left            =   3030
         Picture         =   "GesCtb_frm_222.frx":168A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consultar"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   4230
         Picture         =   "GesCtb_frm_222.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Reversa 
         Height          =   585
         Left            =   3630
         Picture         =   "GesCtb_frm_222.frx":1C9E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Reversa"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   6000
      Left            =   45
      TabIndex        =   16
      Top             =   2310
      Width           =   16395
      _Version        =   65536
      _ExtentX        =   28910
      _ExtentY        =   10583
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   14940
         TabIndex        =   17
         Top             =   60
         Width           =   1130
         _Version        =   65536
         _ExtentX        =   1993
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
            Left            =   900
            TabIndex        =   18
            Top             =   0
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grd_Listad 
         Height          =   5595
         Left            =   30
         TabIndex        =   19
         Top             =   360
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   30
         Cols            =   28
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin Threed.SSPanel pnl_Tit_DocIde 
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   60
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
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
      Begin Threed.SSPanel pnl_Tit_NumSol 
         Height          =   285
         Left            =   1110
         TabIndex        =   21
         Top             =   60
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
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
      Begin Threed.SSPanel pnl_Tit_NomCli 
         Height          =   285
         Left            =   3450
         TabIndex        =   22
         Top             =   60
         Width           =   3465
         _Version        =   65536
         _ExtentX        =   6112
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Proveedor"
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
      Begin Threed.SSPanel pnl_Tit_FecSol 
         Height          =   285
         Left            =   6900
         TabIndex        =   23
         Top             =   60
         Width           =   2595
         _Version        =   65536
         _ExtentX        =   4577
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Descripción"
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
      Begin Threed.SSPanel pnl_Tit_Produc 
         Height          =   285
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   1065
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
      Begin Threed.SSPanel pnl_Tit_SitIns 
         Height          =   285
         Left            =   10380
         TabIndex        =   25
         Top             =   60
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2205
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Total a Pagar"
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
      Begin Threed.SSPanel pnl_Tit_IngIns 
         Height          =   285
         Left            =   9480
         TabIndex        =   26
         Top             =   60
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1605
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   285
         Left            =   11610
         TabIndex        =   27
         Top             =   60
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Procesado CxP"
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
         Left            =   12780
         TabIndex        =   28
         Top             =   60
         Width           =   1090
         _Version        =   65536
         _ExtentX        =   1923
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Fecha Pago"
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
         Left            =   13860
         TabIndex        =   29
         Top             =   60
         Width           =   1090
         _Version        =   65536
         _ExtentX        =   1923
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Código Pago"
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
      Left            =   40
      TabIndex        =   30
      Top             =   60
      Width           =   16395
      _Version        =   65536
      _ExtentX        =   28910
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
      Begin Threed.SSPanel pnl_TitPri 
         Height          =   315
         Left            =   750
         TabIndex        =   31
         Top             =   180
         Width           =   8565
         _Version        =   65536
         _ExtentX        =   15108
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Registros de Cuentas por Pagar"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
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
         Left            =   90
         Picture         =   "GesCtb_frm_222.frx":1FA8
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   8450
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   16600
      _Version        =   65536
      _ExtentX        =   29281
      _ExtentY        =   14905
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   40
         TabIndex        =   33
         Top             =   1440
         Width           =   16390
         _Version        =   65536
         _ExtentX        =   28910
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8160
            TabIndex        =   4
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
            TabIndex        =   2
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
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   5310
            TabIndex        =   37
            Top             =   450
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   450
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   35
            Top             =   120
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   34
            Top             =   120
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CtaPag_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 7)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 10) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 7)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 10) = "X"
             End If
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
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
   moddat_g_int_FlgGrb = 1
   frm_Ctb_CtaPag_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 7
   If UCase(Trim(grd_Listad.Text)) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar el registro, esta procesado.", vbExclamation, modgen_g_str_NomPlt
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
   g_str_Parame = g_str_Parame & " USP_CNTBL_CTAPAG_BORRAR ( "
   g_str_Parame = g_str_Parame & CLng(moddat_g_str_Codigo) & ", " 'CTAPAG_CODPAG
   g_str_Parame = g_str_Parame & "2, " 'CTAPAG_TIPTAB
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro de cuenta por pagar se eliminó correctamente.", vbInformation, modgen_g_str_NomPlt
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
   grd_Listad.Col = 19
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 20
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 0
   frm_Ctb_CtaPag_02.Show 1
   
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
   grd_Listad.Col = 19
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 20
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   grd_Listad.Col = 7
   If (UCase(grd_Listad.Text) = "SI") Then
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "No se pudo editar el registro, esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_CtaPag_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
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
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 19
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 20
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   'procesado por Compensasion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL((SELECT COMAUT_CODEST FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "              Where A.COMAUT_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODEST IN (1,2,4,5)  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODOPE = " & CLng(moddat_g_str_Codigo) & ")  "
   g_str_Parame = g_str_Parame & "           ,0) AS CODEST,  "
   g_str_Parame = g_str_Parame & "        (SELECT CTAPAG_FLGCTB FROM CNTBL_CTAPAG B  "
   g_str_Parame = g_str_Parame & "          WHERE B.CTAPAG_CODPAG = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "            AND B.CTAPAG_TIPTAB = 2)  "
   g_str_Parame = g_str_Parame & "          AS CXP_CTB "
   g_str_Parame = g_str_Parame & "   FROM DUAL  "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If g_rst_Princi!CODEST <> 0 Then
      Select Case g_rst_Princi!CODEST
             Case 1: MsgBox "El registro se encuentra como pendiente en modulo de compensación, no se puede revertir.", vbExclamation, modgen_g_str_NomPlt
             Case 2: MsgBox "El registro se encuentra como aprobado en modulo de compensación, no se puede revertir.", vbExclamation, modgen_g_str_NomPlt
             Case 4: MsgBox "El registro se encuentra como aplicado en modulo de compensación, no se puede revertir.", vbExclamation, modgen_g_str_NomPlt
             Case 5: MsgBox "El registro se encuentra como pagado en modulo de compensación, no se puede revertir.", vbExclamation, modgen_g_str_NomPlt
      End Select
      Exit Sub
   Else
      'procesado por CxP
      If CInt(g_rst_Princi!CXP_CTB) = 0 Then
         MsgBox "Solo se puede dar reversa a los registros que hayan sido procesados.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
       
      moddat_g_int_FlgGrb = 3
      Call gs_RefrescaGrid(grd_Listad)
      frm_Ctb_CtaPag_02.Show 1
      
      Call gs_SetFocus(grd_Listad)
   End If
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
Dim r_bol_EstTca        As Boolean
Dim r_str_CtaDeb        As String
Dim r_str_CtaHab        As String
Dim r_str_CadDin        As String
Dim r_str_CadTca        As String
Dim r_bol_Estado        As Boolean
Dim r_dbl_TipSbs        As Double
Dim r_int_Contar        As Integer

   r_bol_Estado = False
   For r_int_Contar = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contar, 7) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contar, 10) = "X" Then
             r_bol_Estado = True
             Exit For
          End If
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No se ha seleccionado ningún registro.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'validar existencia de la dinamica
   r_str_CadDin = ""
   r_str_CadTca = ""
   r_dbl_TipSbs = 0
   r_bol_EstTca = True

   For r_int_Contar = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contar, 7) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contar, 10) = "X" Then
             'valida que todos tengan tipo cambio SBS(2) - VENTA(1)
             r_dbl_TipSbs = 0
             r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(grd_Listad.TextMatrix(r_int_Contar, 1), "yyyymmdd"), 1)
             If r_dbl_TipSbs = 0 Then
                r_bol_EstTca = False
                r_str_CadTca = r_str_CadTca & " - " & Trim(grd_Listad.TextMatrix(r_int_Contar, 0))
             End If
          End If
       End If
   Next
      
   If r_bol_EstTca = False Then
      MsgBox "Falta definir el tipo de cambio SBS de los siguientes registros" & vbCrLf & "Código: " & r_str_CadTca, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro que desea procesar los registros seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiento
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub fs_GeneraAsiento()
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumAsi    As Integer
Dim r_int_NumIte    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
Dim r_str_AsiGen    As String
Dim r_int_NumAux    As Integer
Dim r_str_CadAux    As String
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_Contar    As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   
   For r_int_Contar = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contar, 7) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contar, 10) = "X" Then
             r_int_NumAsi = 0
             r_int_NumIte = 0
             r_str_FechaC = Format(grd_Listad.TextMatrix(r_int_Contar, 1), "yyyymmdd")
             r_str_FechaL = grd_Listad.TextMatrix(r_int_Contar, 1) 'FECHA
             
             'r_int_PerAno = modctb_int_PerAno 'Year(grd_Listad.TextMatrix(r_int_Contar, 1)) 'FECHA
             'r_int_PerMes = modctb_int_PerMes 'Month(grd_Listad.TextMatrix(r_int_Contar, 1)) 'FECHA
             r_int_PerAno = Year(grd_Listad.TextMatrix(r_int_Contar, 1)) 'FECHA
             r_int_PerMes = Month(grd_Listad.TextMatrix(r_int_Contar, 1)) 'FECHA
             
             'Obteniendo Nro. de Asiento
             r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
             r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
            
             'TIPO CAMBIO SBS(2) - VENTA(1)
             r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, r_str_FechaC, 1)
             r_str_Glosa = ""
             r_str_Glosa = Trim(CStr(grd_Listad.TextMatrix(r_int_Contar, 2))) & "/" & Trim(CStr(grd_Listad.TextMatrix(r_int_Contar, 4)))  'TIPO_OPERACION/DESCRIPCION
             r_str_Glosa = Mid("PAGO " & r_str_Glosa, 1, 60) 'NRO_DOC/DESCRIPCION
            
             'Insertar en CABECERA
             Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                           r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FechaL, "1")
                                                                          
             'Insertar en DETALLES - IMPORTE TOTAL - DEBE
             r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
             r_int_NumIte = 0: r_str_CtaDeb = ""
             If CDbl(grd_Listad.TextMatrix(r_int_Contar, 12)) > 0 Then
                Call fs_Convertir(grd_Listad.TextMatrix(r_int_Contar, 18), r_dbl_TipSbs, grd_Listad.TextMatrix(r_int_Contar, 12), r_dbl_MtoSol, r_dbl_MtoDol)
                r_str_CtaDeb = Trim(grd_Listad.TextMatrix(r_int_Contar, 24)) 'nro cuenta
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                                     r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                                     r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
             End If
             r_dbl_MtoSol = 0: r_dbl_MtoDol = 0: r_str_CtaDeb = ""
             If CDbl(grd_Listad.TextMatrix(r_int_Contar, 13)) > 0 Then
                Call fs_Convertir(grd_Listad.TextMatrix(r_int_Contar, 18), r_dbl_TipSbs, grd_Listad.TextMatrix(r_int_Contar, 13), r_dbl_MtoSol, r_dbl_MtoDol)
                r_str_CtaDeb = Trim(grd_Listad.TextMatrix(r_int_Contar, 25)) 'nro cuenta
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                                     r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                                     r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
             End If
             r_dbl_MtoSol = 0: r_dbl_MtoDol = 0: r_str_CtaDeb = ""
             If CDbl(grd_Listad.TextMatrix(r_int_Contar, 14)) > 0 Then
                Call fs_Convertir(grd_Listad.TextMatrix(r_int_Contar, 18), r_dbl_TipSbs, grd_Listad.TextMatrix(r_int_Contar, 14), r_dbl_MtoSol, r_dbl_MtoDol)
                r_str_CtaDeb = Trim(grd_Listad.TextMatrix(r_int_Contar, 26)) 'nro cuenta
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                                     r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                                     r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
             End If
             r_dbl_MtoSol = 0: r_dbl_MtoDol = 0: r_str_CtaDeb = ""
             If CDbl(grd_Listad.TextMatrix(r_int_Contar, 15)) > 0 Then
                Call fs_Convertir(grd_Listad.TextMatrix(r_int_Contar, 18), r_dbl_TipSbs, grd_Listad.TextMatrix(r_int_Contar, 15), r_dbl_MtoSol, r_dbl_MtoDol)
                r_str_CtaDeb = Trim(grd_Listad.TextMatrix(r_int_Contar, 27)) 'nro cuenta
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                                     r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                                     r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
             End If
                                      
             r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
             r_str_CtaHab = ""
             If CDbl(grd_Listad.TextMatrix(r_int_Contar, 11)) > 0 Then
                Call fs_Convertir(grd_Listad.TextMatrix(r_int_Contar, 18), r_dbl_TipSbs, grd_Listad.TextMatrix(r_int_Contar, 11), r_dbl_MtoSol, r_dbl_MtoDol)
                If CInt(grd_Listad.TextMatrix(r_int_Contar, 18)) = 1 Then
                   r_str_CtaHab = "251419010109" 'soles
                Else
                   r_str_CtaHab = "252419010109" 'dolares
                End If
                r_int_NumIte = r_int_NumIte + 1
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                                     r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FechaL), _
                                                     r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
             End If
            
             r_str_CadAux = ""
             r_str_CadAux = r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
             'Actualiza flag de contabilizacion
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " UPDATE CNTBL_CTAPAG  "
             g_str_Parame = g_str_Parame & "    SET CTAPAG_DATCTB = '" & r_str_CadAux & "',  "
             g_str_Parame = g_str_Parame & "        CTAPAG_FLGCTB = 1 ,  "
             g_str_Parame = g_str_Parame & "        CTAPAG_FECCTB = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ",  "
             g_str_Parame = g_str_Parame & "        CTAPAG_TIPCAM = " & CDbl(r_dbl_TipSbs)
             g_str_Parame = g_str_Parame & "  WHERE CTAPAG_CODPAG = " & CLng(grd_Listad.TextMatrix(r_int_Contar, 0))
             g_str_Parame = g_str_Parame & "    AND CTAPAG_TIPTAB = 2  "
             
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                Exit Sub
             End If
             
             'Enviar a la tabla de autorizaciones
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
             g_str_Parame = g_str_Parame & " " & CLng(grd_Listad.TextMatrix(r_int_Contar, 0)) & ", " 'COMAUT_CODOPE
             g_str_Parame = g_str_Parame & " " & Format(grd_Listad.TextMatrix(r_int_Contar, 1), "yyyymmdd") & ", " 'COMAUT_FECOPE
             g_str_Parame = g_str_Parame & " " & grd_Listad.TextMatrix(r_int_Contar, 19) & ", "      'COMAUT_TIPDOC
             g_str_Parame = g_str_Parame & " '" & grd_Listad.TextMatrix(r_int_Contar, 20) & "', "    'COMAUT_NUMDOC
             g_str_Parame = g_str_Parame & " " & grd_Listad.TextMatrix(r_int_Contar, 18) & ", "      'COMAUT_CODMON
             g_str_Parame = g_str_Parame & " " & CDbl(grd_Listad.TextMatrix(r_int_Contar, 6)) & ", " 'COMAUT_IMPPAG
             g_str_Parame = g_str_Parame & " " & grd_Listad.TextMatrix(r_int_Contar, 21) & ", "  'COMAUT_CODBNC
             g_str_Parame = g_str_Parame & " '" & grd_Listad.TextMatrix(r_int_Contar, 22) & "', "  'COMAUT_CTACRR
             g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB
             g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
             g_str_Parame = g_str_Parame & " '" & Trim(grd_Listad.TextMatrix(r_int_Contar, 4)) & "',  " 'COMAUT_DESCRIPCION
             g_str_Parame = g_str_Parame & " " & 1 & ",  " 'COMAUT_TIPOPE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "  'SEGUSUCRE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "  'SEGPLTCRE
             g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "  'SEGSUCCRE
                                                                                                                                                                                                                             
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
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

Private Sub fs_BuscarCtas(p_TipOpe As Integer, p_CodMon As Integer, ByRef p_CtaDeb As String, ByRef p_CtaHab As String)
   p_CtaDeb = ""
   p_CtaHab = ""
   If p_TipOpe = 1 And p_CodMon = 1 Then 'SEGURO DESGRAVAMEN - PEN
      p_CtaDeb = "251602010103"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 1 And p_CodMon = 2 Then 'SEGURO DESGRAVAMEN - USD
      p_CtaDeb = "252602010103"
      p_CtaHab = "252419010109"
   ElseIf p_TipOpe = 2 And p_CodMon = 1 Then 'SEGURO INCENDIO - PEN
      p_CtaDeb = "251602010104"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 2 And p_CodMon = 2 Then 'SEGURO INCENDIO - USD
      p_CtaDeb = "252602010104"
      p_CtaHab = "252419010109"
   ElseIf p_TipOpe = 3 And p_CodMon = 1 Then 'PREPAGOS COFIDE - PEN
      p_CtaDeb = "191807010101"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 4 And p_CodMon = 1 Then 'COMISION FMV - PEN
      p_CtaDeb = "191807020101"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 4 And p_CodMon = 2 Then 'COMISION FMV - USD
      p_CtaDeb = "192807020101"
      p_CtaHab = "252419010109"
   ElseIf p_TipOpe = 5 And p_CodMon = 1 Then 'DEVOLUCION EXCEDENTE AFP <<CAMBIO>> DEVOLUCION GASTOS DE CIERRE - PEN
      p_CtaDeb = "291807010114" '"291807010112"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 6 And p_CodMon = 1 Then 'DEVOLUCION PLAN AHORRO - PEN
      p_CtaDeb = "291807010111"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 7 And p_CodMon = 1 Then 'PAGO MENSUAL COFIDE - PEN
      p_CtaDeb = "251419010111" '"191807010101"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 8 And p_CodMon = 1 Then 'DEVOLUCION EXCEDENTE PREPAGOS - PEN
      p_CtaDeb = "111301060201"
      p_CtaHab = "251419010109"
   ElseIf p_TipOpe = 8 And p_CodMon = 2 Then 'DEVOLUCION EXCEDENTE PREPAGOS - USD
      p_CtaDeb = "112301060202"
      p_CtaHab = "252419010109"
   ElseIf p_TipOpe = 9 And p_CodMon = 1 Then 'DEVOLUCION TOTAL DE AFP - PEN
      p_CtaDeb = "291807010114"
      p_CtaHab = "251419010109"
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   grd_Listad.ColWidth(0) = 1050 'CODIGO
   grd_Listad.ColWidth(1) = 1050 'FECHA
   grd_Listad.ColWidth(2) = 1290 'NRO DOCUMENTO
   grd_Listad.ColWidth(3) = 3450 'PROVEEDOR
   grd_Listad.ColWidth(4) = 2580 'TIPO OPERACION
   grd_Listad.ColWidth(5) = 900  'MONEDA
   grd_Listad.ColWidth(6) = 1230 'TOTAL A PAGAR
   grd_Listad.ColWidth(7) = 1190 'PROCESADO CXP
   grd_Listad.ColWidth(8) = 1060 'FECHA PAGO
   grd_Listad.ColWidth(9) = 1080 'CODIGO PAGO
   
   grd_Listad.ColWidth(10) = 1120 'SELECCION
   grd_Listad.ColWidth(11) = 0   'CTAPAG_IMPPAG
   grd_Listad.ColWidth(12) = 0   'CTAPAG_IMPPG1
   grd_Listad.ColWidth(13) = 0   'CTAPAG_IMPPG2
   grd_Listad.ColWidth(14) = 0   'CTAPAG_IMPPG3
   grd_Listad.ColWidth(15) = 0   'CTAPAG_IMPPG4
   grd_Listad.ColWidth(16) = 0   'CTAPAG_TIPCAM
   grd_Listad.ColWidth(17) = 0   'CTAPAG_TIPOPE
   grd_Listad.ColWidth(18) = 0   'CTAPAG_CODMON
   grd_Listad.ColWidth(19) = 0   'CTAPAG_TIPDOC
   grd_Listad.ColWidth(20) = 0   'CTAPAG_NUMDOC
   grd_Listad.ColWidth(21) = 0   'CTAPAG_CODBCO
   grd_Listad.ColWidth(22) = 0   'CTAPAG_CTACRR
   grd_Listad.ColWidth(23) = 0   'DESCRIPCION
   grd_Listad.ColWidth(24) = 0   'CTAPAG_NUCTA1
   grd_Listad.ColWidth(25) = 0   'CTAPAG_NUCTA2
   grd_Listad.ColWidth(26) = 0   'CTAPAG_NUCTA3
   grd_Listad.ColWidth(27) = 0   'CTAPAG_NUCTA4
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
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

Public Sub fs_Buscar()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CTAPAG_CODPAG, A.CTAPAG_FECOPE, A.CTAPAG_TIPDOC, A.CTAPAG_NUMDOC, "
   g_str_Parame = g_str_Parame & "       TRIM(B.MAEPRV_RAZSOC) AS MAEPRV_RAZSOC, TRIM(D.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "       A.CTAPAG_FLGCTB , A.CTAPAG_FLGCOM, A.CTAPAG_TIPOPE, A.CTAPAG_CODMON, CTAPAG_TIPCAM,  "
   g_str_Parame = g_str_Parame & "       CTAPAG_CODBCO, CTAPAG_CTACRR, F.COMPAG_FECPAG, F.COMPAG_CODCOM, A.CTAPAG_DESCRP,  "
   g_str_Parame = g_str_Parame & "       NVL(CTAPAG_IMPPG1,0) + NVL(CTAPAG_IMPPG2,0) + NVL(CTAPAG_IMPPG3,0) + NVL(CTAPAG_IMPPG4,0) AS IMPORTE,"
   g_str_Parame = g_str_Parame & "       CTAPAG_IMPPG1, CTAPAG_IMPPG2, CTAPAG_IMPPG3, CTAPAG_IMPPG4, "
   g_str_Parame = g_str_Parame & "       CTAPAG_NUCTA1, CTAPAG_NUCTA2, CTAPAG_NUCTA3, CTAPAG_NUCTA4  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_CTAPAG A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CTAPAG_TIPDOC AND B.MAEPRV_NUMDOC = A.CTAPAG_NUMDOC "
   g_str_Parame = g_str_Parame & "   AND A.CTAPAG_TIPTAB = 2 "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.CTAPAG_CODMON  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_COMDET E ON E.COMDET_CODOPE = A.CTAPAG_CODPAG AND E.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_COMPAG F ON F.COMPAG_CODCOM = E.COMDET_CODCOM AND F.COMPAG_SITUAC = 1 AND F.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & " WHERE A.CTAPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   AND A.CTAPAG_FECOPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   'g_str_Parame = g_str_Parame & "   AND A.CTAPAG_TIPTAB = 1  "
   g_str_Parame = g_str_Parame & " ORDER BY CTAPAG_CODPAG, CTAPAG_FECOPE ASC  "
   
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
      grd_Listad.Text = Format(CStr(g_rst_Princi!CTAPAG_CODPAG), "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CTAPAG_FECOPE)

      grd_Listad.Col = 2
      grd_Listad.Text = g_rst_Princi!CTAPAG_TIPDOC & "-" & Trim(g_rst_Princi!CTAPAG_NUMDOC & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_Listad.Col = 4 'CStr(g_rst_Princi!TIPOPERACION & "")
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_DESCRP & "")
                  
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 6 'IMPORTE A PAGAR
      grd_Listad.Text = Format(g_rst_Princi!IMPORTE, "###,###,###,##0.00")
            
      grd_Listad.Col = 7
      grd_Listad.Text = IIf(g_rst_Princi!CTAPAG_FLGCTB = 1, "SI", "NO")
      
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 8 'IIf(g_rst_Princi!CTAPAG_FLGCOM = 1, "SI", "NO")
         grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_Listad.Col = 9
         grd_Listad.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
      
      'CTAPAG_IMPPG1, CTAPAG_IMPPG2, CTAPAG_IMPPG3, CTAPAG_IMPPG4, "
      
      grd_Listad.Col = 11
      grd_Listad.Text = g_rst_Princi!IMPORTE
      'importes
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!CTAPAG_IMPPG1
      grd_Listad.Col = 13
      grd_Listad.Text = g_rst_Princi!CTAPAG_IMPPG2
      grd_Listad.Col = 14
      grd_Listad.Text = g_rst_Princi!CTAPAG_IMPPG3
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!CTAPAG_IMPPG4
      
      grd_Listad.Col = 16 'se guarda cuando se genera el asiento
      grd_Listad.Text = g_rst_Princi!CTAPAG_TIPCAM
            
      grd_Listad.Col = 17
      grd_Listad.Text = "" 'g_rst_Princi!CTAPAG_TIPOPE
      grd_Listad.Col = 18
      grd_Listad.Text = g_rst_Princi!CTAPAG_CODMON
      grd_Listad.Col = 19
      grd_Listad.Text = g_rst_Princi!CTAPAG_TIPDOC
      grd_Listad.Col = 20
      grd_Listad.Text = g_rst_Princi!CTAPAG_NUMDOC
      
      grd_Listad.Col = 21
      grd_Listad.Text = g_rst_Princi!CTAPAG_CODBCO
      grd_Listad.Col = 22
      grd_Listad.Text = g_rst_Princi!CTAPAG_CTACRR
      grd_Listad.Col = 23
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_DESCRP & "")
      
      grd_Listad.Col = 24
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_NUCTA1 & "")
      grd_Listad.Col = 25
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_NUCTA2 & "")
      grd_Listad.Col = 26
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_NUCTA3 & "")
      grd_Listad.Col = 27
      grd_Listad.Text = Trim(g_rst_Princi!CTAPAG_NUCTA4 & "")
   
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
Dim r_int_Contar        As Integer
Dim r_int_Fila          As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE CUENTAS POR PAGAR"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter

      r_int_NumFil = 4
      .Cells(r_int_NumFil, 2) = "CÓDIGO"
      .Cells(r_int_NumFil, 3) = "FECHA"
      .Cells(r_int_NumFil, 4) = "NRO DOCUMENTO"
      .Cells(r_int_NumFil, 5) = "PROVEEDOR"
      .Cells(r_int_NumFil, 6) = "DESCRIPCION"
      .Cells(r_int_NumFil, 7) = "MONEDA"
      .Cells(r_int_NumFil, 8) = "TOTAL A PAGAR"
      .Cells(r_int_NumFil, 9) = "PROCESADO CxP"
      .Cells(r_int_NumFil, 10) = "FECHA PAGO"
         
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12 'fecha
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18 'nro documento
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 45 'proveedor
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 30 'tipo operacion
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 8 'moneda
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 17 'total a pagar
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 15 'Procesado CxP
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15 'FECHA PAGO
      .Columns("J").HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 5
      For r_int_Fila = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(r_int_Fila, 0)
         .Cells(r_int_NumFil, 3) = "'" & grd_Listad.TextMatrix(r_int_Fila, 1)
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(r_int_Fila, 2)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(r_int_Fila, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(r_int_Fila, 4)
         .Cells(r_int_NumFil, 7) = grd_Listad.TextMatrix(r_int_Fila, 5)
         .Cells(r_int_NumFil, 8) = grd_Listad.TextMatrix(r_int_Fila, 6)
         .Cells(r_int_NumFil, 9) = grd_Listad.TextMatrix(r_int_Fila, 7)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(r_int_Fila, 8)
         
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(r_int_NumFil, 3), .Cells(r_int_NumFil, 10)).HorizontalAlignment = xlHAlignCenter
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 7
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 10
         If grd_Listad.Text = "X" Then
             grd_Listad.Text = ""
         Else
              grd_Listad.Text = "X"
         End If
      End If
      Call gs_RefrescaGrid(grd_Listad)
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
