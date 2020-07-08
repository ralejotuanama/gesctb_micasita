VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptSun_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17670
   Icon            =   "GesCtb_frm_929.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   17670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17655
      _Version        =   65536
      _ExtentX        =   31141
      _ExtentY        =   16880
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
         Height          =   645
         Left            =   60
         TabIndex        =   1
         Top             =   810
         Width           =   17520
         _Version        =   65536
         _ExtentX        =   30903
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
         Begin VB.CommandButton cmd_EnvMail 
            Height          =   585
            Left            =   3480
            Picture         =   "GesCtb_frm_929.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Envio de Correo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   2310
            Picture         =   "GesCtb_frm_929.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Exportar documentos Electrónicos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_AbrArc 
            Height          =   585
            Left            =   2890
            Picture         =   "GesCtb_frm_929.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1740
            Picture         =   "GesCtb_frm_929.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_GenNCr 
            Height          =   585
            Left            =   1170
            Picture         =   "GesCtb_frm_929.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Generar Nota Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   600
            Picture         =   "GesCtb_frm_929.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16905
            Picture         =   "GesCtb_frm_929.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_929.frx":18FA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   840
         Left            =   60
         TabIndex        =   4
         Top             =   1500
         Width           =   17520
         _Version        =   65536
         _ExtentX        =   30903
         _ExtentY        =   1482
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
         Begin VB.CheckBox Chk_CorEnv 
            Caption         =   "Correo Enviado"
            Height          =   255
            Left            =   5400
            TabIndex        =   37
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cmb_TipRsp 
            Height          =   315
            Left            =   5370
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   90
            Width           =   2280
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1575
            TabIndex        =   6
            Top             =   90
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            Left            =   1575
            TabIndex        =   7
            Top             =   420
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin:"
            Height          =   195
            Left            =   135
            TabIndex        =   10
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio:"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Criterio:"
            Height          =   195
            Left            =   4170
            TabIndex        =   8
            Top             =   150
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   17520
         _Version        =   65536
         _ExtentX        =   30903
         _ExtentY        =   1244
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   660
            TabIndex        =   12
            Top             =   60
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   660
            TabIndex        =   13
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Facturas Electrónicas"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   16890
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   16320
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_929.frx":1C04
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7095
         Left            =   60
         TabIndex        =   14
         Top             =   2385
         Width           =   17520
         _Version        =   65536
         _ExtentX        =   30903
         _ExtentY        =   12515
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   6600
            Width           =   2595
         End
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5700
            MaxLength       =   100
            TabIndex        =   30
            Top             =   6630
            Width           =   6975
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6060
            Left            =   45
            TabIndex        =   15
            Top             =   405
            Width           =   17430
            _ExtentX        =   30745
            _ExtentY        =   10689
            _Version        =   393216
            Rows            =   26
            Cols            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   315
            Left            =   7710
            TabIndex        =   17
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel pnl_Tit_FecPro 
            Height          =   315
            Left            =   10155
            TabIndex        =   18
            Top             =   90
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "F. Proc."
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
         Begin Threed.SSPanel pnl_Tit_DoiCli 
            Height          =   315
            Left            =   1290
            TabIndex        =   19
            Top             =   90
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
            Height          =   315
            Left            =   2565
            TabIndex        =   20
            Top             =   90
            Width           =   4290
            _Version        =   65536
            _ExtentX        =   7567
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
         Begin Threed.SSPanel pnl_Tit_FecEmi 
            Height          =   315
            Left            =   9060
            TabIndex        =   21
            Top             =   90
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "F. Emisión"
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
         Begin Threed.SSPanel pnl_Seleccionar 
            Height          =   315
            Left            =   13560
            TabIndex        =   22
            Top             =   90
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1244
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   " Sel."
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
               Left            =   390
               TabIndex        =   23
               Top             =   15
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_Refer 
            Height          =   315
            Left            =   12330
            TabIndex        =   27
            Top             =   90
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Refer."
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
         Begin Threed.SSPanel pnl_Tit_Moneda 
            Height          =   315
            Left            =   6840
            TabIndex        =   28
            Top             =   90
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
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
         Begin Threed.SSPanel pnl_Tit_FecPag 
            Height          =   315
            Left            =   11250
            TabIndex        =   35
            Top             =   90
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "F. Pago"
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
         Begin Threed.SSPanel pnl_Tit_Correo 
            Height          =   315
            Left            =   14250
            TabIndex        =   38
            Top             =   90
            Width           =   2910
            _Version        =   65536
            _ExtentX        =   5133
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Correo"
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4710
            TabIndex        =   33
            Top             =   6690
            Width           =   825
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   6690
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents r_chi_sftp  As ChilkatSFtp
Attribute r_chi_sftp.VB_VarHelpID = -1
Dim l_rst_FacEle           As ADODB.Recordset
Dim l_str_RutaLg           As String
Dim l_str_RutFacEnt        As String
Dim l_str_RutaLgEma        As String
Dim l_str_NomLOG           As String
Dim l_int_NumLOG           As Integer
Dim l_fsobj                As Scripting.FileSystemObject
Dim l_txtStr               As TextStream
Dim l_str_FecCar           As String
Dim l_str_RutaArc          As String
Dim l_str_Percon_mail      As String

Private Sub cmb_Buscar_Click()
   If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
      txt_Buscar.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   Else
      txt_Buscar.Enabled = True
      Call gs_SetFocus(txt_Buscar)
   End If
   txt_Buscar.Text = ""
End Sub

Private Sub cmd_AbrArc_Click()
Dim r_str_NumFac     As String
Dim r_str_NotCre     As String
Dim r_str_TipDoc     As String
   
   If grd_Listad.Row = -1 Then Exit Sub
   
   r_str_NotCre = grd_Listad.TextMatrix(grd_Listad.Row, 8)
   r_str_NumFac = grd_Listad.TextMatrix(grd_Listad.Row, 16)
   
   If r_str_NotCre = "" Then
      If InStr(r_str_NumFac, "B") > 0 Then
         r_str_TipDoc = "03"
      Else
         r_str_TipDoc = "01"
      End If
   Else
      r_str_TipDoc = "07"
   End If
   
   ShellExecute Me.hwnd, "open", moddat_g_str_RutFac & "\reportes\aceptados\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf", "", "", 4
   
End Sub

Private Sub cmd_Buscar_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   If cmb_TipRsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de respuesta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRsp)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_EnvMail_Click()
Dim r_str_NomPdf     As String
Dim r_str_NomXml     As String
Dim r_str_Cadena     As String
Dim r_int_Contad     As Integer
Dim r_int_ConSel     As Integer
Dim r_str_NotCre     As String
Dim r_str_NumFac     As String
Dim r_str_NomCli     As String
Dim r_str_TipDoc     As String
Dim r_str_NumOpe     As String
Dim r_str_NumDoc     As String
Dim r_str_RutFic     As String
Dim r_str_RLgEma     As String
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado documentos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_str_NomPdf = ""
   r_str_NomXml = ""
   l_str_Percon_mail = ""
   
   If MsgBox("¿Está seguro de enviar el correo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      r_str_NotCre = ""
      r_str_NumFac = ""
      r_str_NomCli = ""
      r_str_NumOpe = ""
      r_str_NumDoc = ""
      r_str_Cadena = ""
      r_str_NomPdf = ""
      r_str_NomXml = ""
      r_str_RutFic = moddat_g_str_RutFac & "\reportes\aceptados\"
      
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
      
         r_str_NotCre = grd_Listad.TextMatrix(r_int_Contad, 8)                      'Número de Nota de Crédito
         r_str_NumFac = grd_Listad.TextMatrix(r_int_Contad, 16)                     'Número de Factura
         r_str_NomCli = grd_Listad.TextMatrix(r_int_Contad, 2)                      'Nombre del Cliente
         r_str_NumOpe = Replace(grd_Listad.TextMatrix(r_int_Contad, 0), "-", "")    'Número de Operación
         r_str_NumDoc = CStr(grd_Listad.TextMatrix(r_int_Contad, 1))                'Número del Documento de Identidad del Cluente
         
         'Determina Tipo Documento Electrónico
         If r_str_NotCre = "" Then
            If InStr(r_str_NumFac, "B") > 0 Then
               r_str_TipDoc = "03"
            Else
               r_str_TipDoc = "01"
            End If
         Else
            r_str_TipDoc = "07"
         End If
         DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents

         'Obtenemos el Correo del Cliente seleccionado
         l_str_Percon_mail = fs_ObtCor(r_str_TipDoc, r_str_NumFac, r_str_NumOpe)
         DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
         
         'Ruta y nombre del archico PDF y XML
         Set l_fsobj = New FileSystemObject
         
         r_str_NomPdf = "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf"
         r_str_NomXml = "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".xml"
                   
         If gf_Existe_Archivo(r_str_RutFic, r_str_NomPdf) And gf_Existe_Archivo(r_str_RutFic, r_str_NomXml) Then
            
            If Trim(r_str_NomPdf) <> "" And Trim(r_str_NomXml) <> "" And l_str_Percon_mail <> "" And r_str_TipDoc <> "07" Then
               
               'Ruta para almacenar los archivos generados de email enviados
               'Crear Archivo LOG del Proceso
               l_str_NomLOG = UCase(App.EXEName) & "_" & r_str_NumFac & "_" & Format(date, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".LOG"
               l_int_NumLOG = FreeFile
                              
               r_str_RLgEma = Replace(moddat_g_str_RutFac, "\Fact", "\Logs\EMAIL_CONFORMIDAD\")
               
               'Crear la Carpeta email
               Set l_fsobj = New FileSystemObject
               If l_fsobj.FolderExists(r_str_RLgEma) = False Then
                  l_fsobj.CreateFolder (r_str_RLgEma)
               End If
               
               If gf_Existe_Archivo(r_str_RLgEma & "\", l_str_NomLOG) Then
                  Kill r_str_RLgEma & "\" & l_str_NomLOG
                  DoEvents
               End If
               
               l_str_RutaLgEma = r_str_RLgEma & "\" & l_str_NomLOG
               Open l_str_RutaLgEma For Output As l_int_NumLOG
               Close #l_int_NumLOG
               
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Proceso           : " & modgen_g_str_NomPlt)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Nombre Ejecutable : " & UCase(App.EXEName))
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Número Revisión   : " & modgen_g_str_NumRev)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Nombre PC         : " & modgen_g_str_NombPC)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Origen Datos      : " & moddat_g_str_NomEsq & " - " & moddat_g_str_EntDat)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Inicio Proceso    : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")
                              
               
               'Enviando Correo Electrónico
               'modgen_g_str_Mail_Asunto = "ENVIO DE FACTURA ELECTRONICA (Cliente: " & CStr(r_str_NumDoc) & " - " & r_str_NomCli & ")"
'               modgen_g_str_Mail_Mensaj = ""
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE OPERACION" & vbTab & " : " & vbTab & r_str_NumOpe & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE         " & vbTab & vbTab & vbTab & " : " & vbTab & CStr(r_str_NumDoc) & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE     " & vbTab & vbTab & " : " & vbTab & r_str_NomCli & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA              " & vbTab & vbTab & vbTab & " : " & vbTab & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA               " & vbTab & vbTab & vbTab & " : " & vbTab & Format(Time, "hh:mm:ss") & Chr(13)
'
               modgen_g_str_Mail_Asunto = "DOCUMENTO ELECTRONICO - EDPYME MICASITA S.A. [" & IIf(r_str_TipDoc = "03", "BOLETA", "FACTURA") & "] : " & r_str_NumFac
               modgen_g_str_Mail_Mensaj = ""
               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ESTIMADO CLIENTE: " & Chr(13)
               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & r_str_NomCli & Chr(13)
               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "" & Chr(13)
               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "Por medio de la presente, le informamos que le hemos emitido el siguiente documento electrónico:" & Chr(13)
               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "[" & IIf(r_str_TipDoc = "03", "BOLETA", "FACTURA") & "]: " & r_str_NumFac

               r_str_Cadena = l_str_Percon_mail
               
               'Destinatarios de Correo
               ReDim moddat_g_arr_Genera(0)

               If (Len(Trim(r_str_Cadena)) > 0) Then
                   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
               End If
               DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
               
               Call fs_EnvCor(mps_Sesion, mps_Mensaj, r_str_NumFac, r_str_NumOpe, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, r_str_NomPdf, r_str_NomXml, r_str_RutFic)
               
               DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
               
               'Actualiza Flag de Envío de Correo
               modprc_g_str_CadEje = ""
               modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CNTBL_DOCELE SET DOCELE_ENVCOR = 1 "
               modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DOCELE_IDE_SERNUM = '" & CStr(r_str_NumFac) & "' "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DOCELE_NUMOPE = '" & CStr(r_str_NumOpe) & "' "
               
               If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
'                  Exit Sub
               End If
               
               'Cerrando Archivo LOG del Proceso
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Se envió el documento electrónico " & CStr(r_str_NumFac) & " a : ")
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")
               Call fs_Escribir_Linea(l_str_RutaLgEma, "N° DE OPERACION   : " & r_str_NumOpe)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "ID CLIENTE        : " & CStr(r_str_NumDoc))
               Call fs_Escribir_Linea(l_str_RutaLgEma, "NOMBRE CLIENTE    : " & r_str_NomCli)
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Fecha Proceso     : " & Format(date, "dd/mm/yyyy"))
               Call fs_Escribir_Linea(l_str_RutaLgEma, "Fin Proceso       : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
               Call fs_Escribir_Linea(l_str_RutaLgEma, "")

               DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
            End If
         End If
       End If
   Next r_int_Contad
   
   MsgBox "Se enviaron satisfactoriamente el(los) documento(s) electrónico(s).", vbInformation, modgen_g_str_NomPlt
   
   chkSeleccionar.Value = 0
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Function fs_ObtCor(ByVal p_TipDoc As String, ByVal p_NumFac As String, ByVal p_NumOpe As String) As String
Dim r_rst_Princi      As ADODB.Recordset
Dim r_obj_Excel       As Excel.Application
Dim r_int_NumFil      As Integer
Dim r_str_Parame      As String
Dim r_str_RutFic      As String
   
   fs_ObtCor = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT B.DATGEN_TIPDOC, B.DATGEN_NUMDOC, B.DATGEN_DIRELE , "
   r_str_Parame = r_str_Parame & "         TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOM_CLIENTE "
   r_str_Parame = r_str_Parame & "    FROM CNTBL_DOCELE A "
   r_str_Parame = r_str_Parame & "         INNER JOIN CLI_DATGEN B ON TRIM(A.DOCELE_REC_TIPDOC) = B.DATGEN_TIPDOC AND A.DOCELE_REC_NUMDOC = TRIM(B.DATGEN_NUMDOC) "
   r_str_Parame = r_str_Parame & "   WHERE A.DOCELE_IDE_SERNUM = '" & p_NumFac & "'"
   r_str_Parame = r_str_Parame & "     AND A.DOCELE_NUMOPE = '" & p_NumOpe & "'"

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   fs_ObtCor = Trim(r_rst_Princi!DatGen_DirEle & "")
   
   If fs_ObtCor <> "" Then
      'Actualiza Correo electrónico en CNTBL_DOCELE
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CNTBL_DOCELE SET DOCELE_REC_CORREC = '" & CStr(fs_ObtCor) & "' "
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DOCELE_IDE_SERNUM = '" & CStr(p_NumFac) & "' "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DOCELE_NUMOPE = '" & CStr(p_NumOpe) & "' "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
         Exit Function
      End If
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   'Screen.MousePointer = 0
End Function

Private Sub fs_EnvCor(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, ByVal p_NumFac As String, ByVal p_NumOpe As String, p_Arregl() As moddat_tpo_Genera, p_Asunto As String, p_Contenido As String, Optional ByVal p_NomFil As String, Optional ByVal p_NomFil2 As String, Optional ByVal p_RutFil As String)
Dim r_int_Contad  As Integer

   On Error GoTo moddat_gf_EnvCor
     
   'Inicializa
   p_Sesion.DownLoadMail = False
   p_Sesion.NewSession = True
   p_Sesion.SignOn
   p_Mensaje.SessionID = p_Sesion.SessionID

   'Envío
   p_Mensaje.Compose

   For r_int_Contad = 0 To UBound(p_Arregl) - 1
      If Len(Trim(p_Arregl(r_int_Contad + 1).Genera_Codigo)) > 0 Then
         p_Mensaje.RecipIndex = r_int_Contad
         p_Mensaje.RecipDisplayName = p_Arregl(r_int_Contad + 1).Genera_Codigo
      End If
   Next r_int_Contad

   p_Mensaje.MsgSubject = p_Asunto
   p_Mensaje.MsgNoteText = p_Contenido
   
   If Len(Trim(p_NomFil)) > 0 Then
      p_Mensaje.AttachmentIndex = 0
      p_Mensaje.AttachmentName = p_NomFil
      p_Mensaje.AttachmentPathName = p_RutFil & p_NomFil
      p_Mensaje.AttachmentPosition = 0
      p_Mensaje.AttachmentType = mapData
   End If

   If Len(Trim(p_NomFil2)) > 0 Then
      p_Mensaje.AttachmentIndex = 1
      p_Mensaje.AttachmentName = p_NomFil2
      p_Mensaje.AttachmentPathName = p_RutFil & p_NomFil2
      p_Mensaje.AttachmentPosition = 1
      p_Mensaje.AttachmentType = mapData
   End If

   p_Mensaje.send
   DoEvents

  'Cierra la sesión
  p_Sesion.SignOff
  Exit Sub

moddat_gf_EnvCor:
   p_Sesion.SignOff
   
   Call fs_Escribir_Linea(l_str_RutaLgEma, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_EnvCor")
      
   'Actualiza Observaciones por envio de correo en CNTBL_DOCELE
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CNTBL_DOCELE SET DOCELE_OBSCOR = '" & Trim(Err.Description) & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DOCELE_IDE_SERNUM = '" & CStr(p_NumFac) & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DOCELE_NUMOPE = '" & CStr(p_NumOpe) & "' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
      Exit Sub
   End If
End Sub

Private Sub cmd_ExpArc_Click()
Dim r_str_NumFac     As String
Dim r_str_NotCre     As String
Dim r_str_TipDoc     As String
Dim r_str_NomCli     As String
Dim r_int_Contad     As Long
Dim r_int_ConSel     As Long
Dim r_str_RutFic     As String
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado documentos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirma
   If MsgBox("¿Está seguro de exportar los documentos seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
   
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
         r_str_NotCre = grd_Listad.TextMatrix(r_int_Contad, 8)
         r_str_NumFac = grd_Listad.TextMatrix(r_int_Contad, 16)
         r_str_NomCli = grd_Listad.TextMatrix(r_int_Contad, 2)
         
         If r_str_NotCre = "" Then
            If InStr(r_str_NumFac, "B") > 0 Then
               r_str_TipDoc = "03"
            Else
               r_str_TipDoc = "01"
            End If
         Else
            r_str_TipDoc = "07"
         End If
            
         r_str_NomCli = "C:\SBSMIC\" & r_str_NomCli
         
         'Crear la Carpeta para almacenar la(s) factura(s) seleccionada(s)
         Set l_fsobj = New FileSystemObject
         If l_fsobj.FolderExists(r_str_NomCli) = False Then
            l_fsobj.CreateFolder (r_str_NomCli)
         End If
         
         Set l_fsobj = New FileSystemObject
         r_str_RutFic = "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac
                 
         If gf_Existe_Archivo(moddat_g_str_RutFac & "\reportes\aceptados\", r_str_RutFic & ".pdf") Then
            FileCopy moddat_g_str_RutFac & "\reportes\aceptados\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf", r_str_NomCli & "\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf"
            DoEvents
         End If
         
         If gf_Existe_Archivo(moddat_g_str_RutFac & "\reportes\aceptados\", r_str_RutFic & ".xml") Then
            FileCopy moddat_g_str_RutFac & "\reportes\aceptados\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".xml", r_str_NomCli & "\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".xml"
            DoEvents
         End If
         
'         FileCopy moddat_g_str_RutFac & "\reportes\aceptados\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf", r_str_NomCli & "\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".pdf"
'         FileCopy moddat_g_str_RutFac & "\reportes\aceptados\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".xml", r_str_NomCli & "\" & "20511904162-" & r_str_TipDoc & "-" & r_str_NumFac & ".xml"
'
      End If
   Next r_int_Contad
   
   MsgBox "Se exportaron los documentos seleccionados, en: " & r_str_NomCli, vbExclamation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_GenNCr_Click()
Dim r_int_Contad           As Integer
Dim r_int_ConSel           As Integer
Dim r_str_NumOpe           As String
Dim r_str_NumFac           As String
Dim r_str_NumMov           As String
Dim r_str_Fecemi           As String
Dim r_arr_NumFac()         As moddat_tpo_Genera

   'valida selección
   r_int_ConSel = 0
   r_str_NumFac = ""
   
   ReDim r_arr_NumFac(0)
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
         If grd_Listad.TextMatrix(r_int_Contad, 14) = 1 And grd_Listad.TextMatrix(r_int_Contad, 15) = 0 Then
            MsgBox "La Factura no se encuentra aceptada: " & grd_Listad.TextMatrix(r_int_Contad, 0), vbInformation, modgen_g_str_NomPlt '
            Exit Sub
         ElseIf cmb_TipRsp.ItemData(cmb_TipRsp.ListIndex) = 4 Then
            MsgBox "No se puede generar Nota de Crédito.", vbInformation, modgen_g_str_NomPlt 'La Factura no se encuentra aceptada
            Exit Sub
         ElseIf (grd_Listad.TextMatrix(r_int_Contad, 8)) <> "" Then
            MsgBox "No se puede generar Nota de Crédito: " & grd_Listad.TextMatrix(r_int_Contad, 0), vbInformation, modgen_g_str_NomPlt 'es Nota de Crédito
            Exit Sub
         End If
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Facturas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirma
   If MsgBox("¿Está seguro de generar Nota de Crédito a las factura(s) seleccionada(s) ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
      
         ReDim Preserve r_arr_NumFac(UBound(r_arr_NumFac) + 1)
         r_str_NumOpe = grd_Listad.TextMatrix(r_int_Contad, 10)
         r_str_Fecemi = grd_Listad.TextMatrix(r_int_Contad, 11)
         r_str_NumFac = grd_Listad.TextMatrix(r_int_Contad, 16)
         r_str_NumMov = grd_Listad.TextMatrix(r_int_Contad, 17)
         
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_Codigo = r_str_NumOpe
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_Nombre = r_str_NumFac
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_ConHip = r_str_NumMov
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_Refere = r_str_Fecemi
      End If
   Next
   
   For r_int_Contad = 0 To UBound(r_arr_NumFac)
      If Len(Trim(r_arr_NumFac(r_int_Contad).Genera_Codigo)) > 0 Then
         Call fs_Generar_NotaCredito(r_arr_NumFac(r_int_Contad).Genera_Refere, r_arr_NumFac(r_int_Contad).Genera_Codigo, r_arr_NumFac(r_int_Contad).Genera_Nombre, r_arr_NumFac(r_int_Contad).Genera_ConHip)
      End If
   Next r_int_Contad
      
   chkSeleccionar.Value = False
   Call fs_Limpia
   MsgBox "Proceso Finalizado.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub fs_Generar_NotaCredito(ByVal p_FecEmi As String, ByVal p_NumOpe As String, ByVal p_NumFac As String, ByVal p_NumMov As String)
Dim r_lng_Contad           As Long
Dim r_int_SerFac           As Integer
Dim r_lng_NumFac           As Long

   On Error GoTo MyError
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "     SELECT DOCELE_CODIGO                                                                                       AS CODIGO, "
   g_str_Parame = g_str_Parame & "            'IDE'                                                                                               AS CAMPO_IDE_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_IDE_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_IDE_02, "
   End If
   
   g_str_Parame = g_str_Parame & "            DOCELE_IDE_FECEMI                                                                                   AS CAMPO_IDE_03, "
   g_str_Parame = g_str_Parame & "            DOCELE_IDE_HOREMI                                                                                   AS CAMPO_IDE_04, "
   g_str_Parame = g_str_Parame & "            '07'                                                                                                AS CAMPO_IDE_05, "
   g_str_Parame = g_str_Parame & "            DOCELE_IDE_TIPMON                                                                                   AS CAMPO_IDE_06, "
   g_str_Parame = g_str_Parame & "            'EMI'                                                                                               AS CAMPO_EMI_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_EMI_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_EMI_02, "
   End If
   
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_TIPDOC                                                                                   AS CAMPO_EMI_03, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_NUMDOC                                                                                   AS CAMPO_EMI_04, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_NOMCOM                                                                                   AS CAMPO_EMI_05, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_DENOMI                                                                                   AS CAMPO_EMI_06, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_UBIGEO                                                                                   AS CAMPO_EMI_07, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_DIRCOM                                                                                   AS CAMPO_EMI_08, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_URBANI                                                                                   AS CAMPO_EMI_09, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_PROVIN                                                                                   AS CAMPO_EMI_10, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_DEPART                                                                                   AS CAMPO_EMI_11, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_DISTRI                                                                                   AS CAMPO_EMI_12, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_CODPAI                                                                                   AS CAMPO_EMI_13, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_TELEMI                                                                                   AS CAMPO_EMI_14, "
   g_str_Parame = g_str_Parame & "            DOCELE_EMI_COREMI                                                                                   AS CAMPO_EMI_15, "
   g_str_Parame = g_str_Parame & "            'REC'                                                                                               AS CAMPO_REC_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_REC_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_REC_02, "
   End If
   
   g_str_Parame = g_str_Parame & "            DOCELE_REC_TIPDOC                                                                                   AS CAMPO_REC_03, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_NUMDOC                                                                                   AS CAMPO_REC_04, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_DENOMI                                                                                   AS CAMPO_REC_05, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_DIRCOM                                                                                   AS CAMPO_REC_06, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_DEPART                                                                                   AS CAMPO_REC_07, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_PROVIN                                                                                   AS CAMPO_REC_08, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_DISTRI                                                                                   AS CAMPO_REC_09, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_CODPAI                                                                                   AS CAMPO_REC_10, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_TELREC                                                                                   AS CAMPO_REC_11, "
   g_str_Parame = g_str_Parame & "            DOCELE_REC_CORREC                                                                                   AS CAMPO_REC_12, "
   g_str_Parame = g_str_Parame & "            'DRF'                                                                                               AS CAMPO_DRF_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_DRF_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_DRF_02, "
   End If
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            '03'                                                                                             AS CAMPO_DRF_03, "
   Else
      g_str_Parame = g_str_Parame & "            '01'                                                                                             AS CAMPO_DRF_03, "
   End If
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_04, "
   g_str_Parame = g_str_Parame & "            '01'                                                                                                AS CAMPO_DRF_05, "
   g_str_Parame = g_str_Parame & "            TRIM(V.CATSUN_DESCRI)                                                                               AS CAMPO_DRF_06, "
   g_str_Parame = g_str_Parame & "            'CAB'                                                                                               AS CAMPO_CAB_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_CAB_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_CAB_02, "
   End If
   
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_CODIGO_OPEGRV                                                                            AS CAMPO_CAB_03, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TOTVTA_OPEGRV                                                                            AS CAMPO_CAB_04, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_CODIGO_OPEINA                                                                            AS CAMPO_CAB_05, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TOTVTA_OPEINA                                                                            AS CAMPO_CAB_06, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_CODIGO_OPEEXO                                                                            AS CAMPO_CAB_07, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TOTVTA_OPEEXO                                                                            AS CAMPO_CAB_08, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_CODIMP                                                                                   AS CAMPO_CAB_18_1, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_MTOIMP                                                                                   AS CAMPO_CAB_18_2, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_OTRCAR                                                                                   AS CAMPO_CAB_19, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_CODIGO_TOTDSC                                                                            AS CAMPO_CAB_20, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TOTDSC                                                                                   AS CAMPO_CAB_21, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_IMPTOT_DOCUME                                                                            AS CAMPO_CAB_22, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TOTANT                                                                                   AS CAMPO_CAB_25, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_TIPOPE                                                                                   AS CAMPO_CAB_26, "
   g_str_Parame = g_str_Parame & "            DOCELE_CAB_LEYEND                                                                                   AS CAMPO_CAB_27, "
   g_str_Parame = g_str_Parame & "            'ADI'                                                                                               AS CAMPO_ADI_01, "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_ADI_02, "
   Else
      g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_ADI_02, "
   End If
   
   g_str_Parame = g_str_Parame & "            DOCELE_ADI_TITADI                                                                                   AS CAMPO_ADI_03, "
   g_str_Parame = g_str_Parame & "            DOCELE_ADI_VALADI                                                                                   AS CAMPO_ADI_04, "
   g_str_Parame = g_str_Parame & "            DOCELE_NUMOPE                                                                                       AS OPERACION, "
   g_str_Parame = g_str_Parame & "            DOCELE_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO "
   g_str_Parame = g_str_Parame & "       FROM CNTBL_DOCELE "
   g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN V ON V.CATSUN_NROCAT = 9 AND V.CATSUN_CODIGO = '01' "
   g_str_Parame = g_str_Parame & "      WHERE DOCELE_IDE_SERNUM = '" & CStr(p_NumFac) & "' "
   
   If InStr(p_NumFac, "B") > 0 Then
      g_str_Parame = g_str_Parame & "        AND DOCELE_IDE_TIPDOC = '03' "
   Else
      g_str_Parame = g_str_Parame & "        AND DOCELE_IDE_TIPDOC = '01' "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal en CNTBL_DOCELE , procedimiento: fs_Generar_NotaCredito")
      Exit Sub
   End If
 
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
  
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      
      Call fs_Obtener_Codigo("07", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
      
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
      g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
      g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
      g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
      g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
      g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
      g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
      g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
      g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
      g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
      g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
      g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
      g_str_Parame = g_str_Parame & " VALUES ( "
      g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
      g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
      g_str_Parame = g_str_Parame & " NULL, "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
      g_str_Parame = g_str_Parame & " NULL, "
      g_str_Parame = g_str_Parame & " NULL, "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_REC_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
      End If
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_05 & "'                  , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
      End If
            
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & " NULL, " '
      g_str_Parame = g_str_Parame & " NULL, " '

      If IsNull(g_rst_Princi!CAMPO_CAB_25) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_25 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
      g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
      g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
      g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      
      DoEvents: DoEvents: DoEvents
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "     SELECT 'DET'                                                                                               AS CAMPO_DET_01, "
      
      If InStr(p_NumFac, "B") > 0 Then
         g_str_Parame = g_str_Parame & "            'B'                                                                                              AS CAMPO_DET_02, "
      Else
         g_str_Parame = g_str_Parame & "            'F'                                                                                              AS CAMPO_DET_02, "
      End If
      
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_NUMITE                                                                                AS CAMPO_DET_03, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODPRD                                                                                AS CAMPO_DET_04, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_DESPRD                                                                                AS CAMPO_DET_05, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CANTID                                                                                AS CAMPO_DET_06, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_UNIDAD                                                                                AS CAMPO_DET_07, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_VALUNI                                                                                AS CAMPO_DET_08, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_PUNVTA                                                                                AS CAMPO_DET_09, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIMP                                                                                AS CAMPO_DET_10_1, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_MTOIMP                                                                                AS CAMPO_DET_10_2, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_TIPAFE                                                                                AS CAMPO_DET_10_3, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_VALVTA                                                                                AS CAMPO_DET_11, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_VALREF                                                                                AS CAMPO_DET_12, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODSUN                                                                                AS CAMPO_DET_15, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODCON                                                                                AS CAMPO_DET_16, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_NROCON                                                                                AS CAMPO_DET_17, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIGO_FECOTO                                                                         AS CAMPO_DET_18, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_FECOTO                                                                                AS CAMPO_DET_19, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIGO_TIPPRE                                                                         AS CAMPO_DET_20, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_TIPPRE                                                                                AS CAMPO_DET_21, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIGO_PARREG                                                                         AS CAMPO_DET_22, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_PARREG                                                                                AS CAMPO_DET_23, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIGO_PRIVIV                                                                         AS CAMPO_DET_24, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_PRIVIV                                                                                AS CAMPO_DET_25, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODIGO_DIRCOM                                                                         AS CAMPO_DET_26, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_DIRCOM                                                                                AS CAMPO_DET_27, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODUBI                                                                                AS CAMPO_DET_28, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_UBIGEO                                                                                AS CAMPO_DET_29, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODURB                                                                                AS CAMPO_DET_30, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_URBANI                                                                                AS CAMPO_DET_31, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODDPT                                                                                AS CAMPO_DET_32, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_DEPART                                                                                AS CAMPO_DET_33, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODPRV                                                                                AS CAMPO_DET_34, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_PROVIN                                                                                AS CAMPO_DET_35, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_CODDIS                                                                                AS CAMPO_DET_36, "
      g_str_Parame = g_str_Parame & "            DOCELEDET_DET_DISTRI                                                                                AS CAMPO_DET_37 "
      g_str_Parame = g_str_Parame & "       FROM CNTBL_DOCELEDET "
      g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN V ON V.CATSUN_NROCAT = 9 AND V.CATSUN_CODIGO = '01' "
      g_str_Parame = g_str_Parame & "      WHERE DOCELEDET_CODIGO = " & CStr(g_rst_Princi!CODIGO) & " "
              
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal en CNTBL_DOCELEDET , procedimiento: fs_Generar_NotaCredito")
         Exit Sub
      End If
      
      If g_rst_Listas.BOF And g_rst_Listas.EOF Then
         g_rst_Listas.Close
         Set g_rst_Listas = Nothing
         Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro en CNTBL_DOCELEDET, procedimiento: fs_Generar_NotaCredito")
         Exit Sub
      End If
      
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Listas!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_10_3 & "'                   , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Listas!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & " NULL, " '
         g_str_Parame = g_str_Parame & " NULL, " '
      
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_15 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_16 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_17 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_18 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_19 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_20 & "'                     , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_21 & "   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_22 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_23 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_24 & "'                     , "
         
         If IsNull(g_rst_Listas!CAMPO_DET_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Listas!CAMPO_DET_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_28 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_29 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_30 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_31 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_32 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_33 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_34 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_35 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_36 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Listas!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
      
         g_rst_Listas.MoveNext
      Loop
    
    
      'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 0 "
      g_str_Parame = g_str_Parame & "  WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "    AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      
      'ACTUALIZA EL CAMPO DOCELE_NFANCR - FACTURA DE LA NOTA DE CREDITO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " UPDATE CNTBL_DOCELE SET DOCELE_NFANCR = '" & p_NumFac & "', "
      g_str_Parame = g_str_Parame & "        DOCELE_DRF_NUMDOC = '" & p_NumFac & "', "
      g_str_Parame = g_str_Parame & "        DOCELE_FECNCR =  " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " "
      g_str_Parame = g_str_Parame & "  WHERE DOCELE_CODIGO = " & r_lng_Contad & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar DOCELE_NFANCR de la tabla CNTBL_DOCELE, procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      
      'ACTUALIZA EL CAMPO DOCELE_SITUAC - DE LA FACTURA A LA QUE SE LE GENERÓ LA NOTA DE CRÉDITO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " UPDATE CNTBL_DOCELE SET DOCELE_SITUAC = 2 "
      g_str_Parame = g_str_Parame & "  WHERE DOCELE_IDE_SERNUM = '" & p_NumFac & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar DOCELE_NFANCR de la tabla CNTBL_DOCELE, procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      
      r_lng_Contad = r_lng_Contad + 1
      r_lng_NumFac = r_lng_NumFac + 1
      
      g_rst_Princi.MoveNext
      
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub

Private Sub cmd_ExpExc_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
    
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_TipRsp.Clear
   
   cmb_TipRsp.AddItem "- TODOS -"
   cmb_TipRsp.ItemData(cmb_TipRsp.NewIndex) = 3
   cmb_TipRsp.AddItem "ACEPTADAS"
   cmb_TipRsp.ItemData(cmb_TipRsp.NewIndex) = 1
   cmb_TipRsp.AddItem "RECHAZADAS"
   cmb_TipRsp.ItemData(cmb_TipRsp.NewIndex) = 0
   cmb_TipRsp.AddItem "NOTA DE CREDITO"
   cmb_TipRsp.ItemData(cmb_TipRsp.NewIndex) = 4
   cmb_TipRsp.ListIndex = 0
   
    'Búsqueda
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "DOCUMENTO"
   cmb_Buscar.AddItem "APELLIDOS Y NOMBRES"
   cmb_Buscar.AddItem "NRO OPERACION"
   cmb_Buscar.ListIndex = 0
   
   l_str_FecCar = Format(date, "yyyymmdd")
   
   grd_Listad.Cols = 22
   grd_Listad.ColWidth(0) = 1230  'Operacion
   grd_Listad.ColWidth(1) = 1260  'id Cliente
   grd_Listad.ColWidth(2) = 4260  'Cliente
   grd_Listad.ColWidth(3) = 900   'Moneda
   grd_Listad.ColWidth(4) = 1350  'Importe
   grd_Listad.ColWidth(5) = 1080  'F.Emisión
   grd_Listad.ColWidth(6) = 1080  'F.Proc
   grd_Listad.ColWidth(7) = 1080  'F.Pago
   grd_Listad.ColWidth(8) = 1260  'Refer
   grd_Listad.ColWidth(9) = 660  'Seleccionar
   grd_Listad.ColWidth(10) = 0    'numero de operacion
   grd_Listad.ColWidth(11) = 0    'fecha de emisión sin formato
   grd_Listad.ColWidth(12) = 0    'fecha de proceso sin formato
   grd_Listad.ColWidth(13) = 0    'importe de Pago
   grd_Listad.ColWidth(14) = 0    'Flag de Enviado
   grd_Listad.ColWidth(15) = 0    'Flag de Respuesta
   grd_Listad.ColWidth(16) = 0    'Número Factura
   grd_Listad.ColWidth(17) = 0    'Número Movimiento
   grd_Listad.ColWidth(18) = 0    'Interés
   grd_Listad.ColWidth(19) = 0    'Otros_importes
   grd_Listad.ColWidth(20) = 0    'Tipo de Comprobante
   grd_Listad.ColWidth(21) = 3000 'Correo
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(21) = flexAlignLeftCenter
   
   'Crear Archivo LOG del Proceso - Nota de Crédito
   l_str_NomLOG = UCase(App.EXEName) & "_NC_" & Format(date, "yyyymmdd") & ".LOG"
   l_int_NumLOG = FreeFile
    
   If gf_Existe_Archivo(g_str_RutLog & "\", l_str_NomLOG) Then
      Kill g_str_RutLog & "\" & l_str_NomLOG
      DoEvents
   End If
    
   l_str_RutaLg = g_str_RutLog & "\" & l_str_NomLOG
   l_str_RutFacEnt = moddat_g_str_RutFac & "\entrada\" 'moddat_g_str_RutLoc
   
   'Crear la Carpeta entrada
   Set l_fsobj = New FileSystemObject
   If l_fsobj.FolderExists(l_str_RutFacEnt) = False Then
      l_fsobj.CreateFolder (l_str_RutFacEnt)
   End If
     
End Sub

Private Sub fs_Limpia()
   ipp_FecIni.Text = "01/01/" & Format(Year(date), "0000")
   ipp_FecFin.Text = Format(date, "DD/MM/YYYY")
   cmb_TipRsp.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
   chkSeleccionar.Value = 0
   
   cmb_Buscar.ListIndex = 0
   txt_Buscar.Text = Empty
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmb_TipRsp.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
   cmd_GenNCr.Enabled = Not p_Activa
   cmd_ExpArc.Enabled = Not p_Activa
   cmd_AbrArc.Enabled = Not p_Activa
   cmd_EnvMail.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()

   Call gs_LimpiaGrid(grd_Listad)
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELE_NUMOPE, DOCELE_REC_TIPDOC, DOCELE_REC_NUMDOC,  TRIM(DOCELE_REC_DENOMI) AS CLIENTE, DOCELE_IDE_TIPMON, "
   g_str_Parame = g_str_Parame & "        REPLACE(A.DOCELE_IDE_FECEMI,'-','') AS FEC_EMISION, A.DOCELE_FECPRO AS FEC_PROCESO, DOCELE_CAB_IMPTOT_DOCUME AS MONTO_PAGO, "
   g_str_Parame = g_str_Parame & "        DOCELE_NUMMOV AS NUMMOV, DOCELE_NFANCR AS NUM_FACTURA, REPLACE(DOCELE_FECPAG,'-','') AS FEC_PAGO, DOCELE_REC_CORREC AS CORREO, "
   g_str_Parame = g_str_Parame & "        DOCELE_FLGENV,  DOCELE_FLGRPT, DOCELE_IDE_SERNUM AS NUMERO_FACTURA, DOCELE_IDE_TIPDOC AS TIPO_COMPROBANTE, "
   g_str_Parame = g_str_Parame & "        NVL((SELECT DOCELEDET_DET_VALUNI "
   g_str_Parame = g_str_Parame & "               FROM CNTBL_DOCELEDET "
   g_str_Parame = g_str_Parame & "              WHERE DOCELEDET_CODIGO = DOCELE_CODIGO AND DOCELEDET_DET_NUMITE = '001'),0) AS INTERES, "
   g_str_Parame = g_str_Parame & "        NVL((SELECT DOCELEDET_DET_VALUNI "
   g_str_Parame = g_str_Parame & "               FROM CNTBL_DOCELEDET "
   g_str_Parame = g_str_Parame & "              WHERE DOCELEDET_CODIGO = DOCELE_CODIGO AND DOCELEDET_DET_NUMITE = '002'),0) AS OTROS_IMPORTES "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELE A "
   g_str_Parame = g_str_Parame & "  WHERE TO_NUMBER(REPLACE(A.DOCELE_IDE_FECEMI,'-','')) >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "    AND TO_NUMBER(REPLACE(A.DOCELE_IDE_FECEMI,'-','')) <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   If cmb_TipRsp.ListIndex <> 0 Then
      If cmb_TipRsp.ItemData(cmb_TipRsp.ListIndex) = 4 Then
         g_str_Parame = g_str_Parame & "    AND A.DOCELE_IDE_TIPDOC = '07'"
      Else
         g_str_Parame = g_str_Parame & "    AND (A.DOCELE_IDE_TIPDOC = '01' OR A.DOCELE_IDE_TIPDOC = '03') "
         g_str_Parame = g_str_Parame & "    AND DOCELE_FLGENV = 1 "
         g_str_Parame = g_str_Parame & "    AND A.DOCELE_FLGRPT = " & cmb_TipRsp.ItemData(cmb_TipRsp.ListIndex)
         If cmb_TipRsp.ItemData(cmb_TipRsp.ListIndex) = 0 Then
            g_str_Parame = g_str_Parame & "    AND A.DOCELE_OBSERV IS NOT NULL "
         End If
      End If
   End If
   g_str_Parame = g_str_Parame & "   AND DOCELE_SITUAC = 1 "
   
   If cmb_Buscar.ListIndex > 0 Then
      If cmb_Buscar.ListIndex = 1 Then    'NRO DOCUMENTO
         g_str_Parame = g_str_Parame & "   AND DOCELE_REC_NUMDOC = '" & Trim(txt_Buscar.Text) & "' "
      ElseIf cmb_Buscar.ListIndex = 2 Then 'APELLIDOS Y NOMBRES
         g_str_Parame = g_str_Parame & "   AND TRIM(DOCELE_REC_DENOMI) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%' "
      ElseIf cmb_Buscar.ListIndex = 3 Then 'NRO OPERACION
         g_str_Parame = g_str_Parame & "   AND DOCELE_NUMOPE = '" & Trim(txt_Buscar.Text) & "' "
      End If
   End If
4
   If Chk_CorEnv.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND DOCELE_ENVCOR = 0 "
   Else
      g_str_Parame = g_str_Parame & "   AND DOCELE_ENVCOR = 1 "
   End If
   
   g_str_Parame = g_str_Parame & " ORDER BY A.DOCELE_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_FacEle, 3) Then
      Exit Sub
   End If

   If l_rst_FacEle.BOF And l_rst_FacEle.EOF Then
      l_rst_FacEle.Close
      Set l_rst_FacEle = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Activa(True)
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   l_rst_FacEle.MoveFirst
   
   Do While Not l_rst_FacEle.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'numero operacion (formateado)
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(Trim(l_rst_FacEle!DOCELE_NUMOPE & ""))
      
      'tipo de documento
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(l_rst_FacEle!DOCELE_REC_TIPDOC) & "-" & Trim(l_rst_FacEle!DOCELE_REC_NUMDOC & "")
      
      'nombre del cliente
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(l_rst_FacEle!CLIENTE & "")
      
      'Moneda
      grd_Listad.Col = 3
      If l_rst_FacEle!DOCELE_IDE_TIPMON = "PEN" Then
         grd_Listad.Text = "S/.   "
      Else
         grd_Listad.Text = "US$   "
      End If
      
      'importe
      grd_Listad.Col = 4
      grd_Listad.Text = Format(l_rst_FacEle!MONTO_PAGO, "###,###,###,##0.00")
             
      'fecha de Emisión (formateado)
      grd_Listad.Col = 5
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_FacEle!FEC_EMISION))
      
      'fecha de proceso. (formateado)
      grd_Listad.Col = 6
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_FacEle!FEC_PROCESO))
      
      'fecha de Pago
      grd_Listad.Col = 7
      If Not IsNull(l_rst_FacEle!FEC_PAGO) Then
         grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_FacEle!FEC_PAGO))
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 8
      If Not IsNull(l_rst_FacEle!NUM_FACTURA) Then
         grd_Listad.Text = Trim(CStr(l_rst_FacEle!NUM_FACTURA))
      Else
         grd_Listad.Text = ""
      End If
      
      '9 es seleccion
      grd_Listad.Col = 9
      grd_Listad.Text = ""
      
      'numero de operacion
      grd_Listad.Col = 10
      grd_Listad.Text = Trim(l_rst_FacEle!DOCELE_NUMOPE & "")
      
      'fecha de Pago
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(l_rst_FacEle!FEC_EMISION)
      
      'fecha de proceso
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(l_rst_FacEle!FEC_PROCESO)
      
      'importe de emisión
      grd_Listad.Col = 13
      grd_Listad.Text = l_rst_FacEle!MONTO_PAGO
      
      'Flag de Enviado
      grd_Listad.Col = 14
      grd_Listad.Text = l_rst_FacEle!DOCELE_FLGENV
      
      'Flag de Respuesta
      grd_Listad.Col = 15
      grd_Listad.Text = l_rst_FacEle!DOCELE_FLGRPT
      
      'Número Factura
      grd_Listad.Col = 16
      grd_Listad.Text = l_rst_FacEle!NUMERO_FACTURA
      
      'Número Movimiento
      grd_Listad.Col = 17
      If Not IsNull(l_rst_FacEle!NUMMOV) Then
         grd_Listad.Text = l_rst_FacEle!NUMMOV
      Else
         grd_Listad.Text = ""
      End If
      
      'Interés
      grd_Listad.Col = 18
      If Not IsNull(l_rst_FacEle!INTERES) Then
         grd_Listad.Text = l_rst_FacEle!INTERES
      Else
         grd_Listad.Text = 0
      End If
      
      'Otros_Importes
      grd_Listad.Col = 19
      If Not IsNull(l_rst_FacEle!OTROS_IMPORTES) Then
         grd_Listad.Text = l_rst_FacEle!OTROS_IMPORTES
      Else
         grd_Listad.Text = 0
      End If
      
      'Otros_Importes
      grd_Listad.Col = 20
      If Not IsNull(l_rst_FacEle!TIPO_COMPROBANTE) Then
         If l_rst_FacEle!TIPO_COMPROBANTE = "01" Then
            grd_Listad.Text = "F"
         ElseIf l_rst_FacEle!TIPO_COMPROBANTE = "03" Then
            grd_Listad.Text = "B"
         ElseIf l_rst_FacEle!TIPO_COMPROBANTE = "07" Then
            grd_Listad.Text = "NC"
         End If
      Else
         grd_Listad.Text = ""
      End If
      
      'Correo
      grd_Listad.Col = 21
      If IsNull(l_rst_FacEle!CORREO) Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = Trim(l_rst_FacEle!CORREO)
      End If
      
      l_rst_FacEle.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel            As Excel.Application
Dim r_int_NroFil           As Integer
Dim r_int_nroaux           As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      r_int_NroFil = 2
      .Cells(r_int_NroFil, 1) = "REPORTE DE DOCUMENTOS ELECTRONICOS (" & CStr(ipp_FecIni.Text) & " - " & CStr(ipp_FecFin.Text) & ")"
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Font.Size = 12
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Merge
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "ITEM"
      .Cells(r_int_NroFil, 2) = "OPERACION"
      .Cells(r_int_NroFil, 3) = "ID CLIENTE"
      .Cells(r_int_NroFil, 4) = "CLIENTE"
      .Cells(r_int_NroFil, 5) = "FECHA EMISION"
      .Cells(r_int_NroFil, 6) = "FECHA PROCESO"
      .Cells(r_int_NroFil, 7) = "FECHA PAGO"
      .Cells(r_int_NroFil, 8) = "TIPO COMPROBANTE"
      .Cells(r_int_NroFil, 9) = "N° COMPROBANTE"
      .Cells(r_int_NroFil, 10) = "MONEDA"
      .Cells(r_int_NroFil, 11) = "INTERES"
      .Cells(r_int_NroFil, 12) = "OTROS IMPORTES"
      .Cells(r_int_NroFil, 13) = "TOTAL_VENTA"
      .Cells(r_int_NroFil, 14) = "NUM_COMPR_ANULADO"
      .Cells(r_int_NroFil, 15) = "CORREO"

      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 18
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 18
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 45
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 19
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 16
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 16
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("N").ColumnWidth = 23
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 45
      .Columns("O").HorizontalAlignment = xlHAlignLeft
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = Format(r_int_NroFil - 4, "00#")                        'ITEM
         .Cells(r_int_NroFil, 2) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 0)           'OPERACION
         .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 1)           'TIPO DOCUMENTO
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 2)                 'CLIENTE
         .Cells(r_int_NroFil, 5) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 5))     'FECHA EMISION
         .Cells(r_int_NroFil, 6) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 6))     'FECHA PROCESO
         .Cells(r_int_NroFil, 7) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 7))     'FECHA PAGO
         .Cells(r_int_NroFil, 8) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 20))    'TIPO COMPROBANTE
         .Cells(r_int_NroFil, 9) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 16))    'N° COMPROBANTE
         .Cells(r_int_NroFil, 10) = "'" & CStr(grd_Listad.TextMatrix(r_int_nroaux, 3))    'MONEDA
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_nroaux, 18)               'INTERES
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_nroaux, 19)               'OTROS_IMPORTES
         .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_nroaux, 4)                'TOTAL_VENTAS
         .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_nroaux, 8)                'REFER
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_nroaux, 21)               'CORREO
         
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      .Range(.Cells(1, 1), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(4, 1), .Cells(r_int_NroFil - 1, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(r_int_NroFil - 1, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(r_int_NroFil - 1, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(r_int_NroFil - 1, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(r_int_NroFil - 1, 15)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub chkSeleccionar_Click()
Dim r_int_Fila             As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_int_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_int_Fila, 9) = ""
         Next r_int_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_int_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_int_Fila, 9) = "X"
         Next r_int_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub



Private Sub grd_Listad_DblClick()
 If grd_Listad.Rows > 0 Then
      If grd_Listad.TextMatrix(grd_Listad.Row, 9) = "X" Then
         grd_Listad.TextMatrix(grd_Listad.Row, 9) = ""
      Else
         grd_Listad.TextMatrix(grd_Listad.Row, 9) = "X"
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
      Call gs_SetFocus(cmb_TipRsp)
   End If
End Sub

Private Sub cmb_TipRsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub pnl_Tit_Moneda_Click()
   If Len(Trim(pnl_Tit_Moneda.Tag)) = 0 Or pnl_Tit_Moneda.Tag = "D" Then
      pnl_Tit_Moneda.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

'Reordenar
Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 10, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 10, "C-")
   End If
End Sub

Private Sub pnl_Tit_DoiCli_Click()
   If Len(Trim(pnl_Tit_DoiCli.Tag)) = 0 Or pnl_Tit_DoiCli.Tag = "D" Then
      pnl_Tit_DoiCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DoiCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecEmi_Click()
   If Len(Trim(pnl_Tit_FecEmi.Tag)) = 0 Or pnl_Tit_FecEmi.Tag = "D" Then
      pnl_Tit_FecEmi.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 11, "C")
   Else
      pnl_Tit_FecEmi.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 11, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecPro_Click()
   If Len(Trim(pnl_Tit_FecPro.Tag)) = 0 Or pnl_Tit_FecPro.Tag = "D" Then
      pnl_Tit_FecPro.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 12, "C")
   Else
      pnl_Tit_FecPro.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 12, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 13, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 13, "N-")
   End If
End Sub

Private Sub pnl_Tit_Refer_Click()
  If Len(Trim(pnl_Tit_Refer.Tag)) = 0 Or pnl_Tit_Refer.Tag = "D" Then
      pnl_Tit_Refer.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_Tit_Refer.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "C-")
   End If
End Sub

Private Sub fs_Escribir_Linea(p_ruta As String, p_texto As String)

On Error GoTo MyError

   'Escribir en archivo según se ubique
   Set l_fsobj = New FileSystemObject
   Set l_txtStr = l_fsobj.OpenTextFile(p_ruta, ForAppending, False)
   l_txtStr.WriteLine (p_texto)
   l_txtStr.Close
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Escribir_Linea")
End Sub
Private Function fs_Obtener_Codigo(ByVal p_TipDoc As String, ByRef p_CodIte As Long, ByRef p_SerFac As Integer, ByRef p_NumFac As Long)
Dim r_str_Parame           As String
Dim r_rst_Codigo           As ADODB.Recordset
Dim r_int_InsUpd           As Integer
   
   p_CodIte = 0
   p_SerFac = 0
   p_NumFac = 0
   r_int_InsUpd = 0
   
   
   'Código Máximo de CNTBL_DOCELE
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT MAX(DOCELE_CODIGO) AS CODIGO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_DOCELE "
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      r_rst_Codigo.Close
      Set r_rst_Codigo = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontro ningún registro en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
      If IsNull(r_rst_Codigo!CODIGO) Then
         p_CodIte = 0
      Else
         p_CodIte = r_rst_Codigo!CODIGO
      End If
   End If
   
   p_CodIte = p_CodIte + 1
   
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing
   
   
   'Número de Serie
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL(DOCELE_NUMSER,0) AS SERIE "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_FOLIOS_DOCELE "
   r_str_Parame = r_str_Parame & "  WHERE DOCELE_TIPDOC = '" & p_TipDoc & "' "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If

   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      p_SerFac = 1
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
   
      If IsNull(r_rst_Codigo!SERIE) Then
         p_SerFac = 1
      Else
         p_SerFac = r_rst_Codigo!SERIE
      End If
   End If
   
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing
   
   'Número de Factura
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DOCELE_NUMCOR AS CORRELATIVO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_FOLIOS_DOCELE "
   r_str_Parame = r_str_Parame & "  WHERE DOCELE_TIPDOC = '" & p_TipDoc & "'"
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      p_NumFac = 0
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
      If IsNull(r_rst_Codigo!CORRELATIVO) Then
         p_NumFac = 0
      Else
         p_NumFac = r_rst_Codigo!CORRELATIVO
      End If
      r_int_InsUpd = 1
   End If
   
   p_NumFac = p_NumFac + 1
    
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing
  
   If p_NumFac = 99999999 Then
      p_NumFac = 0
      p_SerFac = p_SerFac + 1
   End If

   'Actualizando Folio
   If r_int_InsUpd = 0 Then
      'Insert
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " INSERT INTO CNTBL_FOLIOS_DOCELE ("
      r_str_Parame = r_str_Parame & "        DOCELE_TIPDOC, "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMSER, "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMCOR, "
      r_str_Parame = r_str_Parame & "        SEGUSUCRE, "
      r_str_Parame = r_str_Parame & "        SEGFECCRE, "
      r_str_Parame = r_str_Parame & "        SEGHORCRE, "
      r_str_Parame = r_str_Parame & "        SEGPLTCRE, "
      r_str_Parame = r_str_Parame & "        SEGTERCRE, "
      r_str_Parame = r_str_Parame & "        SEGSUCCRE) "
      r_str_Parame = r_str_Parame & " VALUES ("
      r_str_Parame = r_str_Parame & " '" & CStr(p_TipDoc) & "', "
      r_str_Parame = r_str_Parame & " '" & Format(CStr(p_SerFac), "000") & "', "
      r_str_Parame = r_str_Parame & " '" & Format(CStr(p_NumFac), "00000000") & "', "
      r_str_Parame = r_str_Parame & " '" & modgen_g_str_CodUsu & "' ,"
      r_str_Parame = r_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      r_str_Parame = r_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      r_str_Parame = r_str_Parame & " '" & UCase(App.EXEName) & "'                            , "
      r_str_Parame = r_str_Parame & " '" & modgen_g_str_NombPC & "'                           , "
      r_str_Parame = r_str_Parame & " '" & modgen_g_str_CodSuc & "')"

      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para insertar en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
         Exit Function
      End If
      
   Else
      'Update
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " UPDATE CNTBL_FOLIOS_DOCELE SET "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMSER = '" & Format(CStr(p_SerFac), "000") & "', "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMCOR = '" & Format(CStr(p_NumFac), "00000000") & "', "
      r_str_Parame = r_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "', "
      r_str_Parame = r_str_Parame & "        SEGFECACT = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      r_str_Parame = r_str_Parame & "        SEGHORACT = " & Format(Time, "HHMMSS") & ", "
      r_str_Parame = r_str_Parame & "        SEGPLTACT = '" & UCase(App.EXEName) & "', "
      r_str_Parame = r_str_Parame & "        SEGTERACT = '" & modgen_g_str_NombPC & "', "
      r_str_Parame = r_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "' "
      r_str_Parame = r_str_Parame & "  WHERE "
      r_str_Parame = r_str_Parame & "  DOCELE_TIPDOC = '" & CStr(p_TipDoc) & "' "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para actualizar en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
         Exit Function
      End If
   End If
   
   Set r_rst_Codigo = Nothing
   
End Function
Private Function fs_NroEnLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True) As String

Dim dblCentavos            As Double
Dim lngContDec             As Long
Dim lngContCent            As Long
Dim lngContMil             As Long
Dim lngContMillon          As Long
Dim strNumLetras           As String
Dim strNumero              As Variant
Dim strDecenas             As Variant
Dim strCentenas            As Variant
Dim blnNegativo            As Boolean
Dim blnPlural              As Boolean
                
    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If
    
    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop
    
    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop
    
    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If
    
    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
            strNumLetras = fs_NroEnLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON" & _
                                                                    IIf(blnPlural, "ES ", " ")
    End If
    
    If lngContMil > 0 Then
        
        If lngContMil = 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & fs_NroEnLetras(lngContMil, False)
            lngContMil = 0
            
        End If
        If lngContMil > 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & fs_NroEnLetras(lngContMil, False)
            lngContMil = 0
        End If
        'MsgBox strNumLetras
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
        'MsgBox strNumLetras
    End If
    
    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If
    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If
        
        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
    'MsgBox "Por Aqui"
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0# Then
            
                strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            Else

                'strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            End If
            fs_NroEnLetras = strNumLetras
            Exit Function
        End If
    End If
    
    If curNumero > 0# Then
    
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If
    
    If dblCentavos > 0# Then
        strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    'Else
    End If
    'If dblCentavos = 0# Then
        'MsgBox strNumLetras
        'strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    'End If
    
    fs_NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
    
End Function
Private Sub fs_Convertir_Utf8NoBom(p_sFile)

Dim UTFStream              As New ADODB.Stream
Dim ANSIStream             As New ADODB.Stream
Dim BinaryStream           As New ADODB.Stream

On Error GoTo MyError
    ANSIStream.Type = adTypeText
    ANSIStream.Mode = adModeReadWrite
    ANSIStream.Charset = "iso-8859-1"
    ANSIStream.Open
    ANSIStream.LoadFromFile p_sFile   'ANSI File
    
    UTFStream.Type = adTypeText
    UTFStream.Mode = adModeReadWrite
    UTFStream.Charset = "UTF-8"
    UTFStream.Open
    ANSIStream.CopyTo UTFStream
    

    UTFStream.Position = 3 'skip BOM
    BinaryStream.Type = adTypeBinary
    BinaryStream.Mode = adModeReadWrite
    BinaryStream.Open

    'Strips BOM (first 3 bytes)
    UTFStream.CopyTo BinaryStream

    BinaryStream.SaveToFile p_sFile, adSaveCreateOverWrite
    BinaryStream.Flush
    BinaryStream.Close
    Exit Sub
    
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", Convertir archivo a UTF8, procedimiento: fs_Genera_Utf8NoBom")
End Sub
'''Private Function fs_Cargar_Archivo(p_sRuta, p_sFile) As Boolean
'''
'''Dim r_str_success          As Long
'''Dim r_dbl_port             As Long
'''Dim r_str_HostName         As String
'''Dim r_str_RutServ          As String
'''Dim r_key                  As New ChilkatSshKey
'''Dim r_privKey              As String
'''
'''   On Error GoTo MyError
'''
'''   fs_Cargar_Archivo = False
'''
'''   Set r_chi_sftp = New ChilkatSFtp
'''
'''   r_str_success = r_chi_sftp.UnlockComponent("30")
'''   If (r_str_success <> 1) Then
'''       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''       Exit Function
'''   End If
'''
'''   'Set some timeouts, in milliseconds:
'''   r_chi_sftp.ConnectTimeoutMs = 5000
'''   r_chi_sftp.IdleTimeoutMs = 10000
'''
'''   '  Connect to the SSH server.
'''   '  The standard SSH port = 22
'''   '  The hostname may be a hostname or IP address.
'''
'''   '  Producción:
'''   '  Sftp.escondatagate.net (puerto 6022)
'''   '  Calidad:
'''   '  Sftpqa.escondatagate.net (puerto 3022)
'''
'''      r_str_HostName = " Sftp.escondatagate.net"
'''      r_dbl_port = 6022
'''
''''   r_str_HostName = "Sftpqa.escondatagate.net"
''''   r_dbl_port = 3022
'''
'''   r_str_success = r_chi_sftp.Connect(r_str_HostName, r_dbl_port)
'''   If (r_str_success <> 1) Then
'''       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''       Exit Function
'''   End If
'''
'''   'clave pública
'''   r_privKey = r_key.LoadText(moddat_g_str_RutFac & "\" & "id_rsa.ppk")
'''
'''   If (r_key.LastMethodSuccess <> 1) Then
'''      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''      Exit Function
'''   End If
'''
'''   r_str_success = r_key.FromOpenSshPrivateKey(r_privKey)
'''   If (r_str_success <> 1) Then
'''      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''      Exit Function
'''   End If
'''
'''   '  Authenticate with the SSH server.  Chilkat SFTP supports
'''   '  both password-based authenication as well as public-key
'''   '  authentication.  This example uses password authenication.
'''   r_str_success = r_chi_sftp.AuthenticatePw("micasi02", "Micasi2018*") 'r_chi_sftp.AuthenticatePw("micasi02", r_key)
'''
'''   If (r_str_success <> 1) Then
'''       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''       Exit Function
'''   End If
'''
'''   '  After authenticating, the SFTP subsystem must be initialized:
'''   r_str_success = r_chi_sftp.InitializeSftp()
'''   If (r_str_success <> 1) Then
'''       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''       Exit Function
'''   End If
'''
'''   r_str_RutServ = "/WWW/entrada/" & p_sFile
'''
'''   r_str_success = r_chi_sftp.UploadFileByName(r_str_RutServ, p_sRuta)
'''   If (r_str_success <> 1) Then
'''       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
'''       Exit Function
'''   End If
'''
'''   fs_Cargar_Archivo = True
'''   Exit Function
'''
'''MyError:
'''   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Cargar_Archivo")
'''End Function
Private Sub fs_Leer_Archivo_EnvSFTP(p_sFile)
Dim r_str_Cadena           As String
Dim r_arr_NumFac()         As moddat_tpo_Genera
Dim r_lng_Contad           As Long
Dim r_str_Parame           As String
Dim r_str_CadAux           As String

On Error GoTo Err

   ReDim r_arr_NumFac(0)
   Open p_sFile For Input As #1


   Do While Not EOF(1)
      Line Input #1, r_str_Cadena
      
      If InStr(r_str_Cadena, "IDE") > 0 Then

         ReDim Preserve r_arr_NumFac(UBound(r_arr_NumFac) + 1)
      
         r_str_CadAux = Trim(Mid(r_str_Cadena, InStr(r_str_Cadena, "IDE|") + Len("IDE|")))
         r_str_CadAux = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
         
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_Codigo = r_str_CadAux
      End If
   Loop
   
   Close #1

   'Actualiza el campo FACELE_FLGENV para identificar cuales se enviaron en el archivo
   For r_lng_Contad = 0 To UBound(r_arr_NumFac)
      If Len(Trim(r_arr_NumFac(r_lng_Contad).Genera_Codigo)) > 0 Then
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & "UPDATE CNTBL_FACELE SET FACELE_FLGENV = 1 "
         r_str_Parame = r_str_Parame & " WHERE FACELE_IDE_SERNUM = '" & r_arr_NumFac(r_lng_Contad).Genera_Codigo & "' "
         
         If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar FACELE_FLGENV de la tabla CNTBL_FACELE, procedimiento: fs_Leer_Archivo_EnvSFTP")
            Exit Sub
         End If
      End If
   Next
   Exit Sub
   
Err:
Close #1
Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Leer_Archivo_EnvSFTP")
Err.Clear

End Sub
Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   Else
      If cmb_Buscar.ListIndex = 1 Then
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub
