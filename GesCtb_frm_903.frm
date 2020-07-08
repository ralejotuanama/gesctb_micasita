VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_ComCie_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   8985
   ClientLeft      =   3615
   ClientTop       =   2445
   ClientWidth     =   11460
   Icon            =   "GesCtb_frm_903.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _Version        =   65536
      _ExtentX        =   20241
      _ExtentY        =   15901
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   600
            TabIndex        =   33
            Top             =   60
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Comerciales"
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
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   330
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Datos del Comercial"
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
            Picture         =   "GesCtb_frm_903.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   645
         Left            =   30
         TabIndex        =   35
         Top             =   720
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_903.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10800
            Picture         =   "GesCtb_frm_903.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_903.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Modificar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_903.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   495
         Left            =   30
         TabIndex        =   36
         Top             =   3390
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   873
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
         Begin EditLib.fpDoubleSingle ipp_TasInt 
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Top             =   90
            Width           =   3000
            _Version        =   196608
            _ExtentX        =   5292
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_TaInMo 
            Height          =   315
            Left            =   9300
            TabIndex        =   7
            Top             =   90
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin VB.Label Label5 
            Caption         =   "Tasa Interés Moratorio:"
            Height          =   285
            Left            =   7200
            TabIndex        =   38
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label22 
            Caption         =   "Tasa Interés:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   120
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   525
         Left            =   30
         TabIndex        =   39
         Top             =   1830
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   926
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   5100
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   345
            Left            =   9300
            MaxLength       =   11
            TabIndex        =   2
            Top             =   90
            Width           =   1920
         End
         Begin VB.Label Label31 
            Caption         =   "Número Documento:"
            Height          =   315
            Left            =   7200
            TabIndex        =   41
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label Label30 
            Caption         =   "Documento Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   150
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   30
         TabIndex        =   42
         Top             =   1380
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   1800
            TabIndex        =   43
            Top             =   60
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "2011-04"
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
         Begin VB.Label Label99 
            Caption         =   "Periodo:"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   120
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   45
         Top             =   4770
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin EditLib.fpDateTime ipp_FecDes 
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   90
            Width           =   3000
            _Version        =   196608
            _ExtentX        =   5292
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
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   1800
            TabIndex        =   14
            Top             =   420
            Width           =   3000
            _Version        =   196608
            _ExtentX        =   5292
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_TotPre 
            Height          =   315
            Left            =   9300
            TabIndex        =   15
            Top             =   420
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin EditLib.fpDateTime ipp_UltVct 
            Height          =   315
            Left            =   9300
            TabIndex        =   13
            Top             =   90
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
         Begin VB.Label Label9 
            Caption         =   "Fecha Ult. Vencimiento:"
            Height          =   315
            Left            =   7200
            TabIndex        =   65
            Top             =   120
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Monto Préstamo:"
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   450
            Width           =   1665
         End
         Begin VB.Label Label6 
            Caption         =   "Total Préstamo:"
            Height          =   285
            Left            =   7200
            TabIndex        =   62
            Top             =   450
            Width           =   1845
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Desembolso:"
            Height          =   315
            Left            =   90
            TabIndex        =   46
            Top             =   120
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1125
         Left            =   30
         TabIndex        =   47
         Top             =   6480
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   1984
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
         Begin VB.ComboBox cmb_NueCre 
            Height          =   315
            ItemData        =   "GesCtb_frm_903.frx":11AE
            Left            =   1800
            List            =   "GesCtb_frm_903.frx":11B0
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   90
            Width           =   5100
         End
         Begin VB.ComboBox cmb_SitCre 
            Height          =   315
            ItemData        =   "GesCtb_frm_903.frx":11B2
            Left            =   1800
            List            =   "GesCtb_frm_903.frx":11B4
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   420
            Width           =   3000
         End
         Begin VB.CheckBox chk_FlgJud 
            Caption         =   "Judicial"
            Height          =   285
            Left            =   6000
            TabIndex        =   22
            Top             =   780
            Width           =   1665
         End
         Begin VB.CheckBox chk_FlgCas 
            Caption         =   "Castigo"
            Height          =   285
            Left            =   9300
            TabIndex        =   23
            Top             =   780
            Width           =   1665
         End
         Begin VB.CheckBox chk_FlgRef 
            Caption         =   "Refinanciado"
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Top             =   780
            Width           =   1665
         End
         Begin VB.Label Label8 
            Caption         =   "Clasificación Crédito:"
            Height          =   315
            Left            =   90
            TabIndex        =   68
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Situación Crédito:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   450
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo de Crédito:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   780
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   855
         Left            =   30
         TabIndex        =   49
         Top             =   3900
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_MonGar 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   3000
         End
         Begin VB.ComboBox cmb_TipGar 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   90
            Width           =   3000
         End
         Begin EditLib.fpDoubleSingle ipp_Mtogar 
            Height          =   315
            Left            =   9300
            TabIndex        =   11
            Top             =   420
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_ValRea 
            Height          =   315
            Left            =   9300
            TabIndex        =   9
            Top             =   90
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Realización:"
            Height          =   195
            Left            =   7200
            TabIndex        =   73
            Top             =   150
            Width           =   1500
         End
         Begin VB.Label Label3 
            Caption         =   "Monto Garantía:"
            Height          =   315
            Left            =   7200
            TabIndex        =   52
            Top             =   480
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda Garantía:"
            Height          =   315
            Left            =   90
            TabIndex        =   51
            Top             =   450
            Width           =   1905
         End
         Begin VB.Label Label17 
            Caption         =   "Tipo Garantía:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   120
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   495
         Left            =   30
         TabIndex        =   53
         Top             =   2370
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   873
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
         Begin VB.ComboBox cmb_CodPrd 
            Height          =   315
            ItemData        =   "GesCtb_frm_903.frx":11B6
            Left            =   1800
            List            =   "GesCtb_frm_903.frx":11B8
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   90
            Width           =   5100
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   9300
            TabIndex        =   4
            Top             =   90
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            Caption         =   "Número de Operación:"
            Height          =   315
            Left            =   7200
            TabIndex        =   58
            Top             =   150
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   54
            Top             =   150
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   495
         Left            =   30
         TabIndex        =   55
         Top             =   8460
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   873
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            ItemData        =   "GesCtb_frm_903.frx":11BA
            Left            =   1800
            List            =   "GesCtb_frm_903.frx":11BC
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   90
            Width           =   3000
         End
         Begin VB.Label Label2 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   120
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   825
         Left            =   30
         TabIndex        =   57
         Top             =   7620
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin EditLib.fpDoubleSingle ipp_SalCap 
            Height          =   315
            Left            =   1800
            TabIndex        =   25
            Top             =   420
            Width           =   3000
            _Version        =   196608
            _ExtentX        =   5292
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_IntDev 
            Height          =   315
            Left            =   9300
            TabIndex        =   26
            Top             =   420
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_LinCre 
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   90
            Width           =   3000
            _Version        =   196608
            _ExtentX        =   5292
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
            MinValue        =   "0"
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
         Begin VB.Label Label13 
            Caption         =   "Saldo Linea Credito:"
            Height          =   285
            Left            =   90
            TabIndex        =   67
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label Label11 
            Caption         =   "Int. Devengado:"
            Height          =   315
            Left            =   7200
            TabIndex        =   66
            Top             =   480
            Width           =   1665
         End
         Begin VB.Label Label18 
            Caption         =   "Saldo Capital:"
            Height          =   315
            Left            =   90
            TabIndex        =   64
            Top             =   480
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   495
         Left            =   30
         TabIndex        =   59
         Top             =   2880
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   873
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
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            ItemData        =   "GesCtb_frm_903.frx":11BE
            Left            =   1800
            List            =   "GesCtb_frm_903.frx":11C0
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   90
            Width           =   9435
         End
         Begin VB.Label Label15 
            Caption         =   "Actividad Económica:"
            Height          =   315
            Left            =   90
            TabIndex        =   60
            Top             =   120
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   855
         Left            =   30
         TabIndex        =   69
         Top             =   5610
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_CreInd 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   90
            Width           =   3000
         End
         Begin VB.ComboBox cmb_MonCre 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   420
            Width           =   3000
         End
         Begin EditLib.fpDoubleSingle ipp_MtoCre 
            Height          =   315
            Left            =   9300
            TabIndex        =   18
            Top             =   420
            Width           =   1920
            _Version        =   196608
            _ExtentX        =   3387
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
            MinValue        =   "0"
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
         Begin VB.Label Label21 
            Caption         =   "Credito Indirecto:"
            Height          =   315
            Left            =   90
            TabIndex        =   72
            Top             =   120
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Moneda Cred. Ind.:"
            Height          =   315
            Left            =   90
            TabIndex        =   71
            Top             =   450
            Width           =   1905
         End
         Begin VB.Label Label19 
            Caption         =   "Monto Cred. Indirecto:"
            Height          =   315
            Left            =   7200
            TabIndex        =   70
            Top             =   480
            Width           =   1905
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_ComCie_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()   As moddat_tpo_Genera
Dim r_lng_FecCie     As Long
Dim r_dbl_TipCam     As Double
Dim r_int_PerGra     As Integer
Dim r_dbl_PrvGen     As Double
Dim r_dbl_PrvEsp     As Double
Dim r_dbl_PrvCam     As Double
Dim r_dbl_PrvCic     As Double
Dim r_dbl_PrvAdc     As Double
Dim r_dbl_FecDev     As Double
Dim r_dbl_DevVig     As Double
Dim r_dbl_DevVen     As Double
Dim r_dbl_AcuDvg     As Double
Dim r_dbl_AcuDvc     As Double
Dim r_int_Exporc     As Integer
Dim l_str_Cadena     As String

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   cmb_CodPrd.ListIndex = -1
   msk_NumOpe.Text = ""
   cmb_CodCiu.ListIndex = -1
   ipp_TasInt.Text = 0
   ipp_TaInMo.Text = 0
   cmb_TipGar.ListIndex = -1
   ipp_ValRea.Text = 0
   cmb_MonGar.ListIndex = -1
   ipp_Mtogar.Text = 0
   ipp_FecDes.Text = Format(date, "dd/mm/yyyy")
   ipp_UltVct.Text = Format(date, "dd/mm/yyyy")
   ipp_MtoPre.Text = 0
   ipp_TotPre.Text = 0
   cmb_CreInd.ListIndex = -1
   cmb_MonCre.ListIndex = -1
   ipp_MtoCre.Text = 0
   cmb_NueCre.ListIndex = -1
   cmb_SitCre.ListIndex = -1
   chk_FlgRef.Value = 0
   chk_FlgJud.Value = 0
   chk_FlgCas.Value = 0
   ipp_SalCap.Text = 0
   ipp_LinCre.Text = 0
   cmb_Situac.ListIndex = -1
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")
   Call gs_Buscar_Produc(cmb_CodPrd, l_arr_Produc, 1)
   Call gs_Buscar_CodCiu(cmb_CodCiu, 1, "102")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipGar, 1, "241")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonGar, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonCre, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call gs_Carga_TipCre(cmb_NueCre)
      
   cmb_CreInd.Clear
   cmb_CreInd.AddItem "CARTA FIANZA"
   cmb_CreInd.ItemData(cmb_CreInd.NewIndex) = 4
   cmb_CreInd.ListIndex = -1
End Sub

Private Sub gs_Carga_TipCre(p_Combo As ComboBox)
   p_Combo.Clear
   l_str_Cadena = "SELECT * FROM CTB_TIPCRE "
   l_str_Cadena = l_str_Cadena & "ORDER BY TIPCRE_CODIGO ASC "
   
   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!TIPCRE_CODIGO) & " - " & Trim$(g_rst_Genera!TIPCRE_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!TIPCRE_CODIGO)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmd_Editar_Click()
   Call fs_Activa(True)
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   cmb_CodPrd.Enabled = False
   msk_NumOpe.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_TipCre As Integer
   
   If moddat_g_int_FlgGrb = 1 Then
      r_lng_FecCie = 0
      r_dbl_TipCam = 0
      r_int_PerGra = 0
      r_dbl_PrvGen = 0
      r_dbl_PrvEsp = 0
      r_dbl_PrvCam = 0
      r_dbl_PrvCic = 0
      r_dbl_PrvAdc = 0
      r_dbl_DevVig = 0
      r_dbl_DevVen = 0
      r_dbl_AcuDvc = 0
      r_int_Exporc = 1
   End If
   
   'validacion general
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   If cmb_CodPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodPrd)
      Exit Sub
   End If
   If Len(Trim(msk_NumOpe.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(msk_NumOpe)
      Exit Sub
   End If
   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   If cmb_TipGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipGar)
      Exit Sub
   End If
   If cmb_NueCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NueCre)
      Exit Sub
   End If
   If cmb_SitCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitCre)
      Exit Sub
   End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   'validacion de tipo de garantia
   Select Case cmb_TipGar.ItemData(cmb_TipGar.ListIndex)
      'Cuando es HIPOTECA
      Case 1
         If ipp_TasInt.Text = 0 Then
            MsgBox "Debe ingresar el Tasa de Interés.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_TasInt)
            Exit Sub
         End If
         If ipp_TaInMo.Text = 0 Then
            MsgBox "Debe ingresar el Tasa de Interés Moratorio.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_TaInMo)
            Exit Sub
         End If
         If cmb_TipGar.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Garantía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipGar)
            Exit Sub
         End If
         If cmb_MonGar.ListIndex = -1 Then
            MsgBox "Debe seleccionar la moneda de la Garantía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_MonGar)
            Exit Sub
         End If
         'If ipp_Mtogar.Text = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Garantía.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Mtogar)
         '   Exit Sub
         'End If
         If chk_FlgRef.Value = 1 And chk_FlgJud.Value = 1 And chk_FlgCas.Value = 1 Then
            MsgBox "Solo se puede seleccionar un Tipo de Crédito a la vez.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         ElseIf chk_FlgRef.Value = 1 And chk_FlgJud.Value = 1 Then
            MsgBox "No se puede seleccionar Refinanciado y Judicial a la vez.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         ElseIf chk_FlgRef.Value = 1 And chk_FlgCas.Value = 1 Then
            MsgBox "No se puede seleccionar Refinanciado y Castigo a la vez.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         ElseIf chk_FlgJud.Value = 1 And chk_FlgCas.Value = 1 Then
            MsgBox "No se puede seleccionar Judicial y Castigo a la vez.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         If ipp_SalCap.Text = 0 Then
            MsgBox "Debe ingresar el Saldo Capital.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_SalCap)
            Exit Sub
         End If
         If ipp_LinCre.Text = 0 Then
            MsgBox "El campo 'Saldo Linea de Credito' esta sin data.", vbExclamation, modgen_g_str_NomPlt
         End If
         
         'Valida Nivel de Endeudamiento
         If moddat_gf_Consulta_NivelEndeudamiento(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, Right(modsec_g_str_Period, 2), Left(modsec_g_str_Period, 4), ipp_MtoPre.Value) = True Then
            MsgBox "El Valor del préstamo sobrepasa el nivel de endeudamiento permitido según norma SBS en Créditos Comerciales y/o Cartas Fianza." & vbCrLf & "Favor consultar con el área de Riesgos", vbExclamation, modgen_g_str_NomPlt
            If MsgBox("¿Desea continuar a pesar de sobrepasar el limite permitido por norma SBS?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Call gs_SetFocus(ipp_MtoPre)
               Exit Sub
            End If
         End If
   
         If moddat_g_int_FlgGrb = 1 Then
            If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            End If
         Else
            If MsgBox("¿Está seguro de modificar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            End If
         End If
         
         If cmb_SitCre.ItemData(cmb_SitCre.ListIndex) = 1 Then
            r_int_TipCre = 1
         ElseIf cmb_SitCre.ItemData(cmb_SitCre.ListIndex) = 5 And chk_FlgRef.Value = 0 And chk_FlgJud.Value = 0 And chk_FlgCas.Value = 0 Then
            r_int_TipCre = 1
         End If
         If chk_FlgRef.Value = 1 Then
            r_int_TipCre = 4
         ElseIf chk_FlgJud.Value = 1 Then
            r_int_TipCre = 6
         ElseIf chk_FlgCas.Value = 1 Then
            r_int_TipCre = 3
         End If
         
      'Cuando es CARTA FIANZA
      Case 4
         If cmb_CreInd.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Credito Indirecto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_CreInd)
            Exit Sub
         End If
         If cmb_MonCre.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de moneda del Credito Indirecto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_MonCre)
            Exit Sub
         End If
         If ipp_MtoCre.Text = 0 Then
            MsgBox "Debe ingresar el Monto del Credito del indirecto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoCre)
            Exit Sub
         End If
      
      'Cuando es PAGARE
      Case 10
      
      
      'Cuando es OTRO
      Case Else
         MsgBox "Solo se permite la selección de HIPOTECA ó CARTA FIANZA.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
         
   End Select
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = "USP_CRE_COMCIE ("
   g_str_Parame = g_str_Parame & Right(modsec_g_str_Period, 2) & ", "                                                            'comcie_permes
   g_str_Parame = g_str_Parame & Left(modsec_g_str_Period, 4) & ", "                                                             'comcie_perano
   g_str_Parame = g_str_Parame & "'" & Trim(msk_NumOpe.Text) & "', "                                                             'comcie_numope
   g_str_Parame = g_str_Parame & "'" & "000001" & "', "                                                                          'comcie_codemp
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_codtit
   g_str_Parame = g_str_Parame & r_lng_FecCie & ", "                                                                             'comcie_feccie
   g_str_Parame = g_str_Parame & r_dbl_TipCam & ", "                                                                             'comcie_tipcam
   If l_arr_Produc(cmb_CodPrd.ListIndex + 1).Genera_Codigo = "008" Then
      g_str_Parame = g_str_Parame & 1 & ", "                                                                                     'comcie_tipmon
   ElseIf l_arr_Produc(cmb_CodPrd.ListIndex + 1).Genera_Codigo = "005" Then
      g_str_Parame = g_str_Parame & 2 & ", "                                                                                     'comcie_tipmon
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_MtoPre.Text) & ", "                                                                    'comcie_mtopre
   g_str_Parame = g_str_Parame & CDbl(ipp_TotPre.Text) & ", "                                                                    'comcie_totpre
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_plames
   g_str_Parame = g_str_Parame & "'" & gs_Buscar_ClaPrd(l_arr_Produc(cmb_CodPrd.ListIndex + 1).Genera_Codigo) & "', "            'comcie_claprd
   g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_CodPrd.ListIndex + 1).Genera_Codigo & "', "                              'comcie_codprd
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_codsub
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_ubigeo
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_codpry
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_prymcs
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_acteco
   If cmb_CodCiu.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "'0', "
   Else
      g_str_Parame = g_str_Parame & "'" & cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex) & "', "                                      'comcie_codciu
   End If
   g_str_Parame = g_str_Parame & "'" & "', "                                                                                     'comcie_sececo
   g_str_Parame = g_str_Parame & r_int_PerGra & ", "                                                                             'comcie_pergra
   g_str_Parame = g_str_Parame & CDbl(ipp_TasInt.Text) & ", "                                                                    'comcie_cosefe
   g_str_Parame = g_str_Parame & CDbl(ipp_TasInt.Text) & ", "                                                                    'comcie_tasint
   g_str_Parame = g_str_Parame & CDbl(ipp_TaInMo.Text) & ", "                                                                    'comcie_tasmor
   If cmb_TipGar.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_tipgar
   Else
      g_str_Parame = g_str_Parame & cmb_TipGar.ItemData(cmb_TipGar.ListIndex) & ", "                                             'comcie_tipgar
   End If
   If cmb_MonGar.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_mongar
   Else
      g_str_Parame = g_str_Parame & cmb_MonGar.ItemData(cmb_MonGar.ListIndex) & ", "                                             'comcie_mongar
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_Mtogar.Text) & ", "                                                                    'comcie_mtogar
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecDes.Text), "yyyymmdd") & ", "                                               'comcie_fecdes
   If cmb_Situac.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_situac
   Else
      g_str_Parame = g_str_Parame & cmb_Situac.ItemData(cmb_Situac.ListIndex) & ", "                                             'comcie_situac
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_SalCap.Text) & ", "                                                                    'comcie_salcap
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_diamor
   If cmb_SitCre.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_sitcre
   Else
      g_str_Parame = g_str_Parame & cmb_SitCre.ItemData(cmb_SitCre.ListIndex) & ", "                                             'comcie_sitcre
   End If
   g_str_Parame = g_str_Parame & r_int_TipCre & ", "                                                                             'comcie_tipcre
   g_str_Parame = g_str_Parame & chk_FlgRef.Value & ", "                                                                         'comcie_flgref
   g_str_Parame = g_str_Parame & chk_FlgJud.Value & ", "                                                                         'comcie_flgjud
   g_str_Parame = g_str_Parame & chk_FlgCas.Value & ", "                                                                         'comcie_flgcas
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_clacre
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_clacli
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_claali
   g_str_Parame = g_str_Parame & CDbl(r_dbl_PrvGen) & ", "                                                                       'comcie_prvgen
   g_str_Parame = g_str_Parame & CDbl(r_dbl_PrvEsp) & ", "                                                                       'comcie_prvesp
   g_str_Parame = g_str_Parame & CDbl(r_dbl_PrvCam) & ", "                                                                       'comcie_prvcam
   g_str_Parame = g_str_Parame & CDbl(r_dbl_PrvCic) & ", "                                                                       'comcie_prvcic
   g_str_Parame = g_str_Parame & CDbl(r_dbl_PrvAdc) & ", "                                                                       'comcie_prvadc
   g_str_Parame = g_str_Parame & CDbl(r_dbl_FecDev) & ", "                                                                       'comcie_fecdev
   g_str_Parame = g_str_Parame & CDbl(r_dbl_DevVig) & ", "                                                                       'comcie_devvig
   g_str_Parame = g_str_Parame & CDbl(r_dbl_DevVen) & ", "                                                                       'comcie_devven
   g_str_Parame = g_str_Parame & CDbl(ipp_IntDev.Text) & ", "                                                                    'comcie_acudvg
   g_str_Parame = g_str_Parame & CDbl(r_dbl_AcuDvc) & ", "                                                                       'comcie_acudvc
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "                                                'comcie_tdocli
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "                                                             'comcie_ndocli
   If cmb_SitCre.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_claprv
   Else
      g_str_Parame = g_str_Parame & cmb_SitCre.ItemData(cmb_SitCre.ListIndex) - 1 & ", "                                         'comcie_claprv
   End If
   
   g_str_Parame = g_str_Parame & r_int_Exporc & ", "                                                                             'comcie_exporc
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_capamo
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_fecamo
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecDes.Text), "yyyymmdd") & ", "                                               'comcie_aprcre
   g_str_Parame = g_str_Parame & Format(CDate(ipp_UltVct.Text), "yyyymmdd") & ", "                                               'comcie_ultvct
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_cuoatr
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_cuopen
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_tippag
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_cuopag
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_intdif
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_capven
   g_str_Parame = g_str_Parame & CDbl(ipp_SalCap.Text) & ", "                                                                    'comcie_capvig
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_vctant
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_prxvct
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_ultpag
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_fvgant
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_intcom
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_intmor
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_gascob
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_otrgas
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_ucppag
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_imovig
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_gcovig
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_otgvig
   g_str_Parame = g_str_Parame & 0 & ", "                                                                                        'comcie_acudif
   If cmb_NueCre.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "'000', "                                                                                    'comcie_nuecre
   Else
      g_str_Parame = g_str_Parame & "'" & Format(cmb_NueCre.ItemData(cmb_NueCre.ListIndex), "000") & "', "                       'comcie_nuecre
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_LinCre.Text) & ", "                                                                    'comcie_lincre
   If cmb_CreInd.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_creind
   Else
      g_str_Parame = g_str_Parame & cmb_CreInd.ItemData(cmb_CreInd.ListIndex) & ", "                                             'comcie_creind
   End If
   If cmb_MonCre.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "0, "                                                                                        'comcie_moncre
   Else
      g_str_Parame = g_str_Parame & cmb_MonCre.ItemData(cmb_MonCre.ListIndex) & ", "                                             'comcie_moncre
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_MtoCre.Text) & ", "                                                                    'comcie_mtocre
   g_str_Parame = g_str_Parame & CDbl(ipp_ValRea.Text) & ", "                                                                    'comcie_mtorea
 
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & moddat_g_int_FlgGrb & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CRE_COMCIE.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   Else
      MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   'moddat_g_int_FlgGrb = 2
   'Call fs_Activa(False)
   Call cmd_Salida_Click
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Cargar_Datos
   Call gs_SetFocus(cmd_Editar)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Period.Caption = modsec_g_str_Period
   
   Call gs_CentraForm(Me)
   Call fs_Limpia
   Call fs_Inicia
   
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_Activa(True)
      cmd_Cancel.Enabled = False
   ElseIf moddat_g_int_FlgGrb = 2 Then
      Call fs_Activa(False)
      Call fs_Cargar_Datos
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub gs_Carga_SitCre(p_Combo As ComboBox, ByVal p_ClaCre As Integer)
   p_Combo.Clear
   l_str_Cadena = "SELECT * FROM CTB_SITCRE "
   l_str_Cadena = l_str_Cadena & "WHERE SITCRE_CLACRE = " & p_ClaCre & " "
   l_str_Cadena = l_str_Cadena & "ORDER BY SITCRE_CODSIT ASC "
   
   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!SITCRE_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!SITCRE_CODSIT)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Function gs_Buscar_ClaPrd(ByVal p_CodPrd As String) As Integer
   l_str_Cadena = "SELECT * FROM CRE_PRODUC "
   l_str_Cadena = l_str_Cadena & "WHERE PRODUC_CODIGO = '" & p_CodPrd & "' "
      
   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      gs_Buscar_ClaPrd = CInt(g_rst_Genera!PRODUC_CODCLA)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub gs_Buscar_Produc(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipCre As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   l_str_Cadena = "SELECT * FROM CRE_PRODUC WHERE PRODUC_SITUAC = 1 "
   
   If p_TipCre <> 99 Then
      l_str_Cadena = l_str_Cadena & "AND PRODUC_CODCLA = " & CStr(p_TipCre) & " "
   End If
   
   l_str_Cadena = l_str_Cadena & "ORDER BY PRODUC_CODIGO ASC"

   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!Produc_Codigo) & " - " & Trim$(g_rst_Listas!PRODUC_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Listas!Produc_Codigo)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Listas!PRODUC_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Listas!PRODUC_CODCLA
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub gs_Buscar_CodCiu(p_Combo As ComboBox, ByVal p_TipTab As Integer, ByVal p_CodGrp As String, Optional ByVal p_TipOrd As Integer)
   'p_TipTab   =  1  -  MNT$PARDES - Parámetros de Descripción
   'p_TipTab   =  2  -  MNT$PARVAL - Parámetros de Valor
   
   p_Combo.Clear
   Select Case p_TipTab
      Case 1
         l_str_Cadena = "SELECT * FROM MNT_PARDES WHERE "
         l_str_Cadena = l_str_Cadena & "PARDES_CODGRP = '" & p_CodGrp & "' AND "
         l_str_Cadena = l_str_Cadena & "PARDES_CODITE <> '000000' AND "
         l_str_Cadena = l_str_Cadena & "PARDES_SITUAC = 1 "
         If p_TipOrd = 1 Then
            l_str_Cadena = l_str_Cadena & "ORDER BY PARDES_DESCRI ASC"
         Else
            l_str_Cadena = l_str_Cadena & "ORDER BY PARDES_CODITE ASC"
         End If
      
      Case 2
         l_str_Cadena = "SELECT * FROM MNT_PARVAL WHERE "
         l_str_Cadena = l_str_Cadena & "PARVAL_CODGRP = '" & p_CodGrp & "' AND "
         l_str_Cadena = l_str_Cadena & "PARVAL_CODITE <> '000000' AND "
         l_str_Cadena = l_str_Cadena & "PARVAL_SITUAC = 1 "
         If p_TipOrd = 1 Then
            l_str_Cadena = l_str_Cadena & "ORDER BY PARVAL_DESCRI ASC"
         Else
            l_str_Cadena = l_str_Cadena & "ORDER BY PARVAL_CODITE ASC"
         End If
   End Select

   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      Select Case p_TipTab
         Case 1
            p_Combo.AddItem Trim$(g_rst_Genera!PARDES_CODITE) & " - " & Trim$(g_rst_Genera!PARDES_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
         Case 2
            p_Combo.AddItem Trim$(g_rst_Genera!PARVAL_CODITE) & " - " & Trim$(g_rst_Genera!PARVAL_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARVAL_CODITE)
      End Select
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Cargar_Datos()
   l_str_Cadena = "SELECT * FROM CRE_COMCIE WHERE "
   l_str_Cadena = l_str_Cadena & "COMCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   l_str_Cadena = l_str_Cadena & "COMCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND "
   l_str_Cadena = l_str_Cadena & "COMCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "
         
   If Not gf_EjecutaSQL(l_str_Cadena, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!COMCIE_TDOCLI)
      txt_NumDoc.Text = Trim(g_rst_Princi!COMCIE_NDOCLI)
      cmb_CodPrd.ListIndex = gf_Busca_Arregl(l_arr_Produc, g_rst_Princi!comcie_codprd) - 1
      'Call gs_BuscarCombo_Item(cmb_CodPrd, g_rst_Princi!comcie_codprd)
      msk_NumOpe.Text = Trim(g_rst_Princi!COMCIE_NUMOPE)
      Call gs_BuscarCombo_Item(cmb_CodCiu, g_rst_Princi!comcie_codciu)
      ipp_TasInt.Text = Trim(g_rst_Princi!COMCIE_TASINT)
      ipp_TaInMo.Text = Trim(g_rst_Princi!comcie_tasmor)
      Call gs_BuscarCombo_Item(cmb_TipGar, g_rst_Princi!comcie_tipgar)
      Call gs_BuscarCombo_Item(cmb_MonGar, g_rst_Princi!COMCIE_MONGAR)
      ipp_Mtogar.Text = Trim(g_rst_Princi!COMCIE_MTOGAR)
      Call gs_BuscarCombo_Item(cmb_CreInd, g_rst_Princi!COMCIE_CREIND)
      Call gs_BuscarCombo_Item(cmb_MonCre, g_rst_Princi!COMCIE_MONCRE)
      ipp_MtoCre.Text = Trim(g_rst_Princi!COMCIE_MTOCRE)
      Call gs_BuscarCombo_Item(cmb_NueCre, g_rst_Princi!COMCIE_NUECRE)
      Call gs_BuscarCombo_Item(cmb_SitCre, g_rst_Princi!COMCIE_SITCRE)
      ipp_FecDes.Text = Right(CStr(g_rst_Princi!COMCIE_FECDES), 2) & "/" & Mid(CStr(g_rst_Princi!COMCIE_FECDES), 5, 2) & "/" & Left(CStr(g_rst_Princi!COMCIE_FECDES), 4)
      ipp_UltVct.Text = Right(CStr(g_rst_Princi!COMCIE_ULTVCT), 2) & "/" & Mid(CStr(g_rst_Princi!COMCIE_ULTVCT), 5, 2) & "/" & Left(CStr(g_rst_Princi!COMCIE_ULTVCT), 4)
      ipp_MtoPre.Text = Trim(g_rst_Princi!COMCIE_MTOPRE)
      ipp_TotPre.Text = Trim(g_rst_Princi!comcie_totpre)
      chk_FlgRef.Value = Trim(g_rst_Princi!COMCIE_FLGREF)
      chk_FlgJud.Value = Trim(g_rst_Princi!comcie_flgjud)
      chk_FlgCas.Value = Trim(g_rst_Princi!comcie_flgcas)
      ipp_SalCap.Text = Trim(g_rst_Princi!COMCIE_SALCAP)
      ipp_LinCre.Text = Trim(g_rst_Princi!COMCIE_LINCRE)
      ipp_IntDev.Text = Trim(g_rst_Princi!COMCIE_ACUDVG)
      Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!COMCIE_SITUAC)
      ipp_ValRea.Text = Trim(g_rst_Princi!COMCIE_MTOREA)
      
      r_lng_FecCie = Trim(g_rst_Princi!comcie_feccie)
      r_dbl_TipCam = Trim(g_rst_Princi!COMCIE_TIPCAM)
      r_int_PerGra = Trim(g_rst_Princi!COMCIE_PERGRA)
      r_dbl_PrvGen = Trim(g_rst_Princi!COMCIE_PRVGEN)
      r_dbl_PrvEsp = Trim(g_rst_Princi!COMCIE_PRVESP)
      r_dbl_PrvCam = Trim(g_rst_Princi!COMCIE_PRVCAM)
      r_dbl_PrvCic = Trim(g_rst_Princi!COMCIE_PRVCIC)
      r_dbl_PrvAdc = Trim(g_rst_Princi!COMCIE_PRVADC)
      r_dbl_FecDev = Trim(g_rst_Princi!comcie_fecdev)
      r_dbl_DevVig = Trim(g_rst_Princi!comcie_devvig)
      r_dbl_DevVen = Trim(g_rst_Princi!comcie_devven)
      r_dbl_AcuDvg = Trim(g_rst_Princi!COMCIE_ACUDVG)
      r_dbl_AcuDvc = Trim(g_rst_Princi!COMCIE_ACUDVC)
      r_int_Exporc = Trim(g_rst_Princi!comcie_exporc)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   cmb_CodPrd.Enabled = p_Habilita
   msk_NumOpe.Enabled = p_Habilita
   cmb_CodCiu.Enabled = p_Habilita
   ipp_TasInt.Enabled = p_Habilita
   ipp_TaInMo.Enabled = p_Habilita
   cmb_TipGar.Enabled = p_Habilita
   ipp_ValRea.Enabled = p_Habilita
   cmb_MonGar.Enabled = p_Habilita
   ipp_Mtogar.Enabled = p_Habilita
   cmb_CreInd.Enabled = p_Habilita
   cmb_MonCre.Enabled = p_Habilita
   ipp_MtoCre.Enabled = p_Habilita
   ipp_FecDes.Enabled = p_Habilita
   ipp_MtoPre.Enabled = p_Habilita
   ipp_TotPre.Enabled = p_Habilita
   cmb_SitCre.Enabled = p_Habilita
   chk_FlgRef.Enabled = p_Habilita
   chk_FlgJud.Enabled = p_Habilita
   chk_FlgCas.Enabled = p_Habilita
   ipp_SalCap.Enabled = p_Habilita
   ipp_LinCre.Enabled = p_Habilita
   cmb_NueCre.Enabled = p_Habilita
   cmb_Situac.Enabled = p_Habilita
   ipp_IntDev.Enabled = p_Habilita
   ipp_UltVct.Enabled = p_Habilita
   cmd_Editar.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & " ")
   End If
End Sub

Private Sub ipp_ValRea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonGar)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodPrd)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CodPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(msk_NumOpe)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasInt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub ipp_TasInt_GotFocus()
   Call gs_SelecTodo(ipp_TasInt)
End Sub

Private Sub ipp_TasInt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TaInMo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub ipp_TaInMo_GotFocus()
   Call gs_SelecTodo(ipp_TaInMo)
End Sub

Private Sub ipp_TaInMo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipGar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub cmb_TipGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub cmb_MonGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Mtogar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub ipp_Mtogar_GotFocus()
   Call gs_SelecTodo(ipp_Mtogar)
End Sub

Private Sub ipp_Mtogar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecDes)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub ipp_FecDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_UltVct)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub ipp_UltVct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoPre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub ipp_MtoPre_GotFocus()
   Call gs_SelecTodo(ipp_MtoPre)
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TotPre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ".,")
   End If
End Sub

Private Sub ipp_TotPre_GotFocus()
   Call gs_SelecTodo(ipp_TotPre)
End Sub

Private Sub ipp_TotPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CreInd)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub cmb_CreInd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonCre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub cmb_MonCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoCre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ".,")
   End If
End Sub

Private Sub ipp_MtoCre_GotFocus()
   Call gs_SelecTodo(ipp_MtoCre)
End Sub

Private Sub ipp_MtoCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NueCre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ".,")
   End If
End Sub

Private Sub cmb_NueCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SitCre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub cmb_SitCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_FlgRef)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub chk_FlgRef_Click()
   Call gs_SetFocus(ipp_LinCre)
End Sub

Private Sub chk_FlgJud_Click()
   Call gs_SetFocus(ipp_LinCre)
End Sub

Private Sub chk_FlgCas_Click()
   Call gs_SetFocus(ipp_LinCre)
End Sub

Private Sub ipp_LinCre_GotFocus()
   Call gs_SelecTodo(ipp_LinCre)
End Sub

Private Sub ipp_LinCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SalCap)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ".,")
   End If
End Sub

Private Sub ipp_SalCap_GotFocus()
   Call gs_SelecTodo(ipp_SalCap)
End Sub

Private Sub ipp_SalCap_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntDev)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub ipp_IntDev_GotFocus()
   Call gs_SelecTodo(ipp_IntDev)
End Sub

Private Sub ipp_IntDev_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & ",.")
   End If
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub cmb_NueCre_Click()
   If cmb_NueCre.ListIndex <> -1 Then
      Call gs_Carga_SitCre(cmb_SitCre, cmb_NueCre.ItemData(cmb_NueCre.ListIndex))
   End If
End Sub

