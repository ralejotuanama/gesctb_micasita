VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_GesPer_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "GesCtb_frm_202.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6020
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8090
      _Version        =   65536
      _ExtentX        =   14270
      _ExtentY        =   10619
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   19
         Top             =   60
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   300
            Left            =   660
            TabIndex        =   20
            Top             =   150
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Gestión Personal"
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
            Picture         =   "GesCtb_frm_202.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   21
         Top             =   780
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7320
            Picture         =   "GesCtb_frm_202.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_202.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_Datos 
         Height          =   3975
         Left            =   60
         TabIndex        =   22
         Top             =   1500
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   7011
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
         Begin VB.TextBox txt_Observ 
            Height          =   930
            Left            =   1590
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "GesCtb_frm_202.frx":0B9A
            Top             =   3360
            Width           =   5850
         End
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            Top             =   1050
            Width           =   5910
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   5910
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2370
            Width           =   2880
         End
         Begin VB.ComboBox cmb_TipOpe 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2040
            Width           =   5910
         End
         Begin Threed.SSPanel pnl_CodPla 
            Height          =   315
            Left            =   1590
            TabIndex        =   3
            Top             =   1380
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin EditLib.fpDateTime ipp_FecOpe 
            Height          =   315
            Left            =   1590
            TabIndex        =   4
            Top             =   1710
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin EditLib.fpDoubleSingle ipp_Importe 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   2700
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   315
            Left            =   5970
            TabIndex        =   5
            Top             =   1710
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin EditLib.fpDateTime ipp_FecIniVac 
            Height          =   315
            Left            =   4260
            TabIndex        =   10
            Top             =   2700
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin EditLib.fpDateTime ipp_FecFinVac 
            Height          =   315
            Left            =   6480
            TabIndex        =   11
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
         Begin Threed.SSPanel pnl_Saldo 
            Height          =   315
            Left            =   5910
            TabIndex        =   12
            Top             =   2370
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
            Alignment       =   4
         End
         Begin EditLib.fpLongInteger ipp_DiaSol 
            Height          =   315
            Left            =   1590
            TabIndex        =   9
            Top             =   3030
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
            ButtonStyle     =   1
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
            Text            =   "7"
            MaxValue        =   "90"
            MinValue        =   "1"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label lbl_Dia_Sol 
            AutoSize        =   -1  'True
            Caption         =   "Dias Solicitado:"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label lbl_Descrip 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   3510
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   -1410
            TabIndex        =   38
            Top             =   3300
            Width           =   570
         End
         Begin VB.Label lbl_Hasta 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   5910
            TabIndex        =   37
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label lbl_TipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   4890
            TabIndex        =   36
            Top             =   1770
            Width           =   930
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   810
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código Interno:"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código Planilla:"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   28
            Top             =   1470
            Width           =   1080
         End
         Begin VB.Label lbl_Importe 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   2760
            Width           =   570
         End
         Begin VB.Label lbl_Moneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   2430
            Width           =   630
         End
         Begin VB.Label lbl_FecOpe 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operación:"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   1770
            Width           =   1275
         End
         Begin VB.Label lbl_TipOpe 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operación:"
            Height          =   255
            Left            =   150
            TabIndex        =   24
            Top             =   2070
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   1140
            Width           =   600
         End
      End
      Begin Threed.SSPanel pnl_Abono 
         Height          =   1215
         Left            =   60
         TabIndex        =   32
         Top             =   4710
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   2143
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
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   750
            Width           =   2880
         End
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   420
            Width           =   2880
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   180
            TabIndex        =   41
            Top             =   450
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Abono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   90
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   34
            Top             =   -2160
            Width           =   510
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   795
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_GesPer_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_int_Contar   As Integer

Private Sub cmb_Proveedor_Click()
    Call fs_Buscar_prov
End Sub

Private Sub cmb_Proveedor_LostFocus()
    'Call fs_Buscar_prov
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
End Sub

Private Sub fs_Saldo_Devengado(ByRef p_Saldo As Integer, ByRef p_Vencidos As Integer)
Dim r_rst_Princi     As ADODB.Recordset
Dim r_str_Parame     As String

  p_Saldo = 0
  p_Vencidos = -1
  
  r_str_Parame = ""
  'r_str_Parame = r_str_Parame & " SELECT DECODE(A.MAEPRV_FECING,'',0, ROUND((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/360),0)) - "
  'r_str_Parame = r_str_Parame & "        (NVL(C.GESPER_DIAGOZ,0) + ROUND(NVL(D.DIAS_GOZADOS,0),0)) AS DIAS_SALDO, "
  r_str_Parame = r_str_Parame & " SELECT TRUNC(DECODE(A.MAEPRV_FECING,'',0, ABS((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/365))) - "
  r_str_Parame = r_str_Parame & "        (NVL(C.GESPER_DIAGOZ,0) + ROUND(NVL(D.DIAS_GOZADOS,0),0)),0) AS DIAS_SALDO, "
  r_str_Parame = r_str_Parame & "        trunc(months_between(SYSDATE, to_date(A.MAEPRV_FECING,'YYYYMMDD'))/12) * 30 AS DIAS_VENCIDOS "
  r_str_Parame = r_str_Parame & "   FROM CNTBL_MAEPRV A "
  r_str_Parame = r_str_Parame & "   LEFT JOIN CNTBL_GESPER C ON C.GESPER_TIPDOC = A.MAEPRV_TIPDOC AND C.GESPER_NUMDOC = A.MAEPRV_NUMDOC "
  r_str_Parame = r_str_Parame & "         AND C.GESPER_SITUAC = 1 AND C.GESPER_TIPTAB = 3 " '--MAESTRO
  r_str_Parame = r_str_Parame & "   LEFT JOIN (SELECT SUM(NVL(H.GESPER_IMPORT,0)) AS DIAS_GOZADOS , H.GESPER_TIPDOC, H.GESPER_NUMDOC "
  r_str_Parame = r_str_Parame & "                FROM CNTBL_GESPER H "
  r_str_Parame = r_str_Parame & "               Where H.GESPER_TIPTAB = 2 " '--AUTORIZADOS
  r_str_Parame = r_str_Parame & "                 AND H.GESPER_SITUAC = 2 " '--APROBADOS
  r_str_Parame = r_str_Parame & "               GROUP BY H.GESPER_TIPDOC, H.GESPER_NUMDOC) D "
  r_str_Parame = r_str_Parame & "     ON A.MAEPRV_TIPDOC = D.GESPER_TIPDOC AND A.MAEPRV_NUMDOC = D.GESPER_NUMDOC "
  r_str_Parame = r_str_Parame & "  WHERE A.MAEPRV_TIPPER = 2 "   '--PERSONAL INTERNOS
  r_str_Parame = r_str_Parame & "    AND A.MAEPRV_SITUAC = 1 "
  r_str_Parame = r_str_Parame & "    AND A.MAEPRV_TIPDOC = " & moddat_g_int_TipDoc
  r_str_Parame = r_str_Parame & "    AND A.MAEPRV_NUMDOC = '" & moddat_g_str_NumDoc & "'"
   
  If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
     Screen.MousePointer = 0
     Exit Sub
  End If
   
  If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
     p_Saldo = r_rst_Princi!DIAS_SALDO
     p_Vencidos = r_rst_Princi!DIAS_VENCIDOS
  End If
   
  r_rst_Princi.Close
  Set r_rst_Princi = Nothing
End Sub
     
Private Sub cmd_Grabar_Click()
Dim r_bol_Estado As Boolean

   If (cmb_TipDoc.ListIndex = -1) Then
       MsgBox "Debe de seleccionar un tipo de documento.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipDoc)
       Exit Sub
   End If
        
   If Len(Trim(cmb_Proveedor.Text)) = 0 Then
       MsgBox "Tiene que ingresar un proveedor.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   Else
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(cmb_Proveedor.Text), "-") > 0 Then
              For l_int_Contar = 1 To UBound(l_arr_MaePrv)
                  If Trim(Mid(cmb_Proveedor.Text, 1, InStr(Trim(cmb_Proveedor.Text), "-") - 1)) = Trim(l_arr_MaePrv(l_int_Contar).Genera_Codigo) Then
                     r_bol_Estado = True
                     Exit For
                  End If
              Next
           End If
           If r_bol_Estado = False Then
              MsgBox "El Proveedor no se encuentra en la lista.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(cmb_Proveedor)
              Exit Sub
           End If
       End If
   End If
    
   If Trim(pnl_CodPla.Caption) = "" Then
       MsgBox "No se puede registrar si no hay un código de planilla.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   End If
    
   If Trim(pnl_CodPla.Tag) <> fs_NumDoc(cmb_Proveedor.Text) Then
       MsgBox "El Nro de Documento no coincide con la busqueda inicial.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   End If
        
   If cmb_TipOpe.ListIndex = -1 Then
       MsgBox "Debe de seleccionar un tipo de operación.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipOpe)
       Exit Sub
   End If

   If moddat_g_int_TipRec = 1 Then
      'gestion de pagos
      Call fs_Guardar_Pag
   Else
      'gestion de vacaciones
      Call fs_Guardar_Vac
   End If
End Sub

Private Sub fs_Guardar_Pag()
   If CDbl(pnl_TipCambio.Caption) = 0 Then
      MsgBox "El tipo de cambio no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecOpe)
      Exit Sub
   End If

   If cmb_Moneda.ListIndex = -1 Then
       MsgBox "Debe de seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Moneda)
       Exit Sub
   End If
    
   If CDbl(ipp_Importe.Text) <= 0 Then
       MsgBox "Debe de ingresar un monto mayor a cero.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_Importe)
       Exit Sub
   End If
    
   If cmb_Banco.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Banco)
      Exit Sub
   End If
   
   If cmb_CtaCte.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un nro cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCte)
      Exit Sub
   End If
    
   If (Format(ipp_FecOpe.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(ipp_FecOpe.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FecOpe)
       Exit Sub
   End If
   
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Private Sub fs_Guardar_Vac()
Dim r_str_Parame   As String
Dim r_rst_Genera   As ADODB.Recordset
Dim r_str_FecFin   As String
Dim r_str_FecSig   As String
Dim r_str_CadAux   As String
Dim r_bol_Estado   As Boolean
Dim r_int_Saldo    As Integer
Dim r_int_Vencid   As Integer

   If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 2 And cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 5 Then
      If Format(ipp_FecIniVac.Text, "yyyymmdd") >= Format(ipp_FecFinVac.Text, "yyyymmdd") Then
         MsgBox "Los rangos de fechas estan mal ingresados, favor verificar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIniVac)
         Exit Sub
      End If
   Else
      If Format(ipp_FecIniVac.Text, "yyyymmdd") > Format(ipp_FecFinVac.Text, "yyyymmdd") Then
         MsgBox "Los rangos de fechas estan mal ingresados, favor verificar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIniVac)
         Exit Sub
      End If
   End If
   
   If Trim(ipp_DiaSol.Text) = 0 Then
      MsgBox "No se puede solicitar menos de 1 día de vacaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_DiaSol)
      Exit Sub
   End If
         
   If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 2 And cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 4 And cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 5 Then
      If CLng(ipp_DiaSol.Text) < 7 Then
         MsgBox "No se puede solicitar menos de 7 días de vacaciones.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_DiaSol)
         Exit Sub
      End If
   End If
   If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 1 Or cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 3 Or cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 2 Then
      'vacaciones pendientes
      If Trim(pnl_Saldo.Caption) = "" Then
         MsgBox "No cuenta con días disponibles.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipOpe)
         Exit Sub
      End If
      
      If CLng(pnl_Saldo.Caption) <= 0 Then
         MsgBox "No cuenta con saldo disponible.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipOpe)
         Exit Sub
      End If
        
      If CLng(pnl_Saldo.Caption) < CLng(ipp_DiaSol.Text) Then
         MsgBox "No cuenta con saldo disponible.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_DiaSol)
         Exit Sub
      End If
            
      If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 2 Then
         'a cuenta
         If CInt(ipp_DiaSol.Text) > 6 Then
            MsgBox "Solo se puede solicitar de 1 a 6 días.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_DiaSol)
            Exit Sub
         End If
      End If
      
      If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 2 And Trim(txt_Observ.Text) = "" Then
         'a cuenta
         MsgBox "Debe ingresar un comentario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Observ)
         Exit Sub
      End If
   Else
      If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 5 Then
         'Adelanto de vacaciones
         r_int_Saldo = 0
         r_int_Vencid = 0
         Call fs_Saldo_Devengado(r_int_Saldo, r_int_Vencid)
         'If r_int_Vencid <> 0 Then
         '   MsgBox "El trabajador no tiene que haber cumplido el año.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_DiaSol)
         '   Exit Sub
         'End If
         If r_int_Saldo < ipp_DiaSol.Text Then
            MsgBox "No cuenta con saldo disponible.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_DiaSol)
            Exit Sub
         End If
      End If
      
      'a cuenta extraordinaria
      If Trim(txt_Observ.Text) = "" Then
         MsgBox "Debe ingresar un comentario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Observ)
         Exit Sub
      End If
   End If
   
   If fs_validar_fecha(Trim(pnl_Codigo.Caption)) = False Then
      MsgBox "La fecha ya esta registrada, ingrese otro rango.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIniVac)
      Exit Sub
   End If
   
   'VALIDACION FECHA CONTABLE
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT LPAD(PERMES_CODMES,2,'0') AS CODMES, PERMES_CODANO "
   r_str_Parame = r_str_Parame & "   FROM CTB_PERMES A "
   r_str_Parame = r_str_Parame & "  WHERE A.PERMES_CODEMP = '000001'"
   r_str_Parame = r_str_Parame & "    AND A.PERMES_TIPPER = 1 "
   r_str_Parame = r_str_Parame & "    AND A.PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If

   r_str_FecFin = ""
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_str_CadAux = DateAdd("m", 1, "01/" & Format(r_rst_Genera!CODMES, "00") & "/" & r_rst_Genera!PERMES_CODANO)
      'r_str_FecSig = Format(ff_Ultimo_Dia_Mes(Month(r_str_FecFin), Year(r_str_FecFin)), "00")
      r_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   End If

   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
'   If (Format(date, "yyyymmdd") < Format(r_str_FecFin, "yyyymmdd") Or Format(date, "yyyymmdd") > Format(r_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FecIniVac)
'       Exit Sub
'   End If
   r_bol_Estado = False
   r_str_CadAux = Format(r_str_FecFin, "yyyymm22")
   If Format(ipp_FecOpe.Text, "yyyymmdd") < r_str_CadAux Then
      r_bol_Estado = True
   Else
      r_bol_Estado = False
      If Format(ipp_FecIniVac.Text, "yyyymmdd") > Format(r_str_FecFin, "yyyymmdd") Then
         r_bol_Estado = True
      End If
   End If
   
   If Format(ipp_FecOpe.Text, "yyyymm") < Format(r_str_FecFin, "yyyymm") Then
      r_bol_Estado = False
   End If

   If frm_Ctb_GesPer_01.fs_UserEjecutivo(modgen_g_str_CodUsu, "313") = "" Then 'administrador vacaciones
      If r_bol_Estado = False Then
         MsgBox "Intenta solicitar vacaciones fuera de fecha, favor de comunicarse con el analista de administración al personal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIniVac)
         Exit Sub
      End If
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT COUNT(*) AS CONTEO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPTAB = 2 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_NUMDOC = '" & moddat_g_str_NumDoc & "'"
   r_str_Parame = r_str_Parame & "    AND A.GESPER_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPOPE = 5 " 'ADELANTO DE VACACIONES
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If r_rst_Genera!CONTEO <> 0 Then
      MsgBox "Se tiene un adelanto pendiente de vacaciones.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   Call fs_Grabar
   Screen.MousePointer = 0
End Sub

Private Function fs_validar_fecha(p_Codigo As String) As Boolean
Dim r_int_TotFec   As Integer
Dim r_int_FilFec   As Integer
Dim r_str_FecAu1   As String
Dim r_str_FecAu2   As String

Dim r_str_Parame   As String
Dim r_int_TotAux   As Integer
Dim r_int_filAux   As Integer
Dim r_rst_Princi   As ADODB.Recordset
Dim r_rst_Genera   As ADODB.Recordset
Dim r_str_FecAux   As String

   fs_validar_fecha = True
            
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT "
   r_str_Parame = r_str_Parame & "        NVL((SELECT COUNT(*) "
   r_str_Parame = r_str_Parame & "               FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "              WHERE A.GESPER_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   r_str_Parame = r_str_Parame & "                AND A.GESPER_NUMDOC = '" & fs_NumDoc(cmb_Proveedor.Text) & "'"
   r_str_Parame = r_str_Parame & "                AND A.GESPER_TIPTAB = 2 "
   If Len(Trim(p_Codigo)) > 0 Then
      r_str_Parame = r_str_Parame & "                AND A.GESPER_CODGES <> " & p_Codigo
   End If
   r_str_Parame = r_str_Parame & "                AND A.GESPER_SITUAC IN (1,2)),0) AS CONTEO " 'PENDIENTE Y APROBADO
   r_str_Parame = r_str_Parame & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Function
   End If
      
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      If r_rst_Princi!CONTEO > 0 Then
         r_int_TotFec = DateDiff("d", ipp_FecIniVac.Text, ipp_FecFinVac.Text) + 1
         r_str_FecAux = Format(ipp_FecIniVac.Text, "yyyymmdd")
         For r_int_FilFec = 1 To r_int_TotFec
                 
             r_str_Parame = ""
             r_str_Parame = r_str_Parame & " SELECT "
             r_str_Parame = r_str_Parame & "        NVL((SELECT COUNT(*) "
             r_str_Parame = r_str_Parame & "               FROM CNTBL_GESPER A "
             r_str_Parame = r_str_Parame & "              WHERE A.GESPER_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
             r_str_Parame = r_str_Parame & "                AND A.GESPER_NUMDOC = '" & fs_NumDoc(cmb_Proveedor.Text) & "'"
             r_str_Parame = r_str_Parame & "                AND A.GESPER_TIPTAB = 2 "
             r_str_Parame = r_str_Parame & "                AND A.GESPER_SITUAC IN (1,2) " 'PENDIENTE Y APROBADO
             If Len(Trim(p_Codigo)) > 0 Then
                r_str_Parame = r_str_Parame & "                AND A.GESPER_CODGES <> " & p_Codigo
             End If
             r_str_Parame = r_str_Parame & "                AND (GESPER_FECHA1 <= " & r_str_FecAux & " AND " & r_str_FecAux & " <= GESPER_FECHA2)),0) AS CONTEO "
             r_str_Parame = r_str_Parame & "   FROM DUAL "
             
             If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
                Exit Function
             End If
             r_rst_Genera.MoveFirst
             If r_rst_Genera!CONTEO <> 0 Then
                r_rst_Princi.Close
                Set r_rst_Princi = Nothing
                fs_validar_fecha = False
                Exit Function
             End If
             r_str_FecAux = Format(DateAdd("d", 1, gf_FormatoFecha(r_str_FecAux)), "yyyymmdd")
         Next
                     
      End If
   End If
   
   fs_validar_fecha = True
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Function

Private Function fs_ValNumDoc() As Boolean
Dim r_str_NumDoc  As String
Dim r_bol_Estado  As Boolean

   fs_ValNumDoc = True
   r_str_NumDoc = ""

   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then 'DNI - 8
       If Len(Trim(r_str_NumDoc)) <> 8 Then
          MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then 'RUC - 11
       If Not gf_Valida_RUC(Trim(r_str_NumDoc), Mid(Trim(r_str_NumDoc), Len(Trim(r_str_NumDoc)), 1)) Then
          MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   Else 'OTROS
       If Len(Trim(cmb_Proveedor.Text)) = 0 Then
          MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   End If
   
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then
      'CONSULTAR
      If moddat_g_int_TipRec = 2 Or moddat_g_int_TipRec = 3 Then
         pnl_Titulo.Caption = "Gestión Vacaciones - Consultar"
         Call fs_Cargar_Vac
      Else
         pnl_Titulo.Caption = "Gestión Pagos - Consultar"
         Call fs_Cargar_Pag
      End If
      Call fs_Desabilitar(False)
      cmd_Grabar.Visible = False
   ElseIf moddat_g_int_FlgGrb = 1 Then
      'INSERTAR
      If moddat_g_int_TipRec = 2 Then
         pnl_Titulo.Caption = "Gestión Vacaciones - Adicionar"
      Else
         pnl_Titulo.Caption = "Gestión Pagos - Adicionar"
      End If
   ElseIf moddat_g_int_FlgGrb = 2 Then
      'EDITAR
      If moddat_g_int_TipRec = 2 Then
         pnl_Titulo.Caption = "Gestión Vacaciones - Editar"
         Call fs_Cargar_Vac
      End If
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   If moddat_g_int_TipRec = 2 Or moddat_g_int_TipRec = 3 Then
      'GESTION DE VACACIONES
      If frm_Ctb_GesPer_01.fs_UserEjecutivo(modgen_g_str_CodUsu, "313") <> "" Or frm_Ctb_GesPer_01.fs_UserEjecutivo(modgen_g_str_CodUsu, "314") <> "" Then 'administrador vacaciones, evaluador
         Call moddat_gs_Carga_LisIte_Combo(cmb_TipOpe, 1, "140")
      Else
         If moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2 Then
            Call fs_CargaMntPardes(cmb_TipOpe, "140")
         Else
            Call moddat_gs_Carga_LisIte_Combo(cmb_TipOpe, 1, "140")
         End If
      End If
   Else
     'GESTION DE PAGOS
      Call moddat_gs_Carga_LisIte_Combo(cmb_TipOpe, 1, "130")
   End If
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call fs_CargaMntPardes(cmb_TipDoc, "118")
End Sub

Private Sub fs_FechaMax()
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL((SELECT MAX(A.GESPER_FECHA2) "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_SITUAC IN (1,2) "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPTAB IN (2) "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_NUMDOC = " & moddat_g_str_NumDoc & "),0) AS FECHA_MAX "
   r_str_Parame = r_str_Parame & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
     r_rst_Genera.Close
     Set r_rst_Genera = Nothing
     Exit Sub
   End If
   
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
       
       If r_rst_Genera!FECHA_MAX = 0 Then
          ipp_FecIniVac.Text = moddat_g_str_FecSis
          ipp_FecFinVac.Text = DateAdd("d", 6, moddat_g_str_FecSis)
       Else
          ipp_FecIniVac.Text = DateAdd("d", 1, gf_FormatoFecha(r_rst_Genera!FECHA_MAX))
          ipp_FecFinVac.Text = DateAdd("d", 7, gf_FormatoFecha(r_rst_Genera!FECHA_MAX))
       End If
       
      r_rst_Genera.MoveNext
   Loop
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Private Sub fs_Limpiar()
Dim r_int_DiaGoz As Integer
Dim r_int_DiaGan As Integer

   pnl_Codigo.Caption = ""
   cmb_TipDoc.ListIndex = 0
   cmb_Proveedor.Text = ""
   pnl_CodPla.Caption = ""
   ipp_FecOpe.Text = moddat_g_str_FecSis
   'TIPO CAMBIO SBS(2) - VENTA(1)
   pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecOpe.Text, "yyyymmdd"), 1)
   pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
   cmb_TipOpe.ListIndex = 0
   cmb_Moneda.ListIndex = 0
   ipp_Importe.Text = "0.00"
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   pnl_Saldo.Caption = ""
   ipp_DiaSol.Text = 7
   txt_Observ.Text = ""
   
   If moddat_g_int_TipRec = 2 Or moddat_g_int_TipRec = 3 Then
      'Gestion de Vacaciones
      Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_int_TipDoc)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, moddat_g_str_NumDoc, 0)
      cmb_TipDoc.Enabled = False
      cmb_Proveedor.Enabled = False
      ipp_FecOpe.Enabled = False
      
      ipp_FecFinVac.Enabled = False
      lbl_TipCam.Visible = False
      pnl_TipCambio.Visible = False
      cmb_Moneda.Visible = False
      pnl_Abono.Visible = False
      cmb_Moneda.Visible = False
      ipp_Importe.Visible = False
      
      lbl_Moneda.Caption = "Saldo Dias:"
      pnl_Saldo.Left = 1590
      ipp_FecIniVac.Left = 1590
      ipp_FecFinVac.Left = 3850
      lbl_Hasta.Left = 3250
      lbl_Dia_Sol.Visible = True
      ipp_DiaSol.Visible = True
      lbl_Descrip.Visible = True
      txt_Observ.Visible = True
      pnl_Datos.Height = 4440
      
      ipp_DiaSol.Top = 2700
      ipp_FecIniVac.Top = 3030
      lbl_Hasta.Top = 3030
      ipp_FecFinVac.Top = 3030
      lbl_Importe.Caption = "Dias Solicitado:"
      lbl_Dia_Sol.Caption = "Desde:"
      
      pnl_Saldo.Alignment = 7
      r_int_DiaGoz = 0
      r_int_DiaGan = 0
            
      Call fs_FechaMax
      If Weekday(ipp_FecFinVac.Text) = vbFriday Then
         ipp_DiaSol.Text = ipp_DiaSol.Text + 2
         ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
      End If
   
      Call frm_Ctb_GesPer_01.fs_SaldoDias(moddat_g_int_TipDoc, moddat_g_str_NumDoc, r_int_DiaGan, r_int_DiaGoz)
      pnl_Saldo.Caption = CStr(r_int_DiaGan - r_int_DiaGoz)
   Else
      'Gestion de Pago
      pnl_Saldo.Visible = False
      ipp_FecIniVac.Visible = False
      ipp_FecFinVac.Visible = False
      lbl_Hasta.Visible = False
      lbl_Dia_Sol.Visible = False
      ipp_DiaSol.Visible = False
      lbl_Descrip.Visible = False
      txt_Observ.Visible = False
      pnl_Datos.Height = 3160
   End If
End Sub

Private Sub fs_Desabilitar(p_Estado As Boolean)
   cmb_TipDoc.Enabled = p_Estado
   cmb_Proveedor.Enabled = p_Estado
   ipp_FecOpe.Enabled = p_Estado
   cmb_TipOpe.Enabled = p_Estado
   cmb_Moneda.Enabled = p_Estado
   ipp_Importe.Enabled = p_Estado
   ipp_DiaSol.Enabled = p_Estado
   cmb_Banco.Enabled = p_Estado
   cmb_CtaCte.Enabled = p_Estado
   
   ipp_FecIniVac.Enabled = p_Estado
   ipp_FecFinVac.Enabled = p_Estado
   txt_Observ.Enabled = p_Estado
End Sub

Private Sub ipp_DiaSol_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       'Call gs_SetFocus(txt_Observ)
       Call gs_SetFocus(ipp_FecIniVac)
   End If
End Sub

Private Sub ipp_DiaSol_LostFocus()
   Call ipp_FecIniVac_LostFocus
   
   'If Weekday(ipp_FecFinVac.Text) = vbFriday Then
   '   ipp_DiaSol.Text = ipp_DiaSol.Text + 2
   '   MsgBox "Se adicionara 2 días a su solicitud, por ser viernes la fecha termino.", vbExclamation, modgen_g_str_NomPlt
   '   Call fs_DiasSolicitados
   'End If
End Sub

Private Sub ipp_FecIniVac_LostFocus()
   Call fs_DiasSolicitados
End Sub

Private Sub fs_DiasSolicitados()
Dim r_str_Msg As String

   r_str_Msg = ""
   
   ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
   If Weekday(ipp_FecFinVac.Text) = vbFriday Then
      ipp_DiaSol.Text = ipp_DiaSol.Text + 2
      ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
      MsgBox "Se adicionara 2 días a su solicitud, por ser viernes la fecha termino.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 4 Then
      If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 2 Then
         If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) <> 5 Then
            If CInt(ipp_DiaSol.Text) < 7 Then
               ipp_DiaSol.Text = 7
               ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
      
               If Weekday(ipp_FecFinVac.Text) = vbFriday Then
                  ipp_DiaSol.Text = ipp_DiaSol.Text + 2
                  ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
                  MsgBox "Minimo a solicitar son 7 dias y se adicionara 2 días, por ser viernes la fecha termino.", vbExclamation, modgen_g_str_NomPlt
               Else
                  MsgBox "Mínimo a solicitar son 7 días.", vbExclamation, modgen_g_str_NomPlt
               End If
            End If
         End If
      End If
   End If
   
   If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 2 Then
      If CInt(ipp_DiaSol.Text) > 6 Then
         ipp_DiaSol.Text = 6
         ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)

         If Weekday(ipp_FecFinVac.Text) = vbFriday Then
            ipp_DiaSol.Text = ipp_DiaSol.Text + 2
            ipp_FecFinVac.Text = DateAdd("d", CInt(ipp_DiaSol.Text) - 1, ipp_FecIniVac.Text)
            'MsgBox "Minimo a solicitar son 7 dias y se adicionara 2 días, por ser viernes la fecha termino.", vbExclamation, modgen_g_str_NomPlt
         End If
         
      End If
   End If
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub ipp_FecFinVac_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub ipp_FecIniVac_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       'Call gs_SetFocus(ipp_DiaSol)
       Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub ipp_FecOpe_LostFocus()
   'TIPO CAMBIO SBS(2) - VENTA(1)
   pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecOpe.Text, "yyyymmdd"), 1)
   pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
End Sub

Private Sub ipp_FecOpe_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipOpe)
   End If
End Sub

Private Sub cmb_TipOpe_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If moddat_g_int_TipRec = 1 Then
          'gestion personal
          Call gs_SetFocus(cmb_Moneda)
       Else
          'gestion vacaciones
          'Call gs_SetFocus(ipp_FecIniVac)
          Call gs_SetFocus(ipp_DiaSol)
       End If
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_Importe)
   End If
End Sub

Private Sub ipp_Importe_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
   End If
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_Proveedor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FecOpe)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub fs_CargaMntPardes(p_Combo As ComboBox, ByVal p_CodGrp As String)
Dim r_str_Parame As String

   p_Combo.Clear
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM MNT_PARDES A "
   r_str_Parame = r_str_Parame & " WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   If Trim(p_CodGrp) = "118" Then
      r_str_Parame = r_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000007') "
   ElseIf Trim(p_CodGrp) = "140" Then
      r_str_Parame = r_str_Parame & " AND A.PARDES_CODITE IN ('000001','000002','000003','000005') "
   End If
   r_str_Parame = r_str_Parame & "   AND PARDES_SITUAC = 1 "
   r_str_Parame = r_str_Parame & " ORDER BY PARDES_CODITE ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!PARDES_CODITE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Grabar()
Dim r_str_CodGen As String
Dim r_str_Parame As String
Dim r_rst_Genera As ADODB.Recordset

   r_str_CodGen = ""
      
   If moddat_g_int_TipRec = 1 Then
      'GESTION DE PAGOS
      If moddat_g_int_FlgGrb = 1 Then
         r_str_CodGen = modmip_gf_Genera_CodGen(3, 7)
      Else
         r_str_CodGen = Trim(pnl_Codigo.Caption)
      End If
   
      If Len(Trim(r_str_CodGen)) = 0 Then
         MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER ( "
      r_str_Parame = r_str_Parame & CLng(r_str_CodGen) & ", "
      r_str_Parame = r_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
      r_str_Parame = r_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', "
      r_str_Parame = r_str_Parame & Format(ipp_FecOpe.Text, "yyyymmdd") & ", "
      r_str_Parame = r_str_Parame & CDbl(pnl_TipCambio.Caption) & ", "
      r_str_Parame = r_str_Parame & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "
      r_str_Parame = r_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "
      r_str_Parame = r_str_Parame & CDbl(ipp_Importe.Text) & ", "
      If (cmb_Banco.ListIndex = -1) Then
          r_str_Parame = r_str_Parame & "Null , " 'GESPER_CODBNC
      Else
          r_str_Parame = r_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", " 'GESPER_CODBNC
      End If
      r_str_Parame = r_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', " 'GESPER_CTACRR
      r_str_Parame = r_str_Parame & "1 , " 'GESPER_TIPTAB
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_FECHA1
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_FECHA2
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_DESCRI
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVEN
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVIG
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
      r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      r_str_Parame = r_str_Parame & "1) " 'as_insupd - moddat_g_int_FlgGrb
   Else
      'GESTION DE VACACIONES
      If moddat_g_int_FlgGrb = 1 Then
         r_str_CodGen = modmip_gf_Genera_CodGen(3, 13)
      Else
         r_str_CodGen = Trim(pnl_Codigo.Caption)
      End If
   
      If Len(Trim(r_str_CodGen)) = 0 Then
         MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER ( "
      r_str_Parame = r_str_Parame & CLng(r_str_CodGen) & ", "
      r_str_Parame = r_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
      r_str_Parame = r_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', "
      r_str_Parame = r_str_Parame & Format(ipp_FecOpe.Text, "yyyymmdd") & ", "
      r_str_Parame = r_str_Parame & "NULL, " 'TIPO CAMBIO
      r_str_Parame = r_str_Parame & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "
      r_str_Parame = r_str_Parame & "NULL, " 'TIPO MONEDA
      r_str_Parame = r_str_Parame & CDbl(ipp_DiaSol.Text) & ", "
      r_str_Parame = r_str_Parame & "Null , " 'GESPER_CODBNC
      r_str_Parame = r_str_Parame & "'', " 'GESPER_CTACRR
      r_str_Parame = r_str_Parame & "2 , " 'GESPER_TIPTAB
      r_str_Parame = r_str_Parame & Format(ipp_FecIniVac.Text, "yyyymmdd") & ", " 'GESPER_FECHA1
      r_str_Parame = r_str_Parame & Format(ipp_FecFinVac.Text, "yyyymmdd") & ", " 'GESPER_FECHA2
      r_str_Parame = r_str_Parame & "'" & Trim(txt_Observ.Text) & "', " 'GESPER_DESCRI
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVEN
      r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVIG
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
      r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      r_str_Parame = r_str_Parame & CStr(moddat_g_int_FlgGrb) & ") " 'as_insupd
   End If

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   If moddat_g_int_TipRec = 1 Then
      Call frm_Ctb_GesPer_01.fs_BuscarPag
   Else
      Call frm_Ctb_GesPer_03.fs_BuscarReg
      Call frm_Ctb_GesPer_01.fs_BuscarVac
   End If
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub fs_Cargar_Pag()
Dim r_int_Contad As Integer

   pnl_CodPla.Tag = ""
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT LPAD(A.GESPER_CODGES,10,'0') AS GESPER_CODGES, A.GESPER_TIPDOC, A.GESPER_NUMDOC,  "
   g_str_Parame = g_str_Parame & "         TRIM(B.MAEPRV_RAZSOC) AS MAEPRV_RAZSOC, A.GESPER_FECOPE, A.GESPER_TIPCAM,  "
   g_str_Parame = g_str_Parame & "         A.GESPER_TIPOPE, A.GESPER_CODMON, A.GESPER_IMPORT, B.MAEPRV_CODSIC,  "
   g_str_Parame = g_str_Parame & "         A.GESPER_CODBNC, GESPER_CTACRR  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_GESPER A  "
   g_str_Parame = g_str_Parame & "   INNER JOIN CNTBL_MAEPRV B ON A.GESPER_TIPDOC = B.MAEPRV_TIPDOC AND A.GESPER_NUMDOC = B.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "   WHERE A.GESPER_CODGES = '" & CLng(moddat_g_str_Codigo) & "'  "
   g_str_Parame = g_str_Parame & "     AND A.GESPER_TIPTAB = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Codigo.Caption = CStr(g_rst_Princi!GESPER_CODGES)
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!GESPER_TIPDOC)
      
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!GESPER_NUMDOC & "", 0)
      pnl_CodPla.Tag = Trim(g_rst_Princi!GESPER_NUMDOC & "")
      pnl_CodPla.Caption = Trim(g_rst_Princi!MAEPRV_CODSIC & "")
      
      ipp_FecOpe.Text = gf_FormatoFecha(g_rst_Princi!GESPER_FECOPE)
      pnl_TipCambio.Caption = Format(g_rst_Princi!GESPER_TIPCAM, "###,###,##0.000000") & " "
      Call gs_BuscarCombo_Item(cmb_TipOpe, g_rst_Princi!GESPER_TIPOPE)
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!GESPER_CODMON)
      
      If Not IsNull(g_rst_Princi!GESPER_CODBNC) Then
         Call gs_BuscarCombo_Item(cmb_Banco, g_rst_Princi!GESPER_CODBNC)
      End If
      If Not IsNull(g_rst_Princi!GESPER_CTACRR) Then
         Call gs_BuscarCombo_Text(cmb_CtaCte, g_rst_Princi!GESPER_CTACRR, -1)
      End If
      ipp_Importe.Text = Format(CStr(g_rst_Princi!GESPER_IMPORT), "###,###,##0.00")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Cargar_Vac()
Dim r_int_Contad As Integer

   pnl_CodPla.Tag = ""
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT LPAD(A.GESPER_CODGES,10,'0') AS GESPER_CODGES, A.GESPER_TIPDOC, A.GESPER_NUMDOC, "
   g_str_Parame = g_str_Parame & "         TRIM(B.MAEPRV_RAZSOC) AS MAEPRV_RAZSOC, A.GESPER_FECOPE, "
   g_str_Parame = g_str_Parame & "         A.GESPER_TIPOPE, A.GESPER_IMPORT, B.MAEPRV_CODSIC, A.GESPER_FECHA1, A.GESPER_FECHA2, "
   g_str_Parame = g_str_Parame & "         TRIM(A.GESPER_DESCRI) As GESPER_DESCRI "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_GESPER A "
   g_str_Parame = g_str_Parame & "   INNER JOIN CNTBL_MAEPRV B ON A.GESPER_TIPDOC = B.MAEPRV_TIPDOC AND A.GESPER_NUMDOC = B.MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & "   WHERE A.GESPER_CODGES = '" & CLng(moddat_g_str_Codigo) & "'  "
   g_str_Parame = g_str_Parame & "     AND A.GESPER_TIPTAB = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Codigo.Caption = CStr(g_rst_Princi!GESPER_CODGES)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!GESPER_TIPDOC)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!GESPER_NUMDOC & "", 0)
      
      pnl_CodPla.Tag = Trim(g_rst_Princi!GESPER_NUMDOC & "")
      pnl_CodPla.Caption = Trim(g_rst_Princi!MAEPRV_CODSIC & "")
      ipp_FecOpe.Text = gf_FormatoFecha(g_rst_Princi!GESPER_FECOPE)
      Call gs_BuscarCombo_Item(cmb_TipOpe, g_rst_Princi!GESPER_TIPOPE)
      
      ipp_FecIniVac.Text = gf_FormatoFecha(g_rst_Princi!GESPER_FECHA1)
      ipp_FecFinVac.Text = gf_FormatoFecha(g_rst_Princi!GESPER_FECHA2)
      ipp_DiaSol.Text = CInt(g_rst_Princi!GESPER_IMPORT)
      txt_Observ.Text = Trim(g_rst_Princi!GESPER_DESCRI & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function fs_ComboIndex(p_Combo As ComboBox, Cadena As String, p_Tipo As Integer) As Integer
Dim r_int_Contad As Integer

   fs_ComboIndex = -1
   For r_int_Contad = 0 To p_Combo.ListCount - 1
       If Trim(Cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, InStr(Trim(p_Combo.List(r_int_Contad)), "-") - 1)) Then
          fs_ComboIndex = r_int_Contad
          Exit For
       End If
   Next
End Function

Private Sub fs_CargarPrv()
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
   cmb_Proveedor.Clear
   cmb_Proveedor.Text = ""
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_TipDoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, A.MAEPRV_CODSIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   g_str_Parame = g_str_Parame & "    AND A.MAEPRV_TIPPER = 2  "
 
   If moddat_g_int_FlgGrb = 1 Then 'INSERT
      g_str_Parame = g_str_Parame & " AND A.MAEPRV_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      cmb_Proveedor.AddItem Trim(g_rst_Genera!MAEPRV_NUMDOC & "") & " - " & Trim(g_rst_Genera!MaePrv_RazSoc & "")
      
      ReDim Preserve l_arr_MaePrv(UBound(l_arr_MaePrv) + 1)
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Prefij = Trim(g_rst_Genera!MAEPRV_CODSIC & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Buscar_prov()
Dim r_str_NumDoc As String
Dim r_int_Contar As Integer

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   r_str_NumDoc = ""
   pnl_CodPla.Caption = ""
   pnl_CodPla.Tag = ""
          
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc)
          Exit Sub
       End If
       If cmb_Proveedor.ListIndex = -1 Then
          Exit Sub
       End If
      
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   
   For r_int_Contar = 1 To UBound(l_arr_MaePrv)
       If Trim(l_arr_MaePrv(r_int_Contar).Genera_Codigo) = r_str_NumDoc Then
          pnl_CodPla.Caption = Trim(l_arr_MaePrv(r_int_Contar).Genera_Prefij)
          pnl_CodPla.Tag = Trim(l_arr_MaePrv(r_int_Contar).Genera_Codigo)
          Exit For
       End If
   Next
   '==========================================
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_CODBNC_MN1, A.MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN1, A.MAEPRV_CODBNC_MN2, A.MAEPRV_CTACRR_MN2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN2, A.MAEPRV_CODBNC_MN3, A.MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN3, A.MAEPRV_CODBNC_DL1, A.MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL1, A.MAEPRV_CODBNC_DL2, A.MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL2, A.MAEPRV_CODBNC_DL3, A.MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   'If (moddat_g_int_FlgGrb = 1) Then
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   'End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      MsgBox "No se ha encontrado el proveedor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Proveedor)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   If (moddat_g_int_FlgGrb = 1) Then
       If (g_rst_GenAux!MAEPRV_CONDIC = 2) Then
          MsgBox "El proveedor se encuentra en condición de NO HABIDO, revisar sunat.", vbExclamation, modgen_g_str_NomPlt
          g_rst_GenAux.Close
          Set g_rst_GenAux = Nothing
          Exit Sub
       End If
   End If
      
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)

   If (g_rst_GenAux!MAEPRV_CODBNC_MN1 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN1, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN2 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN2)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN2, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN3 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN3)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN3, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   
   If (g_rst_GenAux!MAEPRV_CODBNC_DL1 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL1, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL2 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL2)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL2, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL3 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL3)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL3, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   
   Call fs_CargarBancos
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Function fs_NumDoc(p_Cadena As String) As String
   fs_NumDoc = ""
   If (cmb_TipDoc.ListIndex > -1) Then
      If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
           If InStr(1, p_Cadena, "-") <= 0 Then
              Exit Function
           End If
           fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
      End If
   End If
End Function

Private Sub fs_CargarBancos()
Dim r_bol_Estado   As Boolean
Dim r_int_File     As Integer

   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   'soles
   If (cmb_Moneda.ListIndex = 0) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
   'dolares
   If (cmb_Moneda.ListIndex = 1) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
End Sub

Private Sub cmb_Banco_Click()
Dim r_str_Cadena  As String
   
   cmb_CtaCte.Clear
   r_str_Cadena = ""
   lbl_Cuenta.Caption = "Cuenta:"
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   If (cmb_Moneda.ListIndex = 0) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
   
   If (cmb_Moneda.ListIndex = 1) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
End Sub

Private Sub cmb_Moneda_Click()
   Call fs_CargarBancos
End Sub

