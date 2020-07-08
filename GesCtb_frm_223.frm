VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CtaPag_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "GesCtb_frm_223.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel7 
      Height          =   7005
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   12356
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1460
         Left            =   60
         TabIndex        =   20
         Top             =   1500
         Width           =   8020
         _Version        =   65536
         _ExtentX        =   14146
         _ExtentY        =   2575
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
         Begin VB.TextBox txt_Descrip 
            Height          =   315
            Left            =   1590
            MaxLength       =   60
            TabIndex        =   2
            Top             =   960
            Width           =   6060
         End
         Begin EditLib.fpDateTime ipp_FchOpe 
            Height          =   315
            Left            =   1590
            TabIndex        =   1
            Top             =   630
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            Top             =   300
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label Label38 
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
            Left            =   120
            TabIndex        =   23
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operación:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   540
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   8020
         _Version        =   65536
         _ExtentX        =   14146
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
            TabIndex        =   25
            Top             =   180
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cuentas por Pagar - Accion"
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
            Left            =   90
            Picture         =   "GesCtb_frm_223.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   3890
         Left            =   60
         TabIndex        =   26
         Top             =   3000
         Width           =   8020
         _Version        =   65536
         _ExtentX        =   14146
         _ExtentY        =   6862
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
         Begin VB.ComboBox cmb_CtaImp_01 
            Height          =   315
            Left            =   3030
            TabIndex        =   9
            Top             =   2400
            Width           =   4620
         End
         Begin VB.ComboBox cmb_CtaImp_02 
            Height          =   315
            Left            =   3030
            TabIndex        =   11
            Top             =   2730
            Width           =   4620
         End
         Begin VB.ComboBox cmb_CtaImp_03 
            Height          =   315
            Left            =   3030
            TabIndex        =   13
            Top             =   3060
            Width           =   4620
         End
         Begin VB.ComboBox cmb_CtaImp_04 
            Height          =   315
            Left            =   3030
            TabIndex        =   15
            Top             =   3390
            Width           =   4620
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            ItemData        =   "GesCtb_frm_223.frx":0316
            Left            =   1590
            List            =   "GesCtb_frm_223.frx":0318
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   4400
         End
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   1590
            TabIndex        =   4
            Top             =   630
            Width           =   6060
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   4400
         End
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1290
            Width           =   4400
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1620
            Width           =   4400
         End
         Begin EditLib.fpDoubleSingle ipp_Import1 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   2400
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_Import2 
            Height          =   315
            Left            =   1590
            TabIndex        =   10
            Top             =   2730
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_Import3 
            Height          =   315
            Left            =   1590
            TabIndex        =   12
            Top             =   3060
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_Import4 
            Height          =   315
            Left            =   1590
            TabIndex        =   14
            Top             =   3390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Importe 4:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   3450
            Width           =   705
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Importe 3:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   3120
            Width           =   705
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Determinación"
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
            Left            =   120
            TabIndex        =   38
            Top             =   2100
            Width           =   1230
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Importe 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Importe 2:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   2790
            Width           =   705
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CUENTAS CONTABLES"
            Height          =   195
            Left            =   4500
            TabIndex        =   35
            Top             =   2160
            Width           =   1770
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1020
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   120
            TabIndex        =   31
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   690
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1350
            Width           =   510
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   32
         Top             =   780
         Width           =   8020
         _Version        =   65536
         _ExtentX        =   14146
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
         Begin VB.CommandButton cmd_Reversa 
            Height          =   600
            Left            =   630
            Picture         =   "GesCtb_frm_223.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7410
            Picture         =   "GesCtb_frm_223.frx":0624
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_223.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CtaPag_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_int_TopNiv        As Integer

Private Sub cmd_Grabar_Click()
Dim r_dbl_ImpAux   As Double
Dim r_bol_Estado   As Boolean
Dim r_int_Contar   As Integer
Dim r_dbl_TipCam   As Double

   If Trim(txt_Descrip.Text) = "" Then
      MsgBox "Tiene que ingresar una descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descrip)
      Exit Sub
   End If
   
   If Len(Trim(cmb_TipDoc.Text)) = 0 Then
       MsgBox "Tiene que selecconar un tipo de documento.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipDoc)
       Exit Sub
   End If
   
   If Len(Trim(cmb_Proveedor.Text)) = 0 Then
       MsgBox "Tiene que ingresar un proveedor.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   Else
       If (fs_ValNumDoc(cmb_TipDoc, cmb_Proveedor) = False) Then
           Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(cmb_Proveedor.Text), "-") > 0 Then
              For r_int_Contar = 1 To UBound(l_arr_MaePrv)
                  If Trim(Mid(cmb_Proveedor.Text, 1, InStr(Trim(cmb_Proveedor.Text), "-") - 1)) = Trim(l_arr_MaePrv(r_int_Contar).Genera_Codigo) Then
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
   '--------------------------------------------
   If CDbl(ipp_Import1.Text) > 0 And Len(Trim(cmb_CtaImp_01.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaImp_01)
      Exit Sub
   End If
   If CDbl(ipp_Import1.Text) = 0 And Len(Trim(cmb_CtaImp_01.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import1)
      Exit Sub
   End If
         
   If CDbl(ipp_Import2.Text) > 0 And Len(Trim(cmb_CtaImp_02.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaImp_02)
      Exit Sub
   End If
   If CDbl(ipp_Import2.Text) = 0 And Len(Trim(cmb_CtaImp_02.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import2)
      Exit Sub
   End If
   
   If CDbl(ipp_Import3.Text) > 0 And Len(Trim(cmb_CtaImp_03.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaImp_03)
      Exit Sub
   End If
   If CDbl(ipp_Import3.Text) = 0 And Len(Trim(cmb_CtaImp_03.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import3)
      Exit Sub
   End If
   
   If CDbl(ipp_Import4.Text) > 0 And Len(Trim(cmb_CtaImp_04.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaImp_04)
      Exit Sub
   End If
   If CDbl(ipp_Import4.Text) = 0 And Len(Trim(cmb_CtaImp_04.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import4)
      Exit Sub
   End If
   '---------------VALIDAR EXISTENCIA DE LA CUENTA-----------------------------
   If Len(Trim(cmb_CtaImp_01.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaImp_01.Text) = False Then
         Call gs_SetFocus(cmb_CtaImp_01)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaImp_02.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaImp_02.Text) = False Then
         Call gs_SetFocus(cmb_CtaImp_02)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaImp_03.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaImp_03.Text) = False Then
         Call gs_SetFocus(cmb_CtaImp_03)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaImp_04.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaImp_04.Text) = False Then
         Call gs_SetFocus(cmb_CtaImp_04)
         Exit Sub
      End If
   End If
   '--------------------------------------------
   If (CDbl(ipp_Import1.Text) + CDbl(ipp_Import2.Text) + CDbl(ipp_Import3.Text) + CDbl(ipp_Import4.Text)) = 0 Then
      MsgBox "En el grupo determinación tiene que ingresar un importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import1)
      Exit Sub
   End If
   
'   If (Format(ipp_FchOpe.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FchOpe.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta Registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FchOpe)
'       Exit Sub
'   End If

   If fs_ValidaPeriodo(ipp_FchOpe.Text) = False Then
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Grabar
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   cmd_Reversa.Visible = False
   cmd_Grabar.Visible = False
   If moddat_g_int_FlgGrb = 0 Then
      'consultar
      pnl_TitPri.Caption = "Cuentas por Pagar - Consultar"
      Call fs_Cargar_Datos
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then
      'registrar
      pnl_TitPri.Caption = "Cuentas por Pagar - Adicionar"
      cmd_Grabar.Visible = True
   ElseIf moddat_g_int_FlgGrb = 2 Then
      'editar
      pnl_TitPri.Caption = "Cuentas por Pagar - Modificar"
      cmd_Grabar.Visible = True
      Call fs_Cargar_Datos
   ElseIf moddat_g_int_FlgGrb = 3 Then
      'reversa
      pnl_TitPri.Caption = "Cuentas por Pagar - Reversa"
      cmd_Reversa.Left = 30
      cmd_Reversa.Visible = True
      Call fs_Desabilitar
      Call fs_Cargar_Datos
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Limpia()
   ipp_FchOpe.Text = moddat_g_str_FecSis
   txt_Descrip.Text = ""
   Call gs_BuscarCombo_Item(cmb_TipDoc, 6)
   cmb_Proveedor.Text = ""
   cmb_Moneda.ListIndex = 0
   cmb_Banco.ListIndex = -1
   cmb_CtaCte.ListIndex = -1

   ipp_Import1.Text = "0.00"
   ipp_Import2.Text = "0.00"
   ipp_Import3.Text = "0.00"
   ipp_Import4.Text = "0.00"
   
   cmb_CtaImp_01.Text = ""
   cmb_CtaImp_02.Text = ""
   cmb_CtaImp_03.Text = ""
   cmb_CtaImp_04.Text = ""
End Sub

Private Sub fs_Inicia()
Dim r_int_Contar     As Integer

   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   
   'cargar las cuentas contables
   l_int_TopNiv = -1
   moddat_g_str_CodEmp = "000001"
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaImp_01, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   
   cmb_CtaImp_02.Clear
   cmb_CtaImp_03.Clear
   cmb_CtaImp_04.Clear
   For r_int_Contar = 1 To UBound(l_arr_CtaCtb)
       cmb_CtaImp_02.AddItem l_arr_CtaCtb(r_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(r_int_Contar).Genera_Nombre
       cmb_CtaImp_03.AddItem l_arr_CtaCtb(r_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(r_int_Contar).Genera_Nombre
       cmb_CtaImp_04.AddItem l_arr_CtaCtb(r_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(r_int_Contar).Genera_Nombre
   Next
   
End Sub

Public Sub fs_Desabilitar()
   ipp_FchOpe.Enabled = False
   txt_Descrip.Enabled = False
   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   cmb_Moneda.Enabled = False
   cmb_Banco.Enabled = False
   cmb_CtaCte.Enabled = False
   ipp_Import1.Enabled = False
   ipp_Import2.Enabled = False
   ipp_Import3.Enabled = False
   ipp_Import4.Enabled = False
   cmb_CtaImp_01.Enabled = False
   cmb_CtaImp_02.Enabled = False
   cmb_CtaImp_03.Enabled = False
   cmb_CtaImp_04.Enabled = False
End Sub

Private Sub fs_Grabar()
Dim r_str_CodGen  As String

    r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 12)
   Else
      r_str_CodGen = Trim(pnl_Codigo.Caption)
   End If
   
   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CTAPAG ( "
   g_str_Parame = g_str_Parame & CLng(r_str_CodGen) & ", " 'CTAPAG_CODPAG
   g_str_Parame = g_str_Parame & Format(ipp_FchOpe.Text, "yyyymmdd") & ", " 'CTAPAG_FECOPE
   g_str_Parame = g_str_Parame & "null, " 'CTAPAG_TIPOPE
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'CTAPAG_CODMON
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'CTAPAG_TIPDOC
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text, cmb_TipDoc) & "', " 'CTAPAG_NUMDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descrip.Text) & "', " 'CTAPAG_DESCRP
   'g_str_Parame = g_str_Parame & CDbl(pnl_TipCambio.Caption) & ", "  'CTAPAG_TIPCAM
   g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", " 'CTAPAG_CODBCO
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', "  'CTAPAG_CTACRR
   g_str_Parame = g_str_Parame & "0, " 'CTAPAG_IMPPAG
   g_str_Parame = g_str_Parame & "1, "  'CTAPAG_SITUAC
   g_str_Parame = g_str_Parame & "2, "  'CTAPAG_TIPTAB
   g_str_Parame = g_str_Parame & CDbl(ipp_Import1.Text) & ", " 'CTAPAG_IMPPG1
   g_str_Parame = g_str_Parame & CDbl(ipp_Import2.Text) & ", " 'CTAPAG_IMPPG2
   g_str_Parame = g_str_Parame & CDbl(ipp_Import3.Text) & ", " 'CTAPAG_IMPPG3
   g_str_Parame = g_str_Parame & CDbl(ipp_Import4.Text) & ", " 'CTAPAG_IMPPG4
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaImp_01.Text = "", "", Mid(cmb_CtaImp_01.Text, 1, l_int_TopNiv)) & "', " 'CTAPAG_NUCTA1
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaImp_02.Text = "", "", Mid(cmb_CtaImp_02.Text, 1, l_int_TopNiv)) & "', " 'CTAPAG_NUCTA2
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaImp_03.Text = "", "", Mid(cmb_CtaImp_03.Text, 1, l_int_TopNiv)) & "', " 'CTAPAG_NUCTA3
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaImp_04.Text = "", "", Mid(cmb_CtaImp_04.Text, 1, l_int_TopNiv)) & "', " 'CTAPAG_NUCTA4
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CtaPag_01.fs_Buscar
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CtaPag_01.fs_Buscar
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub cmd_Reversa_Click()
Dim r_bol_Estado As Boolean
   
   If MsgBox("¿Esta seguro que desea realizar esta operación de reversa?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CTAPAG_REVERSA ( "
   g_str_Parame = g_str_Parame & " " & CLng(Trim(pnl_Codigo.Caption)) & ", "  'CTAPAG_CODPAG
   g_str_Parame = g_str_Parame & " 2, "  'CTAPAG_TIPTAB
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
                                                                                                                                                                                                                 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la operación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (g_rst_Genera!as_resul = 1) Then
       'reversa normal
       MsgBox "Se completo la operación de reversa.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CtaPag_01.fs_Buscar
       Unload Me
   ElseIf (g_rst_Genera!as_resul = 2) Then
       'procesado compensacion
       MsgBox "No se puede revertir, el registro se encuentra en el modulo de contabilidad.", vbExclamation, modgen_g_str_NomPlt
       Call frm_Ctb_CtaPag_01.fs_Buscar
       Unload Me
   ElseIf (g_rst_Genera!as_resul = 3) Then
       'procesado cuentas x pagar
       MsgBox "El registro ya fue revertido.", vbExclamation, modgen_g_str_NomPlt
       Call frm_Ctb_CtaPag_01.fs_Buscar
       Unload Me
   Else
       MsgBox "Favor de verificar la operación de reversa.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Public Sub fs_Cargar_Datos()
Dim r_str_Parame   As String
Dim r_rst_Princi   As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT CTAPAG_CODPAG, CTAPAG_FECOPE, CTAPAG_CODMON,  "
   r_str_Parame = r_str_Parame & "         CTAPAG_TIPDOC, CTAPAG_NUMDOC, CTAPAG_DESCRP, CTAPAG_CODBCO,  "
   r_str_Parame = r_str_Parame & "         CTAPAG_CTACRR, " 'CTAPAG_TIPCAM
   r_str_Parame = r_str_Parame & "         CTAPAG_IMPPG1, CTAPAG_IMPPG2, CTAPAG_IMPPG3, CTAPAG_IMPPG4,  "
   r_str_Parame = r_str_Parame & "         CTAPAG_NUCTA1, CTAPAG_NUCTA2, CTAPAG_NUCTA3, CTAPAG_NUCTA4   "
   r_str_Parame = r_str_Parame & "    FROM CNTBL_CTAPAG A  "
   r_str_Parame = r_str_Parame & "   WHERE A.CTAPAG_CODPAG =  " & CLng(moddat_g_str_Codigo)
   r_str_Parame = r_str_Parame & "     AND A.CTAPAG_TIPTAB = 2  "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      pnl_Codigo.Caption = Format(r_rst_Princi!CTAPAG_CODPAG, "0000000000")
      ipp_FchOpe.Text = gf_FormatoFecha(r_rst_Princi!CTAPAG_FECOPE)
      txt_Descrip.Text = Trim(r_rst_Princi!CTAPAG_DESCRP & "")
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_rst_Princi!CTAPAG_TIPDOC)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, r_rst_Princi!CTAPAG_NUMDOC & "", 0)
      Call gs_BuscarCombo_Item(cmb_Moneda, r_rst_Princi!CTAPAG_CODMON)
      Call gs_BuscarCombo_Item(cmb_Banco, r_rst_Princi!CTAPAG_CODBCO)
      Call gs_BuscarCombo_Text(cmb_CtaCte, r_rst_Princi!CTAPAG_CTACRR, -1)
      
      ipp_Import1.Text = CDbl(r_rst_Princi!CTAPAG_IMPPG1)
      ipp_Import2.Text = CDbl(r_rst_Princi!CTAPAG_IMPPG2)
      ipp_Import3.Text = CDbl(r_rst_Princi!CTAPAG_IMPPG3)
      ipp_Import4.Text = CDbl(r_rst_Princi!CTAPAG_IMPPG4)
      
      cmb_CtaImp_01.ListIndex = fs_ComboIndex(cmb_CtaImp_01, r_rst_Princi!CTAPAG_NUCTA1 & "", l_int_TopNiv)
      cmb_CtaImp_02.ListIndex = fs_ComboIndex(cmb_CtaImp_02, r_rst_Princi!CTAPAG_NUCTA2 & "", l_int_TopNiv)
      cmb_CtaImp_03.ListIndex = fs_ComboIndex(cmb_CtaImp_03, r_rst_Princi!CTAPAG_NUCTA3 & "", l_int_TopNiv)
      cmb_CtaImp_04.ListIndex = fs_ComboIndex(cmb_CtaImp_04, r_rst_Princi!CTAPAG_NUCTA4 & "", l_int_TopNiv)
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
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

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv(cmb_TipDoc, cmb_Proveedor)
End Sub

Private Sub fs_CargarPrv(p_Combo_Tdoc As ComboBox, p_Combo_Nom As ComboBox)
   ReDim l_arr_MaePrv(0) 'BENEFICIARIO(2)
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   p_Combo_Nom.Clear
   p_Combo_Nom.Text = ""
   If (p_Combo_Tdoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & p_Combo_Tdoc.ItemData(p_Combo_Tdoc.ListIndex)
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
      p_Combo_Nom.AddItem Trim(g_rst_Genera!MAEPRV_NUMDOC & "") & " - " & Trim(g_rst_Genera!MaePrv_RazSoc & "")
         
      ReDim Preserve l_arr_MaePrv(UBound(l_arr_MaePrv) + 1)
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmb_Proveedor_Click()
   Call fs_Buscar_Ctas
End Sub

Private Sub cmb_Moneda_Click()
   Call fs_CargarBancos
End Sub

Private Sub fs_Buscar_Ctas()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   r_str_NumDoc = ""
   
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc)
          Exit Sub
       End If
       If cmb_Proveedor.ListIndex = -1 Then
          Exit Sub
       End If
      
       If (fs_ValNumDoc(cmb_TipDoc, cmb_Proveedor) = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text, cmb_TipDoc)
   
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
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   
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
          MsgBox "El beneficiario se encuentra en condición de NO HABIDO, revisar sunat.", vbExclamation, modgen_g_str_NomPlt
          g_rst_GenAux.Close
          Set g_rst_GenAux = Nothing
          Exit Sub
       End If
       'Call gs_SetFocus(txt_Descrip)
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

Private Sub fs_CargarBancos()
Dim r_bol_Estado   As Boolean
Dim r_int_File     As Integer
Dim r_int_Contar   As Integer

   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   'soles
   If (cmb_Moneda.ListIndex = 0) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)
           End If
       Next
   End If
   'dolares
   If (cmb_Moneda.ListIndex = 1) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)
           End If
       Next
   End If
End Sub

Private Sub cmb_Banco_Click()
Dim r_str_Cadena  As String
Dim r_int_Contar  As Integer

   cmb_CtaCte.Clear
   r_str_Cadena = ""
   lbl_Cuenta.Caption = "Cuenta:"
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   If (cmb_Moneda.ListIndex = 0) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
   
   If (cmb_Moneda.ListIndex = 1) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
End Sub

Private Function fs_ValNumDoc(p_ComboTip As ComboBox, p_ComboNom As ComboBox) As Boolean
Dim r_str_NumDoc  As String
Dim r_bol_Estado  As Boolean

   fs_ValNumDoc = True
   r_str_NumDoc = ""

   r_str_NumDoc = fs_NumDoc(p_ComboNom.Text, p_ComboTip)
   If (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 1) Then 'DNI - 8
       If Len(Trim(r_str_NumDoc)) <> 8 Then
          MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   ElseIf (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 6) Then 'RUC - 11
       If Not gf_Valida_RUC(Trim(r_str_NumDoc), Mid(Trim(r_str_NumDoc), Len(Trim(r_str_NumDoc)), 1)) Then
          MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   Else 'OTROS
       If Len(Trim(p_ComboNom.Text)) = 0 Then
          MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   End If
   
End Function

Private Function fs_NumDoc(p_Cadena As String, p_ComboTip As ComboBox) As String
   fs_NumDoc = ""
   If (p_ComboTip.ListIndex > -1) Then
      If (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
           If InStr(1, p_Cadena, "-") <= 0 Then
              Exit Function
           End If
           fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
      End If
   End If
End Function

Private Function fs_ValPlanCta(p_Cuenta As String) As Boolean
   fs_ValPlanCta = True
   
   p_Cuenta = Mid(p_Cuenta, 1, l_int_TopNiv)
   If (Len(Trim(p_Cuenta)) = 0) Then
       MsgBox "Debe de ingresar las cuentas en el grupo determinación.", vbExclamation, modgen_g_str_NomPlt
       fs_ValPlanCta = False
       Exit Function
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT CTAMAE_CODCTA, CTAMAE_REGCOM "
   g_str_Parame = g_str_Parame & "   FROM CTB_CTAMAE "
   g_str_Parame = g_str_Parame & "  WHERE CTAMAE_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & "    AND CTAMAE_CODCTA = '" & Trim(p_Cuenta) & "'"
   'g_str_Parame = g_str_Parame & "    AND CTAMAE_REGCOM = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If (g_rst_Princi!CTAMAE_REGCOM <> 1) Then
          MsgBox "La cuenta " & Trim(p_Cuenta) & ", debe de estar registrada como REGISTRO COMPROBANTE", vbExclamation, modgen_g_str_NomPlt
          fs_ValPlanCta = False
          Exit Function
      End If
   Else
      MsgBox "La cuenta " & Trim(p_Cuenta) & ", no esta registrada en el sistema.", vbExclamation, modgen_g_str_NomPlt
      fs_ValPlanCta = False
      Exit Function
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Sub ipp_FchOpe_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Descrip)
   End If
End Sub

Private Sub txt_Descrip_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipDoc)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
   End If
End Sub

Private Sub cmb_Proveedor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   End If
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_Import1)
   End If
End Sub

Private Sub ipp_Import1_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaImp_01)
   End If
End Sub

Private Sub cmb_CtaImp_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_Import2)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub ipp_Import2_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaImp_02)
   End If
End Sub

Private Sub cmb_CtaImp_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_Import3)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub ipp_Import3_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaImp_03)
   End If
End Sub

Private Sub cmb_CtaImp_03_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_Import4)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub ipp_Import4_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaImp_04)
   End If
End Sub

Private Sub cmb_CtaImp_04_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub



