VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_InvDpf_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "GesCtb_frm_197.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6930
      Left            =   -360
      TabIndex        =   15
      Top             =   0
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   12224
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
         Left            =   420
         TabIndex        =   16
         Top             =   60
         Width           =   8700
         _Version        =   65536
         _ExtentX        =   15346
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
            TabIndex        =   17
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Gestión de Depósito Plazo Fijo"
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
            Picture         =   "GesCtb_frm_197.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4125
         Left            =   420
         TabIndex        =   18
         Top             =   2670
         Width           =   8700
         _Version        =   65536
         _ExtentX        =   15346
         _ExtentY        =   7276
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
         Begin VB.Frame gb_Import_Ori 
            Caption         =   "Importes Originales"
            Height          =   1605
            Left            =   5190
            TabIndex        =   44
            Top             =   2400
            Width           =   3285
            Begin Threed.SSPanel pnl_IntCap_Ori 
               Height          =   315
               Left            =   1470
               TabIndex        =   45
               Top             =   675
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2831
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
            Begin Threed.SSPanel pnl_SalCap_Ori 
               Height          =   315
               Left            =   1470
               TabIndex        =   46
               Top             =   330
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2831
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
            Begin Threed.SSPanel pnl_Total_Ori 
               Height          =   315
               Left            =   1470
               TabIndex        =   49
               Top             =   1080
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2831
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
            Begin VB.Line Line1 
               X1              =   390
               X2              =   3075
               Y1              =   1020
               Y2              =   1020
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   360
               TabIndex        =   50
               Top             =   1155
               Width           =   405
            End
            Begin VB.Label lbl_IntCap_Ref 
               AutoSize        =   -1  'True
               Caption         =   "Rendimiento:"
               Height          =   195
               Left            =   360
               TabIndex        =   48
               Top             =   750
               Width           =   930
            End
            Begin VB.Label lbl_SalCap_Ref 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   360
               TabIndex        =   47
               Top             =   405
               Width           =   570
            End
         End
         Begin VB.TextBox txt_OperRef_Dat 
            Height          =   315
            Left            =   1590
            MaxLength       =   15
            TabIndex        =   12
            Top             =   2685
            Width           =   3400
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1050
            Width           =   3400
         End
         Begin VB.ComboBox cmb_BancOrigen 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2355
            Width           =   3400
         End
         Begin VB.ComboBox cmb_BancDestino 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   3400
         End
         Begin EditLib.fpDateTime ipp_FecOper 
            Height          =   315
            Left            =   1590
            TabIndex        =   3
            Top             =   390
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
         Begin EditLib.fpDoubleSingle ipp_SalCap 
            Height          =   315
            Left            =   1590
            TabIndex        =   6
            Top             =   1380
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
         Begin EditLib.fpDoubleSingle ipp_TasInt 
            Height          =   315
            Left            =   6630
            TabIndex        =   9
            Top             =   1695
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   1695
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
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
            Text            =   "1"
            MaxValue        =   "9999"
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
         Begin Threed.SSPanel pnl_IntCap 
            Height          =   315
            Left            =   1590
            TabIndex        =   10
            Top             =   2025
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin EditLib.fpDateTime ipp_FecCierre 
            Height          =   315
            Left            =   1590
            TabIndex        =   36
            Top             =   3030
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
            AllowNull       =   -1  'True
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
            Text            =   ""
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
         Begin EditLib.fpDateTime ipp_FecQuie 
            Height          =   315
            Left            =   1590
            TabIndex        =   38
            Top             =   3360
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
            AllowNull       =   -1  'True
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
            Text            =   ""
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
         Begin Threed.SSPanel pnl_NumCta_Ref 
            Height          =   315
            Left            =   1590
            TabIndex        =   40
            Top             =   3690
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin EditLib.fpDoubleSingle ipp_IntAjus 
            Height          =   315
            Left            =   6630
            TabIndex        =   7
            Top             =   1380
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ajuste:"
            Height          =   195
            Left            =   5610
            TabIndex        =   42
            Top             =   1470
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta Ref.:"
            Height          =   195
            Left            =   150
            TabIndex        =   41
            Top             =   3750
            Width           =   1200
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Quiebre:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cierre:"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   3090
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Operación de Ref.:"
            Height          =   195
            Left            =   150
            TabIndex        =   28
            Top             =   2760
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Origen de Fondos:"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   2430
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Plazo Días:"
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   1770
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rendimiento:"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   2100
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tasa (%):"
            Height          =   195
            Left            =   5610
            TabIndex        =   24
            Top             =   1770
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   480
            Width           =   1140
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
            TabIndex        =   22
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Institución:"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   1140
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   1470
            Width           =   570
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   1125
         Left            =   420
         TabIndex        =   29
         Top             =   1500
         Width           =   8700
         _Version        =   65536
         _ExtentX        =   15346
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
         Begin VB.ComboBox cmb_TipAccion 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   660
            Width           =   3400
         End
         Begin Threed.SSPanel pnl_Estado 
            Height          =   315
            Left            =   6630
            TabIndex        =   1
            Top             =   330
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
         Begin Threed.SSPanel pnl_NumCta 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            Top             =   330
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   5610
            TabIndex        =   34
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Principal"
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
            TabIndex        =   32
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Datos Renovación"
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
            Left            =   450
            TabIndex        =   31
            Top             =   1290
            Width           =   1590
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Acción:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   720
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   420
         TabIndex        =   35
         Top             =   780
         Width           =   8700
         _Version        =   65536
         _ExtentX        =   15346
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
            Picture         =   "GesCtb_frm_197.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   8070
            Picture         =   "GesCtb_frm_197.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_197.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_InvDpf_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_TipAcc()      As moddat_tpo_Genera
Dim l_str_OpeRef        As String
Dim l_str_SalCap        As String
Dim l_str_IntCap        As String
Dim l_str_IntAju        As String
Dim l_str_FecApe        As String
Dim l_str_TasInt        As String
Dim l_str_PlaDia        As String
Dim l_int_EntInt        As Integer
Dim l_int_EntFon        As Integer
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer

Private Sub cmb_TipAccion_Click()
   If moddat_g_int_Situac = 2 And moddat_g_int_FlgGrb = 3 Then 'VIGENTE - GESTIONAR
      pnl_NumCta.Caption = moddat_g_str_Codigo
      ipp_SalCap.Text = Format(l_str_SalCap, "###,###,##0.00") & " "
      ipp_TasInt.Text = Format(l_str_TasInt, "###,###,##0.00") & " "
      ipp_PlaAno.Text = l_str_PlaDia
      pnl_IntCap.Caption = Format(l_str_IntCap, "###,###,##0.00") & " "
      ipp_IntAjus.Text = Format(l_str_IntAju, "###,###,##0.00") & " "
      'pnl_IntCap_Fin.Caption = Format(CDbl(l_str_IntCap) + CDbl(l_str_IntAju), "###,###,##0.00") & " "
      ipp_FecOper.Text = l_str_FecApe
      ipp_FecOper.Enabled = False
      ipp_PlaAno.Enabled = False
      ipp_IntAjus.Enabled = False
      ipp_TasInt.Enabled = False
      txt_OperRef_Dat.Text = l_str_OpeRef
      ipp_FecOper.Text = l_str_FecApe
      Call gs_BuscarCombo_Item(cmb_BancOrigen, l_int_EntFon)
      'pnl_SalCap_Ref.Enabled = False
      'pnl_IntCap_Ref.Enabled = False
      'lbl_SalCap_Ref.Enabled = False
      'lbl_IntCap_Ref.Enabled = False
      gb_Import_Ori.Visible = False
      pnl_SalCap_Ori.Caption = Format(l_str_SalCap, "###,###,##0.00") & " "
      pnl_IntCap_Ori.Caption = Format(l_str_IntCap, "###,###,##0.00") & " "
      pnl_Total_Ori.Caption = Format(CDbl(l_str_SalCap) + CDbl(l_str_IntCap), "###,###,##0.00") & " "
      If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 2 Then 'CIERRE
         ipp_FecCierre.Enabled = True
         ipp_FecCierre.AllowNull = False
         ipp_FecCierre.Text = moddat_g_str_FecSis
         ipp_FecQuie.Enabled = False
         ipp_FecQuie.AllowNull = True
         ipp_FecQuie.Text = ""
         txt_OperRef_Dat.Enabled = True
         ipp_IntAjus.Enabled = False
      ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 5 Then 'QUIEBRE
         ipp_FecCierre.Enabled = False
         ipp_FecCierre.AllowNull = True
         ipp_FecCierre.Text = ""
         ipp_FecQuie.Enabled = True
         ipp_FecQuie.AllowNull = False
         ipp_FecQuie.Text = moddat_g_str_FecSis
         txt_OperRef_Dat.Enabled = True
         ipp_IntAjus.Enabled = False
      ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 3 Or _
             cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 4 Then 'RENOVAR
         ipp_FecCierre.Enabled = False
         ipp_FecQuie.Enabled = False
         ipp_FecCierre.AllowNull = True
         ipp_FecCierre.Text = ""
         ipp_FecQuie.AllowNull = True
         ipp_FecQuie.Text = ""
         ipp_FecOper.Enabled = True
         ipp_TasInt.Enabled = True
         txt_OperRef_Dat.Enabled = True
         ipp_IntAjus.Enabled = True
         ipp_PlaAno.Enabled = True
         ipp_FecOper.Text = moddat_g_str_FecSis
         gb_Import_Ori.Visible = True
         ipp_IntAjus.Text = "0.00" & " "
         Call gs_BuscarCombo_Item(cmb_BancOrigen, l_int_EntInt)
         If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 3 Then
            'renovacion con capital
            'ipp_SalCap.Text = Format(CDbl(l_str_SalCap) + CDbl(l_str_IntCap) + CDbl(l_str_IntAju), "###,###,##0.00")
            ipp_SalCap.Text = Format(CDbl(l_str_SalCap) + CDbl(l_str_IntCap), "###,###,##0.00")
            Call fs_CalInteres
         End If
         If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 4 Then
            'renovacion sin capital
            ipp_SalCap.Text = Format(CDbl(l_str_SalCap), "###,###,##0.00")
            Call fs_CalInteres
         End If
      End If
   End If
End Sub

Private Sub cmd_Reversa_Click()
Dim r_bol_Estado As Boolean
    
   If Trim(pnl_NumCta_Ref.Caption) = "" Then
      MsgBox "No tiene un nro cuenta de referencia para darle reversa.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_bol_Estado = False
   Call fs_BusDependencia(moddat_g_str_Codigo, r_bol_Estado)
   If r_bol_Estado = True Then
      MsgBox "El DPF " & pnl_NumCta.Caption & ", tiene una renovacion de nivel superior, tiene que darle reversa.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro que desea realizar esta operación de reversa?" & vbCrLf & _
             "Recuerde que debe eliminar el asiento contable manual.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEDPF_REVERSA ( "
   g_str_Parame = g_str_Parame & " " & CLng(Trim(pnl_NumCta.Caption)) & ", "  'MAEDPF_NUMCTA
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
                                                                                                                                                                                                                 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la operación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (g_rst_Genera!as_resul = 1) Then
       MsgBox "Se completo la operación de reversa, recuerde que debe eliminar el asiento contable manual.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_InvDpf_01.fs_BuscarComp
       
       Unload Me
   Else
       MsgBox "Favor de verificar la operación de reversa.", vbInformation, modgen_g_str_NomPlt
   End If
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_bol_Estado    As Boolean

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   l_str_OpeRef = ""
   l_str_SalCap = "0"
   l_str_IntCap = "0"
   l_str_IntAju = "0"
   l_str_TasInt = ""
   l_int_EntInt = -1
   l_int_EntFon = -1
   cmd_Reversa.Visible = False
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then
      pnl_Titulo.Caption = "Depósito Plazo Fijo - Consultar"
      cmd_Grabar.Visible = False
      Call fs_Desabilitar(False)
      Call fs_Cargar_Datos
   ElseIf moddat_g_int_FlgGrb = 1 Then
      pnl_Titulo.Caption = "Depósito Plazo Fijo - Apertura"
      Call gs_BuscarCombo_Item(cmb_TipAccion, 1)
      Call fs_Desabilitar(True)
   ElseIf moddat_g_int_FlgGrb = 2 Then
      pnl_Titulo.Caption = "Depósito Plazo Fijo - Modificar"
      Call fs_Cargar_Datos
      Call fs_Desabilitar(False)
   ElseIf moddat_g_int_FlgGrb = 3 Then
      pnl_Titulo.Caption = "Depósito Plazo Fijo - Gestionar"
      Call fs_Desabilitar(False)
      Call fs_Cargar_Datos
   ElseIf moddat_g_int_FlgGrb = 4 Then 'SOLO RENOVACIONES
      Call fs_Cargar_Datos
      Call fs_Desabilitar(False)
      cmd_Reversa.Visible = True
      cmd_Grabar.Visible = False
      cmd_Reversa.Left = 30
      pnl_Titulo.Caption = "Depósito Plazo Fijo - Reversa"
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contar  As Integer

   Call moddat_gs_Carga_LisIte(cmb_TipAccion, l_arr_TipAcc, 1, 128, 1)
   cmb_TipAccion.Clear
   For r_int_Contar = 1 To UBound(l_arr_TipAcc)
       cmb_TipAccion.AddItem Trim$(l_arr_TipAcc(r_int_Contar).Genera_Nombre)
       cmb_TipAccion.ItemData(cmb_TipAccion.NewIndex) = CLng(l_arr_TipAcc(r_int_Contar).Genera_Codigo)
   Next
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BancDestino, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BancOrigen, 1, "122")
End Sub

Private Sub fs_Limpiar()
Dim r_int_Contar  As Integer

   gb_Import_Ori.Visible = False
   
   pnl_NumCta.Caption = ""
   cmb_TipAccion.ListIndex = -1
   pnl_NumCta.Caption = ""
   ipp_FecOper.Text = moddat_g_str_FecSis
   cmb_BancDestino.ListIndex = -1
   cmb_Moneda.ListIndex = 0
   ipp_SalCap.Text = "0.00"
   pnl_IntCap.Caption = "0.00 "
   ipp_IntAjus.Text = "0.00"
   ipp_PlaAno.Text = 1
   
   ipp_FecCierre.Text = ""
   pnl_NumCta_Ref.Caption = ""
   'pnl_OperRef_Ref.Caption = ""
   'pnl_SalCap.Caption = "0.00 "
   'pnl_IntCap_Ref.Caption = "0.00 "
   cmb_TipAccion.ListIndex = 0
   pnl_Estado.Caption = Trim(moddat_g_str_Situac)
   
   If moddat_g_int_FlgGrb = 3 Then
      If moddat_g_int_Situac = 1 Then
         Call gs_BuscarCombo_Item(cmb_TipAccion, 5)
         cmb_TipAccion.Enabled = False
      Else
         cmb_TipAccion.Clear
         For r_int_Contar = 1 To UBound(l_arr_TipAcc)
             Select Case CLng(l_arr_TipAcc(r_int_Contar).Genera_Codigo)
                    Case 2, 3, 4, 5
                         cmb_TipAccion.AddItem Trim$(l_arr_TipAcc(r_int_Contar).Genera_Nombre)
                         cmb_TipAccion.ItemData(cmb_TipAccion.NewIndex) = CLng(l_arr_TipAcc(r_int_Contar).Genera_Codigo)
             End Select
         Next
         cmb_TipAccion.Enabled = True
      End If
   End If
End Sub

Private Sub fs_CalInteres()
Dim r_dbl_IntCap    As Double
'Dim r_dbl_NueInt    As Double

   If moddat_g_int_FlgGrb = 3 Then  'GESTIONAR
      If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 3 Then  'RENOVAR
         ipp_SalCap.Text = Format((CDbl(l_str_SalCap) + CDbl(l_str_IntCap)) + CDbl(ipp_IntAjus.Text), "###,###,##0.00") & " "
      Else
         ipp_SalCap.Text = Format(CDbl(l_str_SalCap) + CDbl(ipp_IntAjus.Text), "###,###,##0.00") & " "
      End If
   End If
   '=((((1+TASA)^(PLAZO/360))-1)*IMPORTE)
   r_dbl_IntCap = (((1 + (CDbl(ipp_TasInt.Text) / 100)) ^ (CDbl(ipp_PlaAno.Text) / 360)) - 1) * CDbl(ipp_SalCap.Text)
   pnl_IntCap.Caption = Format(r_dbl_IntCap, "###,###,##0.00") & " "
   
   'r_dbl_NueInt = r_dbl_IntCap + CDbl(ipp_IntAjus.Text)
   'pnl_IntCap_Fin.Caption = Format(r_dbl_NueInt, "###,###,##0.00") & " "
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_CtaDeb_1  As String
Dim r_str_CtaHab_1  As String
Dim r_str_CtaHab_2  As String
Dim r_bol_Estado    As Boolean
Dim r_dbl_TipCam    As Double

   If cmb_TipAccion.ListIndex = -1 Then
       MsgBox "Tiene que seleccionar un tipo de acción.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipAccion)
       Exit Sub
   End If
   
   If cmb_BancDestino.ListIndex = -1 Then
       MsgBox "Tiene que seleccionar una institución como destino.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_BancDestino)
       Exit Sub
   End If
   
   If cmb_Moneda.ListIndex = -1 Then
       MsgBox "Tiene que seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Moneda)
       Exit Sub
   End If
   
   If CDbl(ipp_TasInt.Text) <= 0 Then
       MsgBox "Tiene que ingresar la tasa de interés.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_TasInt)
       Exit Sub
   End If
   
   If cmb_BancOrigen.ListIndex = -1 Then
       MsgBox "Tiene que selecionar el origen de los fondos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_BancOrigen)
       Exit Sub
   End If
   
   If Len(Trim(txt_OperRef_Dat.Text)) <= 0 Then
       MsgBox "Tiene que digitar el numero de operación de referencia.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_OperRef_Dat)
       Exit Sub
   End If
   
   
   Dim r_str_FecDpf   As String
   Dim r_dbl_TipSbs   As Double
   r_str_FecDpf = ""
   r_dbl_TipCam = 0
   If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 2 Then
      'Cierre
      r_str_FecDpf = ipp_FecCierre.Text
   ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 5 Then
      'Quiebre
      r_str_FecDpf = ipp_FecQuie.Text
   Else
      r_str_FecDpf = ipp_FecOper.Text
   End If
   
   If fs_ValidaPeriodo(r_str_FecDpf) = False Then
      Exit Sub
   End If
   
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(r_str_FecDpf, "yyyymmdd"), 1)
    If CDbl(r_dbl_TipSbs) = 0 Then
       MsgBox "El tipo de cambio no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
       If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 2 Then
          Call gs_SetFocus(ipp_FecCierre)
       ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 5 Then
          Call gs_SetFocus(ipp_FecQuie)
       Else
          Call gs_SetFocus(ipp_FecOper)
       End If
       Exit Sub
    End If
    
   If (Format(r_str_FecDpf, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(r_str_FecDpf, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
       If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 2 Then
          Call gs_SetFocus(ipp_FecCierre)
       ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 5 Then
          Call gs_SetFocus(ipp_FecQuie)
       Else
          Call gs_SetFocus(ipp_FecOper)
       End If
       Exit Sub
   End If
         
   If moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 3 Then
      Call fs_BusCuentas(r_str_CtaDeb_1, r_str_CtaHab_1, r_str_CtaHab_2, r_bol_Estado)
      If r_bol_Estado = False Then
         MsgBox "No existe ninguna cuenta contable para generar el asiento, debe agregar su dinámica", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
    
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   Screen.MousePointer = 11
   
   If moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2 Then
      Call fs_Grabar
   ElseIf moddat_g_int_FlgGrb = 3 Then
      Call fs_Grabar_Gst
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar_Gst()
Dim r_str_AsiGen   As String
   
   r_str_AsiGen = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEDPF_GESTION ( "
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'MAEDPF_NUMCTA
   g_str_Parame = g_str_Parame & IIf(Trim(ipp_FecCierre.Text) = "", "NULL", Format(ipp_FecCierre.Text, "yyyymmdd")) & ", " 'MAEDPF_FECCIE
   g_str_Parame = g_str_Parame & IIf(Trim(ipp_FecQuie.Text) = "", "NULL", Format(ipp_FecQuie.Text, "yyyymmdd")) & ", " 'MAEDPF_FECQUI
   g_str_Parame = g_str_Parame & IIf(Trim(ipp_FecOper.Text) = "", "NULL", Format(ipp_FecOper.Text, "yyyymmdd")) & ", " 'MAEDPF_FECRNV
   g_str_Parame = g_str_Parame & "'" & Trim(txt_OperRef_Dat.Text) & "', " 'MAEDPF_NUMREF
   g_str_Parame = g_str_Parame & cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) & ", " 'MAEDPF_TIPDPF
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If g_rst_Genera!as_resul = 1 Then
      If cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 2 Then
      'cierre de deposito
         Call fs_GeneraAsiento(pnl_NumCta.Caption, txt_OperRef_Dat.Text, ipp_FecCierre.Text, 2, r_str_AsiGen)
         MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
                "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
         Call frm_Ctb_InvDpf_01.fs_BuscarComp
         Screen.MousePointer = 0
         Unload Me
      ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 5 Then
      'quiebre de deposito
         Call fs_GeneraAsiento(pnl_NumCta.Caption, txt_OperRef_Dat.Text, ipp_FecQuie.Text, 5, r_str_AsiGen)
         MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
                "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
         Call frm_Ctb_InvDpf_01.fs_BuscarComp
         Screen.MousePointer = 0
         Unload Me
      ElseIf cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 3 Or cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) = 4 Then
             'renovacion con capital   -   renovacion sin capital
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " USP_CNTBL_MAEDPF ( "
             g_str_Parame = g_str_Parame & "'" & "" & "', " 'MAEDPF_NUMCTA
             g_str_Parame = g_str_Parame & Format(ipp_FecOper.Text, "yyyymmdd") & ", " 'MAEDPF_FECAPE
             g_str_Parame = g_str_Parame & "1, " 'MAEDPF_TIPDPF - APERTURA DE DEPOSITO
             g_str_Parame = g_str_Parame & cmb_BancDestino.ItemData(cmb_BancDestino.ListIndex) & ", " 'MAEDPF_CODBNC_APE
             g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "  'MAEDPF_CODMOD
             g_str_Parame = g_str_Parame & CDbl(ipp_SalCap.Text) & ", " 'MAEDPF_SALCAP
             g_str_Parame = g_str_Parame & CDbl(pnl_IntCap.Caption) & ", "  'MAEDPF_INTCAP
             g_str_Parame = g_str_Parame & CDbl(ipp_IntAjus.Text) & ", "  'MAEDPF_INTAJU
             g_str_Parame = g_str_Parame & CDbl(ipp_TasInt.Text) & ", "   'MAEDPF_TASINT
             g_str_Parame = g_str_Parame & CLng(ipp_PlaAno) & ", "   'MAEDPF_PLADIA
             g_str_Parame = g_str_Parame & cmb_BancOrigen.ItemData(cmb_BancOrigen.ListIndex) & ", "   'MAEDPF_CODBNC_ORI
             g_str_Parame = g_str_Parame & "'" & Trim(txt_OperRef_Dat.Text) & "', "   'MAEDPF_NUMREF
             g_str_Parame = g_str_Parame & CLng(moddat_g_str_Codigo) & ", " 'MAEDPF_NUMCTA_REF
             g_str_Parame = g_str_Parame & "1, "   'MAEDPF_SITUAC
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
             g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
             g_str_Parame = g_str_Parame & "1) "
   
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
                Exit Sub
             End If
   
             If (g_rst_Genera!RESUL = 1) Then
                 'GRABACION OK -- pnl_NumCta.Caption
                 Call fs_GeneraAsiento(g_rst_Genera!CODIGO, txt_OperRef_Dat.Text, ipp_FecOper.Text, _
                                       cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex), r_str_AsiGen)
                 MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
                        "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
                 Call frm_Ctb_InvDpf_01.fs_BuscarComp
                 Unload Me
             End If
             Screen.MousePointer = 0
      End If
      Call frm_Ctb_InvDpf_01.fs_BuscarComp
   End If
End Sub

Private Sub fs_BusDependencia(ByVal p_NumCta As String, ByRef p_Estado As Boolean)
   p_Estado = False
   
   'extrae el numero de cuenta
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT COUNT(*) CONTEO  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_MAEDPF A  "
   g_str_Parame = g_str_Parame & "   WHERE A.MAEDPF_NUMCTA_REF = " & CStr(CLng(p_NumCta))
   g_str_Parame = g_str_Parame & "     AND A.MAEDPF_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "     AND A.MAEDPF_SITDPF = 1  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      p_Estado = False
      Screen.MousePointer = 0
      Exit Sub
   End If
                  
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      p_Estado = False
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   Else
      If g_rst_Princi!CONTEO > 0 Then
         p_Estado = True
      End If
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_BusCuentas(ByRef p_CtaDeb_1 As String, ByRef p_CtaHab_1 As String, ByRef p_CtaHab_2 As String, ByRef pEstado As Boolean)
   p_CtaDeb_1 = ""
   p_CtaHab_1 = ""
   p_CtaHab_2 = ""
   pEstado = False
   
   'extrae el numero de cuenta
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.CTADPF_CTADEB_01, A.CTADPF_CTADEB_02, A.CTADPF_CTAHAB_01, A.CTADPF_CTAHAB_02  "
   g_str_Parame = g_str_Parame & "    FROM CTB_CTADPF A  "
   g_str_Parame = g_str_Parame & "   WHERE A.CTADPF_CODENT_DES =  " & cmb_BancDestino.ItemData(cmb_BancDestino.ListIndex)
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_TIPDPF = " & cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex)
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_CODENT_ORI = " & cmb_BancOrigen.ItemData(cmb_BancOrigen.ListIndex)
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
               
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      pEstado = False
      Screen.MousePointer = 0
      Exit Sub
   End If
                  
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      pEstado = False
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   Else
      pEstado = True
      p_CtaDeb_1 = Trim(g_rst_Princi!CTADPF_CTADEB_01 & "")
      p_CtaHab_1 = Trim(g_rst_Princi!CTADPF_CTAHAB_01 & "")
      p_CtaHab_2 = Trim(g_rst_Princi!CTADPF_CTAHAB_02 & "")
   End If
End Sub

Private Sub fs_GeneraAsiento(ByVal p_NumCta As String, ByVal p_OperRef As String, p_FecDpf As String, p_TipDpf As Integer, ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_TipSbs        As Double
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_str_CtaDeb_1      As String
Dim r_str_CtaHab_1      As String
Dim r_str_CtaHab_2      As String
Dim r_bol_Estado        As Boolean
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   r_str_CtaDeb_1 = ""
   r_str_CtaHab_1 = ""
   r_str_CtaHab_2 = ""

   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(p_FecDpf, "yyyymmdd")
   r_str_FecPrPgoL = p_FecDpf
   
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(p_FecDpf, "yyyymmdd"), 1)
   
   r_str_Glosa = "OPERACION DPF " & Format(p_NumCta, "00000000")
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
   
   l_int_PerMes = Month(p_FecDpf)
   l_int_PerAno = Year(p_FecDpf)

   Call fs_BusCuentas(r_str_CtaDeb_1, r_str_CtaHab_1, r_str_CtaHab_2, r_bol_Estado)
   '---------------------------------------------------------
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle --- todo es en soles
   r_dbl_ImpSol = 0: r_dbl_ImpDol = 0
   If p_TipDpf = 1 Or p_TipDpf = 3 Or p_TipDpf = 5 Then 'Apertura, RNV_Con, Quiebre
      r_dbl_ImpSol = CDbl(ipp_SalCap.Text)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   ElseIf p_TipDpf = 2 Then 'Cierre deposito
      'r_dbl_ImpSol = CDbl(ipp_SalCap.Text) + CDbl(pnl_IntCap_Fin.Caption)
      r_dbl_ImpSol = CDbl(ipp_SalCap.Text) + CDbl(pnl_IntCap.Caption)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   ElseIf p_TipDpf = 4 Then 'Rnv_Sin
      'r_dbl_ImpSol = CDbl(l_str_IntCap) + CDbl(l_str_IntAju)
      r_dbl_ImpSol = CDbl(l_str_IntCap)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   End If
   If r_str_CtaDeb_1 <> "" And r_dbl_ImpSol > 0 Then
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 Then
         r_dbl_ImpDol = r_dbl_ImpSol
         r_dbl_ImpSol = Format(CDbl(r_dbl_ImpSol * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      End If
      Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                           r_int_NumAsi, 1, r_str_CtaDeb_1, CDate(r_str_FecPrPgoL), _
                                           r_str_Glosa, "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
   End If
   r_dbl_ImpSol = 0: r_dbl_ImpDol = 0
   If p_TipDpf = 1 Or p_TipDpf = 2 Or p_TipDpf = 5 Then 'Apertura, Cierre, Quiebre
      r_dbl_ImpSol = CDbl(ipp_SalCap.Text)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   ElseIf p_TipDpf = 3 Then 'RNV_Con
      'r_dbl_ImpSol = CDbl(l_str_SalCap)
      'r_dbl_ImpSol = Format(CDbl(ipp_SalCap.Text) - CDbl(l_str_IntCap), "###,###,##0.00")
      r_dbl_ImpSol = CDbl(l_str_SalCap)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   ElseIf p_TipDpf = 4 Then 'Rnv_Sin
      'r_dbl_ImpSol = CDbl(l_str_IntCap) + CDbl(l_str_IntAju)
      r_dbl_ImpSol = CDbl(l_str_IntCap)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   End If
   If r_str_CtaHab_1 <> "" And r_dbl_ImpSol > 0 Then
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 Then
         r_dbl_ImpDol = r_dbl_ImpSol
         r_dbl_ImpSol = Format(CDbl(r_dbl_ImpSol * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      End If
      Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                           r_int_NumAsi, 2, r_str_CtaHab_1, CDate(r_str_FecPrPgoL), _
                                           r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
   End If
   r_dbl_ImpSol = 0: r_dbl_ImpDol = 0
   If p_TipDpf = 2 Then 'Cierre
      'r_dbl_ImpSol = CDbl(pnl_IntCap.Caption) + CDbl(ipp_IntAjus.Text)
      r_dbl_ImpSol = CDbl(pnl_IntCap.Caption)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   ElseIf p_TipDpf = 3 Then 'RNV_Con
      'r_dbl_ImpSol = CDbl(l_str_IntCap) + CDbl(l_str_IntAju)
      'r_dbl_ImpSol = CDbl(l_str_IntCap)
      r_dbl_ImpSol = CDbl(l_str_IntCap) + CDbl(ipp_IntAjus.Text)
      r_dbl_ImpDol = Format(CDbl(r_dbl_ImpSol / r_dbl_TipSbs), "###,###,##0.00")
   End If
   If r_str_CtaHab_2 <> "" And r_dbl_ImpSol > 0 Then
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 Then
         r_dbl_ImpDol = r_dbl_ImpSol
         r_dbl_ImpSol = Format(CDbl(r_dbl_ImpSol * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      End If
      Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                           r_int_NumAsi, 3, r_str_CtaHab_2, CDate(r_str_FecPrPgoL), _
                                           r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
   End If
   
   p_AsiGen = r_str_AsiGen
   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_MAEDPF SET "
   If p_TipDpf = 1 Then 'Apertura
      g_str_Parame = g_str_Parame & "     MAEDPF_DATCNT_APE = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   ElseIf p_TipDpf = 2 Then 'Cierre
      g_str_Parame = g_str_Parame & "     MAEDPF_DATCNT_CIE = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   ElseIf p_TipDpf = 3 Then 'RNV_Con
      g_str_Parame = g_str_Parame & "     MAEDPF_DATCNT_RNV = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "', "
      g_str_Parame = g_str_Parame & "     MAEDPF_FECRNV = " & Format(ipp_FecOper.Text, "yyyymmdd")
   ElseIf p_TipDpf = 4 Then 'Rnv_Sin
      g_str_Parame = g_str_Parame & "     MAEDPF_DATCNT_RNV = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "', "
      g_str_Parame = g_str_Parame & "     MAEDPF_FECRNV = " & Format(ipp_FecOper.Text, "yyyymmdd")
   ElseIf p_TipDpf = 5 Then 'Quiebre
      g_str_Parame = g_str_Parame & "     MAEDPF_DATCNT_QUI = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   End If
   g_str_Parame = g_str_Parame & " WHERE MAEDPF_NUMCTA = " & CLng(p_NumCta)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Private Sub fs_Grabar()
Dim r_str_AsiGen   As String
   
   r_str_AsiGen = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEDPF ( "
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = g_str_Parame & "'" & "" & "', " 'MAEDPF_NUMCTA
   Else
      g_str_Parame = g_str_Parame & "'" & CLng(pnl_NumCta.Caption) & "', " 'MAEDPF_NUMCTA
   End If
   g_str_Parame = g_str_Parame & Format(ipp_FecOper.Text, "yyyymmdd") & ", " 'MAEDPF_FECAPE
   g_str_Parame = g_str_Parame & cmb_TipAccion.ItemData(cmb_TipAccion.ListIndex) & ", " 'MAEDPF_TIPDPF
   g_str_Parame = g_str_Parame & cmb_BancDestino.ItemData(cmb_BancDestino.ListIndex) & ", " 'MAEDPF_CODBNC_APE
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "  'MAEDPF_CODMOD
   g_str_Parame = g_str_Parame & CDbl(ipp_SalCap.Text) & ", " 'MAEDPF_SALCAP
   g_str_Parame = g_str_Parame & CDbl(pnl_IntCap.Caption) & ", "  'MAEDPF_INTCAP
   g_str_Parame = g_str_Parame & CDbl(ipp_IntAjus.Text) & ", "  'MAEDPF_INTAJU
   g_str_Parame = g_str_Parame & CDbl(ipp_TasInt.Text) & ", "   'MAEDPF_TASINT
   g_str_Parame = g_str_Parame & CLng(ipp_PlaAno) & ", "   'MAEDPF_PLADIA
   g_str_Parame = g_str_Parame & cmb_BancOrigen.ItemData(cmb_BancOrigen.ListIndex) & ", "   'MAEDPF_CODBNC_ORI
   g_str_Parame = g_str_Parame & "'" & Trim(txt_OperRef_Dat.Text) & "', "   'MAEDPF_NUMREF
   g_str_Parame = g_str_Parame & IIf(Trim(pnl_NumCta_Ref.Caption) = "", "Null", Trim(pnl_NumCta_Ref.Caption)) & ", " 'MAEDPF_NUMCTA_REF
   g_str_Parame = g_str_Parame & "1, "   'MAEDPF_SITUAC
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
       Call fs_GeneraAsiento(g_rst_Genera!CODIGO, txt_OperRef_Dat.Text, ipp_FecOper.Text, 1, r_str_AsiGen)
       MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
              "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_InvDpf_01.fs_BuscarComp
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_InvDpf_01.fs_BuscarComp
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub fs_Cargar_Datos()
Dim r_int_Contad As Integer
   
   l_str_OpeRef = ""
   l_str_SalCap = "0.00 "
   l_str_IntCap = "0.00 "
   l_str_IntAju = "0.00 "
   l_str_FecApe = ""
   l_str_TasInt = "0.00 "
   l_str_PlaDia = "1"
   l_int_EntInt = -1
   l_int_EntFon = -1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT LPAD(A.MAEDPF_NUMCTA,8,'0') AS MAEDPF_NUMCTA, A.MAEDPF_FECAPE, A.MAEDPF_TIPDPF, A.MAEDPF_CODENT_DES,  "
   g_str_Parame = g_str_Parame & "         A.MAEDPF_CODMON, A.MAEDPF_SALCAP, A.MAEDPF_INTCAP, A.MAEDPF_TASINT, A.MAEDPF_PLADIA, A.MAEDPF_CODENT_ORI,  "
   g_str_Parame = g_str_Parame & "         A.MAEDPF_NUMREF, LPAD(A.MAEDPF_NUMCTA_REF,8,'0') AS MAEDPF_NUMCTA_REF, A.MAEDPF_FECCIE, " 'A.MAEDPF_DATCNT,  "
   g_str_Parame = g_str_Parame & "         NVL(B.MAEDPF_SALCAP,0) AS SALCAP_REF, NVL(B.MAEDPF_INTCAP,0) AS INTCAP_REF, TRIM(B.MAEDPF_NUMREF) AS NUMREF_REF, "
   g_str_Parame = g_str_Parame & "         DECODE(A.MAEDPF_TIPDPF,1,  "
   g_str_Parame = g_str_Parame & "                 CASE  "
   g_str_Parame = g_str_Parame & "                  WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VENCIDO'  "
   g_str_Parame = g_str_Parame & "                  WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VIGENTE'  "
   g_str_Parame = g_str_Parame & "         END, 'CERRADO') AS NOM_SITUAC,  "
   g_str_Parame = g_str_Parame & "         B.MAEDPF_TASINT AS TASINT_REF, A.MAEDPF_FECQUI, A.MAEDPF_FECCIE, A.MAEDPF_INTAJU  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_MAEDPF A  "
   'g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_MAEDPF B ON  A.MAEDPF_NUMCTA = B.MAEDPF_NUMCTA_REF  "
   g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_MAEDPF B ON  B.MAEDPF_NUMCTA = A.MAEDPF_NUMCTA_REF  "
   g_str_Parame = g_str_Parame & "   WHERE A.MAEDPF_NUMCTA = " & CInt(moddat_g_str_Codigo) & "  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_NumCta.Caption = CStr(g_rst_Princi!MAEDPF_NUMCTA)
      If moddat_g_int_FlgGrb = 0 Or moddat_g_int_FlgGrb = 2 Then 'Consultar, Editar
         Call gs_BuscarCombo_Item(cmb_TipAccion, g_rst_Princi!MAEDPF_TIPDPF)
      End If
      'pnl_Estado.Caption = Trim(moddat_g_str_Situac)
      pnl_Estado.Caption = Trim(g_rst_Princi!NOM_SITUAC)
      
      ipp_FecOper.Text = gf_FormatoFecha(g_rst_Princi!MAEDPF_FECAPE)
      Call gs_BuscarCombo_Item(cmb_BancDestino, g_rst_Princi!MAEDPF_CODENT_DES)
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!MAEDPF_CODMON)
      
      ipp_SalCap.Text = Format(CStr(g_rst_Princi!MAEDPF_SALCAP), "###,###,##0.00")
      ipp_TasInt.Text = Format(CStr(g_rst_Princi!MAEDPF_TASINT), "###,###,##0.00")
      ipp_PlaAno.Text = CStr(g_rst_Princi!MAEDPF_PLADIA)
      
      pnl_IntCap.Caption = Format(CStr(g_rst_Princi!MAEDPF_INTCAP), "###,###,##0.00") & " "
      ipp_IntAjus.Text = Format(CStr(g_rst_Princi!MAEDPF_INTAJU), "###,###,##0.00")
      'pnl_IntCap_Fin.Caption = Format(g_rst_Princi!MAEDPF_INTCAP + g_rst_Princi!MAEDPF_INTAJU, "###,###,##0.00") & " "
      
      Call gs_BuscarCombo_Item(cmb_BancOrigen, g_rst_Princi!MAEDPF_CODENT_ORI)
      txt_OperRef_Dat.Text = CStr(g_rst_Princi!MAEDPF_NUMREF)
      '-------------
      If Trim(g_rst_Princi!MAEDPF_FECCIE & "") <> "" Then
         ipp_FecCierre.Text = gf_FormatoFecha(g_rst_Princi!MAEDPF_FECCIE)
      End If
      If Trim(g_rst_Princi!MAEDPF_FECQUI & "") <> "" Then
         ipp_FecQuie.Text = gf_FormatoFecha(g_rst_Princi!MAEDPF_FECQUI)
      End If
      
      pnl_NumCta_Ref.Caption = Trim(g_rst_Princi!MAEDPF_NUMCTA_REF & "")
      'pnl_SalCap_Ref.Caption = Format(g_rst_Princi!SALCAP_REF, "###,###,##0.00") & " "
      'pnl_IntCap_Ref.Caption = Format(g_rst_Princi!INTCAP_REF, "###,###,##0.00") & " "
      'pnl_TasInt_Ref.Caption = Format(g_rst_Princi!TASINT_REF, "###,###,##0.00") & " "
      
      l_str_OpeRef = CStr(g_rst_Princi!MAEDPF_NUMREF)
      l_str_SalCap = CDbl(ipp_SalCap.Text)
      l_str_IntCap = CDbl(pnl_IntCap.Caption)
      l_str_IntAju = CDbl(ipp_IntAjus.Text)
      l_str_FecApe = ipp_FecOper.Text
      l_str_TasInt = CDbl(ipp_TasInt.Text)
      l_str_PlaDia = ipp_PlaAno.Text
         
      l_int_EntInt = g_rst_Princi!MAEDPF_CODENT_DES
      l_int_EntFon = g_rst_Princi!MAEDPF_CODENT_ORI
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub fs_Desabilitar(p_Estado As Boolean)
   cmb_TipAccion.Enabled = p_Estado
   ipp_FecOper.Enabled = p_Estado
   cmb_BancDestino.Enabled = p_Estado
   cmb_Moneda.Enabled = p_Estado
   ipp_SalCap.Enabled = p_Estado
   ipp_IntAjus.Enabled = p_Estado
   ipp_PlaAno.Enabled = p_Estado
   cmb_BancOrigen.Enabled = p_Estado
   txt_OperRef_Dat.Enabled = p_Estado
   ipp_TasInt.Enabled = p_Estado
   ipp_FecCierre.Enabled = p_Estado
   ipp_FecQuie.Enabled = p_Estado
   'cmd_Conult.Enabled = True
   ipp_FecQuie.AllowNull = True
   ipp_FecQuie.Text = ""
   ipp_FecCierre.AllowNull = True
   ipp_FecCierre.Text = ""
   
   If moddat_g_int_FlgGrb = 1 Then 'Apertura
      ipp_FecCierre.Enabled = False
      ipp_FecQuie.Enabled = False
      ipp_IntAjus.Enabled = False
      cmb_TipAccion.Enabled = False
      'cmb_Moneda.Enabled = False
   ElseIf moddat_g_int_FlgGrb = 2 Then 'Modificar
      txt_OperRef_Dat.Enabled = True
      cmb_Moneda.Enabled = False
   ElseIf moddat_g_int_FlgGrb = 3 Then 'Gestionar
      cmb_TipAccion.Enabled = True
      txt_OperRef_Dat.Enabled = True
      ipp_IntAjus.Enabled = False
      cmb_Moneda.Enabled = False
      If moddat_g_int_Situac = 1 Then 'VIGENTE - CIERRE
         cmb_TipAccion.Enabled = False
         ipp_FecQuie.Enabled = True
         ipp_FecQuie.AllowNull = False
         ipp_FecQuie.Text = moddat_g_str_FecSis
      End If
   End If
End Sub

Private Sub cmb_TipAccion_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_FecOper.Enabled = True Then
          Call gs_SetFocus(ipp_FecOper)
       ElseIf ipp_IntAjus.Enabled = True Then
          Call gs_SetFocus(ipp_IntAjus)
       ElseIf cmb_BancOrigen.Enabled = True Then
          Call gs_SetFocus(cmb_BancOrigen)
       ElseIf txt_OperRef_Dat.Enabled = True Then
          Call gs_SetFocus(txt_OperRef_Dat)
       End If
   End If
End Sub

Private Sub ipp_FecOper_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If cmb_BancDestino.Enabled = True Then
          Call gs_SetFocus(cmb_BancDestino)
       ElseIf cmb_Moneda.Enabled = True Then
          Call gs_SetFocus(cmb_Moneda)
       ElseIf ipp_SalCap.Enabled = True Then
          Call gs_SetFocus(ipp_SalCap)
       ElseIf ipp_IntAjus.Enabled = True Then
          Call gs_SetFocus(ipp_IntAjus)
       ElseIf ipp_PlaAno.Enabled = True Then
          Call gs_SetFocus(ipp_PlaAno)
       ElseIf ipp_TasInt.Enabled = True Then
          Call gs_SetFocus(ipp_TasInt)
       End If
   End If
End Sub

Private Sub cmb_BancDestino_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If cmb_Moneda.Enabled = False Then
          Call gs_SetFocus(ipp_SalCap)
       Else
          Call gs_SetFocus(cmb_Moneda)
       End If
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_SalCap)
   End If
End Sub

Private Sub ipp_IntCap_LostFocus()
   Call fs_CalInteres
End Sub

Private Sub ipp_IntAjus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_PlaAno.Enabled = True Then
         Call gs_SetFocus(ipp_PlaAno)
      ElseIf cmb_BancOrigen.Enabled = True Then
         Call gs_SetFocus(cmb_BancOrigen)
      ElseIf txt_OperRef_Dat.Enabled = True Then
         Call gs_SetFocus(txt_OperRef_Dat)
      End If
   End If
End Sub

Private Sub ipp_IntAjus_LostFocus()
   Call fs_CalInteres
End Sub

Private Sub ipp_PlaAno_LostFocus()
   Call fs_CalInteres
End Sub

Private Sub ipp_SalCap_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_IntAjus.Enabled = False Then
          Call gs_SetFocus(ipp_PlaAno)
       Else
         Call gs_SetFocus(ipp_IntAjus)
       End If
   End If
End Sub

Private Sub ipp_SalCap_LostFocus()
   Call fs_CalInteres
End Sub

Private Sub ipp_TasInt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_BancOrigen.Enabled = False Then
         Call gs_SetFocus(txt_OperRef_Dat)
      Else
         Call gs_SetFocus(cmb_BancOrigen)
      End If
   End If
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_TasInt)
   End If
End Sub

Private Sub cmb_BancOrigen_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_OperRef_Dat)
   End If
End Sub

Private Sub ipp_TasInt_LostFocus()
   Call fs_CalInteres
End Sub

Private Sub txt_OperRef_Dat_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_FecCierre.Enabled = False And ipp_FecQuie.Enabled = False Then
          Call gs_SetFocus(cmd_Grabar)
       ElseIf ipp_FecCierre.Enabled = True And ipp_FecQuie.Enabled = False Then
           Call gs_SetFocus(ipp_FecCierre)
       ElseIf ipp_FecCierre.Enabled = False And ipp_FecQuie.Enabled = True Then
          Call gs_SetFocus(ipp_FecQuie)
       End If
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub ipp_FecCierre_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_FecQuie.Enabled = False Then
          Call gs_SetFocus(cmd_Grabar)
       Else
          Call gs_SetFocus(ipp_FecQuie)
       End If
   End If
End Sub

Private Sub ipp_FecQuie_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub


