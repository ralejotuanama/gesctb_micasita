VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_AsiCtb_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16935
   Icon            =   "GesCtb_frm_926.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   16935
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8715
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   16935
      _Version        =   65536
      _ExtentX        =   29871
      _ExtentY        =   15372
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
         Height          =   5445
         Left            =   30
         TabIndex        =   17
         Top             =   3090
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
         _ExtentY        =   9604
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
         Begin EditLib.fpDateTime ipp_FecCtb 
            Height          =   315
            Left            =   6030
            TabIndex        =   10
            Top             =   4830
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
         Begin VB.ComboBox cmb_TipMov 
            Height          =   315
            ItemData        =   "GesCtb_frm_926.frx":000C
            Left            =   2430
            List            =   "GesCtb_frm_926.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   4830
            Width           =   2115
         End
         Begin VB.TextBox txt_CtaCtb 
            Height          =   315
            Left            =   450
            TabIndex        =   7
            Text            =   "txt_CtaCtb"
            Top             =   4830
            Width           =   1905
         End
         Begin EditLib.fpDoubleSingle ipp_MtoCta 
            Height          =   315
            Left            =   4620
            TabIndex        =   9
            Top             =   4830
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            MaxValue        =   "9999999"
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
         Begin Threed.SSPanel pnl_Tit_Codigo 
            Height          =   285
            Left            =   420
            TabIndex        =   18
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cta. Contable"
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
            Left            =   1620
            TabIndex        =   19
            Top             =   60
            Width           =   3270
            _Version        =   65536
            _ExtentX        =   5768
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
            Left            =   4890
            TabIndex        =   20
            Top             =   60
            Width           =   4260
            _Version        =   65536
            _ExtentX        =   7514
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
            Left            =   10440
            TabIndex        =   21
            Top             =   60
            Width           =   510
            _Version        =   65536
            _ExtentX        =   900
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
            Left            =   10950
            TabIndex        =   22
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   12150
            TabIndex        =   23
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   13350
            TabIndex        =   24
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   14550
            TabIndex        =   25
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   10950
            TabIndex        =   26
            Top             =   4740
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   12135
            TabIndex        =   27
            Top             =   4740
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   13335
            TabIndex        =   28
            Top             =   4740
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   14535
            TabIndex        =   29
            Top             =   4740
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   10950
            TabIndex        =   30
            Top             =   5040
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   12135
            TabIndex        =   31
            Top             =   5040
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   13335
            TabIndex        =   32
            Top             =   5040
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   14535
            TabIndex        =   33
            Top             =   5040
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Height          =   4275
            Left            =   30
            TabIndex        =   6
            Top             =   360
            Width           =   16785
            _ExtentX        =   29607
            _ExtentY        =   7541
            _Version        =   393216
            Rows            =   6
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Itm"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   9150
            TabIndex        =   54
            Top             =   60
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Contable"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   15750
            TabIndex        =   57
            Top             =   60
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "C. Costo"
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
         Begin VB.Label Label8 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   9465
            TabIndex        =   36
            Top             =   4740
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Diferencia ==>"
            Height          =   285
            Left            =   9465
            TabIndex        =   35
            Top             =   5040
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   37
         Top             =   60
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
            TabIndex        =   38
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
            Picture         =   "GesCtb_frm_926.frx":0010
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   39
         Top             =   780
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
            Left            =   16245
            Picture         =   "GesCtb_frm_926.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ComGra 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_926.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Grabar Comprobante Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetBor 
            Height          =   585
            Left            =   2190
            Picture         =   "GesCtb_frm_926.frx":0B9E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Borrar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetMod 
            Height          =   585
            Left            =   1590
            Picture         =   "GesCtb_frm_926.frx":0EA8
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Editar registro de Detalle Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DetNue 
            Height          =   585
            Left            =   960
            Picture         =   "GesCtb_frm_926.frx":11B2
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Agregar registro de Detalle Contable"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   40
         Top             =   1470
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
         Begin VB.ComboBox cmb_LibCon 
            Height          =   315
            Left            =   8010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecCom 
            Height          =   315
            Left            =   14370
            TabIndex        =   1
            Top             =   390
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
         Begin Threed.SSPanel pnl_Empres 
            Height          =   315
            Left            =   1530
            TabIndex        =   41
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
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
         Begin Threed.SSPanel pnl_Sucurs 
            Height          =   315
            Left            =   1530
            TabIndex        =   42
            Top             =   390
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   8010
            TabIndex        =   43
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
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
         Begin Threed.SSPanel pnl_FecReg 
            Height          =   315
            Left            =   14370
            TabIndex        =   55
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Registro:"
            Height          =   285
            Left            =   12810
            TabIndex        =   56
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label33 
            Caption         =   "Período:"
            Height          =   255
            Left            =   6360
            TabIndex        =   48
            Top             =   90
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   47
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   60
            TabIndex        =   46
            Top             =   420
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Libro Contable:"
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   45
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Comprob.:"
            Height          =   285
            Left            =   12810
            TabIndex        =   44
            Top             =   420
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   49
         Top             =   2280
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
         Begin VB.ComboBox cmb_TipOpe 
            Height          =   315
            ItemData        =   "GesCtb_frm_926.frx":14BC
            Left            =   1530
            List            =   "GesCtb_frm_926.frx":14BE
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   3465
         End
         Begin VB.ComboBox cmb_MonCtb 
            Height          =   315
            Left            =   8010
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   3465
         End
         Begin VB.TextBox txt_GloCab 
            Height          =   315
            Left            =   1530
            MaxLength       =   250
            TabIndex        =   5
            Text            =   "txt_GloCab"
            Top             =   390
            Width           =   15165
         End
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   14370
            TabIndex        =   4
            Top             =   60
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
            Text            =   "0.000000"
            DecimalPlaces   =   6
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
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   6360
            TabIndex        =   53
            Top             =   90
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Cambio:"
            Height          =   255
            Left            =   12810
            TabIndex        =   52
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Glosa Cabecera:"
            Height          =   285
            Left            =   60
            TabIndex        =   51
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Operación:"
            Height          =   225
            Left            =   60
            TabIndex        =   50
            Top             =   90
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_MonCtb()   As moddat_tpo_Genera
Dim l_arr_LogPro()   As modprc_g_tpo_LogPro
Dim l_str_Origen     As String
Dim l_var_ColAnt     As Variant
Dim l_int_GrbDet     As Integer
Dim l_bol_CtaCtb     As Boolean

Private Sub cmd_DetNue_Click()
Dim r_str_ConGlo     As String
Dim r_int_Contad     As Integer

   If CDbl(ipp_TipCam.Value) = 0# Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   
   If fs_ValidaPeriodo(ipp_FecCom.Text) = False Then
      Exit Sub
   End If
   
   If grd_DetAsi.Row >= 0 And fs_ValidarCeldaVacias = True Then
      MsgBox "No deben existir Celdas Vacías.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_DetAsi)
      Exit Sub
   End If
   
   r_int_Contad = fs_ValidarCeldaCero
   If grd_DetAsi.Row >= 0 And (r_int_Contad = 4) Then
      MsgBox "Se debe ingresar Monto para la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoCta)
      Exit Sub
   End If
   
   l_int_GrbDet = 1
   Call fs_ActivaDet(False)
   Call fs_ActivaObj(True)
   
   If l_int_GrbDet = 1 Or l_int_GrbDet = 2 Then
      grd_DetAsi.Rows = grd_DetAsi.Rows + 1
      grd_DetAsi.Row = grd_DetAsi.Rows - 1
   End If
   
   grd_DetAsi.Col = 0:  grd_DetAsi.Text = grd_DetAsi.Row + 1
   grd_DetAsi.Col = 1:  grd_DetAsi.Text = ""
   grd_DetAsi.Col = 2:  grd_DetAsi.Text = ""
   If grd_DetAsi.Row = 0 Then
      grd_DetAsi.Col = 3:  grd_DetAsi.Text = Trim(txt_GloCab.Text)
      grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
   Else
      grd_DetAsi.Col = 3:  grd_DetAsi.Text = grd_DetAsi.TextMatrix(grd_DetAsi.Row - 1, 3)
      grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(CDate(grd_DetAsi.TextMatrix(grd_DetAsi.Row - 1, 4)), "dd/mm/yyyy")
   End If
   grd_DetAsi.Col = 5:  grd_DetAsi.Text = ""
   grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"
   grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"
   
   Call fs_LimpiaDet
   Call fs_TotDebHab
   
   grd_DetAsi.Col = 1
   txt_CtaCtb.Visible = False
   cmb_TipMov.Visible = False
   ipp_MtoCta.Visible = False
   ipp_FecCtb.Visible = False
   
   Call fs_IniciaEdicion
End Sub

Private Sub cmd_DetMod_Click()
   If grd_DetAsi.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_DetAsi)
   Call fs_ActivaDet(False)
   Call gs_SetFocus(grd_DetAsi)
   l_int_GrbDet = 2
End Sub

Private Sub cmd_DetBor_Click()
   If grd_DetAsi.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de eliminar el registro seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      If grd_DetAsi.Rows = 1 Then
         Call gs_LimpiaGrid(grd_DetAsi)
         txt_CtaCtb.Visible = False
         ipp_MtoCta.Visible = False
         ipp_FecCtb.Visible = False
      Else
         grd_DetAsi.RemoveItem grd_DetAsi.Row
         txt_CtaCtb.Text = ""
         txt_CtaCtb.Visible = False
         ipp_MtoCta.Visible = False
         ipp_FecCtb.Visible = False
         grd_DetAsi.Row = 0
      End If
      Call fs_TotDebHab
   End If
End Sub

Private Sub cmd_ComGra_Click()
Dim r_int_CuaAsi        As Integer
Dim r_lng_NumAsi        As Long
Dim r_int_Contad        As Integer
Dim r_str_CtaCtb        As String
Dim r_str_GloDet        As String
Dim r_str_FecCtb        As String
Dim r_str_DebHab        As String
Dim r_dbl_MtoCta_DMN    As Double
Dim r_dbl_MtoCta_HMN    As Double
Dim r_dbl_MtoCta_DME    As Double
Dim r_dbl_MtoCta_HME    As Double

   If cmb_TipOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOpe)
      Exit Sub
   End If
   If cmb_MonCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda del Asiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonCtb)
      Exit Sub
   End If
   If CDbl(ipp_TipCam.Value) = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   If Len(Trim(txt_GloCab.Text)) = 0 Then
      MsgBox "Debe ingresar la Glosa de Cabecera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GloCab)
      Exit Sub
   End If
   If grd_DetAsi.Rows = 0 Then
      MsgBox "Debe ingresar Detalle de Asiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DetNue)
      Exit Sub
   End If
   If fs_ValidarCeldaVacias = True Then
      MsgBox "No deben existir Celdas Vacías.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_DetAsi)
      Exit Sub
   End If
      
   r_int_CuaAsi = 1
   If CDbl(pnl_DifDeb_MN.Caption) > 0 Or CDbl(pnl_DifDeb_ME.Caption) > 0 Or CDbl(pnl_DifHab_MN.Caption) > 0 Or CDbl(pnl_DifHab_ME.Caption) > 0 Then
      If MsgBox("El Asiento se encuentra descuadrado. ¿Desea continuar sin cuadrar el asiento?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Call gs_SetFocus(cmd_DetNue)
         Exit Sub
      End If
      
      r_int_CuaAsi = 2
   End If

   If fs_ValidaPeriodo(ipp_FecCom.Text) = False Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   ReDim l_arr_LogPro(0)
   ReDim l_arr_LogPro(1)
   
   l_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   l_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   l_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   l_arr_LogPro(1).LogPro_NumErr = 0
   l_str_Origen = "LM"
      
   modctb_int_PerMes = Format(ipp_FecCom.Text, "mm")
   modctb_int_PerAno = Format(ipp_FecCom.Text, "yyyy")
   
   'Obtener Número de Asiento
   If moddat_g_int_FlgGrb = 1 Then
       r_lng_NumAsi = modprc_ff_NumAsi(l_arr_LogPro, modctb_int_PerAno, modctb_int_PerMes, l_str_Origen, cmb_LibCon.ItemData(cmb_LibCon.ListIndex))
   Else
      r_lng_NumAsi = modctb_lng_NumAsi
   End If
   
   'Grabando Cabecera de Asiento
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   'Datos Principales
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INGRESO_CNTBL_ASIENTO_1 ("
      g_str_Parame = g_str_Parame & "'" & l_str_Origen & "', "
      g_str_Parame = g_str_Parame & modctb_int_PerAno & ", "
      g_str_Parame = g_str_Parame & modctb_int_PerMes & ", "
      g_str_Parame = g_str_Parame & CInt(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & r_lng_NumAsi & ", "
      g_str_Parame = g_str_Parame & "'" & Right("00" & CStr(CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo)), 3) & "', "
      g_str_Parame = g_str_Parame & CDbl(ipp_TipCam.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(Trim(Mid(cmb_TipOpe.Text, 1, InStr(cmb_TipOpe.Text, "-") - 1))) & "', "
      g_str_Parame = g_str_Parame & "'" & txt_GloCab.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & "', "
      g_str_Parame = g_str_Parame & "'" & LCase(Mid(modgen_g_str_CodUsu, 1, 5)) & "', "
      g_str_Parame = g_str_Parame & "'1', "
      If moddat_g_int_FlgGrb = 1 Then
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
      Else
         g_str_Parame = g_str_Parame & CDbl(pnl_TotDeb_MN.Caption) & ", "
         g_str_Parame = g_str_Parame & CDbl(pnl_TotHab_MN.Caption) & ", "
         g_str_Parame = g_str_Parame & CDbl(pnl_TotDeb_ME.Caption) & ", "
         g_str_Parame = g_str_Parame & CDbl(pnl_TotHab_ME.Caption) & ", "

      End If
      g_str_Parame = g_str_Parame & CInt(moddat_g_int_FlgGrb) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INGRESO_CNTBL_ASIENTO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Eliminar Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " DELETE FROM CNTBL_ASIENTO_DET "
   g_str_Parame = g_str_Parame & "  WHERE ANO = " & CStr(modctb_int_PerAno)
   g_str_Parame = g_str_Parame & "    AND MES = " & CStr(modctb_int_PerMes)
   g_str_Parame = g_str_Parame & "    AND NRO_LIBRO = " & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex))
   g_str_Parame = g_str_Parame & "    AND NRO_ASIENTO = " & CStr(r_lng_NumAsi)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " UPDATE CNTBL_ASIENTO A "
   g_str_Parame = g_str_Parame & "    SET TOT_SOLDEB = 0, "
   g_str_Parame = g_str_Parame & "        TOT_SOLHAB = 0, "
   g_str_Parame = g_str_Parame & "        TOT_DOLDEB = 0, "
   g_str_Parame = g_str_Parame & "        TOT_DOLHAB = 0 "
   g_str_Parame = g_str_Parame & "  WHERE ANO = " & CStr(modctb_int_PerAno)
   g_str_Parame = g_str_Parame & "    AND MES = " & CStr(modctb_int_PerMes)
   g_str_Parame = g_str_Parame & "    AND NRO_LIBRO = " & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex))
   g_str_Parame = g_str_Parame & "    AND NRO_ASIENTO = " & CStr(r_lng_NumAsi)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Ingresando detalle del asiento
   grd_DetAsi.Redraw = False
   For r_int_Contad = 0 To grd_DetAsi.Rows - 1
      grd_DetAsi.Row = r_int_Contad
      
      grd_DetAsi.Col = 1:  r_str_CtaCtb = grd_DetAsi.Text
      grd_DetAsi.Col = 3:  r_str_GloDet = grd_DetAsi.Text
      grd_DetAsi.Col = 4:  r_str_FecCtb = CDate(grd_DetAsi.Text)
      grd_DetAsi.Col = 5:  r_str_DebHab = Left(grd_DetAsi.Text, 1)
      grd_DetAsi.Col = 6:  r_dbl_MtoCta_DMN = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 7:  r_dbl_MtoCta_HMN = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 8:  r_dbl_MtoCta_DME = CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 9:  r_dbl_MtoCta_HME = CDbl(grd_DetAsi.Text)
      
      'Grabando en BD
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_INGRESO_CNTBL_ASI_DET_1 ("
         g_str_Parame = g_str_Parame & "'" & l_str_Origen & "', "
         g_str_Parame = g_str_Parame & CStr(modctb_int_PerAno) & ", "
         g_str_Parame = g_str_Parame & CStr(modctb_int_PerMes) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(r_lng_NumAsi) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_Contad + 1) & ", "

         'Datos de Linea
         g_str_Parame = g_str_Parame & "'" & r_str_CtaCtb & "',"
         g_str_Parame = g_str_Parame & "'" & CDate(r_str_FecCtb) & "', " 'ipp_FecCom.Text
         g_str_Parame = g_str_Parame & "'" & r_str_GloDet & "',"
         g_str_Parame = g_str_Parame & "'" & CStr(r_str_DebHab) & "',"

         If r_str_DebHab = "D" Then
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_DMN) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_DME) & ","
         Else
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_HMN) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCta_HME) & ","
         End If
         'g_str_Parame = g_str_Parame & CInt(moddat_g_int_FlgGrb) & ") "
         g_str_Parame = g_str_Parame & CInt(1) & ") "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If

         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_CTB_ASIDET. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop

   Next r_int_Contad
   
   grd_DetAsi.Redraw = True
   Call gs_RefrescaGrid(grd_DetAsi)
   moddat_g_int_FlgAct = 2
   
   If moddat_g_int_FlgGrb = 1 Then
      MsgBox "El Número de Asiento generado es el: " & CStr(r_lng_NumAsi), vbInformation, modgen_g_str_NomPlt
      
      If MsgBox("¿Desea seguir ingresando Asientos Contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Unload Me
      End If
      
      Call fs_LimpiaCab
      Call fs_LimpiaDet
      Call fs_ActivaDet(True)
      Call gs_SetFocus(cmb_LibCon)
   Else
      MsgBox "El Asiento " & CStr(r_lng_NumAsi) & " se actualizó con éxito", vbInformation, modgen_g_con_PltPar
      Call gs_RefrescaGrid(grd_DetAsi)
      Call fs_ActivaDet(True)
      Call gs_SetFocus(grd_DetAsi)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   l_str_Origen = "LM"
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_LimpiaCab
   Call fs_LimpiaDet
   Call fs_ActivaDet(True)
   
   If moddat_g_int_FlgGrb <> 1 Then
      cmb_LibCon.Enabled = False
      Call fs_LeerCabecera
      Call fs_LeerDetalle
      If grd_DetAsi.Rows > 0 Then
         Call gs_UbiIniGrid(grd_DetAsi)
      End If
      Call fs_TotDebHab
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   l_var_ColAnt = txt_GloCab.BackColor
   Call moddat_gs_Carga_LibCon(cmb_LibCon)
   Call moddat_gs_Carga_NotCtb(cmb_TipOpe)
   
   grd_DetAsi.Font.Size = 7
   grd_DetAsi.ColWidth(0) = 360
   grd_DetAsi.ColWidth(1) = 1200
   grd_DetAsi.ColWidth(2) = 3270
   grd_DetAsi.ColWidth(3) = 4260
   grd_DetAsi.ColWidth(4) = 1290
   grd_DetAsi.ColWidth(5) = 510
   grd_DetAsi.ColWidth(6) = 1200
   grd_DetAsi.ColWidth(7) = 1200
   grd_DetAsi.ColWidth(8) = 1200
   grd_DetAsi.ColWidth(9) = 1200
   grd_DetAsi.ColAlignment(0) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(1) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(2) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(3) = flexAlignLeftCenter
   grd_DetAsi.ColAlignment(4) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(5) = flexAlignCenterCenter
   grd_DetAsi.ColAlignment(6) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(7) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(8) = flexAlignRightCenter
   grd_DetAsi.ColAlignment(9) = flexAlignRightCenter
   With grd_DetAsi
      .FixedRows = 1
      .FixedCols = 1
   End With
   
   cmb_TipMov.Clear
   cmb_TipMov.AddItem "DEBE"
   cmb_TipMov.AddItem "HABER"
   cmb_TipMov.ListIndex = -1

   cmb_LibCon.ListIndex = -1
   cmb_TipOpe.ListIndex = -1
   cmb_MonCtb.ListIndex = -1
    
   'Tamaño de la celda de grd_DetAsi
   grd_DetAsi.RowHeightMin = txt_CtaCtb.Height
   
   'la fuente utilizada para mostrar texto
   txt_CtaCtb.FontName = grd_DetAsi.FontName
   cmb_TipMov.FontName = grd_DetAsi.FontName
   ipp_MtoCta.FontName = grd_DetAsi.FontName
   ipp_FecCtb.FontName = grd_DetAsi.FontName
   
   'el tamaño de la fuente que se va a utilizar
   txt_CtaCtb.FontSize = grd_DetAsi.FontSize
   cmb_TipMov.FontSize = grd_DetAsi.FontSize
   ipp_MtoCta.FontSize = grd_DetAsi.FontSize
   ipp_FecCtb.FontSize = grd_DetAsi.FontSize
   
   txt_CtaCtb.Visible = False
   cmb_TipMov.Visible = False
   ipp_MtoCta.Visible = False
   ipp_FecCtb.Visible = False
   
   'sin borde
   txt_CtaCtb.BorderStyle = vbBSNone
End Sub

Private Sub fs_LimpiaCab()
   Call moddat_gs_FecSis
   
   pnl_Empres.Caption = modctb_str_NomEmp
   pnl_Period.Caption = moddat_gf_Consulta_ParDes("033", CStr(modctb_int_PerMes)) & " " & Format(modctb_int_PerAno, "0000")
   pnl_Sucurs.Caption = modctb_str_NomSuc

   Call moddat_gs_Carga_ParEmp(modctb_str_CodEmp, "102", cmb_MonCtb, l_arr_MonCtb)
   ipp_FecCom.Text = moddat_g_str_FecSis
    
   ipp_FecCom.DateMin = Format(CDate(modctb_str_FecIni), "yyyymmdd")
   ipp_FecCom.DateMax = Format(CDate(modctb_str_FecFin), "yyyymmdd")
   
   ipp_FecCtb.DateMin = Format(CDate(modctb_str_FecIni), "yyyymmdd")
   ipp_FecCtb.DateMax = Format(CDate(modctb_str_FecFin), "yyyymmdd")
   
   Call gs_BuscarCombo_Item(cmb_LibCon, modctb_int_CodLib)
   
   txt_GloCab.Text = ""
   cmb_MonCtb.ListIndex = gf_Busca_Arregl(l_arr_MonCtb, Format("1", "000000")) - 1
   
   ipp_TipCam.Value = 0
   ipp_FecCtb.Value = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
   
   pnl_TotDeb_MN.Caption = "0.00 "
   pnl_TotHab_MN.Caption = "0.00 "
   pnl_DifDeb_MN.Caption = "0.00 "
   pnl_DifHab_MN.Caption = "0.00 "
   pnl_TotDeb_ME.Caption = "0.00 "
   pnl_TotHab_ME.Caption = "0.00 "
   pnl_DifDeb_ME.Caption = "0.00 "
   pnl_DifHab_ME.Caption = "0.00 "
   
   ipp_FecCtb.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
   
   Call gs_LimpiaGrid(grd_DetAsi)
End Sub

Private Sub fs_LimpiaDet()
   cmb_TipMov.ListIndex = -1
   txt_CtaCtb.Text = ""
   ipp_MtoCta.Value = 0
   ipp_FecCtb.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
End Sub

Private Sub fs_ActivaDet(ByVal p_Activa As Integer)
   grd_DetAsi.Enabled = Not p_Activa
End Sub

Private Sub fs_LeerCabecera()
   'Leyendo Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT FECHA_CNTBL, FEC_REGISTRO, DESC_GLOSA, COD_MONEDA, TASA_CAMBIO, CA.TIPO_NOTA, TN.DESCRIPCION  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO CA INNER JOIN TIPO_NOTA_CNTBL TN ON TRIM(CA.TIPO_NOTA) = TRIM(TN.TIPO_NOTA) "
   g_str_Parame = g_str_Parame & " WHERE ORIGEN = '" & l_str_Origen & "'  "
   g_str_Parame = g_str_Parame & "   AND ANO = " & CStr(modctb_int_PerAno) & "  "
   g_str_Parame = g_str_Parame & "   AND MES = " & CStr(modctb_int_PerMes) & "  "
   g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = " & CStr(modctb_int_CodLib) & "  "
   g_str_Parame = g_str_Parame & "   AND NRO_ASIENTO = " & CStr(modctb_lng_NumAsi) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      ipp_FecCom.Text = g_rst_Princi!FECHA_CNTBL
      pnl_FecReg.Caption = Format(CDate(g_rst_Princi!FEC_REGISTRO), "DD/MM/YYYY")
      txt_GloCab.Text = Trim(g_rst_Princi!DESC_GLOSA & "")
      cmb_MonCtb.ListIndex = gf_Busca_Arregl(l_arr_MonCtb, Format(g_rst_Princi!COD_MONEDA, "000000")) - 1
      cmb_TipOpe.Text = g_rst_Princi!TIPO_NOTA & " - " & g_rst_Princi!DESCRIPCION
      ipp_TipCam.Value = g_rst_Princi!TASA_CAMBIO
   End If
   
   ipp_FecCom.Enabled = False
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_LeerDetalle()
Dim r_int_NumCla  As Integer
   'Leyendo Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ITEM, CNTA_CTBL, DET_GLOSA, FECHA_CNTBL, FLAG_DEBHAB, IMP_MOVSOL, IMP_MOVDOL"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET "
   g_str_Parame = g_str_Parame & " WHERE ORIGEN = '" & l_str_Origen & "'  "
   g_str_Parame = g_str_Parame & "   AND ANO = " & CStr(modctb_int_PerAno) & "  "
   g_str_Parame = g_str_Parame & "   AND MES = " & CStr(modctb_int_PerMes) & "  "
   g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = " & CStr(modctb_int_CodLib) & "  "
   g_str_Parame = g_str_Parame & "   AND NRO_ASIENTO = " & CStr(modctb_lng_NumAsi) & " "
   g_str_Parame = g_str_Parame & " ORDER BY ITEM ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      grd_DetAsi.Redraw = False
      
      Do While Not g_rst_Princi.EOF
         grd_DetAsi.Rows = grd_DetAsi.Rows + 1
         grd_DetAsi.Row = grd_DetAsi.Rows - 1
         grd_DetAsi.Col = 0:  grd_DetAsi.Text = g_rst_Princi!Item
         grd_DetAsi.Col = 1:  grd_DetAsi.Text = Trim(g_rst_Princi!CNTA_CTBL)
         grd_DetAsi.Col = 2:  grd_DetAsi.Text = moddat_gf_Consulta_CtaCtb(Trim(g_rst_Princi!CNTA_CTBL)) 'moddat_gf_Consulta_NomCtaCtb(modctb_str_CodEmp, Trim(g_rst_Princi!CNTA_CTBL)) 'ASIDET_CODCTA
         grd_DetAsi.Col = 3:  grd_DetAsi.Text = Trim(g_rst_Princi!DET_GLOSA & "")
         grd_DetAsi.Col = 4:  grd_DetAsi.Text = Format(CDate(g_rst_Princi!FECHA_CNTBL), "dd/mm/yyyy")
         grd_DetAsi.Col = 5:  grd_DetAsi.Text = CStr(g_rst_Princi!FLAG_DEBHAB)
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"
         grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"
         
         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVSOL), 0, g_rst_Princi!IMP_MOVSOL), "###,###,##0.00")                           'Debe MN
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVDOL), 0, g_rst_Princi!IMP_MOVDOL), "###,###,##0.00") 'Format(g_rst_Princi!IMP_MOVSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")  'Debe ME
         Else
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVSOL), 0, g_rst_Princi!IMP_MOVSOL), "###,###,##0.00")                             'Haber MN
            grd_DetAsi.Col = 9:  grd_DetAsi.Text = Format(IIf(IsNull(g_rst_Princi!IMP_MOVDOL), 0, g_rst_Princi!IMP_MOVDOL), "###,###,##0.00") 'Format(g_rst_Princi!IMP_MOVSOL / CDbl(pnl_TipCam.Caption), "###,###,##0.00")  'Haber ME
         End If
   
         g_rst_Princi.MoveNext
      Loop
      grd_DetAsi.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
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
      
      grd_DetAsi.Col = 6:  r_dbl_TDebMN = r_dbl_TDebMN + CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 7:  r_dbl_THabMN = r_dbl_THabMN + CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 8:  r_dbl_TDebME = r_dbl_TDebME + CDbl(grd_DetAsi.Text)
      grd_DetAsi.Col = 9:  r_dbl_THabME = r_dbl_THabME + CDbl(grd_DetAsi.Text)
   Next r_int_Contad
     
   grd_DetAsi.Redraw = True
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






Private Sub cmb_LibCon_GotFocus()
   cmb_LibCon.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_LibCon_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
          Call ipp_FecCom_LostFocus
          Call gs_SetFocus(ipp_FecCom)
   End Select
End Sub

Private Sub cmb_LibCon_LostFocus()
   cmb_LibCon.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_MonCtb_GotFocus()
   cmb_MonCtb.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_MonCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call gs_SetFocus(ipp_TipCam)
   End If
End Sub

Private Sub cmb_MonCtb_LostFocus()
   cmb_MonCtb.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_TipMov_GotFocus()
   cmb_TipMov.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_TipMov_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r_int_FilSel As Integer
Dim r_int_ColSel As Integer

    r_int_FilSel = grd_DetAsi.Row
    r_int_ColSel = grd_DetAsi.Col

    Select Case KeyCode
    ' keycode conjunto de constantes que se presiona ejm.f1, f2 ,space
        Case vbKeyEscape
            'salgo del Dtp sin cambiar su valor
            cmb_TipMov.Visible = False
            grd_DetAsi.SetFocus

        Case vbKeyReturn
            'Finalizo la captura o entrada de datos
            If cmb_TipMov.ListIndex <> -1 Then
               Call fs_EndEditCmb(r_int_FilSel, r_int_ColSel)
               Call fs_Calcular(r_int_FilSel, r_int_ColSel)
               
               If CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 1 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 2 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 3 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 7 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 8 Then
                  If CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 3, 1)) = 1 Then
                     If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
                        grd_DetAsi.Col = 6
                     Else
                        grd_DetAsi.Col = 7
                     End If
                     grd_DetAsi.SetFocus
                  Else
                     If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
                        grd_DetAsi.Col = 8
                     Else
                        grd_DetAsi.Col = 9
                     End If
                     grd_DetAsi.SetFocus
                  End If
               ElseIf CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 4 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1)) = 5 Then
                   If CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 3, 1)) = 1 Or CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 3, 1)) = 2 Then
                      If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
                         grd_DetAsi.Col = 6
                      Else
                         grd_DetAsi.Col = 7
                      End If
                   End If
               End If
            Else
               MsgBox "Debe seleccionar Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_TipMov)
            End If
            
        Case vbKeyDown
            ' Me muevo una fila hacia abajo
            grd_DetAsi.SetFocus
            DoEvents
            If grd_DetAsi.Row < grd_DetAsi.Rows - 1 Then
                grd_DetAsi.Row = grd_DetAsi.Row + 1
            End If
            
        Case vbKeyUp
            'Me muevo una fila hacia arriba
            grd_DetAsi.SetFocus
            DoEvents
            If grd_DetAsi.Row > grd_DetAsi.FixedRows Then
                grd_DetAsi.Row = grd_DetAsi.Row - 1
            End If
    End Select
End Sub

Private Sub cmb_TipMov_LostFocus()
   cmb_TipMov.BackColor = l_var_ColAnt
End Sub

Private Sub cmb_TipOpe_GotFocus()
   cmb_TipOpe.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_TipOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Call gs_SetFocus(cmb_MonCtb)
End Sub

Private Sub cmb_TipOpe_LostFocus()
   cmb_TipOpe.BackColor = l_var_ColAnt
End Sub

Private Function fs_ValidarCeldaVacias() As Boolean
Dim r_int_ContCol    As Integer
Dim r_int_ContFil   As Integer

   fs_ValidarCeldaVacias = False
   With grd_DetAsi
      For r_int_ContCol = 1 To .Cols - 1
         .Col = r_int_ContCol
         For r_int_ContFil = 0 To .Rows - 1
            .Row = r_int_ContFil
            If .Text = "" Then
               fs_ValidarCeldaVacias = True
            End If
         Next r_int_ContFil
      Next r_int_ContCol
   End With
End Function

Private Function fs_ValidarCeldaCero() As Integer
Dim r_int_ContCol    As Integer
Dim r_int_ContFil   As Integer

   fs_ValidarCeldaCero = 0
   With grd_DetAsi
      If .Row >= 0 Then
         For r_int_ContCol = 6 To .Cols - 1
            .Col = r_int_ContCol
            For r_int_ContFil = .Row To .Rows - 1
               .Row = r_int_ContFil
               If .Text = 0 Then
                  fs_ValidarCeldaCero = fs_ValidarCeldaCero + 1
               End If
            Next r_int_ContFil
         Next r_int_ContCol
      End If
   End With
End Function

Private Sub fs_Calcular(ByVal r_int_FilSel As Integer, ByVal r_int_ColSel As Integer)
Dim r_int_NumCla As Integer

   If grd_DetAsi.TextMatrix(r_int_FilSel, 1) = "" Then Exit Sub
   If Len(grd_DetAsi.TextMatrix(r_int_FilSel, 1)) < 12 Then Exit Sub
   r_int_NumCla = CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 1, 1))
   
   If r_int_NumCla = 1 Or r_int_NumCla = 2 Or r_int_NumCla = 3 Or r_int_NumCla = 7 Or r_int_NumCla = 8 Then
      If CInt(Mid(grd_DetAsi.TextMatrix(r_int_FilSel, 1), 3, 1)) = 1 Then
         If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"
            grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"
         Else
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"                                                                     'Debe MN
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"                                                                     'Debe ME
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")                                 'Haber MN
            grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"                                                                     'Haber ME
         End If
      Else
         If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) * CDbl(ipp_TipCam.Value), "###,###,##0.00")  'Debe MN
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")                                 'Debe ME
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"                                                                     'Haber MN
            grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"                                                                     'Haber ME
         Else
            grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"                                                                     'Debe MN
            grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"                                                                     'Debe ME
            grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(CDbl(ipp_MtoCta.Value) * CDbl(ipp_TipCam.Value), "###,###,##0.00")  'Haber MN
            grd_DetAsi.Col = 9:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")                                 'Haber ME
         End If
      End If
   ElseIf r_int_NumCla = 4 Or r_int_NumCla = 5 Then
      If grd_DetAsi.TextMatrix(r_int_FilSel, 5) = "D" Then
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")                                    'Debe MN
         grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"                                                                        'Debe ME
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = "0.00"                                                                        'Haber MN
         grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"
      Else
         grd_DetAsi.Col = 6:  grd_DetAsi.Text = "0.00"                                                                        'Debe MN
         grd_DetAsi.Col = 7:  grd_DetAsi.Text = Format(ipp_MtoCta.Value, "###,###,##0.00")                                    'Debe ME
         grd_DetAsi.Col = 8:  grd_DetAsi.Text = "0.00"                                                                        'Haber MN
         grd_DetAsi.Col = 9:  grd_DetAsi.Text = "0.00"
      End If
   End If
End Sub

Private Sub moddat_gs_Carga_NotCtb(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TIPO_NOTA_CNTBL "
   g_str_Parame = g_str_Parame & "ORDER BY TIPO_NOTA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem CStr(g_rst_Listas!TIPO_NOTA) & " - " & UCase(Trim(g_rst_Listas!DESCRIPCION))
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_LibCtb(p_CodLib As Integer) As String
   moddat_gf_Consulta_LibCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_LIBCON  "
   g_str_Parame = g_str_Parame & "WHERE LIBCON_CODIGO = " & CStr(p_CodLib) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_LibCtb = Trim(g_rst_Listas!LIBCON_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function moddat_gf_Consulta_CtaCtb(ByVal p_CtaCtb As String) As String
   moddat_gf_Consulta_CtaCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DESC_CNTA, FLAG_PERMITE_MOV FROM CNTBL_CNTA  "
   g_str_Parame = g_str_Parame & " WHERE CNTA_CTBL = " & CStr(p_CtaCtb) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      If g_rst_Listas!FLAG_PERMITE_MOV = 1 Then
         l_bol_CtaCtb = True
      Else
         l_bol_CtaCtb = False
      End If
      moddat_gf_Consulta_CtaCtb = Trim(g_rst_Listas!DESC_CNTA)
   Else
      l_bol_CtaCtb = False
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_ActivaObj(ByVal p_Activa As Integer)
   txt_CtaCtb.Enabled = p_Activa
   cmb_TipMov.Enabled = p_Activa
   ipp_MtoCta.Enabled = p_Activa
   ipp_FecCtb.Enabled = p_Activa
End Sub

Private Sub grd_DetAsi_Click()
   Call fs_IniciaEdicion
End Sub

Private Sub grd_DetAsi_GotFocus()
   If txt_CtaCtb.Visible Then
      If grd_DetAsi.Col = 1 Then
         If Trim(txt_CtaCtb.Text) <> "" Then
            grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = moddat_gf_Consulta_CtaCtb(Trim(txt_CtaCtb.Text))
            If grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) <> "" And l_bol_CtaCtb = True Then
               Call fs_EndEditTxt(grd_DetAsi.Row, grd_DetAsi.Col)
               Call fs_SiguienteCelda
            Else
               If l_bol_CtaCtb = False Then
                  MsgBox "La cuenta no permite movimiento", vbInformation, modgen_g_str_NomPlt
               Else
                  MsgBox "Cuenta no existe", vbInformation, modgen_g_str_NomPlt
               End If
               txt_CtaCtb.Text = Empty
               grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col) = Empty
               grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = Empty
               Call gs_SelecTodo(txt_CtaCtb)
            End If
         End If
      ElseIf grd_DetAsi.Col = 3 Then
         Call fs_EndEditTxt(grd_DetAsi.Row, grd_DetAsi.Col)
         Call fs_SiguienteCelda
      End If
      txt_CtaCtb.Visible = False
   End If

   If cmb_TipMov.Visible Then
      If cmb_TipMov.ListIndex <> -1 Then
         grd_DetAsi.Text = UCase(Trim(Mid(cmb_TipMov.Text, 1, 1)))
      Else
         MsgBox "Debe seleccionar Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipMov)
      End If
      cmb_TipMov.Visible = False
   End If

   If ipp_MtoCta.Visible Then
      grd_DetAsi.Text = ipp_MtoCta.Value
      ipp_MtoCta.Visible = False
   End If
End Sub

Private Sub grd_DetAsi_KeyPress(KeyAscii As Integer)
   Select Case grd_DetAsi.Col
      Case 1, 3
          fs_GridEditText KeyAscii
      Case 4
          fs_GridEditDate
      Case 5
          fs_GridEditCombo
      Case 6, 7, 8, 9
          fs_GridEditFp KeyAscii
   End Select
End Sub

Private Sub grd_DetAsi_LeaveCell()
   If txt_CtaCtb.Visible Then
      If grd_DetAsi.Col = 1 Then
         If Trim(txt_CtaCtb.Text) <> "" Then
            grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = moddat_gf_Consulta_CtaCtb(Trim(txt_CtaCtb.Text))
            If grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) <> "" And l_bol_CtaCtb = True Then
               Call fs_EndEditTxt(grd_DetAsi.Row, grd_DetAsi.Col)
               Call fs_SiguienteCelda
            Else
               If l_bol_CtaCtb = False Then
                  MsgBox "La cuenta no permite movimiento", vbInformation, modgen_g_str_NomPlt
               Else
                  MsgBox "Cuenta no existe", vbInformation, modgen_g_str_NomPlt
               End If
               txt_CtaCtb.Text = Empty
               grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = Empty
               Call gs_SelecTodo(txt_CtaCtb)
            End If
         End If
      ElseIf grd_DetAsi.Col = 3 Then
         Call fs_EndEditTxt(grd_DetAsi.Row, grd_DetAsi.Col)
         Call fs_SiguienteCelda
      End If
      txt_CtaCtb.Visible = False
   End If

   If cmb_TipMov.Visible Then
      grd_DetAsi.Text = UCase(Trim(Mid(cmb_TipMov.Text, 1, 1)))
      cmb_TipMov.Visible = False
   End If
   If ipp_FecCtb.Visible Then
      grd_DetAsi.Text = ipp_FecCtb.Text
      ipp_FecCtb.Visible = False
   End If
   If ipp_MtoCta.Visible Then
      grd_DetAsi.Text = ipp_MtoCta.Value
      ipp_MtoCta.Visible = False
   End If
End Sub

Private Sub ipp_FecCom_GotFocus()
   ipp_FecCom.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_FecCom_InvalidData(NextWnd As Long)
   If CDate(ipp_FecCom.Text) < CDate(modctb_str_FecIni) Then
      ipp_FecCom.Text = modctb_str_FecIni
   ElseIf CDate(ipp_FecCom.Text) > CDate(modctb_str_FecFin) Then
      ipp_FecCom.Text = modctb_str_FecFin
   End If
End Sub

Private Sub ipp_FecCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipOpe)
   End If
End Sub

Private Sub ipp_FecCom_LostFocus()
Dim r_int_TipMon     As Integer
   ipp_FecCom.BackColor = l_var_ColAnt
   
   If cmb_LibCon.ListIndex > -1 And cmb_MonCtb.ListIndex > -1 Then
      r_int_TipMon = CInt(l_arr_MonCtb(cmb_MonCtb.ListIndex + 1).Genera_Codigo)
      If r_int_TipMon = 1 Then
         r_int_TipMon = 2
      End If
   
      Select Case cmb_LibCon.ItemData(cmb_LibCon.ListIndex)
         Case 8:     ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(3, r_int_TipMon, Format(CDate(ipp_FecCom.Text), "yyyymmdd"), 2)
         Case 9:     ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(3, r_int_TipMon, Format(CDate(ipp_FecCom.Text), "yyyymmdd"), 1)
         Case Else:: ipp_TipCam.Value = moddat_gf_ObtieneTipCamDia(2, r_int_TipMon, Format(CDate(ipp_FecCom.Text), "yyyymmdd"), 2)
      End Select
   End If
End Sub

Private Sub ipp_FecCtb_GotFocus()
   ipp_FecCtb.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_FecCtb_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r_int_FilSel As Integer
Dim r_int_ColSel As Integer

    r_int_FilSel = grd_DetAsi.Row
    r_int_ColSel = grd_DetAsi.Col

    Select Case KeyCode
        'keycode conjunto de constantes que se presiona ejm.f1, f2 ,space
        Case vbKeyEscape
            ' Leave the text unchanged.
            'salgo del ipp_FecCtb_KeyDown sin cambiar su valor
            ipp_FecCtb.Visible = False
            grd_DetAsi.SetFocus

        Case vbKeyReturn
            'Finalizo la captura o entrada de datos
            Call fs_EndEditFec(r_int_FilSel, r_int_ColSel)
            Call fs_SiguienteCelda
            
        Case vbKeyDown
            ' Me muevo una fila hacia abajo
            grd_DetAsi.SetFocus
            DoEvents
            If grd_DetAsi.Row < grd_DetAsi.Rows - 1 Then
                grd_DetAsi.Row = grd_DetAsi.Row + 1
            End If
            
        Case vbKeyUp
            'Me muevo una fila hacia arriba
            grd_DetAsi.SetFocus
            DoEvents
            If grd_DetAsi.Row > grd_DetAsi.FixedRows Then
               grd_DetAsi.Row = grd_DetAsi.Row - 1
            End If
    End Select
End Sub

Private Sub ipp_FecCtb_LostFocus()
    ipp_FecCtb.BackColor = l_var_ColAnt
End Sub

Private Sub ipp_MtoCta_GotFocus()
   ipp_MtoCta.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_MtoCta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r_int_FilSel As Integer
Dim r_int_ColSel As Integer

   r_int_FilSel = grd_DetAsi.Row
   r_int_ColSel = grd_DetAsi.Col

   Select Case KeyCode
      ' keycode conjunto de constantes que se presionan ejm.f1, f2 ,space
       Case vbKeyEscape
           'salgo del Dtp sin cambiar su valor
           ipp_MtoCta.Visible = False
           grd_DetAsi.SetFocus
       Case vbKeyReturn
           'Finalizo la captura o entrada de datos
           Call fs_EndEditIpp(r_int_FilSel, r_int_ColSel)
       Case vbKeyDown
           ' Me muevo una fila hacia abajo
           grd_DetAsi.SetFocus
           DoEvents
           If grd_DetAsi.Row < grd_DetAsi.Rows - 1 Then
               grd_DetAsi.Row = grd_DetAsi.Row + 1
           End If
       Case vbKeyUp
           'Me muevo una fila hacia arriba
           grd_DetAsi.SetFocus
           DoEvents
           If grd_DetAsi.Row > grd_DetAsi.FixedRows Then
               grd_DetAsi.Row = grd_DetAsi.Row - 1
           End If
   End Select
End Sub

Private Sub ipp_MtoCta_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_MtoCta_LostFocus()
   ipp_MtoCta.BackColor = l_var_ColAnt
   Call fs_TotDebHab
End Sub

Private Sub ipp_TipCam_GotFocus()
   ipp_TipCam.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_GloCab)
   End If
End Sub

Private Sub ipp_TipCam_LostFocus()
   ipp_TipCam.BackColor = l_var_ColAnt
End Sub

Private Sub txt_CtaCtb_DblClick()
   If grd_DetAsi.Col = 1 Then
      frm_Ctb_AsiCtb_04.Show 1
   End If
End Sub

Private Sub txt_CtaCtb_GotFocus()
   txt_CtaCtb.BackColor = modgen_g_con_ColAma
   Select Case grd_DetAsi.Col
      Case 1
         txt_CtaCtb.MaxLength = 12
      Case 3
         txt_CtaCtb.MaxLength = 250
   End Select
End Sub

Private Sub txt_CtaCtb_KeyDown(KeyCode As Integer, Shift As Integer)
'constantes de codigo de tecla
Dim r_int_FilSel As Integer
Dim r_int_ColSel As Integer

   r_int_FilSel = grd_DetAsi.Row
   r_int_ColSel = grd_DetAsi.Col
    
   Select Case KeyCode
        Case vbKeyEscape
           'Salgo del texto y no cambia su valor
           txt_CtaCtb.Visible = False
           grd_DetAsi.SetFocus
   
        Case vbKeyReturn
           ' Finalizo el ingreso de datos
           If grd_DetAsi.Col = 1 Then
              If Trim(txt_CtaCtb.Text) <> "" Then
                  grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = moddat_gf_Consulta_CtaCtb(Trim(txt_CtaCtb.Text))
                  If grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) <> "" And l_bol_CtaCtb = True Then
                     Call fs_EndEditTxt(r_int_FilSel, r_int_ColSel)
                     Call fs_SiguienteCelda
                  Else
                     If l_bol_CtaCtb = False Then
                        MsgBox "La cuenta no permite movimiento", vbInformation, modgen_g_str_NomPlt
                     Else
                        MsgBox "Cuenta no existe", vbInformation, modgen_g_str_NomPlt
                     End If
                     txt_CtaCtb.Text = Empty
                     grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col) = Empty
                     grd_DetAsi.TextMatrix(grd_DetAsi.Row, grd_DetAsi.Col + 1) = Empty
                     Call gs_SelecTodo(txt_CtaCtb)
                  End If
              End If
           ElseIf grd_DetAsi.Col = 3 Then
               Call fs_EndEditTxt(r_int_FilSel, r_int_ColSel)
               Call fs_SiguienteCelda
           End If
   
       Case vbKeyDown
         ' Me muevo una fila hacia abajo.
           grd_DetAsi.SetFocus
           DoEvents
            If grd_DetAsi.Row < grd_DetAsi.Rows - 1 Then
                grd_DetAsi.Row = grd_DetAsi.Row + 1
            End If
       Case vbKeyUp
         ' Me muevo 1 fila arriba
            grd_DetAsi.SetFocus
            DoEvents
            If grd_DetAsi.Row > grd_DetAsi.FixedRows Then
                grd_DetAsi.Row = grd_DetAsi.Row - 1
            End If
   End Select
End Sub

Private Sub txt_CtaCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Or (KeyAscii = vbKeyTab) Then
      KeyAscii = 0
   Else
      If grd_DetAsi.Col = 1 Then
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & Chr(22))
      ElseIf grd_DetAsi.Col = 3 Then
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
      End If
   End If
End Sub

Private Sub txt_CtaCtb_KeyUp(KeyCode As Integer, Shift As Integer)
   If grd_DetAsi.Col = 1 Then
      KeyCode = gf_ValidaCaracter(KeyCode, modgen_g_con_NUMERO)
   ElseIf grd_DetAsi.Col = 3 Then
      KeyCode = gf_ValidaCaracter(KeyCode, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub txt_CtaCtb_LostFocus()
   txt_CtaCtb.BackColor = l_var_ColAnt
End Sub

Private Sub txt_CtaCtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If grd_DetAsi.Col = 1 Then
      txt_CtaCtb.MousePointer = 99
      txt_CtaCtb.MouseIcon = LoadPicture("C:\Windows\Cursors\harrow.cur")
   Else
      txt_CtaCtb.MousePointer = 0
   End If
End Sub

Private Sub txt_GloCab_GotFocus()
   txt_GloCab.BackColor = modgen_g_con_ColAma
   Call gs_SelecTodo(txt_GloCab)
End Sub

Private Sub txt_GloCab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If grd_DetAsi.Rows = 0 Then
         Call gs_SetFocus(cmd_DetNue)
      Else
         Call gs_SetFocus(grd_DetAsi)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
   'Call gs_SetFocus(cmd_DetNue)
End Sub

Private Sub txt_GloCab_LostFocus()
   txt_GloCab.BackColor = l_var_ColAnt
End Sub

'------------- PARA PERMITIR LA EDICION DEL GRID cuando se selecciona el Txt_CtaCtb -----------
Private Sub fs_GridEditText(ByVal KeyAscii As Integer)
   'posiciona el textbox encima de la celda
   If grd_DetAsi.Col = 1 Or grd_DetAsi.Col = 3 Then
      txt_CtaCtb.Left = grd_DetAsi.CellLeft + grd_DetAsi.Left
      txt_CtaCtb.Top = grd_DetAsi.CellTop + grd_DetAsi.Top
      txt_CtaCtb.Width = grd_DetAsi.CellWidth
      txt_CtaCtb.Height = grd_DetAsi.CellHeight
      txt_CtaCtb.Visible = True
      txt_CtaCtb.Enabled = True
      txt_CtaCtb.SetFocus
   End If
   Select Case grd_DetAsi.Col
      Case 1
         txt_CtaCtb.MaxLength = 12
      Case 3
         txt_CtaCtb.MaxLength = 250
   End Select
   Select Case KeyAscii
      Case 0 To Asc(" ")                              'para cualquier caracter extraño que se quiera introducir
          txt_CtaCtb.Text = grd_DetAsi.Text
          txt_CtaCtb.SelStart = Len(txt_CtaCtb.Text)  'donde se ubica el punto inicial del txt
      Case Else
          txt_CtaCtb.Text = Chr$(KeyAscii)
          txt_CtaCtb.SelStart = 1                     'coloca el cursor despues del valor valido
   End Select
End Sub

'------------- PARA PERMITIR LA EDICION DEL GRID cuando se selecciona el Ipp_FecCtb -----------
Private Sub fs_GridEditDate()
   ipp_FecCtb.Left = grd_DetAsi.CellLeft + grd_DetAsi.Left
   ipp_FecCtb.Top = grd_DetAsi.CellTop + grd_DetAsi.Top
   ipp_FecCtb.Width = grd_DetAsi.CellWidth
   ipp_FecCtb.Height = grd_DetAsi.CellHeight
   ipp_FecCtb.Visible = True
   ipp_FecCtb.SetFocus
     
   If grd_DetAsi.Text = "" Then
      ipp_FecCtb.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
   Else
      'ipp_FecCtb.Text = grd_DetAsi.Text
      'If fs_ValidaPeriodo(ipp_FecCtb.Text) = False Then
      '   ipp_FecCtb.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
      'End If
      'ipp_FecCtb.Text = Format(CDate(ipp_FecCom.Text), "dd/mm/yyyy")
   End If
End Sub

'------------- PARA PERMITIR LA EDICION DEL GRID cuando se selecciona el Cmb_TipMov -----------
Private Sub fs_GridEditCombo()
   cmb_TipMov.Left = grd_DetAsi.CellLeft + grd_DetAsi.Left
   cmb_TipMov.Top = grd_DetAsi.CellTop + grd_DetAsi.Top
   cmb_TipMov.Width = grd_DetAsi.CellWidth
   '   cmb_TipMov.Height = grd_DetAsi.CellHeight
   cmb_TipMov.Visible = True
   cmb_TipMov.Enabled = True
   cmb_TipMov.SetFocus
   
   If grd_DetAsi.Text = "H" Then
       cmb_TipMov.Text = "HABER" 'Left(cmb_TipMov.Text, 1)
   ElseIf grd_DetAsi.Text = "D" Then
       cmb_TipMov.Text = "DEBE" 'grd_DetAsi.Text
   End If
End Sub

'------------- PARA PERMITIR LA EDICION DEL GRID cuando se selecciona ipp_MtoCta -----------
Private Sub fs_GridEditFp(ByVal KeyAscii As Integer)
    'posiciona el ipp encima de la celda
    If grd_DetAsi.Col = 6 Or grd_DetAsi.Col = 7 Or grd_DetAsi.Col = 8 Or grd_DetAsi.Col = 9 Then
        ipp_MtoCta.Left = grd_DetAsi.CellLeft + grd_DetAsi.Left
        ipp_MtoCta.Top = grd_DetAsi.CellTop + grd_DetAsi.Top
        ipp_MtoCta.Width = grd_DetAsi.CellWidth
        ipp_MtoCta.Height = grd_DetAsi.CellHeight
        ipp_MtoCta.Visible = True
        ipp_MtoCta.Enabled = True
        ipp_MtoCta.SetFocus
    End If
    Select Case KeyAscii
        Case 0 To Asc(" ")                               'para cualquier caracter extraño que se quiera introducir
            ipp_MtoCta.Value = grd_DetAsi.Text
            ipp_MtoCta.SelStart = Len(ipp_MtoCta.Text)   'donde se ubica el punto inicial del txt
        Case Else
            ipp_MtoCta.Value = Chr(KeyAscii)
            ipp_MtoCta.SelStart = 1                      'coloca el cursor despues del valor valido
    End Select
End Sub

'------------- PARA PERMITIR FINALIZAR LA EDICIÓN DEL GRID cuando se selecciona txt_CtaCtb -----------
Private Sub fs_EndEditTxt(ByVal r_int_FilSel As Integer, r_int_ColSel As Integer)
   'termina la edicion de datos pasando su valor al grd_DetAsi
   If txt_CtaCtb.Visible Then
      grd_DetAsi.TextMatrix(r_int_FilSel, r_int_ColSel) = txt_CtaCtb.Text
      grd_DetAsi.SetFocus
      txt_CtaCtb.Text = Empty
      txt_CtaCtb.Visible = False
   End If
End Sub

'------------- PARA PERMITIR FINALIZAR LA EDICIÓN DEL GRID cuando se selecciona cmb_TipMov -----------
Private Sub fs_EndEditCmb(ByVal r_int_FilSel As Integer, ByVal r_int_ColSel As Integer)
   If cmb_TipMov.Visible Then
     grd_DetAsi.TextMatrix(r_int_FilSel, r_int_ColSel) = Left(cmb_TipMov.Text, 1)
     grd_DetAsi.SetFocus
     cmb_TipMov.Visible = False
   End If
End Sub

'------------- PARA PERMITIR FINALIZAR LA EDICIÓN DEL GRID cuando se selecciona ipp_MtoCta -----------
Private Sub fs_EndEditIpp(ByVal r_int_FilSel As Integer, ByVal r_int_ColSel As Integer)
   If ipp_MtoCta.Visible Then
     grd_DetAsi.Col = 4
     If grd_DetAsi.Text <> Empty Then
         grd_DetAsi.TextMatrix(r_int_FilSel, r_int_ColSel) = ipp_MtoCta.Value
         grd_DetAsi.SetFocus
         ipp_MtoCta.Visible = False
         Call fs_Calcular(r_int_FilSel, r_int_ColSel)
         Call fs_SiguienteCelda
     Else
         MsgBox "Ingrese Tipo de Movimiento", vbInformation, modgen_g_str_NomPlt
         grd_DetAsi.TextMatrix(r_int_FilSel, r_int_ColSel) = 0#
         ipp_MtoCta.Value = 0#
         Call fs_Calcular(r_int_FilSel, r_int_ColSel)
         grd_DetAsi.Col = 4
         Call fs_IniciaEdicion
     End If
   End If
End Sub

'------------- PARA PERMITIR FINALIZAR LA EDICIÓN DEL GRID cuando se selecciona ipp_FecCtb -----------
Private Sub fs_EndEditFec(ByVal r_int_FilSel As Integer, ByVal r_int_ColSel As Integer)
    If ipp_FecCtb.Visible Then
      grd_DetAsi.TextMatrix(r_int_FilSel, r_int_ColSel) = ipp_FecCtb.Text
      grd_DetAsi.SetFocus
      ipp_FecCtb.Visible = False
   End If
End Sub

Private Sub fs_IniciaEdicion()
   'empieza la edicion del grid verificando de que control se hará uso (Txt, cmb, date y ipp)
   If grd_DetAsi.Col = 1 Or grd_DetAsi.Col = 3 Then
      fs_GridEditText Asc(" ")
   ElseIf grd_DetAsi.Col = 4 Then
      fs_GridEditDate
   ElseIf grd_DetAsi.Col = 5 Then
      fs_GridEditCombo
   ElseIf grd_DetAsi.Col = 6 Or grd_DetAsi.Col = 7 Or grd_DetAsi.Col = 8 Or grd_DetAsi.Col = 9 Then
      fs_GridEditFp Asc(" ")
   End If
   
End Sub

Private Sub fs_SiguienteCelda()
   If grd_DetAsi.Col < grd_DetAsi.Cols - 1 Then
      If grd_DetAsi.Col = 1 Then
          grd_DetAsi.Col = grd_DetAsi.Col + 2
      Else
          grd_DetAsi.Col = grd_DetAsi.Col + 1
      End If
      grd_DetAsi.SetFocus
   Else
      Call gs_SetFocus(cmd_DetNue)
   End If
End Sub
