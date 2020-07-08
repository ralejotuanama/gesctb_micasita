VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_TarCre_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   Icon            =   "GesCtb_frm_220.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   5100
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   8996
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
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
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
            TabIndex        =   14
            Top             =   150
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Tarjeta Crédito Corporativo"
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
            Picture         =   "GesCtb_frm_220.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_220.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7620
            Picture         =   "GesCtb_frm_220.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3480
         Left            =   60
         TabIndex        =   16
         Top             =   1500
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
         _ExtentY        =   6138
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2310
            Width           =   6330
         End
         Begin VB.ComboBox cmb_Asigna 
            Height          =   315
            Left            =   1710
            TabIndex        =   8
            Text            =   "cmb_Asigna"
            Top             =   2640
            Width           =   6330
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   660
            Width           =   1635
         End
         Begin VB.ComboBox cmb_Respon 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2970
            Width           =   6330
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1320
            Width           =   1635
         End
         Begin Threed.SSPanel pnl_NumRen 
            Height          =   315
            Left            =   1710
            TabIndex        =   17
            Top             =   330
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
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
         Begin EditLib.fpDoubleSingle ipp_ImpAsig 
            Height          =   315
            Left            =   1710
            TabIndex        =   5
            Top             =   1650
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
         Begin EditLib.fpDoubleSingle ipp_PagTarj 
            Height          =   315
            Left            =   1710
            TabIndex        =   6
            Top             =   1980
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   5910
            TabIndex        =   1
            Top             =   660
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "2017"
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
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   5910
            TabIndex        =   3
            Top             =   990
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin EditLib.fpDateTime ipp_FecDia 
            Height          =   315
            Left            =   1710
            TabIndex        =   2
            Top             =   990
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   4620
            TabIndex        =   28
            Top             =   1050
            Width           =   930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   2370
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Asignación:"
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   2700
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Año Rendición:"
            Height          =   195
            Left            =   4620
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   3030
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Rendición:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Numero Rendición:"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   420
            Width           =   1365
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
            TabIndex        =   21
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   1380
            Width           =   630
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Monto Asignado:"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   1710
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pago de Tarjeta:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   2040
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_TarCre_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Respon()      As moddat_tpo_Genera
Dim l_arr_Asigna()      As moddat_tpo_Genera

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then 'consultar
      pnl_Titulo.Caption = "Registro de Tarjeta Crédito Corporativo - Consultar"
      cmd_Grabar.Visible = False
      Call fs_CargarDatos
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then 'insertar
      pnl_Titulo.Caption = "Registro de Tarjeta Crédito Corporativo - Adicionar"
      ipp_ImpAsig.Text = "13,000.00"
      ipp_ImpAsig.Enabled = False
   ElseIf moddat_g_int_FlgGrb = 2 Then 'modificar
      pnl_Titulo.Caption = "Registro de Tarjeta Crédito Corporativo - Modificar"
      Call fs_CargarDatos
      Call fs_Desabilitar
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
      
   If moddat_g_int_FlgGrb = 1 Then 'insertar
      Call moddat_gs_Carga_EjecMC(cmb_Respon, l_arr_Respon, 133, 1)
   Else 'editar y insertar
      Call moddat_gs_Carga_EjecMC(cmb_Respon, l_arr_Respon, 133, 2)
   End If
End Sub

Private Sub fs_Limpiar()
   cmb_TipDoc.ListIndex = 0
   cmb_Asigna.Text = ""
   cmb_Respon.ListIndex = -1
   
   pnl_NumRen.Caption = ""
   Call gs_BuscarCombo_Item(cmb_PerMes, Month(date))
   
   If Year(date) >= 2017 Then
      ipp_PerAno.Text = Year(date)
   Else
      ipp_PerAno.Text = 2017
   End If
   ipp_FecDia.Text = moddat_g_str_FecSis
   Call ipp_FecDia_LostFocus 'TIPO CAMBIO SBS(2) - VENTA(1)
   
   cmb_Moneda.ListIndex = 0
   ipp_ImpAsig.Text = "0.00"
   ipp_PagTarj.Text = "0.00"

   cmb_TipDoc.ListIndex = 0
   cmb_Asigna.Text = ""
   cmb_Respon.ListIndex = -1
End Sub

Private Sub fs_Desabilitar()
   cmb_PerMes.Enabled = False
   ipp_PerAno.Enabled = False
   cmb_Moneda.Enabled = False
   ipp_ImpAsig.Enabled = False
   ipp_PagTarj.Enabled = False
   cmb_TipDoc.Enabled = False
   cmb_Asigna.Enabled = False
   cmb_Respon.Enabled = False
   ipp_FecDia.Enabled = False
   ipp_TipCam.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_CargarPrv(p_Combo_Tdoc As ComboBox, p_Combo_Nom As ComboBox, p_Tipo As Integer)
   ReDim l_arr_Asigna(0) 'RESPONSABLE(1)
   
   p_Combo_Nom.Clear
   p_Combo_Nom.Text = ""
   If (p_Combo_Tdoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & p_Combo_Tdoc.ItemData(p_Combo_Tdoc.ListIndex)
   g_str_Parame = g_str_Parame & "    AND A.MAEPRV_SITTAR = 1 "
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
      'RESPONSABLE
      ReDim Preserve l_arr_Asigna(UBound(l_arr_Asigna) + 1)
      l_arr_Asigna(UBound(l_arr_Asigna)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
      l_arr_Asigna(UBound(l_arr_Asigna)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmb_TipDoc_Click()
   'RESPONSABLE
   Call fs_CargarPrv(cmb_TipDoc, cmb_Asigna, 1)
End Sub

Private Sub cmd_Grabar_Click()
Dim r_lng_FecAux  As Long
Dim r_dbl_Import  As Double
    r_dbl_Import = 0
    r_lng_FecAux = 0
    
    If cmb_PerMes.ListIndex = -1 Then
       MsgBox "Debe seleccionar un periodo.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_PerMes)
       Exit Sub
    End If
    
    If ipp_PerAno.Value = 0 Then
       MsgBox "Debe de ingresar un año.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_PerAno)
       Exit Sub
    End If
    
    'TIPO SBS
    If CDbl(moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecDia.Text, "yyyymmdd"), 1)) = 0 Then
       MsgBox "Tiene que registrar el tipo de cambio sbs del día.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_PerMes)
       Exit Sub
    End If
    
    If cmb_Moneda.ListIndex = -1 Then
       MsgBox "Debe seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Moneda)
       Exit Sub
    End If
        
    If CDbl(ipp_ImpAsig.Text) <= 0 Then
       MsgBox "El monto asignado debe de ser mayor a cero.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_ImpAsig)
       Exit Sub
    End If
    
    If CDbl(ipp_PagTarj.Text) > CDbl(ipp_ImpAsig.Text) Then
        MsgBox "El Pago de tarjeta no puede ser mayor al monto asignado.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(ipp_PagTarj)
        Exit Sub
    End If
            
    'validacion del asignacion
    If cmb_TipDoc.ListIndex = -1 Then
       MsgBox "Debe de seleccionar el tipo documento del asignado.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipDoc)
       Exit Sub
    End If
            
    If fs_Valida_LstPrv(cmb_TipDoc, cmb_Asigna, "asignación", l_arr_Asigna) = False Then
       Exit Sub
    End If
            
    If cmb_Respon.ListIndex = -1 Then
       MsgBox "Debe de seleccionar el responsable.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Respon)
       Exit Sub
    End If
     
    If Format(ipp_FecDia.Text, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
       MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
             
       If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Call gs_SetFocus(ipp_FecDia)
          Exit Sub
       End If
    End If
     
'    If (Format(ipp_FecDia.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'        Format(ipp_FecDia.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'        MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'        Call gs_SetFocus(ipp_FecDia)
'        Exit Sub
'    End If
   
   '--ipp_FecDia.Text
'   If Format(moddat_g_str_FecSis, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > modctb_int_PerAno & Format(modctb_int_PerMes, "00") & Format(moddat_g_int_PerLim, "00")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FecDia)
'          Exit Sub
'      End If
'      MsgBox "Los asiento a generar perteneceran al periodo anterior.", vbExclamation, modgen_g_str_NomPlt
'   Else
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FecDia)
'          Exit Sub
'      End If
'   End If
              
    If fs_ValidaPeriodo(ipp_FecDia.Text) = False Then
       Exit Sub
    End If
        
    'If fs_Valida_Reg = True Then
    '   MsgBox Trim(cmb_Asigna.Text) & ", ya tiene un registro para este periodo.", vbExclamation, modgen_g_str_NomPlt
    '   Call gs_SetFocus(cmb_PerMes)
    '   Exit Sub
    'End If

    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Public Function fs_Valida_Reg() As Boolean
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset
      
   fs_Valida_Reg = True
   Screen.MousePointer = 11
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT COUNT(*)  AS CONTEO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_CAJCHC A "
   r_str_Parame = r_str_Parame & "  WHERE A.CAJCHC_TIPTAB = 6 "
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_PERMES = " & Format(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), "00") 'CAJCHC_PERMES
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_PERANO = " & ipp_PerAno.Text 'CAJCHC_PERANO
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) 'CAJCHC_TIPDOC
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_NUMDOC = '" & fs_NumDoc(cmb_Asigna.Text, cmb_TipDoc) & "' " 'CAJCHC_NUMDOC
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
   r_str_Parame = r_str_Parame & "    AND A.CAJCHC_SITUAC = 1  "
     
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Screen.MousePointer = 0
   r_rst_Genera.MoveFirst
   
   fs_Valida_Reg = True
   If r_rst_Genera!CONTEO = 0 Then
      fs_Valida_Reg = False
   End If
   
End Function

Private Function fs_Valida_LstPrv(p_ComboTip As ComboBox, p_ComboNom As ComboBox, p_MsgNom As String, p_Arregl() As moddat_tpo_Genera) As Boolean
Dim r_int_Contar  As Integer
Dim r_bol_Estado  As Boolean
   
   fs_Valida_LstPrv = True
   r_bol_Estado = True
   
   If Len(Trim(p_ComboNom.Text)) = 0 Then
       MsgBox "Tiene que ingresar un " & p_MsgNom & ".", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(p_ComboNom)
       r_bol_Estado = False 'Exit Sub
   Else
       If (fs_ValNumDoc(p_ComboTip, p_ComboNom) = False) Then
           r_bol_Estado = False 'Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(p_ComboNom.Text), "-") > 0 Then
              For r_int_Contar = 1 To UBound(p_Arregl)
                  If Trim(Mid(p_ComboNom.Text, 1, InStr(Trim(p_ComboNom.Text), "-") - 1)) = Trim(p_Arregl(r_int_Contar).Genera_Codigo) Then
                     r_bol_Estado = True
                     Exit For
                  End If
              Next
           End If
           If r_bol_Estado = False Then
              MsgBox "El " & p_MsgNom & " no se encuentra en la lista.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(p_ComboNom)
              'Exit Sub
           End If
       End If
   End If
   
   fs_Valida_LstPrv = r_bol_Estado
End Function

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

Public Sub fs_Grabar()
Dim r_str_AsiGen   As String
Dim r_str_CodGen  As String

   r_str_AsiGen = ""
   r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 11)
   Else
      r_str_CodGen = Trim(pnl_NumRen.Caption)
   End If
   
   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_TARCRE ( "
   g_str_Parame = g_str_Parame & CLng(r_str_CodGen) & ", " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & Format(ipp_FecDia.Text, "yyyymmdd") & ", " 'CAJCHC_FECCAJ
   g_str_Parame = g_str_Parame & Format(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), "00") & ", " 'CAJCHC_PERMES
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & ", " 'CAJCHC_PERANO
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'CAJCHC_CODMON
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpAsig.Text) & ", " 'CAJCHC_IMPORT
   g_str_Parame = g_str_Parame & CDbl(ipp_PagTarj.Text) & ", " 'CAJCHC_IMPORT_2
   g_str_Parame = g_str_Parame & "'" & l_arr_Respon(cmb_Respon.ListIndex + 1).Genera_Codigo & "', "  'CAJCHC_RESPON
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'CAJCHC_TIPDOC
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Asigna.Text, cmb_TipDoc) & "', " 'CAJCHC_NUMDOC
   g_str_Parame = g_str_Parame & CDbl(ipp_TipCam.Text) & ", " 'CAJCHC_TIPCAM%TYPE
   g_str_Parame = g_str_Parame & "1, "  'CAJCHC_SITUAC%TYPE
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
       If CDbl(ipp_PagTarj.Text) > 0 Then
          Call fs_GeneraAsiento(Format(g_rst_Genera!CODIGO, "0000000000"), r_str_AsiGen)
          
          MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
                 "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
       Else
          MsgBox "Los datos se grabaron correctamente." & vbCrLf & _
                 "No se genero el asiento por que el pago de tarjeta es cero.", vbInformation, modgen_g_str_NomPlt
       End If
       Call frm_Ctb_TarCre_02.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_TarCre_02.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 3) Then
       MsgBox "El Importe no puede ser menor al total de su detalle: " & Format(g_rst_Genera!TOTDET, "###,###,##0.00") & "", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_ImpAsig)
       Screen.MousePointer = 0
   End If
End Sub

Private Sub fs_GeneraAsiento(ByVal p_Codigo As String, ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_FecPgC        As String
Dim r_str_FecPgL        As String
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer

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
                   
   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPgL = ipp_FecDia.Text
   r_str_FecPgC = Format(ipp_FecDia.Text, "yyyymmdd")

   'TIPO SBS
   r_dbl_TipSbs = CDbl(ipp_TipCam.Text) 'moddat_gf_ObtieneTipCamDia(2, 2, r_str_FecPgC, 1)

   r_str_Glosa = "TARJETA CREDITO " & p_Codigo
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)

   'r_int_PerMes = modctb_int_PerMes 'Month(r_str_FecPgL)
   'r_int_PerAno = modctb_int_PerAno 'Year(r_str_FecPgL)
   r_int_PerMes = Month(r_str_FecPgL)
   r_int_PerAno = Year(r_str_FecPgL)

   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)

   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPgL, "1")

   'Insertar en detalle
   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      r_dbl_ImpSol = CDbl(ipp_PagTarj.Text)
   Else
      r_dbl_ImpSol = CDbl(CDbl(ipp_PagTarj.Text) * r_dbl_TipSbs) 'Importe * CONVERTIDO
      r_dbl_ImpDol = CDbl(ipp_PagTarj.Text)
   End If

   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, "291807010117", CDate(r_str_FecPgL), _
                                        r_str_Glosa, "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPgL))

   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, "111301060102", CDate(r_str_FecPgL), _
                                        r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPgL))
   p_AsiGen = r_str_AsiGen

   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATCNT = '" & r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 6 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Private Sub fs_CargarDatos()
Dim r_rst_Princi     As ADODB.Recordset
Dim r_int_Contad     As Integer

   Call gs_SetFocus(cmb_PerMes)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_PERMES, A.CAJCHC_PERANO,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_CODMON, A.CajChc_Import , A.CAJCHC_IMPORT_2, A.CajChc_Respon,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_TIPDOC, A.CAJCHC_NUMDOC, CAJCHC_TIPCAM  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB = 6   "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      pnl_NumRen.Caption = Format(r_rst_Princi!CajChc_CodCaj, "0000000000")
      Call gs_BuscarCombo_Item(cmb_PerMes, r_rst_Princi!CAJCHC_PERMES)
      ipp_PerAno.Text = r_rst_Princi!CAJCHC_PERANO
      ipp_FecDia.Text = gf_FormatoFecha(r_rst_Princi!CajChc_FecCaj)
      ipp_TipCam.Text = Format(r_rst_Princi!CAJCHC_TIPCAM, "###,###,##0.000000") & " "
      Call gs_BuscarCombo_Item(cmb_Moneda, r_rst_Princi!CAJCHC_CODMON)
      ipp_ImpAsig.Text = r_rst_Princi!CajChc_Import
      ipp_PagTarj.Text = r_rst_Princi!CAJCHC_IMPORT_2
                     
      If Trim(r_rst_Princi!CAJCHC_TIPDOC & "") <> "" Then
         Call gs_BuscarCombo_Item(cmb_TipDoc, r_rst_Princi!CAJCHC_TIPDOC)
      Else
         cmb_TipDoc.ListIndex = -1
      End If
      If Trim(r_rst_Princi!CAJCHC_NUMDOC & "") <> "" Then
         cmb_Asigna.ListIndex = fs_ComboIndex(cmb_Asigna, r_rst_Princi!CAJCHC_NUMDOC & "", 0)
      End If
                              
      cmb_Respon.ListIndex = gf_Busca_Arregl(l_arr_Respon, r_rst_Princi!CajChc_Respon) - 1
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

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_FecDia_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_TipCam)
   End If
End Sub

Private Sub ipp_FecDia_LostFocus()
Dim r_dbl_TipCam   As Double
   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecDia.Text, "yyyymmdd"), 1)
   ipp_TipCam.Text = Format(r_dbl_TipCam, "###,###,##0.000000") & " "
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FecDia)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_ImpAsig.Enabled = False Then
          Call gs_SetFocus(ipp_PagTarj)
       Else
          Call gs_SetFocus(ipp_ImpAsig)
       End If
   End If
End Sub

Private Sub ipp_ImpAsig_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_PagTarj)
   End If
End Sub

Private Sub ipp_PagTarj_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipDoc)
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Asigna)
   End If
End Sub

Private Sub cmb_Asigna_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Respon)
   End If
End Sub

Private Sub cmb_Respon_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   End If
End Sub
