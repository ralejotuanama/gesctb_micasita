VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TipCam_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   6435
   ClientTop       =   765
   ClientWidth     =   7440
   Icon            =   "GesCtb_frm_006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10185
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   17965
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
         Height          =   1095
         Left            =   30
         TabIndex        =   25
         Top             =   9030
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   1931
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
         Begin EditLib.fpDoubleSingle ipp_ValCom 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Top             =   390
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
         Begin EditLib.fpDoubleSingle ipp_ValVta 
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Top             =   720
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
            Left            =   1680
            TabIndex        =   4
            Top             =   60
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
         Begin VB.Label Label2 
            Caption         =   "Valor de Venta:"
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Valor de Compra:"
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   60
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            TabIndex        =   17
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Tipo de Cambio"
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
            Picture         =   "GesCtb_frm_006.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1425
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   2514
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
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   5655
         End
         Begin VB.ComboBox cmb_TipTip 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5655
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   720
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
            Left            =   1680
            TabIndex        =   3
            Top             =   1050
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
         Begin VB.Label Label6 
            Caption         =   "Fecha Inicio:"
            Height          =   285
            Left            =   60
            TabIndex        =   31
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   390
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Seleccione Tipo:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   6045
         Left            =   30
         TabIndex        =   21
         Top             =   2940
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            Height          =   5655
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   9975
            _Version        =   393216
            Rows            =   25
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2430
            TabIndex        =   22
            Top             =   60
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor de Compra"
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
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   4710
            TabIndex        =   24
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor de Ventas"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   29
         Top             =   780
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_006.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_006.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Ver Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_006.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_006.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_006.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   3030
            Picture         =   "GesCtb_frm_006.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_006.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_TipCam_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Moneda_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Moneda_Click
   End If
End Sub

Private Sub cmb_TipTip_Click()
   Call gs_SetFocus(cmb_Moneda)
End Sub

Private Sub cmb_TipTip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipTip_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(ipp_FecDia)
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipTip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipTip)
      Exit Sub
   End If
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call moddat_gs_FecSis
   ipp_FecDia.Text = Format(date, "dd/mm/yyyy")
   ipp_ValCom.Value = 0
   ipp_ValVta.Value = 0
   
   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_VerDet.Enabled = True
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Agrega_TipCam_EDPYMEBANK(CStr(cmb_TipTip.ItemData(cmb_TipTip.ListIndex)), Format(CDate(ipp_FecDia.Text), "yyyymmdd"), moddat_g_int_TipMon, CStr(ipp_ValCom.Value), CStr(ipp_ValVta.Value))
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_OPE_TIPCAM ("
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipTip.ItemData(cmb_TipTip.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecDia.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & moddat_g_int_TipMon & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValVta.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabaci�n de los datos. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
   
   grd_Listad.Enabled = False
   cmd_VerDet.Enabled = False
   cmd_Agrega.Enabled = False
   
   Call moddat_gs_FecSis
   ipp_FecDia.Text = Format(date, "dd/mm/yyyy")
   ipp_ValCom.Value = 0
   ipp_ValVta.Value = 0
   
   Call gs_SetFocus(ipp_FecDia)
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipTip.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   
   Call moddat_gs_FecSis
   ipp_FecDia.Text = Format(date, "dd/mm/yyyy")
   ipp_ValCom.Value = 0
   ipp_ValVta.Value = 0
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   
   Call gs_SetFocus(ipp_FecDia)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerDet_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Con_TipCam_02.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt & " - Tipos de Cambio"
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2360
   grd_Listad.ColWidth(1) = 2270
   grd_Listad.ColWidth(2) = 2270
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipTip, 1, "039")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   cmb_Moneda.RemoveItem (0)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipTip.Enabled = p_Activa
   cmb_Moneda.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_VerDet.Enabled = Not p_Activa
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerDet_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Activa_Editar(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   ipp_FecDia.Enabled = p_Activa
   ipp_ValCom.Enabled = p_Activa
   ipp_ValVta.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_VerDet.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_VerDet.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_int_TipMon = cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
   moddat_g_str_Moneda = cmb_Moneda.Text
   
   moddat_g_str_CodGrp = Format(cmb_TipTip.ItemData(cmb_TipTip.ListIndex), "00")
   moddat_g_str_DesGrp = cmb_TipTip.Text
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_TIPCAM WHERE "
   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = " & cmb_TipTip.ItemData(cmb_TipTip.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_TIPMON = " & moddat_g_int_TipMon & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCAM_FECDIA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!TIPCAM_FECDIA))
      
      grd_Listad.Col = 1
      grd_Listad.Text = Format(g_rst_Princi!TIPCAM_COMPRA, "###,###,##0.000000")
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!TIPCAM_VENTAS, "###,###,##0.000000")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_VerDet.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub ipp_FecDia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom)
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

Private Sub ipp_ValCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValVta)
      
      If cmb_TipTip.ItemData(cmb_TipTip.ListIndex) = 2 Then
         ipp_ValVta.Value = ipp_ValCom.Value
      End If
   End If
End Sub

Private Sub ipp_ValVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub fs_Agrega_TipCam_EDPYMEBANK(ByVal p_TipTip As String, ByVal p_FecDia As String, ByVal p_TipMon As String, ByVal p_ValCom As String, ByVal p_ValVta As String)
         
         
   g_str_Parame = "SELECT * FROM CALENDARIO WHERE "
   g_str_Parame = g_str_Parame & "FECHA = to_date ('" & gf_FormatoFecha(p_FecDia) & "', 'DD/MM/YYYY') "
   g_str_Parame = g_str_Parame & "ORDER BY FECHA DESC "
                  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CALENDARIO SET "
      
      If p_TipTip = 3 Then
         g_str_Parame = g_str_Parame & "VTA_DOL_PROM = " & CDbl(p_ValVta) & ","
         g_str_Parame = g_str_Parame & "CMP_DOL_PROM = " & CDbl(p_ValCom) & " "
      ElseIf p_TipTip = 2 Then
         g_str_Parame = g_str_Parame & "PROM_SBS =" & CDbl(p_ValCom) & " "
      End If
      
      g_str_Parame = g_str_Parame & "WHERE "
      g_str_Parame = g_str_Parame & "FECHA = to_date ('" & gf_FormatoFecha(p_FecDia) & "', 'DD/MM/YYYY') "
         
   Else
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CALENDARIO( "
      g_str_Parame = g_str_Parame & "FECHA, "
      
      If p_TipTip = 3 Then
         g_str_Parame = g_str_Parame & "VTA_DOL_PROM, "
         g_str_Parame = g_str_Parame & "CMP_DOL_PROM) "
      ElseIf p_TipTip = 2 Then
         g_str_Parame = g_str_Parame & "PROM_SBS)  "
      End If
      
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "to_date ('" & gf_FormatoFecha(p_FecDia) & "', 'DD/MM/YYYY'), "
                  
      If p_TipTip = 3 Then
         g_str_Parame = g_str_Parame & CDbl(p_ValVta) & ", "
         g_str_Parame = g_str_Parame & CDbl(p_ValCom) & ") "
      ElseIf p_TipTip = 2 Then
         g_str_Parame = g_str_Parame & CDbl(p_ValCom) & ")  "
      End If
   
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   DoEvents
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
End Sub

