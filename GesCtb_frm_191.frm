VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CajChc_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
   Icon            =   "GesCtb_frm_191.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8175
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   13605
      _Version        =   65536
      _ExtentX        =   23998
      _ExtentY        =   14420
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
         TabIndex        =   37
         Top             =   60
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
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
            TabIndex        =   38
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Detalle de Caja chica"
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
            Picture         =   "GesCtb_frm_191.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_DatPrv 
         Height          =   1185
         Left            =   60
         TabIndex        =   39
         Top             =   2630
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
         _ExtentY        =   2090
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
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   7020
            TabIndex        =   1
            Top             =   360
            Width           =   6180
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   3915
         End
         Begin VB.TextBox txt_Descrip 
            Height          =   315
            Left            =   7020
            MaxLength       =   60
            TabIndex        =   3
            Top             =   690
            Width           =   6180
         End
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1380
            TabIndex        =   2
            Top             =   690
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
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   5520
            TabIndex        =   44
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   420
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   5520
            TabIndex        =   41
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Datos del proveedor"
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
            TabIndex        =   40
            Top             =   60
            Width           =   1740
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   45
         Top             =   765
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
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
            Left            =   12720
            Picture         =   "GesCtb_frm_191.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_191.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_DatCbt 
         Height          =   1750
         Left            =   60
         TabIndex        =   46
         Top             =   3840
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
         _ExtentY        =   3087
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.23
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.TextBox txt_NumPrv 
            Height          =   315
            Left            =   7020
            MaxLength       =   7
            TabIndex        =   7
            Top             =   690
            Width           =   1515
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1350
            Width           =   2175
         End
         Begin VB.ComboBox cmb_TipCbtPrv 
            Height          =   315
            Left            =   7020
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   3915
         End
         Begin VB.TextBox txt_NumSeriePrv 
            Height          =   315
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   6
            Top             =   690
            Width           =   1515
         End
         Begin EditLib.fpDateTime ipp_FchVenc 
            Height          =   315
            Left            =   7020
            TabIndex        =   9
            Top             =   1020
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
         Begin EditLib.fpDateTime ipp_FchEmiPrv 
            Height          =   315
            Left            =   1380
            TabIndex        =   8
            Top             =   1020
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
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   315
            Left            =   7020
            TabIndex        =   11
            Top             =   1350
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
         Begin EditLib.fpDateTime ipp_FchCtb 
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            Top             =   360
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Contable:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   5520
            TabIndex        =   54
            Top             =   1410
            Width           =   930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1410
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vencimiento:"
            Height          =   195
            Left            =   5520
            TabIndex        =   52
            Top             =   1080
            Width           =   1410
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Numero:"
            Height          =   195
            Left            =   5520
            TabIndex        =   50
            Top             =   750
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Serie:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   750
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Datos del comprobante"
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
            TabIndex        =   48
            Top             =   30
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   5520
            TabIndex        =   47
            Top             =   420
            Width           =   990
         End
      End
      Begin Threed.SSPanel pnl_DatDet 
         Height          =   2385
         Left            =   60
         TabIndex        =   55
         Top             =   5640
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   4207
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
         Begin VB.ComboBox cmb_NGrvDH_02 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1320
            Width           =   645
         End
         Begin VB.ComboBox cmb_CtaNoGvd_02 
            Height          =   315
            Left            =   3570
            TabIndex        =   23
            Top             =   1320
            Width           =   4980
         End
         Begin VB.ComboBox cmb_CtaNoGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   20
            Top             =   990
            Width           =   4980
         End
         Begin VB.ComboBox cmb_CtaGvd_02 
            Height          =   315
            Left            =   3570
            TabIndex        =   17
            Top             =   660
            Width           =   4980
         End
         Begin VB.ComboBox cmb_CtaGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   14
            Top             =   330
            Width           =   4980
         End
         Begin VB.ComboBox cmb_NGrvDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   990
            Width           =   645
         End
         Begin VB.ComboBox cmb_GravDH_02 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   660
            Width           =   645
         End
         Begin VB.ComboBox cmb_GravDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   330
            Width           =   645
         End
         Begin EditLib.fpDoubleSingle ipp_ImpGrav_01 
            Height          =   315
            Left            =   1380
            TabIndex        =   12
            Top             =   330
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
         Begin EditLib.fpDoubleSingle ipp_ImpGrav_02 
            Height          =   315
            Left            =   1380
            TabIndex        =   15
            Top             =   660
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
         Begin EditLib.fpDoubleSingle ipp_ImpNGrv_01 
            Height          =   315
            Left            =   1380
            TabIndex        =   18
            Top             =   990
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
         Begin Threed.SSPanel pnl_Igv 
            Height          =   315
            Left            =   3570
            TabIndex        =   26
            Top             =   1650
            Width           =   2670
            _Version        =   65536
            _ExtentX        =   4710
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
         Begin Threed.SSPanel pnl_ImpIgv 
            Height          =   315
            Left            =   1380
            TabIndex        =   24
            Top             =   1650
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
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
         Begin Threed.SSPanel pnl_ImpPpg 
            Height          =   315
            Left            =   1380
            TabIndex        =   27
            Top             =   1980
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_IgvDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   25
            Top             =   1650
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "H"
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
         Begin Threed.SSPanel pnl_PpgDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   28
            Top             =   1980
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
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
         Begin Threed.SSPanel pnl_PorPagar 
            Height          =   315
            Left            =   3570
            TabIndex        =   29
            Top             =   1980
            Width           =   2670
            _Version        =   65536
            _ExtentX        =   4710
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
         Begin EditLib.fpDoubleSingle ipp_ImpNGrv_02 
            Height          =   315
            Left            =   1380
            TabIndex        =   21
            Top             =   1320
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "No Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CUENTAS CONTABLES"
            Height          =   195
            Left            =   5190
            TabIndex        =   64
            Top             =   90
            Width           =   1770
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "D/H"
            Height          =   195
            Left            =   3060
            TabIndex        =   63
            Top             =   90
            Width           =   315
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Por Pagar:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   2010
            Width           =   750
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   390
            Width           =   660
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
            TabIndex        =   59
            Top             =   60
            Width           =   1230
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "No Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1050
            Width           =   915
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1680
            Width           =   270
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "IMPORTE"
            Height          =   195
            Left            =   1860
            TabIndex        =   56
            Top             =   90
            Width           =   735
         End
      End
      Begin Threed.SSPanel pnl_DatFin 
         Height          =   2385
         Left            =   8760
         TabIndex        =   65
         Top             =   5640
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   4207
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
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1170
            Width           =   2880
         End
         Begin VB.TextBox txt_CtrCosto 
            Height          =   315
            Left            =   1590
            MaxLength       =   18
            TabIndex        =   31
            Top             =   840
            Width           =   2880
         End
         Begin VB.ComboBox cmb_CatCtb 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   540
            Width           =   2880
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1500
            Width           =   2880
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   90
            TabIndex        =   70
            Top             =   1230
            Width           =   510
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Datos Financieros"
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
            Left            =   90
            TabIndex        =   69
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Categoría Contable:"
            Height          =   195
            Left            =   90
            TabIndex        =   68
            Top             =   570
            Width           =   1425
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Costos:"
            Height          =   195
            Left            =   90
            TabIndex        =   67
            Top             =   900
            Width           =   1260
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   90
            TabIndex        =   66
            Top             =   1560
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1130
         Left            =   60
         TabIndex        =   71
         Top             =   1470
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
         _ExtentY        =   1993
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.29
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_NumCaja 
            Height          =   315
            Left            =   1380
            TabIndex        =   72
            Top             =   360
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
         Begin Threed.SSPanel pnl_FechaCaja 
            Height          =   315
            Left            =   1380
            TabIndex        =   77
            Top             =   690
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
         Begin Threed.SSPanel pnl_Respon 
            Height          =   315
            Left            =   7020
            TabIndex        =   78
            Top             =   360
            Width           =   6180
            _Version        =   65536
            _ExtentX        =   10901
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   7020
            TabIndex        =   79
            Top             =   690
            Width           =   6180
            _Version        =   65536
            _ExtentX        =   10901
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   5520
            TabIndex        =   80
            Top             =   750
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos de la Caja"
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
            TabIndex        =   76
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label lbl_Numero 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Caja:"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label lbl_Fecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Caja:"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   5520
            TabIndex        =   73
            Top             =   390
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CajChc_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_IniFrm        As Boolean
Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_DebHab()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_int_Contar        As Integer
Dim l_int_TopNiv        As Integer
Dim l_dbl_IGV           As Double
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer
Dim l_str_CtaIGV        As String

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

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_CatCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_CtrCosto)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_CtaGvd_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpGrav_02)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaGvd_01_LostFocus()
'   pnl_Igv.Caption = Mid(cmb_CtaGvd_01.Text, 1, l_int_TopNiv)
   Call fs_Cuenta_IGV
End Sub

Private Sub cmb_CtaGvd_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpNGrv_01)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaGvd_02_LostFocus()
   Call fs_Cuenta_IGV
End Sub

Private Sub cmb_CtaNoGvd_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpNGrv_02)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaNoGvd_01_LostFocus()
   Call fs_Cuenta_IGV
End Sub

Private Sub cmb_CtaNoGvd_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CatCtb)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaNoGvd_02_LostFocus()
   Call fs_Cuenta_IGV
End Sub

Private Sub cmb_GravDH_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaGvd_01)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_GravDH_01_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_GravDH_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaGvd_02)
   End If
End Sub

Private Sub cmb_GravDH_02_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_Moneda_Click()
   Call fs_CargarBancos
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpGrav_01)
   End If
End Sub

Private Sub cmb_NGrvDH_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaNoGvd_01)
   End If
End Sub

Private Sub cmb_NGrvDH_01_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_NGrvDH_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaNoGvd_02)
   End If
End Sub

Private Sub cmb_NGrvDH_02_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_Proveedor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       'Call fs_Buscar_prov
       Call gs_SetFocus(txt_Descrip)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Proveedor_LostFocus()
   If l_dbl_IniFrm = True Then
      Call fs_Buscar_prov
   End If
End Sub

Private Sub cmb_TipCbtPrv_Click()
   Call fs_ActivaRefer
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_TipCbtPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumSeriePrv)
   End If
End Sub

Private Sub cmb_TipCbtPrv_LostFocus()
   Call fs_ActivaRefer
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_ImpAux   As Double
Dim r_bol_Estado   As Boolean
   
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
      
   If (cmb_TipCbtPrv.ListIndex = -1) Then
       MsgBox "Tiene que seleccionar un tipo de comprobante.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipCbtPrv)
       Exit Sub
   End If
   
   If (Len(Trim(txt_NumSeriePrv.Text)) <> 4) Then
       MsgBox "El numero de serie consta de 4 digitos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumSeriePrv)
       Exit Sub
   End If
   
   If (Len(Trim(txt_NumPrv.Text)) <> 7) Then
       MsgBox "El numero del comprobante consta de 7 digitos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumPrv)
       Exit Sub
   End If
   
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 14) Then
       If (ipp_FchVenc.Text = "") Then
           MsgBox "La fecha de vencimiento es obligatoria cuando el comprobante es RECIBOS.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FchVenc)
           Exit Sub
       End If
   End If
   
   If CDbl(pnl_TipCambio.Caption) = 0 Then
      MsgBox "El tipo de cambio no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FchEmiPrv)
      Exit Sub
   End If
   
   If (cmb_CatCtb.ListIndex = -1) Then
       MsgBox "Seleccione una categoria contable, en el grupo datos financieros.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_CatCtb)
       Exit Sub
   End If
    
   If (Format(ipp_FchEmiPrv.Text, "yyyymmdd") > Format(moddat_g_str_FecSis, "yyyymmdd")) Then
       MsgBox "Esta intentando registrar un documento con una fecha futura.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FchEmiPrv)
       Exit Sub
   End If

   If (ipp_FchVenc.Text <> "") Then
       If (Format(ipp_FchVenc.Text, "yyyymmdd") < Format(ipp_FchEmiPrv.Text, "yyyymmdd")) Then
           MsgBox "La fecha de vencimiento no puede ser menor al fecha de emisión.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FchVenc)
           Exit Sub
       End If
   End If
   
   If CLng(Format(ipp_FchCtb.Text, "yyyymmdd")) < CLng(Format(ipp_FchEmiPrv.Text, "yyyymmdd")) Then
       MsgBox "La fecha de emisión no puede ser mayor a la fecha contable.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FchCtb)
       Exit Sub
   End If
   '--------------------------------------------
   If CDbl(ipp_ImpGrav_01.Text) > 0 And Len(Trim(cmb_CtaGvd_01.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable del gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaGvd_01)
      Exit Sub
   End If
   If CDbl(ipp_ImpGrav_01.Text) = 0 And Len(Trim(cmb_CtaGvd_01.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe del gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpGrav_01)
      Exit Sub
   End If
   
   If CDbl(ipp_ImpGrav_02.Text) > 0 And Len(Trim(cmb_CtaGvd_02.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable del gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaGvd_02)
      Exit Sub
   End If
   If CDbl(ipp_ImpGrav_02.Text) = 0 And Len(Trim(cmb_CtaGvd_02.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe del gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpGrav_02)
      Exit Sub
   End If
   
   If CDbl(ipp_ImpNGrv_01.Text) > 0 And Len(Trim(cmb_CtaNoGvd_01.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable del no gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaNoGvd_01)
      Exit Sub
   End If
   If CDbl(ipp_ImpNGrv_01.Text) = 0 And Len(Trim(cmb_CtaNoGvd_01.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe del no gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpNGrv_01)
      Exit Sub
   End If
   
   If CDbl(ipp_ImpNGrv_02.Text) > 0 And Len(Trim(cmb_CtaNoGvd_02.Text)) = 0 Then
      MsgBox "En el grupo determinación es obligatorio la cuenta contable del no gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaNoGvd_02)
      Exit Sub
   End If
   If CDbl(ipp_ImpNGrv_02.Text) = 0 And Len(Trim(cmb_CtaNoGvd_02.Text)) > 0 Then
      MsgBox "En el grupo determinación es obligatorio el importe del no gravado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpNGrv_02)
      Exit Sub
   End If
   '---------------VALIDAR EXISTENCIA DE LA CUENTA-----------------------------
   If Len(Trim(cmb_CtaGvd_01.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaGvd_01.Text) = False Then
         Call gs_SetFocus(cmb_CtaGvd_01)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaGvd_02.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaGvd_02.Text) = False Then
         Call gs_SetFocus(cmb_CtaGvd_02)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaNoGvd_01.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaNoGvd_01.Text) = False Then
         Call gs_SetFocus(cmb_CtaNoGvd_01)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaNoGvd_02.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaNoGvd_02.Text) = False Then
         Call gs_SetFocus(cmb_CtaNoGvd_02)
         Exit Sub
      End If
   End If
   '--------------------------------------------
'   If (Format(ipp_FchCtb.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FchCtb.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FchCtb)
'       Exit Sub
'   End If

   '--ipp_FchCtb.Text
'   If Format(moddat_g_str_FecSis, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > modctb_int_PerAno & Format(modctb_int_PerMes, "00") & Format(moddat_g_int_PerLim, "00")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FchCtb)
'          Exit Sub
'      End If
'      MsgBox "Los asiento a generar perteneceran al periodo anterior.", vbExclamation, modgen_g_str_NomPlt
'   Else
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FchCtb)
'          Exit Sub
'      End If
'   End If

   If fs_ValidaPeriodo(ipp_FchCtb.Text) = False Then
      Exit Sub
   End If

   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Grabar
    
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim r_str_Msg   As String

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   l_dbl_IniFrm = False
   r_str_Msg = ""
   
   Call fs_Inicia
   
   If moddat_g_int_TipEva = 1 Then
      'caja chica
      r_str_Msg = "Caja Chica"
   ElseIf moddat_g_int_TipEva = 6 Then
      'tarjeta credito
      r_str_Msg = "Tarjeta de Credito"
   End If
   
   If moddat_g_int_FlgGrb = 1 Then 'INSERT
      pnl_Titulo.Caption = "Detalle de " & r_str_Msg & " - Adicionar"
      Call fs_Limpiar
   ElseIf moddat_g_int_FlgGrb = 2 Then 'UPDATE
      pnl_Titulo.Caption = "Detalle de " & r_str_Msg & " - Modificar"
      If (Len(Trim(moddat_g_str_Codigo)) > 0) Then
          Call fs_Limpiar
          Call fs_Cargar_Datos
      End If
   ElseIf moddat_g_int_FlgGrb = 0 Then 'CONSULTAR
      pnl_Titulo.Caption = "Detalle de " & r_str_Msg & " - Consultar"
      If (Len(Trim(moddat_g_str_Codigo)) > 0) Then
          cmd_Grabar.Visible = False
          Call fs_Limpiar
          Call fs_Cargar_Datos
          Call fs_Desabilitar
      End If
      cmb_Moneda.Enabled = False
   End If
   
   If moddat_g_int_TipEva = 6 And (moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2) Then
      cmb_Moneda.Enabled = True
   Else
      cmb_Moneda.Enabled = False
   End If
          
   l_dbl_IniFrm = True
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call fs_CargaMntPardes(cmb_TipDoc, "118")
   Call fs_CargaMntPardes(cmb_CatCtb, "124")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCbtPrv, 1, "123")
   'Call fs_CargaMntPardes(cmb_TipCbtPrv, "123")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   
   'cargar deba y haber
   Call moddat_gs_Carga_LisIte(cmb_GravDH_01, l_arr_DebHab, 1, 255, 1)
   cmb_GravDH_01.Clear
   For l_int_Contar = 1 To UBound(l_arr_DebHab)
       cmb_GravDH_01.AddItem Left(l_arr_DebHab(l_int_Contar).Genera_Nombre, 1)
       cmb_GravDH_02.AddItem Left(l_arr_DebHab(l_int_Contar).Genera_Nombre, 1)
       cmb_NGrvDH_01.AddItem Left(l_arr_DebHab(l_int_Contar).Genera_Nombre, 1)
       cmb_NGrvDH_02.AddItem Left(l_arr_DebHab(l_int_Contar).Genera_Nombre, 1)
   Next
      
   'cargar las cuentas contables
   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaGvd_01, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   cmb_CtaGvd_01.Clear
   cmb_CtaGvd_02.Clear
   cmb_CtaNoGvd_01.Clear
   cmb_CtaNoGvd_02.Clear
   For l_int_Contar = 1 To UBound(l_arr_CtaCtb)
       Select Case l_arr_CtaCtb(l_int_Contar).Genera_Codigo
              'Case "451109090101", "451301010102", "451301010105", "451301020108", "451301010201", _
              '     "451301110101", "451301110102", "451301110106", "451301120101", "451301120103", _
              '     "451301130109", "451301290102", "451301290110", "451109060105", "451109060112", _
              '     "451301130101", "191807020101", "291807010112", "451301020102", "451301290109", _
              '     "451301130102", "451301140101", "451109051301", "451109020102", "451301050101", _
              '     "451109020101", "251419010109", "451203020101", "451301020106", "251602010106", _
              '     "451109051301", "451109051302", "451109051303" 'Fotochecks, Publicidad, Uniformes
              Case "191807020101", "291807010112", "251419010109", "251602010106", "151719010104"
                   cmb_CtaGvd_01.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                   cmb_CtaGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                   cmb_CtaNoGvd_01.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                   cmb_CtaNoGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
              Case Else
                   If Mid(l_arr_CtaCtb(l_int_Contar).Genera_Codigo, 1, 2) = "45" Then
                      cmb_CtaGvd_01.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                      cmb_CtaGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                      cmb_CtaNoGvd_01.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                      cmb_CtaNoGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
                   End If
       End Select
   Next

   'cargar igv
   l_dbl_IGV = moddat_gf_Consulta_ParVal("001", "001") 'IGV
   l_dbl_IGV = l_dbl_IGV / 100
   Call moddat_gs_FecSis
                  
   pnl_NumCaja.Caption = Format(moddat_g_str_Codigo, "0000000000")
   pnl_Moneda.Caption = Trim(moddat_g_str_DesMod)
   pnl_FechaCaja.Caption = Trim(moddat_g_str_FecIng)
   pnl_Respon.Caption = Trim(moddat_g_str_Descri)
   
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
End Sub

Private Sub fs_CargaMntPardes(p_Combo As ComboBox, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES A "
   g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   If Trim(p_CodGrp) = "118" Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('009999','000006','000001') "
   End If
   If Trim(p_CodGrp) = "123" Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000002','000005','000012','000014','009999','000088') "
   End If
   If Trim(p_CodGrp) = "124" Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000005') "
   End If
   g_str_Parame = g_str_Parame & "   AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC "
   
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
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!PARDES_CODITE)
            
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Limpiar()
   'DATOS DEL PROVEEDOR
   Call gs_BuscarCombo_Item(cmb_TipDoc, 6)
   cmb_Proveedor.Text = ""
   pnl_Codigo.Caption = ""
   txt_Descrip.Text = ""
   'DATOS DEL COMPROBANTE
   ipp_FchCtb.DateMax = modctb_str_FecFin
   ipp_FchCtb.DateMin = modctb_str_FecIni
   If (Format(moddat_g_str_FecSis, "yyyymmdd") <= Format(modctb_str_FecFin, "yyyymmdd")) Then
       ipp_FchCtb.Text = moddat_g_str_FecSis
   Else
       ipp_FchCtb.Text = modctb_str_FecFin
   End If
   cmb_TipCbtPrv.ListIndex = 0
   txt_NumSeriePrv.Text = ""
   txt_NumPrv.Text = ""
   ipp_FchEmiPrv.Text = moddat_g_str_FecSis
   pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(3, 2, Format(ipp_FchEmiPrv.Text, "yyyymmdd"), 1)
   pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
       
   Call ipp_FchEmiPrv_LostFocus
   
   ipp_FchVenc.Text = ""
   Call gs_BuscarCombo_Item(cmb_Moneda, CInt(moddat_g_str_CodMod))
   'DETERMINACION
   cmb_CtaGvd_01.Text = ""
   cmb_CtaGvd_02.Text = ""
   cmb_CtaNoGvd_01.Text = ""
   cmb_CtaNoGvd_02.Text = ""
   pnl_Igv.Caption = ""
   pnl_PorPagar.Caption = ""
   
   ipp_ImpGrav_01.Text = "0.00"
   ipp_ImpGrav_02.Text = "0.00"
   ipp_ImpNGrv_01.Text = "0.00"
   ipp_ImpNGrv_02.Text = "0.00"
   pnl_ImpIgv.Caption = "0.00 "
   pnl_ImpPpg.Caption = "0.00 "
   
   'DATOS FINANCIEROS
   cmb_CatCtb.ListIndex = -1
   txt_CtrCosto.Text = ""
   cmb_Banco.ListIndex = -1
   cmb_CtaCte.ListIndex = -1
   
   'DATOS STATICOS
   l_str_CtaIGV = "251703020101"
   pnl_Igv.Caption = ""
   
   If moddat_g_int_TipEva = 1 Then
      'caja chica
      pnl_PorPagar.Caption = "111701010101"
   ElseIf moddat_g_int_TipEva = 6 Then
      'tarjeta credito
      pnl_PorPagar.Caption = "291807010117"
   End If
   
   Call ipp_FchEmiPrv_LostFocus
   l_dbl_IniFrm = True
   Call fs_Calcular_Determ 'no olvidar activar
   l_dbl_IniFrm = False
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub fs_Desabilitar()
   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   txt_Descrip.Enabled = False
   
   ipp_FchCtb.Enabled = False
   cmb_TipCbtPrv.Enabled = False
   txt_NumSeriePrv.Enabled = False
   txt_NumPrv.Enabled = False
   ipp_FchEmiPrv.Enabled = False
   ipp_FchVenc.Enabled = False
   cmb_Moneda.Enabled = False
   
   cmb_CtaGvd_01.Enabled = False
   cmb_CtaGvd_02.Enabled = False
   cmb_CtaNoGvd_01.Enabled = False
   cmb_CtaNoGvd_02.Enabled = False
   
   ipp_ImpGrav_01.Enabled = False
   ipp_ImpGrav_02.Enabled = False
   ipp_ImpNGrv_01.Enabled = False
   ipp_ImpNGrv_02.Enabled = False
   cmb_GravDH_01.Enabled = False
   cmb_GravDH_02.Enabled = False
   cmb_NGrvDH_01.Enabled = False
   cmb_NGrvDH_02.Enabled = False
      
   cmb_CatCtb.Enabled = False
   txt_CtrCosto.Enabled = False
   cmb_Banco.Enabled = False
   cmb_CtaCte.Enabled = False
End Sub

Private Sub fs_Grabar()
Dim r_dbl_ImpDeb As Double
Dim r_dbl_ImpHab As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_DET ( "
   g_str_Parame = g_str_Parame & IIf(Trim(pnl_Codigo.Caption) = "", "Null", Trim(pnl_Codigo.Caption)) & ", " 'CAJDET_CODDET
   g_str_Parame = g_str_Parame & CLng(pnl_NumCaja.Caption) & ", "  'CAJDET_CODCAJ
   
   If moddat_g_int_TipEva = 1 Then
      'CAJA CHICA
      g_str_Parame = g_str_Parame & "1, "  'CAJDET_TIPTAB
   ElseIf moddat_g_int_TipEva = 6 Then
      'TARJETA DE CREDITO
      g_str_Parame = g_str_Parame & "6, "  'CAJDET_TIPTAB
   End If
   
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'CAJDET_TIPDOC
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', " 'CAJDET_NUMDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descrip.Text) & "', " 'CAJDET_DESCRP
   g_str_Parame = g_str_Parame & Format(ipp_FchCtb.Text, "yyyymmdd") & ", " 'CAJDET_FECCTB
   g_str_Parame = g_str_Parame & cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) & ", " 'CAJDET_TIPCPB
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumSeriePrv.Text) & "', " 'CAJDET_NSERIE
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumPrv.Text) & "', " 'CAJDET_NROCOM
   g_str_Parame = g_str_Parame & Format(ipp_FchEmiPrv.Text, "yyyymmdd") & ", " 'CAJDET_FECEMI
   g_str_Parame = g_str_Parame & IIf(ipp_FchVenc.Text = "", "Null", Format(ipp_FchVenc.Text, "yyyymmdd")) & ", " 'CAJDET_FECVTO
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'CAJDET_CODMON
   g_str_Parame = g_str_Parame & Trim(pnl_TipCambio.Caption) & ", " 'CAJDET_TIPCAM
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaGvd_01.Text = "", "", Mid(cmb_CtaGvd_01.Text, 1, l_int_TopNiv)) & "', " 'CAJDET_CNT_GRV1
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaGvd_02.Text = "", "", Mid(cmb_CtaGvd_02.Text, 1, l_int_TopNiv)) & "', " 'CAJDET_CNT_GRV2
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaNoGvd_01.Text = "", "", Mid(cmb_CtaNoGvd_01.Text, 1, l_int_TopNiv)) & "', " 'CAJDET_CNT_NGV1
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaNoGvd_02.Text = "", "", Mid(cmb_CtaNoGvd_02.Text, 1, l_int_TopNiv)) & "', " 'CAJDET_CNT_NGV2
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_Igv.Caption) & "', " 'CAJDET_CNT_IGV1
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_PorPagar.Caption) & "', " 'CAJDET_CNT_PPG1
   
   'GRAVADO 1
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_GravDH_01.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpGrav_01.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpGrav_01.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'CAJDET_DEB_GRV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'CAJDET_HAB_GRV1
   'GRAVADO 2
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_GravDH_02.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpGrav_02.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpGrav_02.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'CAJDET_DEB_GRV2
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'CAJDET_HAB_GRV2
   'NO GRAVADO 1
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_NGrvDH_01.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpNGrv_01.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpNGrv_01.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'CAJDET_DEB_NGV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'CAJDET_HAB_NGV1
   'NO GRAVADO 2
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_NGrvDH_02.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpNGrv_02.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpNGrv_02.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'CAJDET_DEB_NGV2
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_NGV2
   'IGV
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_IgvDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpIgv.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpIgv.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'CAJDET_DEB_IGV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_IGV1
   'POR PAGAR
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_PpgDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpPpg.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpPpg.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'CAJDET_DEB_PPG1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_PPG1
   If (cmb_CatCtb.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "Null, " 'CAJDET_CATCTB
   Else
       g_str_Parame = g_str_Parame & cmb_CatCtb.ItemData(cmb_CatCtb.ListIndex) & ", " 'CAJDET_CATCTB
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtrCosto.Text) & "', " 'CAJDET_CNTCST
   If (cmb_Banco.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "Null , " 'CAJDET_CODBNC
   Else
       g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", " 'CAJDET_CODBNC
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', " 'CAJDET_CTACRR
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") " 'as_insupd
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CajChc_03.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CajChc_03.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 3) Then
       MsgBox "El total del comprobante excede al total de la caja asignada, la diferencia es de: " & Format(g_rst_Genera!TOTDIF, "###,###,##0.00"), vbExclamation, modgen_g_str_NomPlt
       Screen.MousePointer = 0
   End If
End Sub

Private Sub fs_Cargar_Datos()
Dim r_dbl_Import     As Double

   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   
   r_dbl_Import = 0
   Call gs_SetFocus(cmb_TipDoc)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJDET_CODDET, A.CAJDET_TIPDOC, A.CAJDET_NUMDOC, A.CAJDET_DESCRP, A.CAJDET_TIPCPB, A.CAJDET_NSERIE, "
   g_str_Parame = g_str_Parame & "        A.CAJDET_NROCOM, A.CAJDET_FECEMI, A.CAJDET_FECVTO, A.CAJDET_CODMON, A.CAJDET_TIPCAM, CAJDET_FECCTB, "
   g_str_Parame = g_str_Parame & "        A.CAJDET_CNT_GRV1, A.CAJDET_CNT_GRV2, A.CAJDET_CNT_NGV1, A.CAJDET_CNT_NGV2, A.CAJDET_CNT_IGV1, A.CAJDET_CNT_PPG1, "
   g_str_Parame = g_str_Parame & "        CAJDET_DEB_GRV1, CAJDET_HAB_GRV1, CAJDET_DEB_GRV2, CAJDET_HAB_GRV2, CAJDET_DEB_NGV1, CAJDET_HAB_NGV1, "
   g_str_Parame = g_str_Parame & "        CAJDET_DEB_NGV2, CAJDET_HAB_NGV2, CAJDET_DEB_IGV1, CAJDET_HAB_IGV1, CAJDET_DEB_PPG1, CAJDET_HAB_PPG1, "
   g_str_Parame = g_str_Parame & "        CAJDET_CATCTB , CAJDET_CNTCST, CAJDET_CODBNC, CAJDET_CTACRR, B.maeprv_RazSoc "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC_DET A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON A.CAJDET_TIPDOC = B.MAEPRV_TIPDOC AND A.CAJDET_NUMDOC = B.MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJDET_CODDET = " & moddat_g_str_CodIte
   g_str_Parame = g_str_Parame & "    AND A.CAJDET_CODCAJ = " & moddat_g_str_Codigo
   If moddat_g_int_TipEva = 1 Then
      'CAJA CHICA
      g_str_Parame = g_str_Parame & "    AND A.CAJDET_TIPTAB = 1 "
   ElseIf moddat_g_int_TipEva = 6 Then
      'TARJETA DE CREDITO
      g_str_Parame = g_str_Parame & "    AND A.CAJDET_TIPTAB = 6 "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!CAJDET_TipDoc)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!CajDet_NumDoc & "", 0)
      
      Call fs_Buscar_prov
           
      pnl_Codigo.Caption = Trim(g_rst_Princi!CajDet_CodDet & "")
      txt_Descrip.Text = Trim(g_rst_Princi!CAJDET_DESCRP & "")
      If Not IsNull(g_rst_Princi!CAJDET_FECCTB) Then
          ipp_FchCtb.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FECCTB)
      End If
      If Not IsNull(g_rst_Princi!CajDet_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtPrv, g_rst_Princi!CajDet_TipCpb)
         'Call fs_ActivaRefer
      End If
      
      txt_NumSeriePrv.Text = Trim(g_rst_Princi!CajDet_Nserie & "")
      txt_NumPrv.Text = Trim(g_rst_Princi!CajDet_NroCom & "")
      ipp_FchEmiPrv.Text = gf_FormatoFecha(g_rst_Princi!CajDet_FecEmi)
      If Not IsNull(g_rst_Princi!CAJDET_FecVto) Then
         ipp_FchVenc.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FecVto)
      End If
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!CAJDET_CODMON)
      pnl_TipCambio.Caption = Format(g_rst_Princi!CAJDET_TipCam, "###,###,##0.000000") & " "
      
      cmb_CtaGvd_01.ListIndex = fs_ComboIndex(cmb_CtaGvd_01, g_rst_Princi!CAJDET_Cnt_Grv1 & "", l_int_TopNiv)
      cmb_CtaGvd_02.ListIndex = fs_ComboIndex(cmb_CtaGvd_02, g_rst_Princi!CAJDET_Cnt_Grv2 & "", l_int_TopNiv)
      cmb_CtaNoGvd_01.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_01, g_rst_Princi!CAJDET_Cnt_Ngv1 & "", l_int_TopNiv)
      cmb_CtaNoGvd_02.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_02, g_rst_Princi!CAJDET_Cnt_Ngv2 & "", l_int_TopNiv)
      pnl_Igv.Caption = Trim(g_rst_Princi!CAJDET_Cnt_Igv1 & "")
      pnl_PorPagar.Caption = Trim(g_rst_Princi!CAJDET_Cnt_Ppg1 & "")
      
      'GRAVADO 1
      If (g_rst_Princi!CAJDET_Deb_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 0
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!CAJDET_Deb_Grv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!CAJDET_Hab_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 1
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!CAJDET_Hab_Grv1, "###,###,##0.00")
      End If
      'GRAVADO 2
      If (g_rst_Princi!CAJDET_Deb_Grv2 > 0) Then
          cmb_GravDH_02.ListIndex = 0
          ipp_ImpGrav_02.Text = Format(g_rst_Princi!CAJDET_Deb_Grv2, "###,###,##0.00")
      End If
      If (g_rst_Princi!CAJDET_Hab_Grv2 > 0) Then
          cmb_GravDH_02.ListIndex = 1
          ipp_ImpGrav_02.Text = Format(g_rst_Princi!CAJDET_Hab_Grv2, "###,###,##0.00")
      End If
      'NO GRAVADO 1
      If (g_rst_Princi!CAJDET_Deb_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 0
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!CAJDET_Deb_Ngv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!CAJDET_Hab_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 1
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!CAJDET_Hab_Ngv1, "###,###,##0.00")
      End If
      'NO GRAVADO 2
      If (g_rst_Princi!CAJDET_Deb_Ngv2 > 0) Then
          cmb_NGrvDH_02.ListIndex = 0
          ipp_ImpNGrv_02.Text = Format(g_rst_Princi!CAJDET_Deb_Ngv2, "###,###,##0.00")
      End If
      If (g_rst_Princi!CAJDET_Hab_Ngv2 > 0) Then
          cmb_NGrvDH_02.ListIndex = 1
          ipp_ImpNGrv_02.Text = Format(g_rst_Princi!CAJDET_Hab_Ngv2, "###,###,##0.00")
      End If
      'IGV
      If (g_rst_Princi!CAJDET_Deb_Igv1 > 0) Then
          pnl_IgvDH.Caption = "D"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!CAJDET_Deb_Igv1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!CAJDET_Hab_Igv1 > 0) Then
          pnl_IgvDH.Caption = "H"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!CAJDET_Hab_Igv1, "###,###,##0.00") & " "
      End If
      'CUENTAS POR PAGAR
      If (g_rst_Princi!CAJDET_Deb_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "D"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!CAJDET_Deb_Ppg1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!CAJDET_Hab_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "H"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!CAJDET_Hab_Ppg1, "###,###,##0.00") & " "
      End If
      '---------------------------------------------------------------------------------------
      If Not IsNull(g_rst_Princi!CAJDET_CatCtb) Then
         Call gs_BuscarCombo_Item(cmb_CatCtb, g_rst_Princi!CAJDET_CatCtb)
      End If
      txt_CtrCosto.Text = Trim(g_rst_Princi!CAJDET_CNTCST & "")
      
      If Not IsNull(g_rst_Princi!CAJDET_CODBNC) Then
         Call gs_BuscarCombo_Item(cmb_Banco, g_rst_Princi!CAJDET_CODBNC)
      End If
      If Not IsNull(g_rst_Princi!CAJDET_CTACRR) Then
         Call gs_BuscarCombo_Text(cmb_CtaCte, g_rst_Princi!CAJDET_CTACRR, -1)
         'cmb_CtaCte.Text = g_rst_Princi!CAJDET_CTACRR
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_prov()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   'pnl_Padron.Caption = ""
   'pnl_Padron.Tag = ""
   r_str_NumDoc = ""
   
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc)
          Exit Sub
       End If
       If cmb_Proveedor.ListIndex = -1 Then
          'MsgBox "Debe de seleccionar un proveedor.", vbExclamation, modgen_g_str_NomPlt
          'Call gs_SetFocus(cmb_Proveedor)
          Exit Sub
       End If
      
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, A.MAEPRV_PADRN1, "
   g_str_Parame = g_str_Parame & "        B.PARDES_DESCRI AS PADRON_1, A.MAEPRV_CODBNC_MN1, A.MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN1, A.MAEPRV_CODBNC_MN2, A.MAEPRV_CTACRR_MN2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN2, A.MAEPRV_CODBNC_MN3, A.MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN3, A.MAEPRV_CODBNC_DL1, A.MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL1, A.MAEPRV_CODBNC_DL2, A.MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL2, A.MAEPRV_CODBNC_DL3, A.MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 121 AND B.PARDES_CODITE = A.MAEPRV_PADRN1 "
   If (moddat_g_int_FlgGrb = 1) Then
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   Else
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & moddat_g_str_TipDoc
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(moddat_g_str_NumDoc) & "' "
   End If
   
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
       
       'txt_NumDoc.Tag = Trim(g_rst_Princi!maeprv_numdoc & "")
       'pnl_RazonSoc.Caption = Trim(g_rst_Princi!maeprv_Razsoc & "")
       'pnl_Padron.Caption = Trim(g_rst_GenAux!PADRON_1 & "")
       'pnl_Padron.Tag = Trim(g_rst_GenAux!MAEPRV_PADRN1 & "")
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

Private Function fs_NumDoc(p_Cadena As String) As String
   fs_NumDoc = ""
   If (cmb_TipDoc.ListIndex > -1) Then
      If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
          'fs_NumDoc = Mid(p_Cadena, 1, 12)
           fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
      End If
   End If
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

Private Function fs_ComboIndex(p_Combo As ComboBox, Cadena As String, p_Tipo As Integer) As Integer
Dim r_int_Contad As Integer
Dim r_int_Largo  As Integer
   
   fs_ComboIndex = -1
   For r_int_Contad = 0 To p_Combo.ListCount - 1
       'If Trim(Cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, r_int_Largo)) Then
        If Trim(Cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, InStr(Trim(p_Combo.List(r_int_Contad)), "-") - 1)) Then
          fs_ComboIndex = r_int_Contad
          Exit For
       End If
   Next
End Function

Private Sub fs_ActivaRefer()
  If cmb_TipCbtPrv.ListIndex > -1 Then
     If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88) Then '(07-N/C) - (88-DEVOLUCIONES)
         cmb_GravDH_01.ListIndex = 1
         cmb_GravDH_02.ListIndex = 1
         cmb_NGrvDH_01.ListIndex = 1
         cmb_NGrvDH_02.ListIndex = 1
         pnl_IgvDH.Caption = "H"
         pnl_PpgDH.Caption = "D"
     Else
         cmb_GravDH_01.ListIndex = 0
         cmb_GravDH_02.ListIndex = 0
         cmb_NGrvDH_01.ListIndex = 0
         cmb_NGrvDH_02.ListIndex = 0
         pnl_IgvDH.Caption = "D"
         pnl_PpgDH.Caption = "H"
     End If
  End If
End Sub

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

Private Sub ipp_FchCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipCbtPrv)
   End If
End Sub

Private Sub ipp_FchCtb_LostFocus()
   If (Format(ipp_FchCtb.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(ipp_FchCtb.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub ipp_FchEmiPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchVenc)
   End If
End Sub

Private Sub ipp_FchEmiPrv_LostFocus()
   If (l_dbl_IniFrm = True) Then
       If (Format(ipp_FchEmiPrv.Text, "yyyymmdd") > Format(moddat_g_str_FecSis, "yyyymmdd")) Then
           MsgBox "Esta intentando registrar un documento con una fecha futura.", vbExclamation, modgen_g_str_NomPlt
       ElseIf (Format(ipp_FchEmiPrv.Text, "yyyy") <> Format(modctb_str_FecFin, "yyyy")) Then
           MsgBox "Esta intentando registrar un documento de un ejercicio anterior.", vbExclamation, modgen_g_str_NomPlt
       End If

       pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(3, 2, Format(ipp_FchEmiPrv.Text, "yyyymmdd"), 1)
       pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
   End If
End Sub

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
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
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
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub ipp_FchVenc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If cmb_Moneda.Enabled = False Then
          Call gs_SetFocus(ipp_ImpGrav_01)
       Else
          Call gs_SetFocus(cmb_Moneda)
       End If
              
       If (ipp_FchVenc.Text <> "") Then
           If (Format(ipp_FchVenc.Text, "yyyymmdd") < Format(ipp_FchEmiPrv.Text, "yyyymmdd")) Then
               MsgBox "La fecha de vencimiento no puede ser menor al fecha de emisión.", vbExclamation, modgen_g_str_NomPlt
           End If
       End If
   End If
End Sub

Private Sub ipp_ImpGrav_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_GravDH_01)
   End If
End Sub

Private Sub ipp_ImpGrav_01_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub ipp_ImpGrav_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_GravDH_02)
   End If
End Sub

Private Sub ipp_ImpGrav_02_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub ipp_ImpNGrv_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_NGrvDH_01)
   End If
End Sub

Private Sub ipp_ImpNGrv_01_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub ipp_ImpNGrv_02_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_NGrvDH_02)
   End If
End Sub

Private Sub ipp_ImpNGrv_02_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub txt_CtrCosto_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_Descrip_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchCtb)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_NumPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchEmiPrv)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumSeriePrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumPrv)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO)
   End If
End Sub


Function fs_HabTxt(p_Objeto As fpDoubleSingle) As String
   If Trim(p_Objeto.Name) = Trim(ipp_ImpGrav_01.Name) Then
      fs_HabTxt = ipp_ImpGrav_01.Text
      If Trim(cmb_GravDH_01.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpGrav_01.Text
      End If
   End If
   
   If Trim(p_Objeto.Name) = Trim(ipp_ImpGrav_02.Name) Then
      fs_HabTxt = ipp_ImpGrav_02.Text
      If Trim(cmb_GravDH_02.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpGrav_02.Text
      End If
   End If
   
   If Trim(p_Objeto.Name) = Trim(ipp_ImpNGrv_01.Name) Then
      fs_HabTxt = ipp_ImpNGrv_01.Text
      If Trim(cmb_NGrvDH_01.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpNGrv_01.Text
      End If
   End If
   
   If Trim(p_Objeto.Name) = Trim(ipp_ImpNGrv_02.Name) Then
      fs_HabTxt = ipp_ImpNGrv_02.Text
      If Trim(cmb_NGrvDH_02.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpNGrv_02.Text
      End If
   End If
End Function

Function fs_HabPnl(p_Objeto As SSPanel) As String
   If Trim(p_Objeto.Name) = Trim(pnl_ImpIgv.Name) Then
      fs_HabPnl = pnl_ImpIgv.Caption
      If Trim(pnl_IgvDH.Caption) = "H" Then
         fs_HabPnl = "-" & pnl_ImpIgv.Caption
      End If
   End If
      
   If Trim(p_Objeto.Name) = Trim(pnl_ImpPpg.Name) Then
      fs_HabPnl = pnl_ImpPpg.Caption
      If Trim(pnl_PpgDH.Caption) = "H" Then
         fs_HabPnl = "-" & pnl_ImpPpg.Caption
      End If
   End If
End Function

Private Sub fs_Calcular_Determ()
Dim r_dbl_Aux     As Double
Dim r_dbl_ImpPp   As Double
Dim r_dbl_ImpAux  As Double
Dim r_bol_Estado  As Boolean
      
   r_bol_Estado = False
   r_dbl_ImpAux = 0
   r_dbl_ImpPp = 0
   
   If (cmb_TipCbtPrv.ListIndex = -1 Or cmb_Moneda.ListIndex = -1) Then
       ipp_ImpGrav_01.Text = "0.00"
       ipp_ImpGrav_02.Text = "0.00"
       ipp_ImpNGrv_01.Text = "0.00"
       ipp_ImpNGrv_02.Text = "0.00"
   End If
   
   If (l_dbl_IniFrm = False) Then
       Exit Sub
   End If
   If (cmb_TipCbtPrv.ListIndex = -1) Then
       Call gs_SetFocus(cmb_TipCbtPrv)
       MsgBox "Debe seleccionar un tipo de comprobante.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   If (cmb_Moneda.ListIndex = -1) Then
       Call gs_SetFocus(cmb_Moneda)
       MsgBox "Debe seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
      
   pnl_ImpIgv.Caption = "0.00 "
   r_dbl_ImpAux = 0
          
   
   'pnl_IgvDH.Caption = "D"
   'pnl_PpgDH.Caption = "H"
   
   'calculo del igv
   r_dbl_ImpAux = Math.Abs((CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02))) * l_dbl_IGV)
   pnl_ImpIgv.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
   
   'calculo por pagar
   r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02)) + _
                  CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + CDbl(fs_HabTxt(ipp_ImpNGrv_02)) + CDbl(fs_HabPnl(pnl_ImpIgv)))
   pnl_ImpPpg.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
End Sub

Private Sub fs_Cuenta_IGV()
Dim r_str_CadAux  As String
Dim r_bol_Estado  As Boolean

   r_str_CadAux = ""
   r_bol_Estado = False
   
   r_str_CadAux = Mid(cmb_CtaGvd_01.Text, 1, l_int_TopNiv)
   If Len(r_str_CadAux) > 0 Then
      r_bol_Estado = True
      pnl_Igv.Caption = r_str_CadAux
   End If
   If r_bol_Estado = False Then
      r_str_CadAux = Mid(cmb_CtaGvd_02.Text, 1, l_int_TopNiv)
      If Len(r_str_CadAux) > 0 Then
         r_bol_Estado = True
         pnl_Igv.Caption = r_str_CadAux
      End If
   End If
   If r_bol_Estado = False Then
      r_str_CadAux = Mid(cmb_CtaNoGvd_01.Text, 1, l_int_TopNiv)
      If Len(r_str_CadAux) > 0 Then
         r_bol_Estado = True
         pnl_Igv.Caption = r_str_CadAux
      End If
   End If
   If r_bol_Estado = False Then
      r_str_CadAux = Mid(cmb_CtaNoGvd_02.Text, 1, l_int_TopNiv)
      If Len(r_str_CadAux) > 0 Then
         r_bol_Estado = True
         pnl_Igv.Caption = r_str_CadAux
      End If
   End If
End Sub

