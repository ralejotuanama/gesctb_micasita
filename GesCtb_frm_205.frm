VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_TrnCta_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "GesCtb_frm_205.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6735
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8235
      _Version        =   65536
      _ExtentX        =   14526
      _ExtentY        =   11880
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
         TabIndex        =   16
         Top             =   60
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registros de Transferencias"
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
            Picture         =   "GesCtb_frm_205.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   18
         Top             =   780
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
            Picture         =   "GesCtb_frm_205.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7350
            Picture         =   "GesCtb_frm_205.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_205.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2145
         Left            =   60
         TabIndex        =   19
         Top             =   2970
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
         _ExtentY        =   3784
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
         Begin VB.ComboBox cmb_NumCta_Cgo 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   660
            Width           =   2800
         End
         Begin VB.ComboBox cmb_BcoCgo 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   330
            Width           =   5415
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1410
            TabIndex        =   8
            Top             =   1350
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
         Begin Threed.SSPanel pnl_TipCta_Cgo 
            Height          =   315
            Left            =   1410
            TabIndex        =   6
            Top             =   1005
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4939
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
         Begin Threed.SSPanel pnl_Moneda_Cgo 
            Height          =   315
            Left            =   5490
            TabIndex        =   7
            Top             =   1050
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         Begin Threed.SSPanel pnl_SalTran_Cgo 
            Height          =   315
            Left            =   1410
            TabIndex        =   39
            Top             =   1680
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Mto Convertido:"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   1740
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   4470
            TabIndex        =   25
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Banco Cargo:"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
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
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   1410
            Width           =   570
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1440
         Left            =   60
         TabIndex        =   26
         Top             =   5160
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
         _ExtentY        =   2540
         _StockProps     =   15
         BevelOuter      =   1
         Begin VB.ComboBox cmb_BcoAbn 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   330
            Width           =   5415
         End
         Begin VB.ComboBox cmb_NumCta_Abn 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   660
            Width           =   2800
         End
         Begin Threed.SSPanel pnl_TipCta_Abn 
            Height          =   315
            Left            =   1410
            TabIndex        =   13
            Top             =   1005
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4939
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
         Begin Threed.SSPanel pnl_Moneda_Abn 
            Height          =   315
            Left            =   5490
            TabIndex        =   14
            Top             =   990
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         Begin VB.Label Label16 
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
            TabIndex        =   36
            Top             =   60
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Banco Abono:"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   390
            Width           =   1020
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   28
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   4470
            TabIndex        =   27
            Top             =   1080
            Width           =   630
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1425
         Left            =   60
         TabIndex        =   31
         Top             =   1500
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
         Begin VB.TextBox txt_NumRef 
            Height          =   315
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   3
            Top             =   990
            Width           =   2800
         End
         Begin Threed.SSPanel pnl_NumMov 
            Height          =   315
            Left            =   1410
            TabIndex        =   1
            Top             =   330
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
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
            Left            =   1410
            TabIndex        =   0
            Top             =   660
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
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   315
            Left            =   5490
            TabIndex        =   2
            Top             =   660
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.000000 "
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nro Referencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label12 
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
            TabIndex        =   35
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nro Movimiento:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   4470
            TabIndex        =   32
            Top             =   720
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_TrnCta_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_CtaBan
   CtaBan_CodBan        As String
   CtaBan_NumCta        As String
   ctaban_TipCta_id     As Long
   CtaBan_TipCta        As String
   ctaban_TipMon_id     As Long
   ctaban_TipMon        As String
   CtaBan_CtaCtb        As String
   ctaban_Descri        As String
   ctaban_Situac        As Integer
   ctaban_SalCap        As Double
End Type

Dim l_arr_TipCta()      As arr_CtaBan

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   cmd_Grabar.Visible = False
   cmd_Reversa.Visible = False
      
   If moddat_g_int_FlgGrb = 0 Then
      pnl_Titulo.Caption = "Transferencia - Consultar"
      Call fs_CargarMov
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then
      pnl_Titulo.Caption = "Transferencia - Adicionar"
      cmd_Grabar.Visible = True
   ElseIf moddat_g_int_FlgGrb = 2 Then
      pnl_Titulo.Caption = "Transferencia - Reversa"
      cmd_Reversa.Left = 30
      cmd_Reversa.Visible = True
      Call fs_CargarMov
      Call fs_Desabilitar
   End If
  
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_BcoCgo, 1, "516")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BcoAbn, 1, "516")
   
   Call fs_BcoCta_Cargar
End Sub

Private Sub fs_Limpia()
   pnl_NumMov.Caption = ""
   ipp_FecOpe.Text = date
   pnl_TipCambio.Caption = "0.000000 "
   Call ipp_FecOpe_LostFocus
   cmb_BcoCgo.ListIndex = -1
   cmb_NumCta_Cgo.ListIndex = -1
   cmb_BcoAbn.ListIndex = -1
   cmb_NumCta_Abn.ListIndex = -1
   txt_NumRef.Text = ""
   pnl_SalTran_Cgo.Caption = "0.00" & " "
End Sub

Private Sub fs_Desabilitar()
   ipp_FecOpe.Enabled = False
   txt_NumRef.Enabled = False
   cmb_BcoCgo.Enabled = False
   cmb_NumCta_Cgo.Enabled = False
   ipp_Import.Enabled = False
   cmb_BcoAbn.Enabled = False
   cmb_NumCta_Abn.Enabled = False
End Sub

Private Function fs_Cargar_Ctas(p_TipMov As Integer) As Integer
Dim r_int_Fila As Integer

    If p_TipMov = 1 Then
    'Cargo Traspaso
       cmb_NumCta_Cgo.Clear
       For r_int_Fila = 1 To UBound(l_arr_TipCta)
           If l_arr_TipCta(r_int_Fila).CtaBan_CodBan = cmb_BcoCgo.ItemData(cmb_BcoCgo.ListIndex) Then
              If moddat_g_int_FlgGrb = 1 Then
                 If l_arr_TipCta(r_int_Fila).ctaban_Situac = 1 Then
                    cmb_NumCta_Cgo.AddItem Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta)
                 End If
              End If
              If moddat_g_int_FlgGrb = 0 Or moddat_g_int_FlgGrb = 2 Then
                 cmb_NumCta_Cgo.AddItem Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta)
              End If
           End If
       Next
    ElseIf p_TipMov = 2 Then
    'Abono Traspaso
       cmb_NumCta_Abn.Clear
       For r_int_Fila = 1 To UBound(l_arr_TipCta)
           If l_arr_TipCta(r_int_Fila).CtaBan_CodBan = cmb_BcoAbn.ItemData(cmb_BcoAbn.ListIndex) Then
              If moddat_g_int_FlgGrb = 1 Then
                 If l_arr_TipCta(r_int_Fila).ctaban_Situac = 1 Then
                    cmb_NumCta_Abn.AddItem Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta)
                 End If
              End If
              If moddat_g_int_FlgGrb = 0 Or moddat_g_int_FlgGrb = 2 Then
                 cmb_NumCta_Abn.AddItem Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta)
              End If
           End If
       Next
    End If
End Function

Private Function fs_Cargar_CtaDet(p_TipMov As Integer) As Integer
Dim r_int_Fila As Integer

    If p_TipMov = 1 Then
    'Cargo Traspaso
       pnl_TipCta_Cgo.Tag = ""
       pnl_Moneda_Cgo.Tag = ""
       pnl_TipCta_Cgo.Caption = ""
       pnl_Moneda_Cgo.Caption = ""
       
       For r_int_Fila = 1 To UBound(l_arr_TipCta)
           If l_arr_TipCta(r_int_Fila).CtaBan_CodBan = cmb_BcoCgo.ItemData(cmb_BcoCgo.ListIndex) And _
              Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta) = Trim(cmb_NumCta_Cgo.Text) Then
              pnl_TipCta_Cgo.Caption = Trim(l_arr_TipCta(r_int_Fila).CtaBan_TipCta)
              pnl_Moneda_Cgo.Caption = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipMon)
              '----
              pnl_TipCta_Cgo.Tag = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipCta_id)
              pnl_Moneda_Cgo.Tag = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipMon_id)
              Exit For
           End If
       Next
    ElseIf p_TipMov = 2 Then
    'Abono Traspaso
       pnl_TipCta_Abn.Tag = ""
       pnl_Moneda_Abn.Tag = ""
       pnl_TipCta_Abn.Caption = ""
       pnl_Moneda_Abn.Caption = ""
       For r_int_Fila = 1 To UBound(l_arr_TipCta)
           If l_arr_TipCta(r_int_Fila).CtaBan_CodBan = cmb_BcoAbn.ItemData(cmb_BcoAbn.ListIndex) And _
              Trim(l_arr_TipCta(r_int_Fila).CtaBan_NumCta) = Trim(cmb_NumCta_Abn.Text) Then
              pnl_TipCta_Abn.Caption = Trim(l_arr_TipCta(r_int_Fila).CtaBan_TipCta)
              pnl_Moneda_Abn.Caption = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipMon)
              '----
              pnl_TipCta_Abn.Tag = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipCta_id)
              pnl_Moneda_Abn.Tag = Trim(l_arr_TipCta(r_int_Fila).ctaban_TipMon_id)
              Exit For
           End If
       Next
    End If
End Function

Private Sub fs_BcoCta_Cargar()
   ReDim l_arr_TipCta(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CTABAN_CODBAN, CTABAN_NUMCTA, CTABAN_TIPCTA, TRIM(B.PARDES_DESCRI) AS TIPOCUENTA,  "
   g_str_Parame = g_str_Parame & "       CTABAN_TIPMON, TRIM(C.PARDES_DESCRI) AS TIPOMONEDA, CTABAN_CTACTB, CTABAN_SITUAC,  "
   g_str_Parame = g_str_Parame & "       TRIM(CTABAN_DESCRI) AS CTABAN_DESCRI  "
   g_str_Parame = g_str_Parame & "  FROM MNT_CTABAN A  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 510 AND B.PARDES_CODITE = A.CTABAN_TIPCTA  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CTABAN_TIPMON  "
   'g_str_Parame = g_str_Parame & " WHERE CTABAN_CODBAN = '000002'  "
 
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
      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_TipCta(UBound(l_arr_TipCta) + 1)
      
      l_arr_TipCta(UBound(l_arr_TipCta)).CtaBan_CodBan = Trim(g_rst_Genera!CtaBan_CodBan & "")
      l_arr_TipCta(UBound(l_arr_TipCta)).CtaBan_NumCta = Trim(g_rst_Genera!CtaBan_NumCta & "")
      l_arr_TipCta(UBound(l_arr_TipCta)).ctaban_TipCta_id = g_rst_Genera!CtaBan_TipCta
      l_arr_TipCta(UBound(l_arr_TipCta)).CtaBan_TipCta = Trim(g_rst_Genera!TIPOCUENTA & "")
      l_arr_TipCta(UBound(l_arr_TipCta)).ctaban_TipMon_id = g_rst_Genera!ctaban_TipMon
      l_arr_TipCta(UBound(l_arr_TipCta)).ctaban_TipMon = Trim(g_rst_Genera!TIPOMONEDA & "")
      l_arr_TipCta(UBound(l_arr_TipCta)).CtaBan_CtaCtb = Trim(g_rst_Genera!CtaBan_CtaCtb & "")
      l_arr_TipCta(UBound(l_arr_TipCta)).ctaban_Situac = g_rst_Genera!ctaban_Situac
      l_arr_TipCta(UBound(l_arr_TipCta)).ctaban_Descri = Trim(g_rst_Genera!ctaban_Descri & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_CtaDeb     As String
Dim r_str_CtaHab     As String

   If Len(Trim(txt_NumRef.Text)) = 0 Then
      MsgBox "Tiene que ingresar el numero de referencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumRef)
      Exit Sub
   End If

   If CDbl(pnl_TipCambio.Caption) = 0 Then
      MsgBox "El Tipo de cambio no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecOpe)
      Exit Sub
   End If
   
   If cmb_BcoCgo.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un banco en el grupo cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BcoCgo)
      Exit Sub
   End If
   
   If cmb_NumCta_Cgo.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar el nro cuenta en el grupo cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumCta_Cgo)
      Exit Sub
   End If
   
   If CDbl(ipp_Import.Text) = 0 Then
      MsgBox "Tiene que ingresar un importe mayor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
   If cmb_BcoAbn.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un banco en el grupo abono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BcoAbn)
      Exit Sub
   End If
   
   If cmb_NumCta_Abn.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar el nro cuenta en el grupo abono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumCta_Abn)
      Exit Sub
   End If
   
   If cmb_BcoCgo.ItemData(cmb_BcoCgo.ListIndex) = cmb_BcoAbn.ItemData(cmb_BcoAbn.ListIndex) And _
      Trim(cmb_NumCta_Cgo.Text) = Trim(cmb_NumCta_Abn.Text) Then
      MsgBox "El nro cuenta del grupo cargo, no puede ser abonada a la misma cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumCta_Cgo)
      Exit Sub
   End If
   
   If CDbl(pnl_SalTran_Cgo.Caption) = 0 Then
      MsgBox "El importe convertido a transferir no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
   r_str_CtaDeb = ""
   r_str_CtaHab = ""
   Call fs_BuscarCtas(r_str_CtaDeb, r_str_CtaHab)
   If r_str_CtaDeb = "" Then
      MsgBox "El nro cuenta del grupo cargo no tiene una cuenta contable asignada.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumCta_Cgo)
      Exit Sub
   End If
   If r_str_CtaHab = "" Then
      MsgBox "El nro cuenta del grupo abono no tiene una cuenta contable asignada.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumCta_Abn)
      Exit Sub
   End If
   
   'If Format(ipp_FecOpe.Text, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
   '   MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
   '
   '   If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
   '      Call gs_SetFocus(ipp_FecOpe)
   '      Exit Sub
   '   End If
   'End If
   
'   If (Format(ipp_FecOpe.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FecOpe.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FecOpe)
'       Exit Sub
'   End If

   If fs_ValidaPeriodo(ipp_FecOpe.Text) = False Then
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Grabar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Grabar()
Dim r_str_AsiGen As String

   r_str_AsiGen = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MOVCTA ( "
   g_str_Parame = g_str_Parame & "" & 1 & ", " '1= TRANSFERENCIAS
   g_str_Parame = g_str_Parame & "" & Format(ipp_FecOpe.Text, "yyyymmdd") & ", " 'MOVCTA_FECOPE
   g_str_Parame = g_str_Parame & "" & CDbl(pnl_TipCambio.Caption) & ", " 'MOVCTA_TIPCAM
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumRef.Text) & "', " 'MOVCTA_NUMREF
   g_str_Parame = g_str_Parame & "" & CDbl(ipp_Import.Text) & ", "  'MOVCTA_SALCAP_CRG
   g_str_Parame = g_str_Parame & "'" & Format(cmb_BcoCgo.ItemData(cmb_BcoCgo.ListIndex), "000000") & "', " 'CODBAN_CRG
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_NumCta_Cgo.Text) & "', " 'NUMCTA_CRG
   g_str_Parame = g_str_Parame & "'" & Format(cmb_BcoAbn.ItemData(cmb_BcoAbn.ListIndex), "000000") & "', " 'CODBAN_ABN
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_NumCta_Abn.Text) & "', " 'NUMCTA_ABN
   g_str_Parame = g_str_Parame & "" & CDbl(pnl_SalTran_Cgo.Caption) & ", "  'MOVCTA_SALCAP_ABN
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'AS_SEGUSU
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'AS_SEGTER
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', " 'AS_SEGPLT
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'AS_SEGSUC
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   '0=nada paso, 1=inserto, 2=monto insuficiente
   If (g_rst_Genera!RESUL = 1) Then
       Call fs_GeneraAsiento(Trim(g_rst_Genera!CODIGO), r_str_AsiGen)
       MsgBox "Operación completada con éxito." & vbCrLf & _
              "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_TrnCta_02.fs_BuscarMov
       Screen.MousePointer = 0
       Unload Me
   End If
   
End Sub

Private Sub cmd_Reversa_Click()
   If MsgBox("¿Esta seguro que desea realizar esta operación de reversa?" & vbCrLf & _
             "Recuerde que debe eliminar el asiento contable manual.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_MOVCTA_REVERSA ( "
      g_str_Parame = g_str_Parame & "" & CLng(pnl_NumMov.Caption) & ",  " 'NUM_MOV
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         MsgBox "No se pudo completar la oepración de reversa.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If (g_rst_Genera!RESUL = 1) Then
         MsgBox "Se completo la operación de reversa, recuerde que debe eliminar el asiento contable manual.", vbInformation, modgen_g_str_NomPlt
         Call frm_Ctb_TrnCta_02.fs_BuscarMov
         Unload Me
      ElseIf (g_rst_Genera!RESUL = 2) Then
         MsgBox "Debe de dar reversa a los movimientos superiores, para dar reversa a este nro movimiento.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If
End Sub

Private Sub fs_GeneraAsiento(ByVal p_Codigo As String, ByRef p_AsiGen As String)
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_int_NumAsi    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_int_Contar    As Integer
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer

   'r_int_PerAno = modctb_int_PerAno 'Year(ipp_FecOpe.Text)
   'r_int_PerMes = modctb_int_PerMes 'Month(ipp_FecOpe.Text)
   r_int_PerAno = Year(ipp_FecOpe.Text)
   r_int_PerMes = Month(ipp_FecOpe.Text)
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   
   r_int_NumAsi = 0
   r_int_NumIte = 0
   r_str_FechaC = Format(ipp_FecOpe.Text, "yyyymmdd")
   r_str_FechaL = ipp_FecOpe.Text
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   p_AsiGen = CStr(r_int_NumAsi)
   
   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecOpe.Text, "yyyymmdd"), 1)
   'r_str_Glosa = Mid("TRANSFERENCIAS PROPIAS" & "/" & Format(Trim(p_Codigo), "00000000"), 1, 60)
   r_str_Glosa = Mid("TRANSFERENCIAS PROPIAS" & "/" & Trim(txt_NumRef), 1, 60)
   
   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                 r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FechaL, "1")
                                 
   Call fs_BuscarCtas(r_str_CtaDeb, r_str_CtaHab)
   
   'Insertar en DETALLE
   'r_dbl_MtoSol = CDbl(ipp_Import.Text)
   'r_dbl_MtoDol = 0
   If CDbl(pnl_Moneda_Cgo.Tag) = 1 Then
      r_dbl_MtoSol = CDbl(ipp_Import.Text)
      r_dbl_MtoDol = Format(CDbl(ipp_Import.Text) / r_dbl_TipSbs, "###,###,##0.00")
   Else
      r_dbl_MtoSol = Format(CDbl(ipp_Import.Text) * r_dbl_TipSbs, "###,###,##0.00")
      r_dbl_MtoDol = CDbl(ipp_Import.Text)
   End If
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, r_str_CtaDeb, CDate(r_str_FechaL), _
                                        r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))

   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, r_str_CtaHab, CDate(r_str_FechaL), _
                                        r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
                                        
   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_MOVCTA "
   g_str_Parame = g_str_Parame & "   SET MOVCTA_DATCNT = '" & r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   g_str_Parame = g_str_Parame & " WHERE MOVCTA_NUMMOV = " & p_Codigo
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Function fs_BuscarCtas(ByRef p_CtaDeb As String, ByRef p_CtaHab As String)
   'extrae el numero de cuenta
   p_CtaDeb = ""
   p_CtaHab = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT (SELECT A.CTABAN_CTACTB  "
   g_str_Parame = g_str_Parame & "           FROM MNT_CTABAN A  "
   g_str_Parame = g_str_Parame & "          WHERE A.CTABAN_CODBAN = '" & Format(cmb_BcoCgo.ItemData(cmb_BcoCgo.ListIndex), "000000") & "' "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_NUMCTA = '" & Trim(cmb_NumCta_Cgo.Text) & "'  "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_TIPCTA = '" & Format(Trim(pnl_TipCta_Cgo.Tag), "000000") & "'  "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_TIPMON = " & CInt(pnl_Moneda_Cgo.Tag) & ") AS CARGO_HABER,  "
   g_str_Parame = g_str_Parame & "        (SELECT A.CTABAN_CTACTB  "
   g_str_Parame = g_str_Parame & "           FROM MNT_CTABAN A  "
   g_str_Parame = g_str_Parame & "          WHERE A.CTABAN_CODBAN = '" & Format(cmb_BcoAbn.ItemData(cmb_BcoAbn.ListIndex), "000000") & "' "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_NUMCTA = '" & Trim(cmb_NumCta_Abn.Text) & "'  "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_TIPCTA = '" & Format(Trim(pnl_TipCta_Abn.Tag), "000000") & "'  "
   g_str_Parame = g_str_Parame & "            AND A.CTABAN_TIPMON = " & CInt(pnl_Moneda_Abn.Tag) & ") AS DEBE_ABONO  "
   g_str_Parame = g_str_Parame & "   FROM DUAL  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If
            
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ninguna cuenta contable para generar el asiento", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   Else
      p_CtaDeb = IIf(IsNull(g_rst_Princi!DEBE_ABONO) = True, "", Trim(g_rst_Princi!DEBE_ABONO & ""))
      p_CtaHab = IIf(IsNull(g_rst_Princi!CARGO_HABER) = True, "", Trim(g_rst_Princi!CARGO_HABER & ""))
   End If
End Function

Private Sub fs_CargarMov()
'moddat_g_int_FlgGrb
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT LPAD(A.MOVCTA_NUMMOV,8,'0') AS MOVCTA_NUMMOV, A.MOVCTA_TIPMOV, A.MOVCTA_FECOPE,  "
   g_str_Parame = g_str_Parame & "       A.MOVCTA_TIPCAM, A.MOVCTA_NUMREF, A.MOVCTA_CODBAN, A.MOVCTA_NUMCTA, A.MOVCTA_IMPORT  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_MOVCTA A  "
   g_str_Parame = g_str_Parame & " WHERE A.MOVCTA_NUMMOV = " & CLng(moddat_g_str_NumOpe)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!MOVCTA_TIPMOV = 1 Then
            pnl_NumMov.Caption = g_rst_Princi!MOVCTA_NUMMOV
            ipp_FecOpe.Text = gf_FormatoFecha(g_rst_Princi!MOVCTA_FECOPE)
            pnl_TipCambio.Caption = Format(g_rst_Princi!MOVCTA_TIPCAM, "###,###,##0.000000") & " "
            txt_NumRef.Text = Trim(g_rst_Princi!MOVCTA_NUMREF & "")
            ipp_Import.Text = Format(CStr(g_rst_Princi!MOVCTA_IMPORT), "###,###,##0.00") & " "
            
            Call gs_BuscarCombo_Item(cmb_BcoCgo, CLng(g_rst_Princi!MOVCTA_CODBAN))
            Call cmb_BcoCgo_Click
            Call gs_BuscarCombo_Text(cmb_NumCta_Cgo, g_rst_Princi!MOVCTA_NUMCTA, -1)
            'cmb_NumCta_Cgo.Text = Trim(g_rst_Princi!MOVCTA_NUMCTA & "")
         End If
         
         If g_rst_Princi!MOVCTA_TIPMOV = 2 Then
            Call gs_BuscarCombo_Item(cmb_BcoAbn, g_rst_Princi!MOVCTA_CODBAN)
            Call cmb_BcoAbn_Click
            Call gs_BuscarCombo_Text(cmb_NumCta_Abn, g_rst_Princi!MOVCTA_NUMCTA, -1)
            'cmb_NumCta_Abn.Text = Trim(g_rst_Princi!MOVCTA_NUMCTA & "")
         End If
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmb_BcoAbn_Click()
   Call fs_Cargar_Ctas(2) 'Abono
End Sub

Private Sub cmb_BcoCgo_Click()
   Call fs_Cargar_Ctas(1) 'Cargo
End Sub

Private Sub cmb_NumCta_Abn_Click()
   Call fs_Cargar_CtaDet(2)
   Call ipp_Import_LostFocus
End Sub

Private Sub cmb_NumCta_Cgo_Click()
   Call fs_Cargar_CtaDet(1)
   Call ipp_Import_LostFocus
End Sub

Private Sub ipp_FecOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumRef)
   End If
End Sub

Private Sub ipp_FecOpe_LostFocus()
  'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
  'TipTip = 1 - Venta / 2 - Compra
   pnl_TipCambio.Caption = Format(moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecOpe.Text, "yyyymmdd"), 1), "###,###,##0.000000") & " "
   Call ipp_Import_LostFocus
End Sub

Private Sub txt_NumRef_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BcoCgo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_BcoCgo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NumCta_Cgo)
   End If
End Sub

Private Sub cmb_NumCta_Cgo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import)
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BcoAbn)
   End If
End Sub

Private Sub ipp_Import_LostFocus()
   pnl_SalTran_Cgo.Caption = "0.00" & " "
   
   If cmb_NumCta_Cgo.ListIndex > -1 And cmb_NumCta_Abn.ListIndex > -1 Then
      If CInt(pnl_Moneda_Cgo.Tag) <> CInt(pnl_Moneda_Abn.Tag) Then
         If CDbl(pnl_TipCambio.Caption) = 0 Then
            Exit Sub
         End If
         If CInt(pnl_Moneda_Cgo.Tag) = 1 Then
            'SOLES
            pnl_SalTran_Cgo.Caption = Format(CDbl(ipp_Import.Text) / CDbl(pnl_TipCambio.Caption), "###,###,##0.00")
         Else
            'DOLAR
            pnl_SalTran_Cgo.Caption = Format(CDbl(ipp_Import.Text) * CDbl(pnl_TipCambio.Caption), "###,###,##0.00")
         End If
      Else
         pnl_SalTran_Cgo.Caption = ipp_Import.Text & " "
      End If
   End If
End Sub

Private Sub cmb_BcoAbn_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NumCta_Abn)
   End If
End Sub

Private Sub cmb_NumCta_Abn_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub



