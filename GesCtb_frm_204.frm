VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegVen_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
   Icon            =   "GesCtb_frm_204.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6900
      Left            =   0
      TabIndex        =   24
      Top             =   -30
      Width           =   13605
      _Version        =   65536
      _ExtentX        =   23998
      _ExtentY        =   12171
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
         TabIndex        =   25
         Top             =   90
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
            TabIndex        =   26
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Ventas"
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
            Picture         =   "GesCtb_frm_204.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_DatPrv 
         Height          =   1185
         Left            =   60
         TabIndex        =   27
         Top             =   1500
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
         Begin VB.TextBox txt_Descrip 
            Height          =   315
            Left            =   7020
            MaxLength       =   60
            TabIndex        =   2
            Top             =   690
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
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   7020
            TabIndex        =   1
            Top             =   360
            Width           =   6180
         End
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1380
            TabIndex        =   28
            Top             =   690
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Datos del Cliente"
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
            TabIndex        =   33
            Top             =   60
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   5520
            TabIndex        =   32
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   420
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   5520
            TabIndex        =   29
            Top             =   390
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   34
         Top             =   800
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   60
            Picture         =   "GesCtb_frm_204.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   12720
            Picture         =   "GesCtb_frm_204.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Reversa 
            Height          =   600
            Left            =   660
            Picture         =   "GesCtb_frm_204.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Reversa del Asiento"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_DatCbt 
         Height          =   2025
         Left            =   60
         TabIndex        =   38
         Top             =   2730
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
         _ExtentY        =   3572
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
         Begin VB.CheckBox chk_RetBien 
            Caption         =   "Retiro de Bienes"
            Height          =   345
            Left            =   10440
            TabIndex        =   11
            Top             =   1320
            Width           =   1875
         End
         Begin VB.TextBox txt_NumSeriePrv 
            Height          =   315
            Left            =   7020
            MaxLength       =   4
            TabIndex        =   5
            Top             =   630
            Width           =   1515
         End
         Begin VB.ComboBox cmb_TipCbtPrv 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   3915
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1300
            Width           =   2175
         End
         Begin VB.TextBox txt_NumPrv 
            Height          =   315
            Left            =   10410
            MaxLength       =   7
            TabIndex        =   6
            Top             =   630
            Width           =   1515
         End
         Begin EditLib.fpDateTime ipp_FchCtb 
            Height          =   315
            Left            =   1380
            TabIndex        =   3
            Top             =   300
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
         Begin EditLib.fpDateTime ipp_FchVenc 
            Height          =   315
            Left            =   7020
            TabIndex        =   8
            Top             =   960
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
            TabIndex        =   7
            Top             =   970
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
            TabIndex        =   10
            Top             =   1300
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
            Alignment       =   4
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   690
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Contable:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1170
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
            TabIndex        =   45
            Top             =   30
            Width           =   1980
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Serie:"
            Height          =   195
            Left            =   5490
            TabIndex        =   44
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Numero:"
            Height          =   195
            Left            =   8910
            TabIndex        =   43
            Top             =   690
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vencimiento:"
            Height          =   195
            Left            =   5490
            TabIndex        =   41
            Top             =   1060
            Width           =   1410
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   1410
            Width           =   630
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   5490
            TabIndex        =   39
            Top             =   1410
            Width           =   930
         End
      End
      Begin Threed.SSPanel pnl_DatDet 
         Height          =   1845
         Left            =   60
         TabIndex        =   48
         Top             =   4800
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   3254
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
         Begin VB.ComboBox cmb_PorCobrar 
            Height          =   315
            Left            =   3570
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1320
            Width           =   2670
         End
         Begin VB.ComboBox cmb_GravDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   330
            Width           =   645
         End
         Begin VB.ComboBox cmb_NGrvDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   660
            Width           =   645
         End
         Begin VB.ComboBox cmb_CtaGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   14
            Top             =   330
            Width           =   4980
         End
         Begin VB.ComboBox cmb_CtaNoGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   17
            Top             =   660
            Width           =   4980
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
         Begin EditLib.fpDoubleSingle ipp_ImpNGrv_01 
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
         Begin Threed.SSPanel pnl_Igv 
            Height          =   315
            Left            =   3570
            TabIndex        =   20
            Top             =   990
            Width           =   2670
            _Version        =   65536
            _ExtentX        =   4710
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
         Begin Threed.SSPanel pnl_ImpIgv 
            Height          =   315
            Left            =   1380
            TabIndex        =   18
            Top             =   990
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpPpg 
            Height          =   315
            Left            =   1380
            TabIndex        =   21
            Top             =   1320
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IgvDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   19
            Top             =   990
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "H"
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
         End
         Begin Threed.SSPanel pnl_PpgDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   22
            Top             =   1320
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1020
            Width           =   270
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "No Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   915
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
            TabIndex        =   53
            Top             =   60
            Width           =   1230
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   390
            Width           =   660
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Por Cobrar:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1350
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "D/H"
            Height          =   195
            Left            =   3060
            TabIndex        =   50
            Top             =   90
            Width           =   315
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CUENTAS CONTABLES"
            Height          =   195
            Left            =   5190
            TabIndex        =   49
            Top             =   90
            Width           =   1770
         End
      End
      Begin Threed.SSPanel pnl_DatRef 
         Height          =   1845
         Left            =   8760
         TabIndex        =   57
         Top             =   4800
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   3254
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
         Begin VB.TextBox txt_NumSerieRef 
            Height          =   315
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   60
            Top             =   990
            Width           =   1425
         End
         Begin VB.ComboBox cmb_TipCbtRef 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   660
            Width           =   2880
         End
         Begin VB.TextBox txt_NumRef 
            Height          =   315
            Left            =   1590
            MaxLength       =   7
            TabIndex        =   58
            Top             =   1320
            Width           =   1425
         End
         Begin EditLib.fpDateTime ipp_FchEmiRef 
            Height          =   315
            Left            =   1590
            TabIndex        =   61
            Top             =   330
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Referencia"
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
            TabIndex        =   66
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   390
            Width           =   1080
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Serie:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   1350
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegVen_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_TopNiv        As Integer
Dim l_dbl_IniFrm        As Boolean
Dim l_arr_DebHab()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_dbl_IGV           As Double

Private Sub chk_RetBien_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpGrav_01)
   End If
End Sub

Private Sub cmb_CtaGvd_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpNGrv_01)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaGvd_01_LostFocus()
   'Call fs_Cuenta_IGV
End Sub

Private Sub cmb_CtaNoGvd_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_PorCobrar)
   End If
End Sub

Private Sub cmb_CtaNoGvd_01_LostFocus()
   'Call fs_Cuenta_IGV
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

Private Sub cmb_Moneda_Click()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(chk_RetBien)
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

Private Sub cmb_PorCobrar_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If ipp_FchEmiRef.Enabled = False Then
          Call gs_SetFocus(cmd_Grabar)
       Else
          Call gs_SetFocus(ipp_FchEmiRef)
       End If
   End If
End Sub

Private Sub cmb_TipCbtPrv_Click()
   Call fs_ActivaRefer
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_TipCbtPrv_LostFocus()
    Call cmb_TipCbtPrv_Click
End Sub

Private Sub cmb_TipCbtRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumSerieRef)
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_ImpAux   As Double
Dim r_bol_Estado   As Boolean
Dim r_int_Contar   As Integer

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
   
   If (Format(ipp_FchCtb.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(ipp_FchCtb.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FchCtb)
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
   
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8) Then '07 - N/C
       If (Len(Trim(ipp_FchEmiRef.Text)) = 0) Then
           MsgBox "La fecha de emisión es obligatoria para el grupo referencia.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FchEmiRef)
           Exit Sub
       End If
       If Format(ipp_FchEmiPrv.Text, "yyyymmdd") < Format(ipp_FchEmiRef.Text, "yyyymmdd") Then
           MsgBox "La fecha de emision del grupo referencia no puede ser mayor a la fecha de emision del grupo comprobante.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FchEmiRef)
           Exit Sub
       End If
       If (cmb_TipCbtRef.ListIndex = -1) Then
           MsgBox "Seleccione un tipo de comprobante, es obligatorio para el grupo referencia.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_TipCbtRef)
           Exit Sub
       End If
       If (Len(Trim(txt_NumSerieRef.Text)) <> 4) Then
           MsgBox "El numero de serie consta de 4 digitos, en el grupo referencia", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(txt_NumSerieRef)
           Exit Sub
       End If
       If (Len(Trim(txt_NumRef.Text)) <> 7) Then
           MsgBox "El numero de serie consta de 7 digitos, en el grupo referencia", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(txt_NumRef)
           Exit Sub
       End If
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
   '---------------VALIDAR EXISTENCIA DE LA CUENTA-----------------------------
   If Len(Trim(cmb_CtaGvd_01.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaGvd_01.Text) = False Then
         Call gs_SetFocus(cmb_CtaGvd_01)
         Exit Sub
      End If
   End If
   If Len(Trim(cmb_CtaNoGvd_01.Text)) > 0 Then
      If fs_ValPlanCta(cmb_CtaNoGvd_01.Text) = False Then
         Call gs_SetFocus(cmb_CtaNoGvd_01)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Registro de Compras
   Call fs_Grabar_RegCom
    
   Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar_RegCom()
Dim r_dbl_ImpDeb As Double
Dim r_dbl_ImpHab As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_REGVEN ( "
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_Codigo.Caption) & "', " 'REGCOM_NUMDOC
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", " 'REGCOM_CODCOM
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', " 'REGCOM_TIPDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descrip.Text) & "', " 'REGCOM_DESCRP
   g_str_Parame = g_str_Parame & Format(ipp_FchCtb.Text, "yyyymmdd") & ", " 'REGCOM_FECCTB
   g_str_Parame = g_str_Parame & cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) & ", " 'REGCOM_TIPCPB
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumSeriePrv.Text) & "', " 'REGCOM_NSERIE
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumPrv.Text) & "', " 'REGCOM_NROCOM
   g_str_Parame = g_str_Parame & Format(ipp_FchEmiPrv.Text, "yyyymmdd") & ", " 'REGCOM_FECEMI
   g_str_Parame = g_str_Parame & IIf(ipp_FchVenc.Text = "", "Null", Format(ipp_FchVenc.Text, "yyyymmdd")) & ", " 'REGCOM_FECVTO
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'REGCOM_CODMON
   g_str_Parame = g_str_Parame & Trim(pnl_TipCambio.Caption) & ", " 'REGCOM_TIPCAM
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaGvd_01.Text = "", "", Mid(cmb_CtaGvd_01.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_GRV1
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaNoGvd_01.Text = "", "", Mid(cmb_CtaNoGvd_01.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_NGV1
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_Igv.Caption) & "', " 'REGCOM_CNT_IGV1
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_PorCobrar.Text) & "', " 'REGCOM_CNT_PPG1
   
   'GRAVADO 1
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_GravDH_01.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpGrav_01.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpGrav_01.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_GRV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'REGCOM_HAB_GRV1
   'NO GRAVADO 1
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(cmb_NGrvDH_01.Text) = "D") Then
       r_dbl_ImpDeb = CDbl(ipp_ImpNGrv_01.Text)
   Else
       r_dbl_ImpHab = CDbl(ipp_ImpNGrv_01.Text)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_NGV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'REGCOM_HAB_NGV1
   'IGV
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_IgvDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpIgv.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpIgv.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_IGV1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_IGV1
   'POR PAGAR
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_PpgDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpPpg.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpPpg.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_PPG1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_PPG1
         
   g_str_Parame = g_str_Parame & IIf(ipp_FchEmiRef.Text = "", "null", Format(ipp_FchEmiRef.Text, "yyyymmdd")) & ", " 'REGCOM_REF_FECEMI
   If (cmb_TipCbtRef.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "Null, " 'REGCOM_REF_TIPCPB
   Else
       g_str_Parame = g_str_Parame & cmb_TipCbtRef.ItemData(cmb_TipCbtRef.ListIndex) & ", " 'REGCOM_REF_TIPCPB
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumSerieRef.Text) & "', " 'REGCOM_REF_NSERIE
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumRef.Text) & "', " 'REGCOM_REF_NROCOM
   g_str_Parame = g_str_Parame & chk_RetBien.Value & ", "  'REGCOM_REF_NROCOM
   g_str_Parame = g_str_Parame & "1, " 'REGCOM_SITUAC
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
       Call frm_Ctb_RegVen_01.fs_BuscarComp
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_RegVen_01.fs_BuscarComp
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub cmd_Reversa_Click()
Dim r_str_Origen   As String
Dim r_str_Ano      As String
Dim r_str_Mes      As String
Dim r_str_Libro    As String
Dim r_str_Asiento  As String
Dim r_str_Resul    As String

   If MsgBox("¿Esta seguro que desea realizar esta operación de reversa?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Reversa de Contabilización(1)
   If moddat_g_int_FlgAct = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_REGVEN_REVERSA ( "
      g_str_Parame = g_str_Parame & " '" & Trim(pnl_Codigo.Caption) & "') " 'REGVEN_CODVEN
                                                                                                                                                                                                                    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         MsgBox "No se pudo completar la operación.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If (g_rst_Genera!as_resul = 1) Then
          MsgBox "Se completo la operación de reversa satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
          'Estados Botones
          cmd_Reversa.Enabled = False
          cmd_Grabar.Enabled = True
          'Habilitar Contenedores
          fs_HabCtrl_Reversa (True)
          Call gs_SetFocus(txt_Descrip)
          Call frm_Ctb_RegVen_01.fs_BuscarComp
          
          Unload Me
      Else
          MsgBox "Favor de verificar la operación de reversa.", vbInformation, modgen_g_str_NomPlt
      End If
   End If

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   l_dbl_IniFrm = False
   cmd_Reversa.Visible = False
   moddat_g_int_Situac = 0
      
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then
      'CONSULTAR
      pnl_Titulo.Caption = "Registro de Ventas - Consultar"
      cmd_Grabar.Visible = False
      Call fs_Cargar_RegCom
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then
      'INSERTAR
      pnl_Titulo.Caption = "Registro de Ventas - Adicionar"
   ElseIf moddat_g_int_FlgGrb = 2 Then
      'EDITAR
      pnl_Titulo.Caption = "Registro de Ventas - Modificar"
      Call fs_Cargar_RegCom
      cmd_Reversa.Visible = False
      cmd_Grabar.Enabled = True
      Select Case moddat_g_int_FlgAct
             Case 0: pnl_Titulo.Caption = "Registro de Ventas - Modificar"
             Case 1: pnl_Titulo.Caption = "Registro de Ventas - Reversa de Contabilización"
                     cmd_Reversa.Visible = True
                     cmd_Grabar.Enabled = False
                     'desabilitar controles
                     fs_HabCtrl_Reversa (False)
      End Select
   End If
   
   l_dbl_IniFrm = True
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_Cadena As String
Dim r_int_Contar As Integer

   l_dbl_IGV = 0
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   Call fs_CargaMntPardes(cmb_TipCbtPrv, "123")
   Call fs_CargaMntPardes(cmb_TipCbtRef, "123")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   
   'cargar deba y haber
   Call moddat_gs_Carga_LisIte(cmb_GravDH_01, l_arr_DebHab, 1, 255, 1)
   cmb_GravDH_01.Clear
   For r_int_Contar = 1 To UBound(l_arr_DebHab)
       cmb_GravDH_01.AddItem Left(l_arr_DebHab(r_int_Contar).Genera_Nombre, 1)
       cmb_NGrvDH_01.AddItem Left(l_arr_DebHab(r_int_Contar).Genera_Nombre, 1)
   Next
      
   'cargar las cuentas contables
   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaGvd_01, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   
   cmb_CtaGvd_01.Clear
   cmb_CtaNoGvd_01.Clear
   For r_int_Contar = 1 To UBound(l_arr_CtaCtb)
       Select Case Trim(l_arr_CtaCtb(r_int_Contar).Genera_Codigo)
              Case "561901010101", "562901010101", "571101010101", "561901010103", "561901010102"
                   cmb_CtaGvd_01.AddItem l_arr_CtaCtb(r_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(r_int_Contar).Genera_Nombre
                   cmb_CtaNoGvd_01.AddItem l_arr_CtaCtb(r_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(r_int_Contar).Genera_Nombre
       End Select
   Next
      
   'cargar igv
   l_dbl_IGV = moddat_gf_Consulta_ParVal("001", "001")
   l_dbl_IGV = l_dbl_IGV / 100
   Call moddat_gs_FecSis
   
   pnl_Igv.Caption = "251703020101"
   
   cmb_PorCobrar.Clear
   cmb_PorCobrar.AddItem ("151719010106")
   cmb_PorCobrar.AddItem ("152719010109")
   cmb_PorCobrar.AddItem ("151719010104")

   ReDim l_arr_MaePrv(0)
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
   cmb_Moneda.ListIndex = 0
   'DETERMINACION
   cmb_CtaGvd_01.Text = ""
   cmb_CtaNoGvd_01.Text = ""
   'pnl_Igv.Caption = ""
   cmb_PorCobrar.ListIndex = 0
   
   ipp_ImpGrav_01.Text = "0.00"
   ipp_ImpNGrv_01.Text = "0.00"
   pnl_ImpIgv.Caption = "0.00 "
   pnl_ImpPpg.Caption = "0.00 "
   
   'REFERENCIA
   ipp_FchEmiRef.Text = ""
   cmb_TipCbtRef.ListIndex = -1
   txt_NumSerieRef.Text = ""
   txt_NumRef.Text = ""
   
   'SE ACTIVARA PARA EL COMPROBANTE 7
   ipp_FchEmiRef.Enabled = False
   cmb_TipCbtRef.Enabled = False
   txt_NumSerieRef.Enabled = False
   txt_NumRef.Enabled = False
   
   'DATOS STATICOS
   'l_str_CtaIGV = "251703020101"
   'pnl_Igv.Caption = ""
   Call ipp_FchEmiPrv_LostFocus
   l_dbl_IniFrm = True
   Call fs_Calcular_Determ
   l_dbl_IniFrm = False
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Function fs_ValNumDoc() As Boolean
Dim r_str_NumDoc As String
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

Private Sub fs_CargaMntPardes(p_Combo As ComboBox, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES A "
   g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   If Trim(p_CodGrp) = "118" Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('009999','000006','000001') "
   End If
   If Trim(p_CodGrp) = "123" Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000003','000007','000008','009999') "
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

Private Function fs_NumDoc(p_Cadena As String) As String
   fs_NumDoc = ""
   If (cmb_TipDoc.ListIndex > -1) Then
      If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
          If p_Cadena <> "" Then
             fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
          End If
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

Private Sub ipp_FchEmiRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipCbtRef)
   End If
End Sub

Private Sub ipp_FchVenc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If cmb_Moneda.Enabled = False Then
          Call gs_SetFocus(chk_RetBien)
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

Private Sub ipp_ImpNGrv_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_NGrvDH_01)
   End If
End Sub

Private Sub ipp_ImpNGrv_01_LostFocus()
   Call fs_Calcular_Determ
End Sub

Private Sub txt_Descrip_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchCtb)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Proveedor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Descrip)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
   End If
End Sub

Private Sub cmb_TipCbtPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumSeriePrv)
   End If
End Sub

Private Sub txt_NumPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchEmiPrv)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
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

Private Sub txt_NumSerieRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumRef)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_ActivaRefer()
  If (cmb_TipCbtPrv.ListIndex > -1) Then
  
      If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7) Then '07 - N/C
          cmb_GravDH_01.ListIndex = 0
          cmb_NGrvDH_01.ListIndex = 0
          pnl_IgvDH.Caption = "D"
          pnl_PpgDH.Caption = "H"
          'SE ACTIVA
          ipp_FchEmiRef.Enabled = True
          cmb_TipCbtRef.Enabled = True
          txt_NumSerieRef.Enabled = True
          txt_NumRef.Enabled = True
      Else
          cmb_GravDH_01.ListIndex = 1
          cmb_NGrvDH_01.ListIndex = 1
          pnl_IgvDH.Caption = "H"
          pnl_PpgDH.Caption = "D"
          If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8 Then
              ipp_FchEmiRef.Enabled = True
              cmb_TipCbtRef.Enabled = True
              txt_NumSerieRef.Enabled = True
              txt_NumRef.Enabled = True
          Else
              'SE DESACTIVA
              ipp_FchEmiRef.Enabled = False
              cmb_TipCbtRef.Enabled = False
              txt_NumSerieRef.Enabled = False
              txt_NumRef.Enabled = False
              ipp_FchEmiRef.Text = ""
              cmb_TipCbtRef.ListIndex = -1
              txt_NumSerieRef.Text = ""
              txt_NumRef.Text = ""
          End If
      End If
      
  End If
End Sub

Private Sub fs_CargarPrv()
   'ReDim l_arr_CtaCteSol(0)
   'ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
   cmb_Proveedor.Clear
   cmb_Proveedor.Text = ""
   'cmb_Banco.Clear
   'cmb_CtaCte.Clear
   
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

Private Sub fs_Calcular_Determ()
Dim r_dbl_Aux     As Double
Dim r_dbl_ImpPp   As Double
Dim r_dbl_ImpAux  As Double
Dim r_bol_Estado  As Boolean
      
   r_bol_Estado = False
   r_dbl_ImpAux = 0
   r_dbl_ImpPp = 0
   
   If cmb_TipCbtPrv.ListIndex = -1 Or cmb_Moneda.ListIndex = -1 Then
       ipp_ImpGrav_01.Text = "0.00"
       ipp_ImpNGrv_01.Text = "0.00"
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
   'COMPROBANTE SIEMPRE 07 - N/C
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7) Then
       'calculo del igv
       pnl_IgvDH.Caption = "D"
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01))) * l_dbl_IGV
       pnl_ImpIgv.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
       
       'calculo por pagar
       pnl_PpgDH.Caption = "H"
       r_dbl_ImpAux = 0
       r_dbl_ImpAux = Math.Abs((fs_HabTxt(ipp_ImpGrav_01)) + _
                      CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + CDbl(fs_HabPnl(pnl_ImpIgv)))
       pnl_ImpPpg.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
   Else 'COMPROBANTE DISTINTO A 07 - N/C
       'calculo del igv
       pnl_IgvDH.Caption = "H"
       r_dbl_ImpAux = Math.Abs((CDbl(fs_HabTxt(ipp_ImpGrav_01))) * l_dbl_IGV)
       pnl_ImpIgv.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
       
       'calculo por pagar
       pnl_PpgDH.Caption = "D"
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + _
                      CDbl(fs_HabPnl(pnl_ImpIgv)))
       pnl_ImpPpg.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
   End If
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
      r_str_CadAux = Mid(cmb_CtaNoGvd_01.Text, 1, l_int_TopNiv)
      If Len(r_str_CadAux) > 0 Then
         r_bol_Estado = True
         pnl_Igv.Caption = r_str_CadAux
      End If
   End If
End Sub

Function fs_HabTxt(p_Objeto As fpDoubleSingle) As String
   If Trim(p_Objeto.Name) = Trim(ipp_ImpGrav_01.Name) Then
      fs_HabTxt = ipp_ImpGrav_01.Text
      If Trim(cmb_GravDH_01.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpGrav_01.Text
      End If
   End If
      
   If Trim(p_Objeto.Name) = Trim(ipp_ImpNGrv_01.Name) Then
      fs_HabTxt = ipp_ImpNGrv_01.Text
      If Trim(cmb_NGrvDH_01.Text) = "H" Then
         fs_HabTxt = "-" & ipp_ImpNGrv_01.Text
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

Private Sub fs_Cargar_RegCom()
Dim r_dbl_Import     As Double

   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   r_dbl_Import = 0
   Call gs_SetFocus(cmb_TipDoc)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT LPAD(REGVEN_CODVEN,10,'0') AS REGVEN_CODVEN, REGVEN_TIPDOC, REGVEN_NUMDOC, REGVEN_DESCRP, REGVEN_FECCTB, REGVEN_TIPCPB, REGVEN_NSERIE, "
   g_str_Parame = g_str_Parame & "        REGVEN_NROCOM, REGVEN_FECEMI, REGVEN_FECVTO, REGVEN_CODMON, REGVEN_TIPCAM, REGVEN_CNT_GRV1, "
   g_str_Parame = g_str_Parame & "        REGVEN_CNT_NGV1, REGVEN_CNT_IGV1, REGVEN_CNT_PPG1, "
   g_str_Parame = g_str_Parame & "        REGVEN_DEB_GRV1, REGVEN_HAB_GRV1, REGVEN_DEB_NGV1, REGVEN_HAB_NGV1, "
   g_str_Parame = g_str_Parame & "        REGVEN_DEB_IGV1, REGVEN_HAB_IGV1, REGVEN_DEB_PPG1, REGVEN_HAB_PPG1, "
   g_str_Parame = g_str_Parame & "        REGVEN_REF_FECEMI, REGVEN_REF_TIPCPB, REGVEN_REF_NSERIE, REGVEN_REF_NROCOM, "
   g_str_Parame = g_str_Parame & "        B.MAEPRV_RAZSOC, REGVEN_RETBIE "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_REGVEN A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON A.REGVEN_TIPDOC = B.MAEPRV_TIPDOC AND A.REGVEN_NUMDOC = B.MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE REGVEN_CODVEN = " & moddat_g_str_Codigo

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!regven_TipDoc)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!regven_NumDoc & "", 0)
           
      pnl_Codigo.Caption = Trim(g_rst_Princi!regven_CodVen & "")
      txt_Descrip.Text = Trim(g_rst_Princi!regven_Descrp & "")
      If Not IsNull(g_rst_Princi!regven_FecCtb) Then
          ipp_FchCtb.Text = gf_FormatoFecha(g_rst_Princi!regven_FecCtb)
      End If
      If Not IsNull(g_rst_Princi!regven_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtPrv, g_rst_Princi!regven_TipCpb)
         'Call fs_ActivaRefer
      End If
      txt_NumSeriePrv.Text = Trim(g_rst_Princi!regven_Nserie & "")
      txt_NumPrv.Text = Trim(g_rst_Princi!regven_NroCom & "")
      ipp_FchEmiPrv.Text = gf_FormatoFecha(g_rst_Princi!regven_FecEmi)
      If Not IsNull(g_rst_Princi!regven_FecVto) Then
         ipp_FchVenc.Text = gf_FormatoFecha(g_rst_Princi!regven_FecVto)
      End If
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!regven_CodMon)
      pnl_TipCambio.Caption = Format(g_rst_Princi!regven_TipCam, "###,###,##0.000000") & " "
         
      chk_RetBien.Value = g_rst_Princi!regven_retbie
         
      cmb_CtaGvd_01.ListIndex = fs_ComboIndex(cmb_CtaGvd_01, g_rst_Princi!regven_Cnt_Grv1 & "", l_int_TopNiv)
      cmb_CtaNoGvd_01.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_01, g_rst_Princi!regven_Cnt_Ngv1 & "", l_int_TopNiv)
      pnl_Igv.Caption = Trim(g_rst_Princi!regven_Cnt_Igv1 & "")
      
      If Not IsNull(g_rst_Princi!regven_Cnt_Ppg1) Then
         cmb_PorCobrar.Text = Trim(g_rst_Princi!regven_Cnt_Ppg1 & "")
      End If
      'GRAVADO 1
      If (g_rst_Princi!regven_Deb_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 0
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!regven_Deb_Grv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!regven_Hab_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 1
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!regven_Hab_Grv1, "###,###,##0.00")
      End If
      'NO GRAVADO 1
      If (g_rst_Princi!regven_Deb_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 0
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!regven_Deb_Ngv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!regven_Hab_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 1
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!regven_Hab_Ngv1, "###,###,##0.00")
      End If
      'IGV
      If (g_rst_Princi!regven_Deb_Igv1 > 0) Then
          pnl_IgvDH.Caption = "D"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!regven_Deb_Igv1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regven_Hab_Igv1 > 0) Then
          pnl_IgvDH.Caption = "H"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!regven_Hab_Igv1, "###,###,##0.00") & " "
      End If
      'CUENTAS POR PAGAR
      If (g_rst_Princi!regven_Deb_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "D"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!regven_Deb_Ppg1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regven_Hab_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "H"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!regven_Hab_Ppg1, "###,###,##0.00") & " "
      End If
      '---------------------------------------------------------------------------------------
      If Not IsNull(g_rst_Princi!regven_Ref_FecEmi) Then
         ipp_FchEmiRef.Text = gf_FormatoFecha(g_rst_Princi!regven_Ref_FecEmi)
      End If
      If Not IsNull(g_rst_Princi!regven_Ref_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtRef, g_rst_Princi!regven_Ref_TipCpb)
      End If
      txt_NumSerieRef.Text = Format(g_rst_Princi!regven_Ref_Nserie & "")
      txt_NumRef.Text = Format(g_rst_Princi!regven_Ref_NroCom & "")
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

Private Sub fs_Desabilitar()
   cmb_TipDoc.Enabled = False
   'txt_NumDoc.Enabled = False
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
   cmb_CtaNoGvd_01.Enabled = False
   cmb_PorCobrar.Enabled = False
   
   ipp_ImpGrav_01.Enabled = False
   ipp_ImpNGrv_01.Enabled = False
   cmb_GravDH_01.Enabled = False
   cmb_NGrvDH_01.Enabled = False
      
   ipp_FchEmiRef.Enabled = False
   cmb_TipCbtRef.Enabled = False
   txt_NumSerieRef.Enabled = False
   txt_NumRef.Enabled = False
   
   chk_RetBien.Enabled = False
End Sub

Private Sub fs_HabCtrl_Reversa(p_Estado As Boolean)
   'PROVEEDOR
   txt_Descrip.Enabled = p_Estado
   'COMPROBANTE
   ipp_FchCtb.Enabled = p_Estado
   cmb_TipCbtPrv.Enabled = p_Estado
   txt_NumSeriePrv.Enabled = p_Estado
   ipp_FchEmiPrv.Enabled = p_Estado
   cmb_Moneda.Enabled = p_Estado
   txt_NumPrv.Enabled = p_Estado
   ipp_FchVenc.Enabled = p_Estado
   chk_RetBien.Enabled = p_Estado
   'DETERMINACION
   ipp_ImpGrav_01.Enabled = p_Estado
   ipp_ImpNGrv_01.Enabled = p_Estado
   '---------
   cmb_GravDH_01.Enabled = p_Estado
   cmb_NGrvDH_01.Enabled = p_Estado
   '---------
   cmb_CtaGvd_01.Enabled = p_Estado
   cmb_CtaNoGvd_01.Enabled = p_Estado
   cmb_PorCobrar.Enabled = p_Estado
   'REFERENCIAS
   If (p_Estado = False) Then
       ipp_FchEmiRef.Enabled = False
       cmb_TipCbtRef.Enabled = False
       txt_NumSerieRef.Enabled = False
       txt_NumRef.Enabled = False
   Else
       If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8) Then '07 - N/C
          ipp_FchEmiRef.Enabled = True
          cmb_TipCbtRef.Enabled = True
          txt_NumSerieRef.Enabled = True
          txt_NumRef.Enabled = True
       Else
          ipp_FchEmiRef.Enabled = False
          cmb_TipCbtRef.Enabled = False
          txt_NumSerieRef.Enabled = False
          txt_NumRef.Enabled = False
       End If
   End If
End Sub


