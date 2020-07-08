VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegCom_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
   Icon            =   "GesCtb_frm_187.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   8130
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   14085
      _Version        =   65536
      _ExtentX        =   24844
      _ExtentY        =   14340
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
         Height          =   630
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Width           =   13950
         _Version        =   65536
         _ExtentX        =   24606
         _ExtentY        =   1111
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
            Left            =   630
            TabIndex        =   53
            Top             =   180
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Cuentas por Pagar"
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
            Picture         =   "GesCtb_frm_187.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_DatPrv 
         Height          =   1155
         Left            =   60
         TabIndex        =   54
         Top             =   1425
         Width           =   13950
         _Version        =   65536
         _ExtentX        =   24606
         _ExtentY        =   2028
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
            Left            =   6780
            TabIndex        =   1
            Top             =   360
            Width           =   6990
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   3850
         End
         Begin VB.TextBox txt_Descrip 
            Height          =   315
            Left            =   6780
            MaxLength       =   60
            TabIndex        =   3
            Top             =   690
            Width           =   6990
         End
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1470
            TabIndex        =   2
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
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   5700
            TabIndex        =   102
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   420
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   5700
            TabIndex        =   60
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
            TabIndex        =   59
            Top             =   60
            Width           =   1740
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   650
         Left            =   60
         TabIndex        =   55
         Top             =   730
         Width           =   13950
         _Version        =   65536
         _ExtentX        =   24606
         _ExtentY        =   1147
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
            Height          =   580
            Left            =   630
            Picture         =   "GesCtb_frm_187.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Reversa del Asiento"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   580
            Left            =   13320
            Picture         =   "GesCtb_frm_187.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   580
            Left            =   30
            Picture         =   "GesCtb_frm_187.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_DatCbt 
         Height          =   2100
         Left            =   60
         TabIndex        =   56
         Top             =   2640
         Width           =   8490
         _Version        =   65536
         _ExtentX        =   14975
         _ExtentY        =   3704
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
         Begin VB.TextBox txt_NumPrv 
            Height          =   315
            Left            =   6780
            MaxLength       =   7
            TabIndex        =   7
            Top             =   930
            Width           =   1515
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1590
            Width           =   2175
         End
         Begin VB.ComboBox cmb_TipCbtPrv 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   600
            Width           =   3850
         End
         Begin VB.TextBox txt_NumSeriePrv 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   6
            Top             =   930
            Width           =   1420
         End
         Begin EditLib.fpDateTime ipp_FchCtb 
            Height          =   315
            Left            =   1470
            TabIndex        =   4
            Top             =   270
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2505
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
            Left            =   6780
            TabIndex        =   9
            Top             =   1260
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
            Left            =   1470
            TabIndex        =   8
            Top             =   1260
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2505
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
            Left            =   6780
            TabIndex        =   11
            Top             =   1590
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   5700
            TabIndex        =   73
            Top             =   1650
            Width           =   930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   1650
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vcto.:"
            Height          =   195
            Left            =   5700
            TabIndex        =   71
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Numero:"
            Height          =   195
            Left            =   5700
            TabIndex        =   69
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Serie:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   990
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
            Left            =   120
            TabIndex        =   63
            Top             =   30
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Contable:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   660
            Width           =   990
         End
      End
      Begin Threed.SSPanel pnl_DatReq 
         Height          =   2100
         Left            =   8610
         TabIndex        =   64
         Top             =   2640
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   3704
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
         Begin VB.ComboBox cmb_CodDetrc 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   960
            Width           =   1500
         End
         Begin VB.ComboBox cmb_AppTrib 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   615
            Width           =   1500
         End
         Begin VB.TextBox txt_DocDetrc 
            Height          =   315
            Left            =   1350
            MaxLength       =   18
            TabIndex        =   15
            Top             =   1290
            Width           =   1500
         End
         Begin EditLib.fpDateTime ipp_FchDetrc 
            Height          =   315
            Left            =   3750
            TabIndex        =   16
            Top             =   1290
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2505
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
         Begin Threed.SSPanel pnl_Padron 
            Height          =   315
            Left            =   1350
            TabIndex        =   12
            Top             =   270
            Width           =   3900
            _Version        =   65536
            _ExtentX        =   6879
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
         Begin Threed.SSPanel pnl_PagCod 
            Height          =   315
            Left            =   1350
            TabIndex        =   106
            Top             =   1635
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
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
         Begin Threed.SSPanel pnl_PagFec 
            Height          =   315
            Left            =   3750
            TabIndex        =   107
            Top             =   1635
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2505
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
            Caption         =   "Código Pago:"
            Height          =   195
            Left            =   90
            TabIndex        =   105
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "F. Pago:"
            Height          =   195
            Left            =   2970
            TabIndex        =   104
            Top             =   1710
            Width           =   600
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "F. Detrac.:"
            Height          =   195
            Left            =   2970
            TabIndex        =   76
            Top             =   1365
            Width           =   750
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Detrac.:"
            Height          =   195
            Left            =   90
            TabIndex        =   75
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Detrac.:"
            Height          =   195
            Left            =   90
            TabIndex        =   74
            Top             =   1020
            Width           =   1110
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Aplic. Tributaria :"
            Height          =   195
            Left            =   90
            TabIndex        =   67
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label lbl_Padron 
            AutoSize        =   -1  'True
            Caption         =   "Padrón:"
            Height          =   195
            Left            =   90
            TabIndex        =   66
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Requisitos"
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
            TabIndex        =   65
            Top             =   60
            Width           =   900
         End
      End
      Begin Threed.SSPanel pnl_DatDet 
         Height          =   3210
         Left            =   60
         TabIndex        =   77
         Top             =   4800
         Width           =   8490
         _Version        =   65536
         _ExtentX        =   14975
         _ExtentY        =   5662
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
         Begin VB.ComboBox cmb_CtaNoGvd_02 
            Height          =   315
            Left            =   3570
            TabIndex        =   28
            Top             =   1350
            Width           =   4750
         End
         Begin VB.ComboBox cmb_CtaNoGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   25
            Top             =   1020
            Width           =   4750
         End
         Begin VB.ComboBox cmb_CtaGvd_02 
            Height          =   315
            Left            =   3570
            TabIndex        =   22
            Top             =   690
            Width           =   4750
         End
         Begin VB.ComboBox cmb_CtaGvd_01 
            Height          =   315
            Left            =   3570
            TabIndex        =   19
            Top             =   360
            Width           =   4750
         End
         Begin VB.ComboBox cmb_NGrvDH_02 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1350
            Width           =   645
         End
         Begin VB.ComboBox cmb_NGrvDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1020
            Width           =   645
         End
         Begin VB.ComboBox cmb_GravDH_02 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   690
            Width           =   645
         End
         Begin VB.ComboBox cmb_GravDH_01 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   645
         End
         Begin VB.ComboBox cmb_PorPagar 
            Height          =   315
            Left            =   3580
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2670
            Width           =   4750
         End
         Begin EditLib.fpDoubleSingle ipp_ImpGrav_01 
            Height          =   315
            Left            =   1470
            TabIndex        =   17
            Top             =   360
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
         Begin EditLib.fpDoubleSingle ipp_ImpGrav_02 
            Height          =   315
            Left            =   1470
            TabIndex        =   20
            Top             =   690
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
         Begin EditLib.fpDoubleSingle ipp_ImpNGrv_01 
            Height          =   315
            Left            =   1470
            TabIndex        =   23
            Top             =   1020
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
         Begin EditLib.fpDoubleSingle ipp_ImpNGrv_02 
            Height          =   315
            Left            =   1470
            TabIndex        =   26
            Top             =   1350
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
         Begin Threed.SSPanel pnl_Igv 
            Height          =   315
            Left            =   3585
            TabIndex        =   31
            Top             =   1680
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
         Begin Threed.SSPanel pnl_Retencion 
            Height          =   315
            Left            =   3585
            TabIndex        =   34
            Top             =   2010
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
         Begin Threed.SSPanel pnl_Detraccion 
            Height          =   315
            Left            =   3585
            TabIndex        =   37
            Top             =   2340
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
            Left            =   1470
            TabIndex        =   29
            Top             =   1680
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_ImpRet 
            Height          =   315
            Left            =   1470
            TabIndex        =   32
            Top             =   2010
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_ImpDet 
            Height          =   315
            Left            =   1470
            TabIndex        =   35
            Top             =   2340
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
            Left            =   1470
            TabIndex        =   38
            Top             =   2670
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
            TabIndex        =   30
            Top             =   1680
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
         Begin Threed.SSPanel pnl_RetDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   33
            Top             =   2010
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
         Begin Threed.SSPanel pnl_DetDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   36
            Top             =   2340
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
         Begin Threed.SSPanel pnl_PpgDH 
            Height          =   315
            Left            =   2910
            TabIndex        =   39
            Top             =   2670
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
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CUENTAS CONTABLES"
            Height          =   195
            Left            =   5190
            TabIndex        =   101
            Top             =   120
            Width           =   1770
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "D/H"
            Height          =   195
            Left            =   3060
            TabIndex        =   100
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Retención:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Detracción:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   2370
            Width           =   825
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Por Pagar:"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   2700
            Width           =   750
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   750
            Width           =   660
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   420
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
            TabIndex        =   82
            Top             =   60
            Width           =   1230
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "No Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "No Gravado:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1410
            Width           =   915
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   1710
            Width           =   270
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "IMPORTE"
            Height          =   195
            Left            =   1860
            TabIndex        =   78
            Top             =   120
            Width           =   735
         End
      End
      Begin Threed.SSPanel pnl_DatFin 
         Height          =   1770
         Left            =   8610
         TabIndex        =   88
         Top             =   4800
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   3122
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
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   990
            Width           =   3090
         End
         Begin VB.TextBox txt_CtrCosto 
            Height          =   315
            Left            =   1350
            MaxLength       =   18
            TabIndex        =   42
            Top             =   660
            Width           =   3090
         End
         Begin VB.ComboBox cmb_CatCtb 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   330
            Width           =   3090
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1320
            Width           =   3090
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   90
            TabIndex        =   99
            Top             =   1050
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
            TabIndex        =   92
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Categ. Contable:"
            Height          =   195
            Left            =   90
            TabIndex        =   91
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Centro Costos:"
            Height          =   195
            Left            =   90
            TabIndex        =   90
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   90
            TabIndex        =   89
            Top             =   1380
            Width           =   555
         End
      End
      Begin Threed.SSPanel pnl_DatRef 
         Height          =   1395
         Left            =   8610
         TabIndex        =   93
         Top             =   6615
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   2469
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
            Left            =   3750
            MaxLength       =   7
            TabIndex        =   48
            Top             =   945
            Width           =   1420
         End
         Begin VB.ComboBox cmb_TipCbtRef 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   600
            Width           =   3840
         End
         Begin VB.TextBox txt_NumSerieRef 
            Height          =   315
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   47
            Top             =   945
            Width           =   1400
         End
         Begin EditLib.fpDateTime ipp_FchEmiRef 
            Height          =   315
            Left            =   1350
            TabIndex        =   45
            Top             =   270
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2469
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
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   2970
            TabIndex        =   98
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Serie:"
            Height          =   195
            Left            =   90
            TabIndex        =   97
            Top             =   990
            Width           =   405
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   90
            TabIndex        =   96
            Top             =   660
            Width           =   990
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   90
            TabIndex        =   95
            Top             =   330
            Width           =   1080
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
            Left            =   90
            TabIndex        =   94
            Top             =   30
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegCom_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_IGV           As Double
Dim l_dbl_IniFrm        As Boolean
Dim l_str_CtaIGV        As String
Dim l_arr_CodDet()      As moddat_tpo_Genera
Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_arr_DebHab()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_int_Contar        As Integer
Dim l_int_TopNiv        As Integer
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer
Dim l_str_NCtaBN        As String

Private Sub cmb_AppTrib_Click()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_AppTrib_LostFocus()
   Call fs_Calcular_Determ
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

Private Sub cmb_CodDetrc_Click()
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_CtaGvd_01_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpGrav_02)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_CtaGvd_01_LostFocus()
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
       Call gs_SetFocus(cmb_PorPagar)
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

Private Sub fs_ActivaRefer()
  If (cmb_TipCbtPrv.ListIndex > -1) Then
       'If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88) Then '(07-N/C) - (88-DEVOLUCIONES)
       If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88) Then '(88-DEVOLUCIONES)
           'SE ACTIVA
           ipp_FchEmiRef.Enabled = True
           cmb_TipCbtRef.Enabled = True
           txt_NumSerieRef.Enabled = True
           txt_NumRef.Enabled = True
           If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88 Then
              ipp_FchEmiRef.Enabled = False
              cmb_TipCbtRef.Enabled = False
              txt_NumSerieRef.Enabled = False
              txt_NumRef.Enabled = False
           End If
       Else
           'If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8 or  Then
           If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Then
              ipp_FchEmiRef.Enabled = True
              cmb_TipCbtRef.Enabled = True
              txt_NumSerieRef.Enabled = True
              txt_NumRef.Enabled = True
              
              'If moddat_g_int_InsAct = 0 Then
              '   'si es registro compras
              '   cmb_FacAsoc.Enabled = True
              '   Call fs_FacAsoc_RegCom
              'ElseIf moddat_g_int_InsAct = 1 Then
              '   'si es ENTREGAS A RENDIR
              '   cmb_FacAsoc.Enabled = True
              '   fs_FacAsoc_EntRen
              'End If
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
              
              'cmb_FacAsoc.Enabled = False
              'cmb_FacAsoc.ListIndex = -1
           End If
       End If
       
       Call fs_Asigna_DebHab(True)
  End If
End Sub

Private Sub fs_Asigna_DebHab(p_Nuevo As Boolean)
    'If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88) Then '(88-DEVOLUCIONES)
    If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Then '(88-DEVOLUCIONES)
        If p_Nuevo = True Then
           cmb_GravDH_01.ListIndex = 1 'H
           cmb_GravDH_02.ListIndex = 1 'H
           cmb_NGrvDH_01.ListIndex = 1 'H
           cmb_NGrvDH_02.ListIndex = 1 'H
        Else
           If cmb_GravDH_01.ListIndex = -1 Then
              cmb_GravDH_01.ListIndex = 1 'H
           End If
           If cmb_GravDH_02.ListIndex = -1 Then
              cmb_GravDH_02.ListIndex = 1 'H
           End If
           If cmb_NGrvDH_01.ListIndex = -1 Then
              cmb_NGrvDH_01.ListIndex = 1 'H
           End If
           If cmb_NGrvDH_02.ListIndex = -1 Then
              cmb_NGrvDH_02.ListIndex = 1 'H
           End If
        End If
        pnl_IgvDH.Caption = "H"
        pnl_RetDH.Caption = "D"
        pnl_DetDH.Caption = "D"
        pnl_PpgDH.Caption = "D"
    Else
        If p_Nuevo = True Then
           cmb_GravDH_01.ListIndex = 0 'D
           cmb_GravDH_02.ListIndex = 0 'D
           cmb_NGrvDH_01.ListIndex = 0 'D
           cmb_NGrvDH_02.ListIndex = 0 'D
        Else
           If cmb_GravDH_01.ListIndex = -1 Then
              cmb_GravDH_01.ListIndex = 0 'D
           End If
           If cmb_GravDH_02.ListIndex = -1 Then
              cmb_GravDH_02.ListIndex = 0 'D
           End If
           If cmb_NGrvDH_01.ListIndex = -1 Then
              cmb_NGrvDH_01.ListIndex = 0 'D
           End If
           If cmb_NGrvDH_02.ListIndex = -1 Then
              cmb_NGrvDH_02.ListIndex = 0 'D
           End If
        End If
        pnl_IgvDH.Caption = "D"
        pnl_RetDH.Caption = "H"
        pnl_DetDH.Caption = "H"
        pnl_PpgDH.Caption = "H"
   End If
End Sub

'Private Sub fs_FacAsoc_RegCom()
'   Screen.MousePointer = 11
'
'   ReDim l_arr_FacAso(0)
'   cmb_FacAsoc.Clear
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT A.REGCOM_CODCOM, DECODE(A.REGCOM_CODMON,1,'S/.','US$') AS MONEDA, REGCOM_DEB_PPG1 + REGCOM_HAB_PPG1 AS PORPAGAR "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_REGCOM A "
'   g_str_Parame = g_str_Parame & "  WHERE A.REGCOM_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
'   g_str_Parame = g_str_Parame & "    AND A.REGCOM_SITUAC = 1 "
'   g_str_Parame = g_str_Parame & "    AND A.REGCOM_FLGCNT = 0 "
'   g_str_Parame = g_str_Parame & "    AND A.REGCOM_TIPTAB = 3 "
'   g_str_Parame = g_str_Parame & "    AND A.REGCOM_CODCOM NOT IN (SELECT B.REGCOM_CODFAC  "
'   g_str_Parame = g_str_Parame & "                                  FROM CNTBL_REGCOM B  "
'   g_str_Parame = g_str_Parame & "                                 WHERE B.REGCOM_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
'   g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_SITUAC = 1  "
'   'g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_FLGCNT = 0  "
'   'g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_CODFAC IS NOT NULL  "
'   If moddat_g_int_FlgGrb = 2 Then
'      g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_CODFAC IS NOT NULL  "
'      g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_CODCOM NOT IN ('" & Trim(pnl_Codigo.Caption) & "'))"
'   Else
'      g_str_Parame = g_str_Parame & "                                   AND B.REGCOM_CODFAC IS NOT NULL)  "
'   End If
'
'   If Trim(pnl_Codigo.Caption) <> "" Then
'      g_str_Parame = g_str_Parame & " AND A.REGCOM_CODCOM NOT IN ('" & Trim(pnl_Codigo.Caption) & "')"
'   End If
'   g_str_Parame = g_str_Parame & "  ORDER BY A.REGCOM_CODCOM ASC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
'      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
'      g_rst_GenAux.Close
'      Set g_rst_GenAux = Nothing
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   ReDim l_arr_FacAso(0)
'   g_rst_GenAux.MoveFirst
'   Do While Not g_rst_GenAux.EOF
'      cmb_FacAsoc.AddItem (g_rst_GenAux!regcom_CodCom & "  -  PPg. " & g_rst_GenAux!Moneda & " " & Format(g_rst_GenAux!PORPAGAR, "###,###,###,##0.00"))
'      ReDim Preserve l_arr_FacAso(UBound(l_arr_FacAso) + 1)
'      l_arr_FacAso(UBound(l_arr_FacAso)).Genera_Codigo = Trim(g_rst_GenAux!regcom_CodCom)
'      g_rst_GenAux.MoveNext
'   Loop
'
'   g_rst_GenAux.Close
'   Set g_rst_GenAux = Nothing
'   Screen.MousePointer = 0
'End Sub

'Private Sub fs_FacAsoc_EntRen()
'   Screen.MousePointer = 11
'
'   ReDim l_arr_FacAso(0)
'   cmb_FacAsoc.Clear
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT A.CAJDET_CODDET, DECODE(A.CAJDET_CODMON,1,'S/.','US$') AS MONEDA, CAJDET_DEB_PPG1 + CAJDET_HAB_PPG1 AS PORPAGAR  "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC_DET A  "
'   g_str_Parame = g_str_Parame & "  WHERE A.CAJDET_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
'   g_str_Parame = g_str_Parame & "    AND A.CAJDET_SITUAC = 1  "
'   g_str_Parame = g_str_Parame & "    AND A.CAJDET_FLGPRC = 0  "
'   g_str_Parame = g_str_Parame & "    AND A.CAJDET_CODCAJ = " & CLng(moddat_g_str_Codigo)
'   g_str_Parame = g_str_Parame & "    AND A.CAJDET_TIPTAB = 2  "
'   g_str_Parame = g_str_Parame & "    AND A.CAJDET_CODDET NOT IN (SELECT B.CAJDET_CODFAC  "
'   g_str_Parame = g_str_Parame & "                                  FROM CNTBL_CAJCHC_DET B  "
'   g_str_Parame = g_str_Parame & "                                 WHERE B.CAJDET_CODMON = " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
'   g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_SITUAC = 1  "
'   'g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_FLGPRC = 0  "
'   g_str_Parame = g_str_Parame & "                                   AND A.CAJDET_CODCAJ = " & CLng(moddat_g_str_Codigo)
'   g_str_Parame = g_str_Parame & "                                   AND A.CAJDET_TIPTAB = 2  "
'   'g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_CODFAC IS NOT NULL  "
'   If moddat_g_int_FlgGrb = 2 Then
'      g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_CODFAC IS NOT NULL  "
'      g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_CODDET NOT IN ('" & Trim(pnl_Codigo.Caption) & "'))"
'   Else
'      g_str_Parame = g_str_Parame & "                                   AND B.CAJDET_CODFAC IS NOT NULL)  "
'   End If
'
'   If Trim(pnl_Codigo.Caption) <> "" Then
'      g_str_Parame = g_str_Parame & " AND A.CAJDET_CODDET NOT IN ('" & Trim(pnl_Codigo.Caption) & "')"
'   End If
'   g_str_Parame = g_str_Parame & "  ORDER BY A.CAJDET_CODDET ASC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
'      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
'      g_rst_GenAux.Close
'      Set g_rst_GenAux = Nothing
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   ReDim l_arr_FacAso(0)
'   g_rst_GenAux.MoveFirst
'   Do While Not g_rst_GenAux.EOF
'
'      cmb_FacAsoc.AddItem (Format(g_rst_GenAux!CajDet_CodDet, "0000000000") & "  -  PPg. " & g_rst_GenAux!Moneda & " " & Format(g_rst_GenAux!PORPAGAR, "###,###,###,##0.00"))
'
'      ReDim Preserve l_arr_FacAso(UBound(l_arr_FacAso) + 1)
'      l_arr_FacAso(UBound(l_arr_FacAso)).Genera_Codigo = Trim(g_rst_GenAux!CajDet_CodDet)
'
'      g_rst_GenAux.MoveNext
'   Loop
'
'   g_rst_GenAux.Close
'   Set g_rst_GenAux = Nothing
'   Screen.MousePointer = 0
'End Sub

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

Private Sub cmb_TipCbtPrv_LostFocus()
   Call fs_ActivaRefer
   Call fs_Calcular_Determ
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
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

Private Sub fs_Buscar_prov()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   pnl_Padron.Caption = ""
   pnl_Padron.Tag = ""
   r_str_NumDoc = ""
   l_str_NCtaBN = ""
   
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
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC, MAEPRV_CTADET "
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
       pnl_Padron.Caption = Trim(g_rst_GenAux!PADRON_1 & "")
       pnl_Padron.Tag = Trim(g_rst_GenAux!MAEPRV_PADRN1 & "")
       Call gs_SetFocus(txt_Descrip)
   End If
      
   l_str_NCtaBN = Trim(g_rst_GenAux!MaePrv_CtaDet & "")
   
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
   l_dbl_IniFrm = True
   Call ipp_FchEmiPrv_LostFocus
   l_dbl_IniFrm = False
   
   ipp_FchVenc.Text = ""
   cmb_Moneda.ListIndex = 0
   'REQUISITOS
   pnl_Padron.Caption = ""
   pnl_Padron.Tag = ""
   l_str_NCtaBN = ""
   cmb_AppTrib.ListIndex = 0
   cmb_CodDetrc.ListIndex = 0
   txt_DocDetrc.Text = ""
   ipp_FchDetrc.Text = ""
   'DETERMINACION
   cmb_CtaGvd_01.Text = ""
   cmb_CtaGvd_02.Text = ""
   cmb_CtaNoGvd_01.Text = ""
   cmb_CtaNoGvd_02.Text = ""
   pnl_Igv.Caption = ""
   pnl_Retencion.Caption = ""
   pnl_Detraccion.Caption = ""
   
   If moddat_g_int_InsAct = 1 Then
      'entragas a rendir
      cmb_PorPagar.ListIndex = 0
   Else
      'Registro de Compras
      cmb_PorPagar.ListIndex = 4
   End If
   ipp_ImpGrav_01.Text = "0.00"
   ipp_ImpGrav_02.Text = "0.00"
   ipp_ImpNGrv_01.Text = "0.00"
   ipp_ImpNGrv_02.Text = "0.00"
   pnl_ImpIgv.Caption = "0.00 "
   pnl_ImpRet.Caption = "0.00 "
   pnl_ImpDet.Caption = "0.00 "
   pnl_ImpPpg.Caption = "0.00 "
   
   'DATOS FINANCIEROS
   cmb_CatCtb.ListIndex = -1
   txt_CtrCosto.Text = ""
   cmb_Banco.ListIndex = -1
   cmb_CtaCte.ListIndex = -1
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
   l_str_CtaIGV = "251703020101"
   pnl_Igv.Caption = ""
   pnl_Retencion.Caption = "251705010103"
   pnl_Detraccion.Caption = "251602010106"
   Call ipp_FchEmiPrv_LostFocus
   l_dbl_IniFrm = True
   Call fs_Calcular_Determ
   l_dbl_IniFrm = False
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub fs_LimpiarDH()
   If CDbl(ipp_ImpGrav_01.Text) = 0 Then
      cmb_GravDH_01.ListIndex = -1
   End If
   If CDbl(ipp_ImpGrav_02.Text) = 0 Then
      cmb_GravDH_02.ListIndex = -1
   End If
   If CDbl(ipp_ImpNGrv_01.Text) = 0 Then
      cmb_NGrvDH_01.ListIndex = -1
   End If
   If CDbl(ipp_ImpNGrv_02.Text) = 0 Then
      cmb_NGrvDH_02.ListIndex = -1
   End If
   
   If CDbl(pnl_ImpIgv.Caption) = 0 Then
      pnl_IgvDH.Caption = ""
   End If
   If CDbl(pnl_ImpRet.Caption) = 0 Then
      pnl_RetDH.Caption = ""
   End If
   If CDbl(pnl_ImpDet.Caption) = 0 Then
      pnl_DetDH.Caption = ""
   End If
   If CDbl(pnl_ImpPpg.Caption) = 0 Then
      pnl_PpgDH.Caption = ""
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
   
   If Len(Trim(pnl_Padron.Tag)) = 0 Then
       MsgBox "El padrón esta vacio en el grupo requisitos, se llena al buscar un proveedor.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
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
   
   If cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) = 3 Then 'retenciones
       MsgBox "El tipo aplicación tributaria de retención no se puede usar.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_AppTrib)
       Exit Sub
   End If
   
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8) Then '(07-N/C) (08-N/D)
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
   
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88) Then '88 - DEVOLUCIONES
       If (cmb_TipCbtRef.ListIndex <> -1) Then
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
   If cmb_Banco.ListIndex = -1 And cmb_Banco.Enabled = True Then
      If MsgBox("¿En el grupo datos financieros no ha seleccionado ningún banco, desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Call gs_SetFocus(cmb_Banco)
         Exit Sub
      End If
   End If
   
   If cmb_Banco.ListIndex > -1 And cmb_Banco.Enabled = True Then
      If cmb_CtaCte.ListIndex = -1 And cmb_Banco.Enabled = True Then
         MsgBox "En el grupo datos financieros debe de seleccionar una cuenta corriente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CtaCte)
         Exit Sub
      End If
   End If
   
   If cmb_Banco.Enabled = True Then
      If CDbl(pnl_ImpDet.Caption) > 0 Then
         If fs_CtaDetrac_Val = False Then
            MsgBox "El beneficiario no tiene registrado su cuenta corriente de detracciones, por favor comunicarse con Contabilidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Descrip)
            Exit Sub
         End If
      End If
   End If
   
   '--------------------------------------------
   r_bol_Estado = True
   'Factura - ninguna (detraccion)
   r_dbl_ImpAux = CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02)) + CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + _
                  CDbl(fs_HabTxt(ipp_ImpNGrv_02)) + CDbl(fs_HabPnl(pnl_ImpIgv))
                  
   If (cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 1 Or CDbl(fs_HabPnl(pnl_ImpIgv)) > 0) Then 'factura, igv=0
   
       If (cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) <> 2) Then 'detraccion
           'If (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 And r_dbl_ImpAux > 700) Then 'soles
           '    If MsgBox("La factura supera los 700 soles y no esta aplicado detracción ¿ desea continuar ?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
           '       Exit Sub
           '    End If
           'ElseIf (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 And (r_dbl_ImpAux * CDbl(pnl_TipCambio.Caption)) > 700) Then 'dolares
           '    If MsgBox("La factura supera los 700 soles y no esta aplicado detracción ¿ desea continuar ?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
           '       Exit Sub
           '    End If
           'End If
           
           If (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 And r_dbl_ImpAux > 400) Then 'soles
               If MsgBox("La factura supera los 400 soles y no esta aplicado detracción ¿ desea continuar ?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               End If
           ElseIf (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 And (r_dbl_ImpAux * CDbl(pnl_TipCambio.Caption)) > 400) Then 'dolares
               If MsgBox("La factura supera los 400 soles y no esta aplicado detracción ¿ desea continuar ?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               End If
           End If
       End If
   End If
   
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
   
   Screen.MousePointer = 11
   
   If moddat_g_int_InsAct = 1 Then
      'Entregas a Rendir
      Call fs_Grabar_EntRen
   Else
      'Registro de Compras
      Call fs_Grabar_RegCom
   End If
    
   Screen.MousePointer = 0
End Sub

Private Function fs_CtaDetrac_Val() As Boolean
Dim r_rst_Princi     As ADODB.Recordset

   Screen.MousePointer = 11
   fs_CtaDetrac_Val = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.MAEPRV_CTADET, A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC   "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_MAEPRV A  "
   g_str_Parame = g_str_Parame & "   WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   g_str_Parame = g_str_Parame & "     AND TRIM(A.MAEPRV_NUMDOC) = TRIM('" & fs_NumDoc(cmb_Proveedor.Text) & "')  "
   g_str_Parame = g_str_Parame & "     AND A.MAEPRV_SITUAC = 1  "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If

   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   If Trim(r_rst_Princi!MaePrv_CtaDet & "") <> "" Then
      fs_CtaDetrac_Val = True
   End If

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
End Function

Private Sub fs_Grabar_EntRen()
Dim r_dbl_ImpDeb As Double
Dim r_dbl_ImpHab As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_ENTREN_DET ( "
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_CodIte) & "', " 'CAJDET_CODDET
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'CAJDET_CODCAJ
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", " 'CAJDET_CODCOM
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', " 'CAJDET_TIPDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descrip.Text) & "', " 'CAJDET_DESCRP
   g_str_Parame = g_str_Parame & Format(ipp_FchCtb.Text, "yyyymmdd") & ", " 'CAJDET_FECCTB
   g_str_Parame = g_str_Parame & cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) & ", " 'CAJDET_TIPCPB
   'If Trim(cmb_FacAsoc.Text) = "" Then
   '   g_str_Parame = g_str_Parame & "NULL, " 'CAJDET_CODFAC
   'Else
   '   g_str_Parame = g_str_Parame & CLng(Left(Trim(cmb_FacAsoc.Text), 10)) & ", " 'CAJDET_CODFAC
   'End If
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
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_Retencion.Caption) & "', " 'CAJDET_CNT_RET1
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_Detraccion.Caption) & "', " 'CAJDET_CNT_DET1
   'g_str_Parame = g_str_Parame & "'" & Trim(cmb_PorPagar.Text) & "', " 'CAJDET_CNT_PPG1
   g_str_Parame = g_str_Parame & "'" & IIf(cmb_PorPagar.Text = "", "", Mid(cmb_PorPagar.Text, 1, l_int_TopNiv)) & "', " 'CAJDET_CNT_PPG1
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
   'RETENCION
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_RetDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpRet.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpRet.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'CAJDET_DEB_RET1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_RET1
   'DETRACCION
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_DetDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpDet.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpDet.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'CAJDET_DEB_DET1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_DET1
   'POR PAGAR
   r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
   If (Trim(pnl_PpgDH.Caption) = "D") Then
       r_dbl_ImpDeb = CDbl(pnl_ImpPpg.Caption)
   Else
       r_dbl_ImpHab = CDbl(pnl_ImpPpg.Caption)
   End If
   g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'CAJDET_DEB_PPG1
   g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'CAJDET_HAB_PPG1
   g_str_Parame = g_str_Parame & Trim(pnl_Padron.Tag) & ", " 'CAJDET_PADRON
   g_str_Parame = g_str_Parame & cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) & ", " 'CAJDET_APPTRB
   g_str_Parame = g_str_Parame & cmb_CodDetrc.ItemData(cmb_CodDetrc.ListIndex) & ", " 'CAJDET_CODDET
   g_str_Parame = g_str_Parame & "'" & Trim(txt_DocDetrc.Text) & "', " 'CAJDET_NUMDET
   
   g_str_Parame = g_str_Parame & IIf(ipp_FchDetrc.Text = "", "Null", Format(ipp_FchDetrc.Text, "yyyymmdd")) & ", " 'CAJDET_FECDET
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
   g_str_Parame = g_str_Parame & IIf(ipp_FchEmiRef.Text = "", "null", Format(ipp_FchEmiRef.Text, "yyyymmdd")) & ", " 'CAJDET_REF_FECEMI
   If (cmb_TipCbtRef.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "Null, " 'CAJDET_REF_TIPCPB
   Else
       g_str_Parame = g_str_Parame & cmb_TipCbtRef.ItemData(cmb_TipCbtRef.ListIndex) & ", " 'CAJDET_REF_TIPCPB
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumSerieRef.Text) & "', " 'CAJDET_REF_NSERIE
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumRef.Text) & "', " 'CAJDET_REF_NROCOM
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
       Call frm_Ctb_EntRen_03.fs_BuscarCaja
       Call frm_Ctb_EntRen_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_EntRen_03.fs_BuscarCaja
       Call frm_Ctb_EntRen_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub fs_Grabar_RegCom()
Dim r_dbl_ImpDeb As Double
Dim r_dbl_ImpHab As Double
Dim r_str_CodGen As String

   r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 6)
   Else
      r_str_CodGen = Trim(pnl_Codigo.Caption)
   End If

   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 2 And moddat_g_int_FlgAct = 3 Or moddat_g_int_FlgAct = 4 Then
      'SOLO CODIGO DETRAC Y FECHA DETRACT
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CNTBL_REGCOM SET  "
      g_str_Parame = g_str_Parame & "       REGCOM_NUMDET = '" & Trim(txt_DocDetrc.Text) & "',  " 'REGCOM_NUMDET
      g_str_Parame = g_str_Parame & "       REGCOM_FECDET = " & IIf(ipp_FchDetrc.Text = "", "Null", Format(ipp_FchDetrc.Text, "yyyymmdd")) & ",  " 'REGCOM_FECDET
      g_str_Parame = g_str_Parame & "       SEGUSUACT = '" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "       SEGFECACT = " & Format(date, "yyyymmdd") & ",  "
      g_str_Parame = g_str_Parame & "       SEGHORACT = " & Format(Time, "HHmmss") & ",  "
      g_str_Parame = g_str_Parame & "       SEGPLTACT = '" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "       SEGTERACT = '" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "       SEGSUCACT = '" & modgen_g_str_CodSuc & "' "
      g_str_Parame = g_str_Parame & " WHERE REGCOM_CODCOM =  '" & Format(Trim(r_str_CodGen), "0000000000") & "'  "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
      Call frm_Ctb_RegCom_03.fs_BuscarComp
      Screen.MousePointer = 0
      Unload Me
   Else
      'PROCESO NORMAL
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_REGCOM ( "
      g_str_Parame = g_str_Parame & "'" & Format(Trim(r_str_CodGen), "0000000000") & "', " 'REGCOM_CODCOM
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", " 'REGCOM_TIPDOC
      g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', " 'REGCOM_NUMDOC
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
      g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaGvd_02.Text = "", "", Mid(cmb_CtaGvd_02.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_GRV2
      g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaNoGvd_01.Text = "", "", Mid(cmb_CtaNoGvd_01.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_NGV1
      g_str_Parame = g_str_Parame & "'" & IIf(cmb_CtaNoGvd_02.Text = "", "", Mid(cmb_CtaNoGvd_02.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_NGV2
      g_str_Parame = g_str_Parame & "'" & Trim(pnl_Igv.Caption) & "', " 'REGCOM_CNT_IGV1
      g_str_Parame = g_str_Parame & "'" & Trim(pnl_Retencion.Caption) & "', " 'REGCOM_CNT_RET1
      g_str_Parame = g_str_Parame & "'" & Trim(pnl_Detraccion.Caption) & "', " 'REGCOM_CNT_DET1
      g_str_Parame = g_str_Parame & "'" & IIf(cmb_PorPagar.Text = "", "", Mid(cmb_PorPagar.Text, 1, l_int_TopNiv)) & "', " 'REGCOM_CNT_PPG1
      'GRAVADO 1
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(cmb_GravDH_01.Text) = "D") Then
          r_dbl_ImpDeb = CDbl(ipp_ImpGrav_01.Text)
      Else
          r_dbl_ImpHab = CDbl(ipp_ImpGrav_01.Text)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_GRV1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'REGCOM_HAB_GRV1
      'GRAVADO 2
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(cmb_GravDH_02.Text) = "D") Then
          r_dbl_ImpDeb = CDbl(ipp_ImpGrav_02.Text)
      Else
          r_dbl_ImpHab = CDbl(ipp_ImpGrav_02.Text)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_GRV2
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'REGCOM_HAB_GRV2
      'NO GRAVADO 1
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(cmb_NGrvDH_01.Text) = "D") Then
          r_dbl_ImpDeb = CDbl(ipp_ImpNGrv_01.Text)
      Else
          r_dbl_ImpHab = CDbl(ipp_ImpNGrv_01.Text)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_NGV1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", " 'REGCOM_HAB_NGV1
      'NO GRAVADO 2
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(cmb_NGrvDH_02.Text) = "D") Then
          r_dbl_ImpDeb = CDbl(ipp_ImpNGrv_02.Text)
      Else
          r_dbl_ImpHab = CDbl(ipp_ImpNGrv_02.Text)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", " 'REGCOM_DEB_NGV2
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_NGV2
      'IGV
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(pnl_IgvDH.Caption) = "D") Then
          r_dbl_ImpDeb = CDbl(pnl_ImpIgv.Caption)
      Else
          r_dbl_ImpHab = CDbl(pnl_ImpIgv.Caption)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_IGV1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_IGV1
      'RETENCION
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(pnl_RetDH.Caption) = "D") Then
          r_dbl_ImpDeb = CDbl(pnl_ImpRet.Caption)
      Else
          r_dbl_ImpHab = CDbl(pnl_ImpRet.Caption)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_RET1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_RET1
      'DETRACCION
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(pnl_DetDH.Caption) = "D") Then
          r_dbl_ImpDeb = CDbl(pnl_ImpDet.Caption)
      Else
          r_dbl_ImpHab = CDbl(pnl_ImpDet.Caption)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_DET1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_DET1
      'POR PAGAR
      r_dbl_ImpDeb = 0: r_dbl_ImpHab = 0
      If (Trim(pnl_PpgDH.Caption) = "D") Then
          r_dbl_ImpDeb = CDbl(pnl_ImpPpg.Caption)
      Else
          r_dbl_ImpHab = CDbl(pnl_ImpPpg.Caption)
      End If
      g_str_Parame = g_str_Parame & r_dbl_ImpDeb & ", "  'REGCOM_DEB_PPG1
      g_str_Parame = g_str_Parame & r_dbl_ImpHab & ", "  'REGCOM_HAB_PPG1
            
      g_str_Parame = g_str_Parame & Trim(pnl_Padron.Tag) & ", " 'REGCOM_PADRON
      g_str_Parame = g_str_Parame & cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) & ", " 'REGCOM_APPTRB
      g_str_Parame = g_str_Parame & cmb_CodDetrc.ItemData(cmb_CodDetrc.ListIndex) & ", " 'REGCOM_CODDET
      g_str_Parame = g_str_Parame & "'" & Trim(txt_DocDetrc.Text) & "', " 'REGCOM_NUMDET
      g_str_Parame = g_str_Parame & IIf(ipp_FchDetrc.Text = "", "Null", Format(ipp_FchDetrc.Text, "yyyymmdd")) & ", " 'REGCOM_FECDET
      If (cmb_CatCtb.ListIndex = -1) Then
          g_str_Parame = g_str_Parame & "Null, " 'REGCOM_CATCTB
      Else
          g_str_Parame = g_str_Parame & cmb_CatCtb.ItemData(cmb_CatCtb.ListIndex) & ", " 'REGCOM_CATCTB
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_CtrCosto.Text) & "', " 'REGCOM_CNTCST
      If (cmb_Banco.ListIndex = -1) Then
          g_str_Parame = g_str_Parame & "Null , " 'REGCOM_CODBNC
      Else
          g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", " 'REGCOM_CODBNC
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', " 'REGCOM_CTACRR
      g_str_Parame = g_str_Parame & IIf(ipp_FchEmiRef.Text = "", "null", Format(ipp_FchEmiRef.Text, "yyyymmdd")) & ", " 'REGCOM_REF_FECEMI
      If (cmb_TipCbtRef.ListIndex = -1) Then
          g_str_Parame = g_str_Parame & "Null, " 'REGCOM_REF_TIPCPB
      Else
          g_str_Parame = g_str_Parame & cmb_TipCbtRef.ItemData(cmb_TipCbtRef.ListIndex) & ", " 'REGCOM_REF_TIPCPB
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NumSerieRef.Text) & "', " 'REGCOM_REF_NSERIE
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NumRef.Text) & "', " 'REGCOM_REF_NROCOM
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
          Call frm_Ctb_RegCom_03.fs_BuscarComp
          Screen.MousePointer = 0
          Unload Me
      ElseIf (g_rst_Genera!RESUL = 2) Then
          MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
          Call frm_Ctb_RegCom_03.fs_BuscarComp
          Screen.MousePointer = 0
          Unload Me
      End If
   End If
End Sub

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

Private Sub cmd_Limpia_Click()
   Call fs_Limpiar
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
   
   'Reversa de Contabilización(1), Reversa de Proceso Caja Chica(2)
   If moddat_g_int_FlgAct = 1 Or moddat_g_int_FlgAct = 2 Then
      g_str_Parame = ""
      If moddat_g_int_FlgAct = 1 Then
         g_str_Parame = g_str_Parame & " USP_CNTBL_REGCOM_REVERSA ( "
         g_str_Parame = g_str_Parame & " '" & Trim(pnl_Codigo.Caption) & "') " 'REGCOM_CODCOM
      ElseIf moddat_g_int_FlgAct = 2 Then
         g_str_Parame = g_str_Parame & " USP_CNTBL_REGCOM_REVPROC ( "
         g_str_Parame = g_str_Parame & " '" & Trim(pnl_Codigo.Caption) & "', " 'REGCOM_CODCOM
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      End If
                                                                                                                                                                                                                    
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
          Call frm_Ctb_RegCom_03.fs_BuscarComp
          
          Unload Me
      ElseIf (g_rst_Genera!as_resul = 2) Then
          If g_rst_Genera!as_tiptab = 1 Then
             MsgBox "Existen registros de la caja Chica que están contabilizados." & _
                    vbCrLf & "Tiene que dar reversa: " & g_rst_Genera!as_codigos, vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          If g_rst_Genera!as_tiptab = 2 Then
             MsgBox "Existen registros de entregas a rendir que están contabilizados." & _
                    vbCrLf & "Tiene que dar reversa: " & g_rst_Genera!as_codigos, vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
      Else
          MsgBox "Favor de verificar la operación de reversa.", vbInformation, modgen_g_str_NomPlt
      End If
   End If
   
End Sub

Private Sub fs_HabCtrl_Reversa(p_Estado As Boolean)
   'PROVEEDOR
   txt_Descrip.Enabled = p_Estado
   'COMPROBANTE
   ipp_FchCtb.Enabled = p_Estado
   cmb_TipCbtPrv.Enabled = p_Estado
   'cmb_FacAsoc.Enabled = p_Estado
   txt_NumSeriePrv.Enabled = p_Estado
   ipp_FchEmiPrv.Enabled = p_Estado
   cmb_Moneda.Enabled = p_Estado
   txt_NumPrv.Enabled = p_Estado
   ipp_FchVenc.Enabled = p_Estado
   'REQUISITOS
   cmb_AppTrib.Enabled = p_Estado
   cmb_CodDetrc.Enabled = p_Estado
   txt_DocDetrc.Enabled = p_Estado
   ipp_FchDetrc.Enabled = p_Estado
   'DETERMINACION
   ipp_ImpGrav_01.Enabled = p_Estado
   ipp_ImpGrav_02.Enabled = p_Estado
   ipp_ImpNGrv_01.Enabled = p_Estado
   ipp_ImpNGrv_02.Enabled = p_Estado
   '---------
   cmb_GravDH_01.Enabled = p_Estado
   cmb_GravDH_02.Enabled = p_Estado
   cmb_NGrvDH_01.Enabled = p_Estado
   cmb_NGrvDH_02.Enabled = p_Estado
   '---------
   cmb_CtaGvd_01.Enabled = p_Estado
   cmb_CtaGvd_02.Enabled = p_Estado
   cmb_CtaNoGvd_01.Enabled = p_Estado
   cmb_CtaNoGvd_02.Enabled = p_Estado
   cmb_PorPagar.Enabled = p_Estado
   'FINANCIEROS
   cmb_CatCtb.Enabled = p_Estado
   txt_CtrCosto.Enabled = p_Estado
   cmb_Banco.Enabled = p_Estado
   cmb_CtaCte.Enabled = p_Estado
   'REFERENCIAS
   If (p_Estado = False) Then
       ipp_FchEmiRef.Enabled = False
       cmb_TipCbtRef.Enabled = False
       txt_NumSerieRef.Enabled = False
       txt_NumRef.Enabled = False
   Else
       'If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 7 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8 Or _
       '   cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88 Then   '(07-N/C) (08-N/D) (88-DEVOLUCIONES)
       If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 8 Or cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88 Then   '(07-N/C) (08-N/D) (88-DEVOLUCIONES)
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

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   l_dbl_IniFrm = False
   cmd_Grabar.Tag = 1 'msg tipo cambio
   cmd_Reversa.Visible = False
   Screen.MousePointer = 11
         
   Call fs_Inicia
   
   If moddat_g_int_FlgGrb = 1 Then
   '---INSERT
      If moddat_g_int_InsAct = 1 Then
         'Entregas a rendir
         pnl_Titulo.Caption = "Registro de Entregas a Rendir - Adicionar"
         Call fs_Limpiar
         Call gs_BuscarCombo_Item(cmb_Moneda, CInt(moddat_g_str_CodMod))
      Else
         'Registro de Compras
         pnl_Titulo.Caption = "Registro de Compras - Adicionar"
         Call fs_Limpiar
      End If
   ElseIf moddat_g_int_FlgGrb = 2 Then
   '---UPDATE
      If moddat_g_int_InsAct = 1 Then
         'Entregas a rendir
         pnl_Titulo.Caption = "Registro de Entregas a Rendir - Modificar"
         Call fs_Limpiar
         Call fs_Cargar_EntRen
      Else
         'Registro de Compras
         Select Case moddat_g_int_FlgAct
                Case 0: pnl_Titulo.Caption = "Registro de Compras - Modificar"
                Case 1: pnl_Titulo.Caption = "Registro de Compras - Reversa de Contabilización"
                Case 2: pnl_Titulo.Caption = "Registro de Compras - Reversa de Proceso"
                        cmd_Reversa.ToolTipText = "Reversa de Proceso"
                Case 3: pnl_Titulo.Caption = "Registro de Compras - Modificar Contabilizado"
                Case 4: pnl_Titulo.Caption = "Registro de Compras - Modificar"
         End Select
         Call fs_Limpiar
         Call fs_Cargar_RegCom
         cmd_Grabar.Visible = True
         If moddat_g_int_FlgAct = 1 Or moddat_g_int_FlgAct = 2 Then 'reversa
            cmd_Reversa.Visible = True: cmd_Reversa.Left = 30
            cmd_Grabar.Visible = False
            'desabilitar controles
            fs_HabCtrl_Reversa (False)
         ElseIf moddat_g_int_FlgAct = 3 Or moddat_g_int_FlgAct = 4 Then
            cmd_Grabar.Visible = True
            'desabilitar controles
            fs_HabCtrl_Reversa (False)
            txt_DocDetrc.Enabled = True
            ipp_FchDetrc.Enabled = True
            Call gs_SetFocus(txt_DocDetrc)
         End If
      End If
   ElseIf moddat_g_int_FlgGrb = 0 Then
   '---CONSULTAR
      cmd_Grabar.Visible = False
      Call fs_Limpiar
      If moddat_g_int_InsAct = 1 Then
         'Entregas a rendir
          pnl_Titulo.Caption = "Registro de Entregas a Rendir - Consultar"
          Call fs_Cargar_EntRen
      Else
         'Registro de Compras
         pnl_Titulo.Caption = "Registro de Compras - Consultar"
         Call fs_Cargar_RegCom
      End If
      Call fs_Desabilitar
   End If
      
   Screen.MousePointer = 0
   cmd_Grabar.Tag = "" 'msg tipo cambio
   l_dbl_IniFrm = True
   Call gs_CentraForm(Me)
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

Private Sub fs_Cargar_RegCom()
Dim r_dbl_Import     As Double

   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   r_dbl_Import = 0
   Call gs_SetFocus(cmb_TipDoc)
   
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT REGCOM_CODCOM, REGCOM_TIPDOC, REGCOM_NUMDOC, REGCOM_DESCRP, REGCOM_FECCTB, REGCOM_TIPCPB, REGCOM_NSERIE, "
'   g_str_Parame = g_str_Parame & "        REGCOM_NROCOM, REGCOM_FECEMI, REGCOM_FECVTO, REGCOM_CODMON, REGCOM_TIPCAM, REGCOM_CNT_GRV1, REGCOM_CNT_GRV2, "
'   g_str_Parame = g_str_Parame & "        REGCOM_CNT_NGV1, REGCOM_CNT_NGV2, REGCOM_CNT_IGV1, REGCOM_CNT_RET1, REGCOM_CNT_DET1, REGCOM_CNT_PPG1, "
'   g_str_Parame = g_str_Parame & "        REGCOM_DEB_GRV1, REGCOM_HAB_GRV1, REGCOM_DEB_GRV2, REGCOM_HAB_GRV2, REGCOM_DEB_NGV1, REGCOM_HAB_NGV1, "
'   g_str_Parame = g_str_Parame & "        REGCOM_DEB_NGV2, REGCOM_HAB_NGV2, REGCOM_DEB_IGV1, REGCOM_HAB_IGV1, REGCOM_DEB_RET1, REGCOM_HAB_RET1, "
'   g_str_Parame = g_str_Parame & "        REGCOM_DEB_DET1, REGCOM_HAB_DET1, REGCOM_DEB_PPG1, REGCOM_HAB_PPG1, REGCOM_PADRON, REGCOM_APPTRB, "
'   g_str_Parame = g_str_Parame & "        REGCOM_CODDET, REGCOM_NUMDET, REGCOM_FECDET, REGCOM_CATCTB, REGCOM_CNTCST, REGCOM_CODBNC, REGCOM_CTACRR, "
'   g_str_Parame = g_str_Parame & "        REGCOM_REF_FECEMI, REGCOM_REF_TIPCPB, REGCOM_REF_NSERIE, REGCOM_REF_NROCOM, TRIM(C.PARDES_DESCRI) AS PADRON_1,  "
'   g_str_Parame = g_str_Parame & "        B.MAEPRV_RAZSOC, REGCOM_CODCAJ_CHC, REGCOM_CODDET_CHC, REGCOM_TIPREG, F.COMPAG_FECPAG, F.COMPAG_CODCOM, "
'   g_str_Parame = g_str_Parame & "        A.REGCOM_CODFAC  "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_REGCOM A "
'   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON A.REGCOM_TIPDOC = B.MAEPRV_TIPDOC AND A.REGCOM_NUMDOC = B.MAEPRV_NUMDOC "
'   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 121 AND C.PARDES_CODITE = A.REGCOM_PADRON "
'   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET E ON E.COMDET_CODOPE = TO_NUMBER(A.REGCOM_CODCOM) AND E.COMDET_SITUAC = 1 AND E.COMDET_TIPOPE = 2  "
'   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG F ON F.COMPAG_CODCOM = E.COMDET_CODCOM AND F.COMPAG_SITUAC = 1 AND F.COMPAG_FLGCTB = 1  "
'   g_str_Parame = g_str_Parame & "  WHERE REGCOM_CODCOM = '" & moddat_g_str_Codigo & "'"
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.REGCOM_CODCOM, A.REGCOM_TIPDOC, A.REGCOM_NUMDOC, A.REGCOM_DESCRP, A.REGCOM_FECCTB, A.REGCOM_TIPCPB, A.REGCOM_NSERIE,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_NROCOM, A.REGCOM_FECEMI, A.REGCOM_FECVTO, A.REGCOM_CODMON, A.REGCOM_TIPCAM, A.REGCOM_CNT_GRV1, A.REGCOM_CNT_GRV2,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_CNT_NGV1, A.REGCOM_CNT_NGV2, A.REGCOM_CNT_IGV1, A.REGCOM_CNT_RET1, A.REGCOM_CNT_DET1, A.REGCOM_CNT_PPG1,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_DEB_GRV1, A.REGCOM_HAB_GRV1, A.REGCOM_DEB_GRV2, A.REGCOM_HAB_GRV2, A.REGCOM_DEB_NGV1, A.REGCOM_HAB_NGV1,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_DEB_NGV2, A.REGCOM_HAB_NGV2, A.REGCOM_DEB_IGV1, A.REGCOM_HAB_IGV1, A.REGCOM_DEB_RET1, A.REGCOM_HAB_RET1,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_DEB_DET1, A.REGCOM_HAB_DET1, A.REGCOM_DEB_PPG1, A.REGCOM_HAB_PPG1, A.REGCOM_PADRON, A.REGCOM_APPTRB,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_CODDET, A.REGCOM_NUMDET, A.REGCOM_FECDET, A.REGCOM_CATCTB, A.REGCOM_CNTCST, A.REGCOM_CODBNC, A.REGCOM_CTACRR,  "
   g_str_Parame = g_str_Parame & "        A.REGCOM_REF_FECEMI, A.REGCOM_REF_TIPCPB, A.REGCOM_REF_NSERIE, A.REGCOM_REF_NROCOM, TRIM(C.PARDES_DESCRI) AS PADRON_1,  "
   g_str_Parame = g_str_Parame & "        B.MAEPRV_RAZSOC, A.REGCOM_CODCAJ_CHC, A.REGCOM_CODDET_CHC, A.REGCOM_TIPREG, F.COMPAG_FECPAG, F.COMPAG_CODCOM  "
   'g_str_Parame = g_str_Parame & "        A.REGCOM_CODFAC  "
   'g_str_Parame = g_str_Parame & "        , H.REGCOM_CODCOM || '  -  PPg. ' || DECODE(H.REGCOM_CODMON,1,'S/.','US$') AS FAC_COD,  "
   'g_str_Parame = g_str_Parame & "        (H.REGCOM_DEB_PPG1 + H.REGCOM_HAB_PPG1) AS FAC_IMP  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_REGCOM A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON A.REGCOM_TIPDOC = B.MAEPRV_TIPDOC AND A.REGCOM_NUMDOC = B.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 121 AND C.PARDES_CODITE = A.REGCOM_PADRON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET E ON E.COMDET_CODOPE = TO_NUMBER(A.REGCOM_CODCOM) AND E.COMDET_SITUAC = 1 AND E.COMDET_TIPOPE = 2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG F ON F.COMPAG_CODCOM = E.COMDET_CODCOM AND F.COMPAG_SITUAC = 1 AND F.COMPAG_FLGCTB = 1  "
   'g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_REGCOM H ON H.REGCOM_CODCOM = A.REGCOM_CODFAC  "
   g_str_Parame = g_str_Parame & "  WHERE A.REGCOM_CODCOM = '" & moddat_g_str_Codigo & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!regcom_TipDoc)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!regcom_NumDoc & "", 0)
      
      Call fs_Buscar_prov
           
      pnl_Codigo.Caption = Trim(g_rst_Princi!regcom_CodCom & "")
      txt_Descrip.Text = Trim(g_rst_Princi!regcom_Descrp & "")
      If Not IsNull(g_rst_Princi!regcom_FecCtb) Then
          ipp_FchCtb.Text = gf_FormatoFecha(g_rst_Princi!regcom_FecCtb)
      End If
      If Not IsNull(g_rst_Princi!regcom_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtPrv, g_rst_Princi!regcom_TipCpb)
      End If
      'If Trim(g_rst_Princi!REGCOM_CODFAC & "") <> "" Then
      '   cmb_FacAsoc.Text = g_rst_Princi!FAC_COD & " " & Format(g_rst_Princi!FAC_IMP, "###,###,##0.00")
      'End If
      
      txt_NumSeriePrv.Text = Trim(g_rst_Princi!regcom_Nserie & "")
      txt_NumPrv.Text = Trim(g_rst_Princi!regcom_NroCom & "")
      ipp_FchEmiPrv.Text = gf_FormatoFecha(g_rst_Princi!regcom_FecEmi)
      If Not IsNull(g_rst_Princi!regcom_FecVto) Then
         ipp_FchVenc.Text = gf_FormatoFecha(g_rst_Princi!regcom_FecVto)
      End If
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!regcom_CodMon)
      pnl_TipCambio.Caption = Format(g_rst_Princi!regcom_TipCam, "###,###,##0.000000") & " "
      pnl_Padron.Tag = Trim(g_rst_Princi!REGCOM_PADRON & "")
      pnl_Padron.Caption = Trim(g_rst_Princi!PADRON_1 & "")
      If Not IsNull(g_rst_Princi!regcom_apptrb) Then
         Call gs_BuscarCombo_Item(cmb_AppTrib, g_rst_Princi!regcom_apptrb)
      End If
      If Not IsNull(g_rst_Princi!regcom_CodDet) Then
         Call gs_BuscarCombo_Item(cmb_CodDetrc, g_rst_Princi!regcom_CodDet)
      End If
      txt_DocDetrc.Text = Trim(g_rst_Princi!regcom_Numdet & "")
      If Not IsNull(g_rst_Princi!regcom_FecDet) Then
         ipp_FchDetrc.Text = gf_FormatoFecha(g_rst_Princi!regcom_FecDet)
      End If
      
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         pnl_PagCod.Caption = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         pnl_PagFec.Caption = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      
      cmb_CtaGvd_01.ListIndex = fs_ComboIndex(cmb_CtaGvd_01, g_rst_Princi!regcom_Cnt_Grv1 & "", l_int_TopNiv)
      cmb_CtaGvd_02.ListIndex = fs_ComboIndex(cmb_CtaGvd_02, g_rst_Princi!regcom_Cnt_Grv2 & "", l_int_TopNiv)
      cmb_CtaNoGvd_01.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_01, g_rst_Princi!regcom_Cnt_Ngv1 & "", l_int_TopNiv)
      cmb_CtaNoGvd_02.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_02, g_rst_Princi!regcom_Cnt_Ngv2 & "", l_int_TopNiv)
      
      pnl_Igv.Caption = Trim(g_rst_Princi!regcom_Cnt_Igv1 & "")
      pnl_Retencion.Caption = Trim(g_rst_Princi!regcom_Cnt_Ret1 & "")
      pnl_Detraccion.Caption = Trim(g_rst_Princi!regcom_Cnt_Det1 & "")
      If Not IsNull(g_rst_Princi!regcom_Cnt_Ppg1) Then
         cmb_PorPagar.ListIndex = fs_ComboIndex(cmb_PorPagar, g_rst_Princi!regcom_Cnt_Ppg1 & "", l_int_TopNiv)
      End If
      'GRAVADO 1
      If (g_rst_Princi!regcom_Deb_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 0
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!regcom_Deb_Grv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!regcom_Hab_Grv1 > 0) Then
          cmb_GravDH_01.ListIndex = 1
          ipp_ImpGrav_01.Text = Format(g_rst_Princi!regcom_Hab_Grv1, "###,###,##0.00")
      End If
      'GRAVADO 2
      If (g_rst_Princi!regcom_Deb_Grv2 > 0) Then
          cmb_GravDH_02.ListIndex = 0
          ipp_ImpGrav_02.Text = Format(g_rst_Princi!regcom_Deb_Grv2, "###,###,##0.00")
      End If
      If (g_rst_Princi!regcom_Hab_Grv2 > 0) Then
          cmb_GravDH_02.ListIndex = 1
          ipp_ImpGrav_02.Text = Format(g_rst_Princi!regcom_Hab_Grv2, "###,###,##0.00")
      End If
      'NO GRAVADO 1
      If (g_rst_Princi!regcom_Deb_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 0
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!regcom_Deb_Ngv1, "###,###,##0.00")
      End If
      If (g_rst_Princi!regcom_Hab_Ngv1 > 0) Then
          cmb_NGrvDH_01.ListIndex = 1
          ipp_ImpNGrv_01.Text = Format(g_rst_Princi!regcom_Hab_Ngv1, "###,###,##0.00")
      End If
      'NO GRAVADO 2
      If (g_rst_Princi!regcom_Deb_Ngv2 > 0) Then
          cmb_NGrvDH_02.ListIndex = 0
          ipp_ImpNGrv_02.Text = Format(g_rst_Princi!regcom_Deb_Ngv2, "###,###,##0.00")
      End If
      If (g_rst_Princi!regcom_Hab_Ngv2 > 0) Then
          cmb_NGrvDH_02.ListIndex = 1
          ipp_ImpNGrv_02.Text = Format(g_rst_Princi!regcom_Hab_Ngv2, "###,###,##0.00")
      End If
      'IGV
      If (g_rst_Princi!regcom_Deb_Igv1 > 0) Then
          pnl_IgvDH.Caption = "D"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!regcom_Deb_Igv1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regcom_Hab_Igv1 > 0) Then
          pnl_IgvDH.Caption = "H"
          pnl_ImpIgv.Caption = Format(g_rst_Princi!regcom_Hab_Igv1, "###,###,##0.00") & " "
      End If
      'RETENCION
      If (g_rst_Princi!regcom_Deb_Ret1 > 0) Then
          pnl_RetDH.Caption = "D"
          pnl_ImpRet.Caption = Format(g_rst_Princi!regcom_Deb_Ret1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regcom_Hab_Ret1 > 0) Then
          pnl_RetDH.Caption = "H"
          pnl_ImpRet.Caption = Format(g_rst_Princi!regcom_Hab_Ret1, "###,###,##0.00") & " "
      End If
      'DETRACCION
      If (g_rst_Princi!regcom_Deb_Det1 > 0) Then
          pnl_DetDH.Caption = "D"
          pnl_ImpDet.Caption = Format(g_rst_Princi!regcom_Deb_Det1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regcom_Hab_Det1 > 0) Then
          pnl_DetDH.Caption = "H"
          pnl_ImpDet.Caption = Format(g_rst_Princi!regcom_Hab_Det1, "###,###,##0.00") & " "
      End If
      'CUENTAS POR PAGAR
      If (g_rst_Princi!regcom_Deb_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "D"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!regcom_Deb_Ppg1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!regcom_Hab_Ppg1 > 0) Then
          pnl_PpgDH.Caption = "H"
          pnl_ImpPpg.Caption = Format(g_rst_Princi!regcom_Hab_Ppg1, "###,###,##0.00") & " "
      End If
      '---------------------------------------------------------------------------------------
      If Not IsNull(g_rst_Princi!regcom_CatCtb) Then
         Call gs_BuscarCombo_Item(cmb_CatCtb, g_rst_Princi!regcom_CatCtb)
      End If
      txt_CtrCosto.Text = Trim(g_rst_Princi!REGCOM_CNTCST & "")
      
      If Not IsNull(g_rst_Princi!regcom_CodBnc) Then
         Call gs_BuscarCombo_Item(cmb_Banco, g_rst_Princi!regcom_CodBnc)
      End If
      If Not IsNull(g_rst_Princi!regcom_CtaCrr) Then
         Call gs_BuscarCombo_Text(cmb_CtaCte, g_rst_Princi!regcom_CtaCrr, -1)
         'cmb_CtaCte.Text = g_rst_Princi!regcom_CtaCrr
      End If
      If Not IsNull(g_rst_Princi!regcom_Ref_FecEmi) Then
         ipp_FchEmiRef.Text = gf_FormatoFecha(g_rst_Princi!regcom_Ref_FecEmi)
      End If
      If Not IsNull(g_rst_Princi!regcom_Ref_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtRef, g_rst_Princi!regcom_Ref_TipCpb)
      End If
      txt_NumSerieRef.Text = Format(g_rst_Princi!regcom_Ref_Nserie & "")
      txt_NumRef.Text = Format(g_rst_Princi!regcom_Ref_NroCom & "")
      
      'si no hay importe - vacío
      Call fs_LimpiarDH
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Cargar_EntRen()
Dim r_dbl_Import     As Double

   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   r_dbl_Import = 0
   Call gs_SetFocus(cmb_TipDoc)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJDET_TIPDOC, A.CAJDET_NUMDOC, A.CAJDET_DESCRP, A.CAJDET_FECCTB, A.CAJDET_TIPCPB, A.CAJDET_NSERIE,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_NROCOM, A.CAJDET_FECEMI, A.CAJDET_FECVTO, A.CAJDET_CODMON, A.CAJDET_TIPCAM, A.CAJDET_CNT_GRV1, A.CAJDET_CNT_GRV2,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_CNT_NGV1, A.CAJDET_CNT_NGV2, A.CAJDET_CNT_IGV1, A.CAJDET_CNT_RET1, A.CAJDET_CNT_DET1, A.CAJDET_CNT_PPG1,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_DEB_GRV1, A.CAJDET_HAB_GRV1, A.CAJDET_DEB_GRV2, A.CAJDET_HAB_GRV2, A.CAJDET_DEB_NGV1, A.CAJDET_HAB_NGV1,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_DEB_NGV2, A.CAJDET_HAB_NGV2, A.CAJDET_DEB_IGV1, A.CAJDET_HAB_IGV1, A.CAJDET_DEB_RET1, A.CAJDET_HAB_RET1,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_DEB_DET1, A.CAJDET_HAB_DET1, A.CAJDET_DEB_PPG1, A.CAJDET_HAB_PPG1, A.CAJDET_PADRON, A.CAJDET_APPTRB,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_CODDET, A.CAJDET_NUMDET, A.CAJDET_FECDET, A.CAJDET_CATCTB, A.CAJDET_CNTCST, A.CAJDET_CODBNC, A.CAJDET_CTACRR,  "
   g_str_Parame = g_str_Parame & "        A.CAJDET_REF_FECEMI, A.CAJDET_REF_TIPCPB, A.CAJDET_REF_NSERIE, A.CAJDET_REF_NROCOM, TRIM(C.PARDES_DESCRI) AS PADRON_1,  "
   g_str_Parame = g_str_Parame & "        B.MAEPRV_RAZSOC, A.CAJDET_CODCAJ, A.CAJDET_CODDET_DET, A.CAJDET_CODDET  "
   'g_str_Parame = g_str_Parame & "        A.CAJDET_CODFAC,  "
   'g_str_Parame = g_str_Parame & "        LPAD(H.CAJDET_CODDET,10,'0') || '  -  PPg. ' || DECODE(H.CAJDET_CODMON,1,'S/.','US$') AS FAC_COD,  "
   'g_str_Parame = g_str_Parame & "        (H.CAJDET_DEB_PPG1 + H.CAJDET_HAB_PPG1) AS FAC_IMP  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_CAJCHC_DET A  "
   g_str_Parame = g_str_Parame & "   INNER JOIN CNTBL_MAEPRV B ON A.CAJDET_TIPDOC = B.MAEPRV_TIPDOC AND A.CAJDET_NUMDOC = B.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 121 AND C.PARDES_CODITE = A.CAJDET_PADRON  "
   'g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_CAJCHC_DET H ON H.CAJDET_CODDET = A.CAJDET_CODFAC  "
   'g_str_Parame = g_str_Parame & "          AND H.CAJDET_TIPTAB = 2 AND H.CAJDET_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "   WHERE A.CAJDET_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "     AND A.CAJDET_CODDET = " & moddat_g_str_CodIte
   g_str_Parame = g_str_Parame & "     AND A.CAJDET_TIPTAB = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!CAJDET_TipDoc)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!CajDet_NumDoc & "", 0)
      
      Call fs_Buscar_prov
           
      pnl_Codigo.Caption = Format(g_rst_Princi!CajDet_CodDet, "0000000000")
      txt_Descrip.Text = Trim(g_rst_Princi!CAJDET_DESCRP & "")
      If Not IsNull(g_rst_Princi!CAJDET_FECCTB) Then
          ipp_FchCtb.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FECCTB)
      End If
      If Not IsNull(g_rst_Princi!CajDet_TipCpb) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtPrv, g_rst_Princi!CajDet_TipCpb)
         'Call fs_ActivaRefer
      End If
      'If Trim(g_rst_Princi!CAJDET_CODFAC & "") <> "" Then
      '   cmb_FacAsoc.Text = g_rst_Princi!FAC_COD & " " & Format(g_rst_Princi!FAC_IMP, "###,###,##0.00")
      'End If
      
      txt_NumSeriePrv.Text = Trim(g_rst_Princi!CajDet_Nserie & "")
      txt_NumPrv.Text = Trim(g_rst_Princi!CajDet_NroCom & "")
      ipp_FchEmiPrv.Text = gf_FormatoFecha(g_rst_Princi!CajDet_FecEmi)
      If Not IsNull(g_rst_Princi!CAJDET_FecVto) Then
         ipp_FchVenc.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FecVto)
      End If
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!CAJDET_CODMON)
      pnl_TipCambio.Caption = Format(g_rst_Princi!CAJDET_TipCam, "###,###,##0.000000") & " "
      pnl_Padron.Tag = Trim(g_rst_Princi!CAJDET_PADRON & "")
      pnl_Padron.Caption = Trim(g_rst_Princi!PADRON_1 & "")
      If Not IsNull(g_rst_Princi!CAJDET_APPTRB) Then
         Call gs_BuscarCombo_Item(cmb_AppTrib, g_rst_Princi!CAJDET_APPTRB)
      End If
      If Not IsNull(g_rst_Princi!CAJDET_CODDET_DET) Then
         Call gs_BuscarCombo_Item(cmb_CodDetrc, g_rst_Princi!CAJDET_CODDET_DET)
      End If
      txt_DocDetrc.Text = Trim(g_rst_Princi!CAJDET_NUMDET & "")
      If Not IsNull(g_rst_Princi!CAJDET_FECDET) Then
         ipp_FchDetrc.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FECDET)
      End If
      
      cmb_CtaGvd_01.ListIndex = fs_ComboIndex(cmb_CtaGvd_01, g_rst_Princi!CAJDET_Cnt_Grv1 & "", l_int_TopNiv)
      cmb_CtaGvd_02.ListIndex = fs_ComboIndex(cmb_CtaGvd_02, g_rst_Princi!CAJDET_Cnt_Grv2 & "", l_int_TopNiv)
      cmb_CtaNoGvd_01.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_01, g_rst_Princi!CAJDET_Cnt_Ngv1 & "", l_int_TopNiv)
      cmb_CtaNoGvd_02.ListIndex = fs_ComboIndex(cmb_CtaNoGvd_02, g_rst_Princi!CAJDET_Cnt_Ngv2 & "", l_int_TopNiv)
      
      pnl_Igv.Caption = Trim(g_rst_Princi!CAJDET_Cnt_Igv1 & "")
      pnl_Retencion.Caption = Trim(g_rst_Princi!CAJDET_CNT_RET1 & "")
      pnl_Detraccion.Caption = Trim(g_rst_Princi!CAJDET_CNT_DET1 & "")
      If Not IsNull(g_rst_Princi!CAJDET_Cnt_Ppg1) Then
         'cmb_PorPagar.Text = Trim(g_rst_Princi!CAJDET_Cnt_Ppg1 & "")
          cmb_PorPagar.ListIndex = fs_ComboIndex(cmb_PorPagar, g_rst_Princi!CAJDET_Cnt_Ppg1 & "", l_int_TopNiv)
      End If
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
      'RETENCION
      If (g_rst_Princi!CAJDET_DEB_RET1 > 0) Then
          pnl_RetDH.Caption = "D"
          pnl_ImpRet.Caption = Format(g_rst_Princi!CAJDET_DEB_RET1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!CAJDET_HAB_RET1 > 0) Then
          pnl_RetDH.Caption = "H"
          pnl_ImpRet.Caption = Format(g_rst_Princi!CAJDET_HAB_RET1, "###,###,##0.00") & " "
      End If
      'DETRACCION
      If (g_rst_Princi!CAJDET_DEB_DET1 > 0) Then
          pnl_DetDH.Caption = "D"
          pnl_ImpDet.Caption = Format(g_rst_Princi!CAJDET_DEB_DET1, "###,###,##0.00") & " "
      End If
      If (g_rst_Princi!CAJDET_HAB_DET1 > 0) Then
          pnl_DetDH.Caption = "H"
          pnl_ImpDet.Caption = Format(g_rst_Princi!CAJDET_HAB_DET1, "###,###,##0.00") & " "
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
      If Not IsNull(g_rst_Princi!CAJDET_REF_FECEMI) Then
         ipp_FchEmiRef.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_REF_FECEMI)
      End If
      If Not IsNull(g_rst_Princi!CAJDET_REF_TIPCPB) Then
         Call gs_BuscarCombo_Item(cmb_TipCbtRef, g_rst_Princi!CAJDET_REF_TIPCPB)
      End If
      txt_NumSerieRef.Text = Format(g_rst_Princi!CAJDET_REF_NSERIE & "")
      txt_NumRef.Text = Format(g_rst_Princi!CAJDET_REF_NROCOM & "")
      
      'si no hay importe - vacío
      Call fs_LimpiarDH
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Desabilitar()
   cmb_TipDoc.Enabled = False
   'txt_NumDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   txt_Descrip.Enabled = False
   
   ipp_FchCtb.Enabled = False
   cmb_TipCbtPrv.Enabled = False
   'cmb_FacAsoc.Enabled = False
   txt_NumSeriePrv.Enabled = False
   txt_NumPrv.Enabled = False
   ipp_FchEmiPrv.Enabled = False
   ipp_FchVenc.Enabled = False
   cmb_Moneda.Enabled = False
   
   cmb_AppTrib.Enabled = False
   cmb_CodDetrc.Enabled = False
   txt_DocDetrc.Enabled = False
   ipp_FchDetrc.Enabled = False
   
   cmb_CtaGvd_01.Enabled = False
   cmb_CtaGvd_02.Enabled = False
   cmb_CtaNoGvd_01.Enabled = False
   cmb_CtaNoGvd_02.Enabled = False
   cmb_PorPagar.Enabled = False
   
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
   
   ipp_FchEmiRef.Enabled = False
   cmb_TipCbtRef.Enabled = False
   txt_NumSerieRef.Enabled = False
   txt_NumRef.Enabled = False
End Sub

Private Sub fs_Inicia()
Dim r_str_Cadena As String
   
   l_dbl_IGV = 0
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   Call moddat_gs_Carga_LisIte_Combo(cmb_AppTrib, 1, "125")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CatCtb, 1, "124")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCbtPrv, 1, "123")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCbtRef, 1, "123")
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
   moddat_g_str_CodEmp = "000001"
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaGvd_01, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   cmb_CtaGvd_02.Clear
   cmb_CtaNoGvd_01.Clear
   cmb_CtaNoGvd_02.Clear
   cmb_PorPagar.Clear
   For l_int_Contar = 1 To UBound(l_arr_CtaCtb)
       cmb_CtaGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
       cmb_CtaNoGvd_01.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
       cmb_CtaNoGvd_02.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
       
       If moddat_g_int_InsAct = 1 Then  'entragas a rendir
          If moddat_g_str_CodMod = 1 Then
             If l_arr_CtaCtb(l_int_Contar).Genera_Codigo = "191807020101" Then
                cmb_PorPagar.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
             End If
          Else
             If l_arr_CtaCtb(l_int_Contar).Genera_Codigo = "192807020101" Then
                cmb_PorPagar.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
             End If
          End If
       Else
          Select Case l_arr_CtaCtb(l_int_Contar).Genera_Codigo
                 Case "252601010101", "251602010101", "252602010101", "191807020101", _
                      "192807020101", "251419010109", "252506010101", "251602010105", _
                      "252602010105", "251704010101", "291807010101", "111701010101", "291807010117"
                      cmb_PorPagar.AddItem l_arr_CtaCtb(l_int_Contar).Genera_Codigo & " - " & l_arr_CtaCtb(l_int_Contar).Genera_Nombre
          End Select
          'cmb_PorPagar.AddItem ("251602010101") 'DEFAULT
       End If
       
   Next
   
   'cargar codigo detraccion
   Call moddat_gs_Carga_LisIte(cmb_CodDetrc, l_arr_CodDet, 1, 126, 1)
   cmb_CodDetrc.Clear
   ReDim l_arr_CtaCteDol(0)
   For l_int_Contar = 1 To UBound(l_arr_CodDet)
       r_str_Cadena = Right(Trim(l_arr_CodDet(l_int_Contar).Genera_Codigo), 3)
       If (Trim(l_arr_CodDet(l_int_Contar).Genera_Codigo) = "999999") Then
          r_str_Cadena = "      "
       End If
       cmb_CodDetrc.AddItem (r_str_Cadena & " - " & _
                             Trim(l_arr_CodDet(l_int_Contar).Genera_Nombre) & " %")
       cmb_CodDetrc.ItemData(cmb_CodDetrc.NewIndex) = CLng(l_arr_CodDet(l_int_Contar).Genera_Codigo)
   Next
   
   'cargar igv
   l_dbl_IGV = moddat_gf_Consulta_ParVal("001", "001") 'IGV
   l_dbl_IGV = l_dbl_IGV / 100
   Call moddat_gs_FecSis
   
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
End Sub

Private Function fs_CodDetrac() As Double
fs_CodDetrac = 0
   If (cmb_CodDetrc.ListIndex > -1) Then
       For l_int_Contar = 1 To UBound(l_arr_CodDet)
           If (Trim(cmb_CodDetrc.ItemData(cmb_CodDetrc.ListIndex)) = CInt(l_arr_CodDet(l_int_Contar).Genera_Codigo)) Then
               fs_CodDetrac = CDbl(Trim(l_arr_CodDet(l_int_Contar).Genera_Nombre))
               Exit Function
           End If
       Next
   End If
End Function

Private Sub fs_CargarPrv()
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
   cmb_Proveedor.Clear
   cmb_Proveedor.Text = ""
   pnl_Padron.Tag = ""
   pnl_Padron.Caption = ""
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


Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
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

Private Sub ipp_FchDetrc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If moddat_g_int_FlgAct = 3 Then
          Call gs_SetFocus(cmd_Grabar)
       Else
          Call gs_SetFocus(ipp_ImpGrav_01)
       End If
   End If
End Sub

Private Sub ipp_FchEmiPrv_LostFocus()
   If l_dbl_IniFrm = True Then
      If moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2 Then
         If cmd_Grabar.Tag = "" Then
            If (Format(ipp_FchEmiPrv.Text, "yyyymmdd") > Format(moddat_g_str_FecSis, "yyyymmdd")) Then
                MsgBox "Esta intentando registrar un documento con una fecha futura.", vbExclamation, modgen_g_str_NomPlt
            ElseIf (Format(ipp_FchEmiPrv.Text, "yyyy") <> Format(modctb_str_FecFin, "yyyy")) Then
                MsgBox "Esta intentando registrar un documento de un ejercicio anterior.", vbExclamation, modgen_g_str_NomPlt
            End If
         End If
      End If
      pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(3, 2, Format(ipp_FchEmiPrv.Text, "yyyymmdd"), 1)
      pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
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

Private Sub txt_Descrip_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchCtb)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_TipCbtPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       'If cmb_FacAsoc.Enabled = False Then
       Call gs_SetFocus(txt_NumSeriePrv)
       'Else
       '   Call gs_SetFocus(cmb_FacAsoc)
       'End If
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

Private Sub ipp_FchEmiPrv_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchVenc)
   End If
End Sub

Private Sub ipp_FchVenc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If cmb_Moneda.Enabled = False Then
          Call gs_SetFocus(cmb_AppTrib)
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

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_AppTrib)
   End If
End Sub

Private Sub txt_Padron_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_AppTrib)
   End If
End Sub

Private Sub cmb_AppTrib_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CodDetrc)
   End If
End Sub

Private Sub cmb_CodDetrc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_DocDetrc)
   End If
End Sub

Private Sub txt_DocDetrc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchDetrc)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_FchDetrc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaGvd_01)
   End If
End Sub

Private Sub cmb_PorPagar_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CatCtb)
   End If
End Sub

Private Sub cmb_CatCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_CtrCosto)
   End If
End Sub

Private Sub txt_CtrCosto_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If (ipp_FchEmiRef.Enabled = False) Then
           Call gs_SetFocus(cmd_Grabar)
       Else
           Call gs_SetFocus(ipp_FchEmiRef)
       End If
   End If
End Sub

Private Sub ipp_FchEmiRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipCbtRef)
   End If
End Sub

Private Sub cmb_TipCbtRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumSerieRef)
   End If
End Sub

Private Sub txt_NumSerieRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumRef)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumRef_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
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
   
   If Trim(p_Objeto.Name) = Trim(pnl_ImpRet.Name) Then
      fs_HabPnl = pnl_ImpRet.Caption
      If Trim(pnl_RetDH.Caption) = "H" Then
         fs_HabPnl = "-" & pnl_ImpRet.Caption
      End If
   End If
   
   If Trim(p_Objeto.Name) = Trim(pnl_ImpDet.Name) Then
      fs_HabPnl = pnl_ImpDet.Caption
      If Trim(pnl_DetDH.Caption) = "H" Then
         fs_HabPnl = "-" & pnl_ImpDet.Caption
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
   
   If (cmb_TipCbtPrv.ListIndex = -1 Or cmb_Moneda.ListIndex = -1 Or cmb_AppTrib.ListIndex = -1 Or cmb_CodDetrc.ListIndex = -1) Then
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
   If (cmb_AppTrib.ListIndex = -1) Then
       Call gs_SetFocus(cmb_AppTrib)
       MsgBox "Debe seleccionar una aplicación tributaria.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   If (cmb_CodDetrc.ListIndex = -1) Then
       Call gs_SetFocus(cmb_CodDetrc)
       MsgBox "Debe seleccionar un código de detracción.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
      
   pnl_ImpIgv.Caption = "0.00 "
   r_dbl_ImpAux = 0
   Call fs_Asigna_DebHab(False)
   
   'COMPROBANTE SIEMPRE 07 - N/C
   If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) = 88 Then
       'calculo del igv
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02))) * l_dbl_IGV
       pnl_ImpIgv.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
       
       pnl_ImpRet.Caption = "0.00 "
       pnl_ImpDet.Caption = "0.00 "
       
       'calculo por pagar
       r_dbl_ImpAux = 0
       r_dbl_ImpAux = Math.Abs((fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02)) + _
                      CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + CDbl(fs_HabTxt(ipp_ImpNGrv_02)) + CDbl(fs_HabPnl(pnl_ImpIgv)))
       pnl_ImpPpg.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
   Else 'COMPROBANTE DISTINTO A 88   //07 - N/C
       'calculo del igv
       r_dbl_ImpAux = Math.Abs((CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02))) * l_dbl_IGV)
       pnl_ImpIgv.Caption = Format(r_dbl_ImpAux, "###,###,##0.00") & " "
       
       'calculo de la retencion
       pnl_ImpRet.Caption = "0.00 "
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + CDbl(fs_HabTxt(ipp_ImpNGrv_02)))
       pnl_ImpRet.Caption = Format(r_dbl_ImpAux * (8 / 100), "###,###,##0.00") & " "  'se quito el redondeo a 2 decimales
       
       r_bol_Estado = True
       If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) <> 2 Then
          r_bol_Estado = False 'distinto a honorarios
       End If
       If cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) <> 3 Then
          r_bol_Estado = False 'distinto a retencion
       End If
       If r_bol_Estado = False Then
          pnl_ImpRet.Caption = "0.00 "
       End If
       
       'calculo de la detraccion
       pnl_ImpDet.Caption = "0.00 "
       r_dbl_Aux = 0
       r_dbl_Aux = fs_CodDetrac / 100
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02)) + _
                      CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + CDbl(fs_HabTxt(ipp_ImpNGrv_02)) + CDbl(fs_HabPnl(pnl_ImpIgv)))
                      
       pnl_ImpDet.Caption = Format(Round(r_dbl_ImpAux * r_dbl_Aux, 0), "###,###,##0.00") & " "
       
       r_bol_Estado = True
       If cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) <> 1 And cmb_TipCbtPrv.ItemData(cmb_TipCbtPrv.ListIndex) <> 7 Then
          r_bol_Estado = False 'distinto a factura
       End If
       If cmb_AppTrib.ItemData(cmb_AppTrib.ListIndex) <> 2 Then
          r_bol_Estado = False 'distinto a detraccion
       End If
       If CDbl(pnl_ImpIgv.Caption) = 0 Then
          r_bol_Estado = False 'igv es cero
       End If
       
       'If (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 And r_dbl_ImpAux < CDbl(700.01)) Then
       '    r_bol_Estado = False
       'ElseIf (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 And (r_dbl_ImpAux * CDbl(pnl_TipCambio.Caption)) < CDbl(700.01)) Then
       '    r_bol_Estado = False
       'End If
       
       If (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 And r_dbl_ImpAux < CDbl(400.01)) Then
          If cmb_CodDetrc.ItemData(cmb_CodDetrc.ListIndex) <> 27 Then
             r_bol_Estado = False
          End If
       ElseIf (cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 And (r_dbl_ImpAux * CDbl(pnl_TipCambio.Caption)) < CDbl(400.01)) Then
          If cmb_CodDetrc.ItemData(cmb_CodDetrc.ListIndex) <> 27 Then
             r_bol_Estado = False
          End If
       End If
       
       If r_bol_Estado = False Then
           pnl_ImpDet.Caption = "0.00 "
       End If
       
       'calculo por pagar
       r_dbl_ImpAux = Math.Abs(CDbl(fs_HabTxt(ipp_ImpGrav_01)) + CDbl(fs_HabTxt(ipp_ImpGrav_02)) + CDbl(fs_HabTxt(ipp_ImpNGrv_01)) + _
                      CDbl(fs_HabTxt(ipp_ImpNGrv_02)) + CDbl(fs_HabPnl(pnl_ImpIgv)) + CDbl(fs_HabPnl(pnl_ImpRet)) + _
                      CDbl(fs_HabPnl(pnl_ImpDet)))
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


