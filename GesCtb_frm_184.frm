VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegCom_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   Icon            =   "GesCtb_frm_184.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   9345
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   16484
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
         Left            =   30
         TabIndex        =   40
         Top             =   60
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
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
            TabIndex        =   41
            Top             =   150
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Proveedor"
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
            Picture         =   "GesCtb_frm_184.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1730
         Left            =   30
         TabIndex        =   42
         Top             =   1395
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
         _ExtentY        =   3052
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
         Begin VB.ComboBox cmb_TarCre 
            Height          =   315
            Left            =   7530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txt_CodSic 
            Height          =   315
            Left            =   7530
            MaxLength       =   6
            TabIndex        =   4
            Top             =   990
            Width           =   1485
         End
         Begin VB.ComboBox cmb_TipPer 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   990
            Width           =   1515
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1590
            MaxLength       =   100
            TabIndex        =   2
            Top             =   660
            Width           =   8830
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7530
            MaxLength       =   12
            TabIndex        =   1
            Top             =   330
            Width           =   1485
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   330
            Width           =   4185
         End
         Begin EditLib.fpDateTime ipp_FecIng 
            Height          =   315
            Left            =   1590
            TabIndex        =   5
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
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso:"
            Height          =   195
            Left            =   150
            TabIndex        =   93
            Top             =   1380
            Width           =   1065
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Tarjeta Crédito:"
            Height          =   195
            Left            =   6240
            TabIndex        =   92
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Código Planilla:"
            Height          =   195
            Left            =   6240
            TabIndex        =   91
            Top             =   1050
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Personal:"
            Height          =   195
            Left            =   150
            TabIndex        =   90
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Registro"
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
            TabIndex        =   84
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   47
            Top             =   390
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nro Documento:"
            Height          =   195
            Left            =   6240
            TabIndex        =   46
            Top             =   390
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   150
            TabIndex        =   45
            Top             =   720
            Width           =   990
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   43
         Top             =   720
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_184.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10890
            Picture         =   "GesCtb_frm_184.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1995
         Left            =   30
         TabIndex        =   48
         Top             =   3165
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
         _ExtentY        =   3510
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
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   6600
            TabIndex        =   9
            Text            =   "cmb_PrvDir"
            Top             =   600
            Width           =   3840
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Text            =   "cmb_DptDir"
            Top             =   600
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1590
            TabIndex        =   10
            Text            =   "cmb_DstDir"
            Top             =   930
            Width           =   3315
         End
         Begin VB.TextBox txt_Telf3 
            Height          =   315
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   13
            Top             =   1260
            Width           =   1485
         End
         Begin VB.TextBox txt_CtaDetrac 
            Height          =   315
            Left            =   6600
            MaxLength       =   11
            TabIndex        =   15
            Top             =   1590
            Width           =   1485
         End
         Begin VB.TextBox txt_Email 
            Height          =   315
            Left            =   1590
            MaxLength       =   40
            TabIndex        =   14
            Top             =   1590
            Width           =   3315
         End
         Begin VB.TextBox txt_Telf2 
            Height          =   315
            Left            =   1590
            MaxLength       =   15
            TabIndex        =   12
            Top             =   1260
            Width           =   1785
         End
         Begin VB.TextBox txt_Telf1 
            Height          =   315
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   11
            Top             =   930
            Width           =   1485
         End
         Begin VB.TextBox txt_DomFisc 
            Height          =   315
            Left            =   1590
            MaxLength       =   100
            TabIndex        =   7
            Top             =   270
            Width           =   8830
         End
         Begin VB.Label lbl_General 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Index           =   53
            Left            =   5220
            TabIndex        =   89
            Top             =   720
            Width           =   705
         End
         Begin VB.Label lbl_General 
            AutoSize        =   -1  'True
            Caption         =   "Departamento:"
            Height          =   195
            Index           =   44
            Left            =   150
            TabIndex        =   88
            Top             =   675
            Width           =   1050
         End
         Begin VB.Label lbl_General 
            AutoSize        =   -1  'True
            Caption         =   "Distrito:"
            Height          =   195
            Index           =   45
            Left            =   150
            TabIndex        =   87
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label Label39 
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
            TabIndex        =   85
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "(Bco Nación)"
            Height          =   195
            Left            =   8250
            TabIndex        =   74
            Top             =   1680
            Width           =   930
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Detractora:"
            Height          =   195
            Left            =   5220
            TabIndex        =   54
            Top             =   1665
            Width           =   1350
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correo:"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   1635
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Telefono 3:"
            Height          =   195
            Left            =   5220
            TabIndex        =   52
            Top             =   1335
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Telefono 2:"
            Height          =   195
            Left            =   150
            TabIndex        =   51
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Telefono 1:"
            Height          =   195
            Left            =   5220
            TabIndex        =   50
            Top             =   1020
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio Fiscal:"
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   345
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1035
         Left            =   30
         TabIndex        =   55
         Top             =   5190
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
         _ExtentY        =   1817
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
         Begin VB.ComboBox cmb_Padron2 
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   630
            Width           =   3840
         End
         Begin VB.ComboBox cmb_Padron1 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   630
            Width           =   3345
         End
         Begin VB.ComboBox cmb_Condicion 
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   3840
         End
         Begin VB.ComboBox cmb_TipCny 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   3345
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Datos Fiscales"
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
            TabIndex        =   86
            Top             =   60
            Width           =   1260
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Padron 2:"
            Height          =   195
            Left            =   5220
            TabIndex        =   59
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Padron 1:"
            Height          =   195
            Left            =   150
            TabIndex        =   58
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   5220
            TabIndex        =   57
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Contribuyente:"
            Height          =   195
            Left            =   150
            TabIndex        =   56
            Top             =   375
            Width           =   1380
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2925
         Left            =   30
         TabIndex        =   60
         Top             =   6255
         Width           =   11520
         _Version        =   65536
         _ExtentX        =   20320
         _ExtentY        =   5151
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
         Begin VB.ComboBox cmb_Banco_DL3 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2520
            Width           =   3345
         End
         Begin VB.TextBox txt_CtaCorr_DL3 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   36
            Top             =   2490
            Width           =   2100
         End
         Begin VB.TextBox txt_CCI_DL3 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   37
            Top             =   2490
            Width           =   2100
         End
         Begin VB.ComboBox cmb_Banco_MN3 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1230
            Width           =   3345
         End
         Begin VB.TextBox txt_CtaCorr_MN3 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   27
            Top             =   1260
            Width           =   2100
         End
         Begin VB.TextBox txt_CCI_MN3 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   28
            Top             =   1260
            Width           =   2100
         End
         Begin VB.TextBox txt_CCI_DL2 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   34
            Top             =   2160
            Width           =   2100
         End
         Begin VB.TextBox txt_CtaCorr_DL2 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   33
            Top             =   2160
            Width           =   2100
         End
         Begin VB.ComboBox cmb_Banco_DL2 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2190
            Width           =   3345
         End
         Begin VB.TextBox txt_CCI_DL1 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1830
            Width           =   2100
         End
         Begin VB.TextBox txt_CtaCorr_DL1 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   30
            Top             =   1830
            Width           =   2100
         End
         Begin VB.ComboBox cmb_Banco_DL1 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1860
            Width           =   3345
         End
         Begin VB.TextBox txt_CCI_MN2 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   25
            Top             =   930
            Width           =   2100
         End
         Begin VB.TextBox txt_CtaCorr_MN2 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   24
            Top             =   930
            Width           =   2100
         End
         Begin VB.ComboBox cmb_Banco_MN2 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   900
            Width           =   3345
         End
         Begin VB.TextBox txt_CCI_MN1 
            Height          =   315
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   22
            Top             =   600
            Width           =   2100
         End
         Begin VB.TextBox txt_CtaCorr_MN1 
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   21
            Top             =   600
            Width           =   2100
         End
         Begin VB.ComboBox cmb_Banco_MN1 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   570
            Width           =   3345
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   83
            Top             =   2565
            Width           =   300
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   82
            Top             =   2550
            Width           =   1230
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Banco 3:"
            Height          =   195
            Left            =   150
            TabIndex        =   81
            Top             =   2580
            Width           =   645
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   80
            Top             =   1335
            Width           =   300
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   79
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Banco 3:"
            Height          =   195
            Left            =   150
            TabIndex        =   78
            Top             =   1290
            Width           =   645
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Banco Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   77
            Top             =   1650
            Width           =   1260
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Banco Soles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   76
            Top             =   330
            Width           =   1080
         End
         Begin VB.Label Label29 
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
            Left            =   150
            TabIndex        =   75
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Banco 1:"
            Height          =   195
            Left            =   150
            TabIndex        =   72
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   71
            Top             =   1890
            Width           =   1230
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   70
            Top             =   1905
            Width           =   300
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Banco 2:"
            Height          =   195
            Left            =   150
            TabIndex        =   69
            Top             =   2250
            Width           =   645
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   68
            Top             =   2220
            Width           =   1230
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   67
            Top             =   2235
            Width           =   300
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Banco 1:"
            Height          =   195
            Left            =   150
            TabIndex        =   66
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   65
            Top             =   660
            Width           =   1230
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   64
            Top             =   660
            Width           =   300
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Banco 2:"
            Height          =   195
            Left            =   150
            TabIndex        =   63
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   5220
            TabIndex        =   62
            Top             =   990
            Width           =   1230
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CCI:"
            Height          =   195
            Left            =   8850
            TabIndex        =   61
            Top             =   1005
            Width           =   300
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cuentas Detractoras:"
      Height          =   195
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frm_Ctb_RegCom_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_PrvEst()   As moddat_tpo_Genera
Dim l_str_CodCiu     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String

Private Sub cmb_Banco_MN1_Click()
    cmb_Banco_MN2.Enabled = True
    txt_CtaCorr_MN1.Enabled = True
    txt_CCI_MN1.Enabled = True
    If (cmb_Banco_MN1.ListIndex > -1) Then
        If (cmb_Banco_MN1.ItemData(cmb_Banco_MN1.ListIndex) = 11) Then
            txt_CtaCorr_MN1.MaxLength = 18
        Else
            txt_CtaCorr_MN1.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_Banco_MN2_Click()
    cmb_Banco_MN3.Enabled = True
    txt_CtaCorr_MN2.Enabled = True
    txt_CCI_MN2.Enabled = True
    If (cmb_Banco_MN2.ListIndex > -1) Then
        If (cmb_Banco_MN2.ItemData(cmb_Banco_MN2.ListIndex) = 11) Then
            txt_CtaCorr_MN2.MaxLength = 18
        Else
            txt_CtaCorr_MN2.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_Banco_MN3_Click()
    txt_CtaCorr_MN3.Enabled = True
    txt_CCI_MN3.Enabled = True
    If (cmb_Banco_MN3.ListIndex > -1) Then
        If (cmb_Banco_MN3.ItemData(cmb_Banco_MN3.ListIndex) = 11) Then
            txt_CtaCorr_MN3.MaxLength = 18
        Else
            txt_CtaCorr_MN3.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_Banco_DL1_Click()
    cmb_Banco_DL2.Enabled = True
    txt_CtaCorr_DL1.Enabled = True
    txt_CCI_DL1.Enabled = True
    If (cmb_Banco_DL1.ListIndex > -1) Then
        If (cmb_Banco_DL1.ItemData(cmb_Banco_DL1.ListIndex) = 11) Then
            txt_CtaCorr_DL1.MaxLength = 18
        Else
            txt_CtaCorr_DL1.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_Banco_DL2_Click()
    cmb_Banco_DL3.Enabled = True
    txt_CtaCorr_DL2.Enabled = True
    txt_CCI_DL2.Enabled = True
    If (cmb_Banco_DL2.ListIndex > -1) Then
        If (cmb_Banco_DL2.ItemData(cmb_Banco_DL2.ListIndex) = 11) Then
            txt_CtaCorr_DL2.MaxLength = 18
        Else
            txt_CtaCorr_DL2.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_Banco_DL3_Click()
    txt_CtaCorr_DL3.Enabled = True
    txt_CCI_DL3.Enabled = True
    If (cmb_Banco_DL3.ListIndex > -1) Then
        If (cmb_Banco_DL3.ItemData(cmb_Banco_DL3.ListIndex) = 11) Then
            txt_CtaCorr_DL3.MaxLength = 18
        Else
            txt_CtaCorr_DL3.MaxLength = 20
        End If
    End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
'      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
'      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   'l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
'      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
'      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
'      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Telf1)
'      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
'   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
'      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
'      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Telf1)
   End If
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
'      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
'      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
'   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
 '     l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
 '     l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
       txt_NumDoc.MaxLength = 8 'DNI
   ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
       txt_NumDoc.MaxLength = 11 'RUC
   Else
       txt_NumDoc.MaxLength = 12
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmb_TipPer_Click()
    If (cmb_TipPer.ListIndex <> 1) Then
        txt_CodSic.Enabled = False
        txt_CodSic.Text = ""
        cmb_TarCre.Enabled = False
        Call gs_BuscarCombo_Item(cmb_TarCre, 2)
        ipp_FecIng.Enabled = False
        ipp_FecIng.Text = ""
    Else
        txt_CodSic.Enabled = True
        cmb_TarCre.Enabled = True
        ipp_FecIng.Enabled = True
        ipp_FecIng.Text = date
    End If
End Sub

Private Sub cmb_TipPer_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If txt_CodSic.Enabled = True Then
           Call gs_SetFocus(txt_CodSic)
        Else
           Call gs_SetFocus(txt_DomFisc)
        End If
    End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If (Len(Trim(moddat_g_str_TipDoc)) > 0 And Len(Trim(moddat_g_str_NumDoc)) > 0) Then
       If moddat_g_int_FlgGrb = 2 Then 'UPDATE
          pnl_Titulo.Caption = "Mantenimiento del Proveedor"
          Call fs_Cargar_Datos
       ElseIf moddat_g_int_FlgGrb = 0 Then 'CONSULTAR
          pnl_Titulo.Caption = "Consulta del Proveedor"
          cmd_Grabar.Visible = False
          Call fs_Cargar_Datos
          Call fs_Desabilitar
       End If
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPer, 1, "127")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCny, 1, "119")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Condicion, 1, "120")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Padron1, 1, "121")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Padron2, 1, "121")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPer, 1, "127")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TarCre, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_MN1, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_MN2, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_MN3, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_DL1, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_DL2, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco_DL3, 1, "122")
   
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub

Private Sub fs_Limpiar()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_RazSoc.Text = ""
   cmb_TipPer.ListIndex = 0
   Call gs_BuscarCombo_Item(cmb_TarCre, 2)
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_DomFisc.Text = ""
   txt_Telf1.Text = ""
   txt_Telf2.Text = ""
   txt_Telf3.Text = ""
   txt_Email.Text = ""
   txt_CtaDetrac.Text = ""
   cmb_TipCny.ListIndex = -1
   cmb_Condicion.ListIndex = -1
   cmb_Padron1.ListIndex = -1
   cmb_Padron2.ListIndex = -1
   
   cmb_Banco_MN1.ListIndex = -1
   txt_CtaCorr_MN1.Text = ""
   txt_CCI_MN1.Text = ""
   cmb_Banco_MN2.ListIndex = -1
   txt_CtaCorr_MN2.Text = ""
   txt_CCI_MN2.Text = ""
   cmb_Banco_MN3.ListIndex = -1
   txt_CtaCorr_MN3.Text = ""
   txt_CCI_MN3.Text = ""
   
   cmb_Banco_DL1.ListIndex = -1
   txt_CtaCorr_DL1.Text = ""
   txt_CCI_DL1.Text = ""
   cmb_Banco_DL2.ListIndex = -1
   txt_CtaCorr_DL2.Text = ""
   txt_CCI_DL2.Text = ""
   cmb_Banco_DL3.ListIndex = -1
   txt_CtaCorr_DL3.Text = ""
   txt_CCI_DL3.Text = ""
   
   cmb_Banco_MN1.ListIndex = -1
   txt_CtaCorr_MN1.Enabled = False
   txt_CCI_MN1.Enabled = False
   cmb_Banco_MN2.Enabled = False
   txt_CtaCorr_MN2.Enabled = False
   txt_CCI_MN2.Enabled = False
   cmb_Banco_MN3.Enabled = False
   txt_CtaCorr_MN3.Enabled = False
   txt_CCI_MN3.Enabled = False
   
   cmb_Banco_DL1.ListIndex = -1
   txt_CtaCorr_DL1.Enabled = False
   txt_CCI_DL1.Enabled = False
   cmb_Banco_DL2.Enabled = False
   txt_CtaCorr_DL2.Enabled = False
   txt_CCI_DL2.Enabled = False
   cmb_Banco_DL3.Enabled = False
   txt_CtaCorr_DL3.Enabled = False
   txt_CCI_DL3.Enabled = False
End Sub

Private Sub fs_Desabilitar()
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   txt_RazSoc.Enabled = False
   cmb_TipPer.Enabled = False
   cmb_TarCre.Enabled = False
   txt_CodSic.Enabled = False
   ipp_FecIng.Enabled = False
   cmb_DptDir.Enabled = False
   
   cmb_PrvDir.Enabled = False
   cmb_DstDir.Enabled = False
   
   txt_DomFisc.Enabled = False
   txt_Telf1.Enabled = False
   txt_Telf2.Enabled = False
   txt_Telf3.Enabled = False
   txt_Email.Enabled = False
   txt_CtaDetrac.Enabled = False
   cmb_TipCny.Enabled = False
   cmb_Condicion.Enabled = False
   cmb_Padron1.Enabled = False
   cmb_Padron2.Enabled = False
   
   cmb_Banco_MN1.Enabled = False
   txt_CtaCorr_MN1.Enabled = False
   txt_CCI_MN1.Enabled = False
   cmb_Banco_MN2.Enabled = False
   txt_CtaCorr_MN2.Enabled = False
   txt_CCI_MN2.Enabled = False
   cmb_Banco_MN3.Enabled = False
   txt_CtaCorr_MN3.Enabled = False
   txt_CCI_MN3.Enabled = False
   
   cmb_Banco_DL1.Enabled = False
   txt_CtaCorr_DL1.Enabled = False
   txt_CCI_DL1.Enabled = False
   cmb_Banco_DL2.Enabled = False
   txt_CtaCorr_DL2.Enabled = False
   txt_CCI_DL2.Enabled = False
   cmb_Banco_DL3.Enabled = False
   txt_CtaCorr_DL3.Enabled = False
   txt_CCI_DL3.Enabled = False
   
   cmb_Banco_MN1.Enabled = False
   txt_CtaCorr_MN1.Enabled = False
   txt_CCI_MN1.Enabled = False
   cmb_Banco_MN2.Enabled = False
   txt_CtaCorr_MN2.Enabled = False
   txt_CCI_MN2.Enabled = False
   cmb_Banco_MN3.Enabled = False
   txt_CtaCorr_MN3.Enabled = False
   txt_CCI_MN3.Enabled = False
   
   cmb_Banco_DL1.Enabled = False
   txt_CtaCorr_DL1.Enabled = False
   txt_CCI_DL1.Enabled = False
   cmb_Banco_DL2.Enabled = False
   txt_CtaCorr_DL2.Enabled = False
   txt_CCI_DL2.Enabled = False
   cmb_Banco_DL3.Enabled = False
   txt_CtaCorr_DL3.Enabled = False
   txt_CCI_DL3.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Cargar_Datos()
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   Call gs_SetFocus(txt_RazSoc)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC, MAEPRV_UBIGEO, MAEPRV_DOMFIS, MAEPRV_TELEF1, "
   g_str_Parame = g_str_Parame & "        MAEPRV_TELEF2, MAEPRV_TELEF3, MAEPRV_CORREO, MAEPRV_CTADET, MAEPRV_TIPCNT, MAEPRV_CONDIC, "
   g_str_Parame = g_str_Parame & "        MAEPRV_PADRN1, MAEPRV_PADRN2, MAEPRV_CODBNC_MN1, MAEPRV_CTACRR_MN1, MAEPRV_NROCCI_MN1, "
   g_str_Parame = g_str_Parame & "        MAEPRV_CODBNC_MN2, MAEPRV_CTACRR_MN2, MAEPRV_NROCCI_MN2, MAEPRV_CODBNC_MN3, MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        MAEPRV_NROCCI_MN3, MAEPRV_CODBNC_DL1, MAEPRV_CTACRR_DL1, MAEPRV_NROCCI_DL1, MAEPRV_CODBNC_DL2, "
   g_str_Parame = g_str_Parame & "        MAEPRV_CTACRR_DL2, MAEPRV_NROCCI_DL2, MAEPRV_CODBNC_DL3, MAEPRV_CTACRR_DL3, MAEPRV_NROCCI_DL3, "
   g_str_Parame = g_str_Parame & "        MAEPRV_TIPPER, MAEPRV_CODSIC, MAEPRV_SITTAR, MAEPRV_FECING  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV "
   g_str_Parame = g_str_Parame & "  WHERE MAEPRV_TIPDOC = " & moddat_g_str_TipDoc
   g_str_Parame = g_str_Parame & "    AND MAEPRV_NUMDOC = " & moddat_g_str_NumDoc

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!MAEPRV_TIPDOC)
      txt_NumDoc.Text = Trim(g_rst_Princi!MAEPRV_NUMDOC & "")
      txt_RazSoc.Text = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      Call gs_BuscarCombo_Item(cmb_TipPer, g_rst_Princi!MAEPRV_TIPPER)
      txt_CodSic.Text = Trim(g_rst_Princi!MAEPRV_CODSIC & "")
      Call gs_BuscarCombo_Item(cmb_TarCre, g_rst_Princi!MAEPRV_SITTAR)
      ipp_FecIng.Text = ""
      If Trim(g_rst_Princi!MAEPRV_FECING & "") <> "" Then
         ipp_FecIng.Text = gf_FormatoFecha(CStr(g_rst_Princi!MAEPRV_FECING))
      End If
      txt_DomFisc.Text = Trim(g_rst_Princi!MAEPRV_DOMFIS & "")
      If (Len(Trim(g_rst_Princi!MAEPRV_Ubigeo & "")) > 0) Then
          Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!MAEPRV_Ubigeo, 2)))
          'Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!MAEPRV_Ubigeo, 2))
          Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!MAEPRV_Ubigeo, 3, 2)))
          'Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!MAEPRV_Ubigeo, 2), Mid(g_rst_Princi!MAEPRV_Ubigeo, 3, 2))
          Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!MAEPRV_Ubigeo, 2)))
          'l_str_UbiGeo = Trim(g_rst_Princi!MAEPRV_Ubigeo & "")
     End If
      txt_Telf1.Text = Trim(g_rst_Princi!MAEPRV_TELEF1 & "")
      txt_Telf2.Text = Trim(g_rst_Princi!MAEPRV_TELEF2 & "")
      txt_Telf3.Text = Trim(g_rst_Princi!MAEPRV_TELEF3 & "")
      txt_Email.Text = Trim(g_rst_Princi!MAEPRV_CORREO & "")
      txt_CtaDetrac.Text = Trim(g_rst_Princi!MaePrv_CtaDet & "")

      Call gs_BuscarCombo_Item(cmb_TipCny, g_rst_Princi!MAEPRV_TIPCNT)
      Call gs_BuscarCombo_Item(cmb_Condicion, g_rst_Princi!MAEPRV_CONDIC)
      Call gs_BuscarCombo_Item(cmb_Padron1, g_rst_Princi!MAEPRV_PADRN1)
      Call gs_BuscarCombo_Item(cmb_Padron2, g_rst_Princi!MAEPRV_PADRN2)

      Call gs_BuscarCombo_Item(cmb_Banco_MN1, Trim(g_rst_Princi!MAEPRV_CODBNC_MN1))
      txt_CtaCorr_MN1.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_MN1 & "")
      txt_CCI_MN1.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_MN1 & "")

      Call gs_BuscarCombo_Item(cmb_Banco_MN2, Trim(g_rst_Princi!MAEPRV_CODBNC_MN2))
      txt_CtaCorr_MN2.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_MN2 & "")
      txt_CCI_MN2.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_MN2 & "")
      
      Call gs_BuscarCombo_Item(cmb_Banco_MN3, Trim(g_rst_Princi!MAEPRV_CODBNC_MN3))
      txt_CtaCorr_MN3.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_MN3 & "")
      txt_CCI_MN3.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_MN3 & "")
      
      Call gs_BuscarCombo_Item(cmb_Banco_DL1, Trim(g_rst_Princi!MAEPRV_CODBNC_DL1))
      txt_CtaCorr_DL1.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_DL1 & "")
      txt_CCI_DL1.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_DL1 & "")
   
      Call gs_BuscarCombo_Item(cmb_Banco_DL2, Trim(g_rst_Princi!MAEPRV_CODBNC_DL2))
      txt_CtaCorr_DL2.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_DL2 & "")
      txt_CCI_DL2.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_DL2 & "")
      
      Call gs_BuscarCombo_Item(cmb_Banco_DL3, Trim(g_rst_Princi!MAEPRV_CODBNC_DL3))
      txt_CtaCorr_DL3.Text = Trim(g_rst_Princi!MAEPRV_CTACRR_DL3 & "")
      txt_CCI_DL3.Text = Trim(g_rst_Princi!MAEPRV_NROCCI_DL3 & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Msg_1 As String
Dim r_str_Msg_2 As String
Dim r_str_Msg_3 As String

    r_str_Msg_1 = "La cuenta corriente del BBVA consta de 18 digitos."
    r_str_Msg_2 = "Tiene que digitar una cuenta corriente."
    r_str_Msg_3 = "La cuenta corriente interbancaria (CCI), consta de 20 digitos."
    
    If (cmb_TipDoc.ListIndex = -1) Then
        MsgBox "Debe seleccionar el tipo de documento.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipDoc)
        Exit Sub
    End If
       
    If (Len(Trim(txt_NumDoc.Text)) = 0) Then
        MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_NumDoc)
        Exit Sub
    End If
    
    If (fs_ValTipoDoc() = False) Then
        Exit Sub
    End If
    
    If (Len(Trim(txt_RazSoc.Text)) = 0) Then
        MsgBox "Debe ingresar la razon social/nombre.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_RazSoc)
        Exit Sub
    End If
    
    If (cmb_TipPer.ListIndex = -1) Then
        MsgBox "Debe seleccionar un tipo de personal.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipPer)
        Exit Sub
    End If
    
    If (cmb_TipPer.ListIndex = 1) Then
        If Len(Trim(txt_CodSic.Text)) <> 6 Then
            MsgBox "Debe ingresar un codigo de planilla de 6 digitos.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_CodSic)
            Exit Sub
        End If
        If (cmb_TarCre.ListIndex = -1) Then
            MsgBox "Debe seleccionar si se le asigna una tarjeta de crédito.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TarCre)
            Exit Sub
        End If
        If (Trim(ipp_FecIng.Text) = "") Then
            MsgBox "Debe de ingresar la fecha de ingreso.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FecIng)
            Exit Sub
        End If
    End If
    
    If (cmb_DptDir.ListIndex <> -1) Then
        If (cmb_PrvDir.ListIndex = -1) Then
            MsgBox "Debe seleccionar una provincia.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PrvDir)
            Exit Sub
        Else
            If (cmb_DstDir.ListIndex = -1) Then
                MsgBox "Debe seleccionar un distrito.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(cmb_DstDir)
                Exit Sub
            End If
        End If
    End If
      
    If Len(Trim(txt_Email.Text)) > 0 Then
       If Not gf_ValidarEmail(txt_Email.Text) Then
          MsgBox "El E-mail del proveedor no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_Email)
          Exit Sub
       End If
    End If
    
    If Trim(txt_CtaDetrac.Text) <> "" Then
       If Len(Trim(txt_CtaDetrac.Text)) <> 11 Then
          MsgBox "La cuenta corriente detractora debe de contar de  11 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_CtaDetrac)
          Exit Sub
       End If
    End If
    
    If (cmb_TipCny.ListIndex = -1) Then
        MsgBox "Debe seleccionar el tipo de contribuyente.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipCny)
        Exit Sub
    End If
    
    If (cmb_Condicion.ListIndex = -1) Then
        MsgBox "Debe seleccionar una condición.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Condicion)
        Exit Sub
    End If
    
    If (cmb_Padron1.ListIndex = -1) Then
        MsgBox "Debe seleccionar un tipo de padron 1.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Padron1)
        Exit Sub
    End If

    If (cmb_Padron2.ListIndex = -1) Then
        MsgBox "Debe seleccionar un tipo de padron 2.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Padron2)
        Exit Sub
    End If
    '-----------------------SOLES------------------------------------
    If (cmb_Banco_MN1.ListIndex > -1) Then
        If (cmb_Banco_MN1.ItemData(cmb_Banco_MN1.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_MN1.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_MN1)
                Exit Sub
            End If
        Else
          If (Len(Trim(txt_CtaCorr_MN1.Text)) = 0) Then
              MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(txt_CtaCorr_MN1)
              Exit Sub
          End If
          If (Len(Trim(txt_CCI_MN1.Text)) <> 20) Then
              MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(txt_CCI_MN1)
              Exit Sub
          End If
        End If
    End If

    If (cmb_Banco_MN2.ListIndex > -1) Then
        If (cmb_Banco_MN2.ItemData(cmb_Banco_MN2.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_MN2.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_MN2)
                Exit Sub
            End If
        Else
            If (Len(Trim(txt_CtaCorr_MN2.Text)) = 0) Then
                MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_MN2)
                Exit Sub
            End If
            If (Len(Trim(txt_CCI_MN2.Text)) <> 20) Then
                MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CCI_MN2)
                Exit Sub
            End If
        End If
    End If
    
    If (cmb_Banco_MN3.ListIndex > -1) Then
        If (cmb_Banco_MN3.ItemData(cmb_Banco_MN3.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_MN3.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_MN3)
                Exit Sub
            End If
        Else
            If (Len(Trim(txt_CtaCorr_MN3.Text)) = 0) Then
                MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_MN3)
                Exit Sub
            End If
            If (Len(Trim(txt_CCI_MN3.Text)) <> 20) Then
                MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CCI_MN3)
                Exit Sub
            End If
        End If
    End If
    '-----------------------DOLARES----------------------------------
    If (cmb_Banco_DL1.ListIndex > -1) Then
        If (cmb_Banco_DL1.ItemData(cmb_Banco_DL1.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_DL1.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL1)
                Exit Sub
            End If
        Else
            If (Len(Trim(txt_CtaCorr_DL1.Text)) = 0) Then
                MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL1)
                Exit Sub
            End If
            If (Len(Trim(txt_CCI_DL1.Text)) <> 20) Then
                MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CCI_DL1)
                Exit Sub
            End If
        End If
    End If
    
    If (cmb_Banco_DL2.ListIndex > -1) Then
        If (cmb_Banco_DL2.ItemData(cmb_Banco_DL2.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_DL2.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL2)
                Exit Sub
            End If
        Else
            If (Len(Trim(txt_CtaCorr_DL2.Text)) = 0) Then
                MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL2)
                Exit Sub
            End If
            If (Len(Trim(txt_CCI_DL2.Text)) <> 20) Then
                MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CCI_DL2)
                Exit Sub
            End If
        End If
    End If
    
    If (cmb_Banco_DL3.ListIndex > -1) Then
        If (cmb_Banco_DL3.ItemData(cmb_Banco_DL3.ListIndex) = 11) Then
            If (Len(Trim(txt_CtaCorr_DL3.Text)) <> 18) Then
                MsgBox r_str_Msg_1, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL3)
                Exit Sub
            End If
        Else
            If (Len(Trim(txt_CtaCorr_DL3.Text)) = 0) Then
                MsgBox r_str_Msg_2, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CtaCorr_DL3)
                Exit Sub
            End If
            If (Len(Trim(txt_CCI_DL3.Text)) <> 20) Then
                MsgBox r_str_Msg_3, vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_CCI_DL3)
                Exit Sub
            End If
            
        End If
    End If
    
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEPRV ( "
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", " 'MAEPRV_TIPDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', " 'MAEPRV_NUMDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_RazSoc.Text) & "', " 'MAEPRV_RAZSOC
   g_str_Parame = g_str_Parame & CStr(cmb_TipPer.ItemData(cmb_TipPer.ListIndex)) & ", " 'AS_MAEPRV_TIPPER
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CodSic.Text) & "', " 'AS_MAEPRV_CODSIC
   g_str_Parame = g_str_Parame & CStr(cmb_TarCre.ItemData(cmb_TarCre.ListIndex)) & ", " 'MAEPRV_TARCRE
   If Trim(ipp_FecIng.Text) = "" Then
      g_str_Parame = g_str_Parame & "NULL, " 'MAEPRV_FECING
   Else
      g_str_Parame = g_str_Parame & Format(Trim(ipp_FecIng.Text), "YYYYMMDD") & ", " 'MAEPRV_FECING
   End If
   If (cmb_DptDir.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "'', " 'MAEPRV_UBIGEO
   Else
       g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', " 'MAEPRV_UBIGEO
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_DomFisc.Text) & "', " 'MAEPRV_DOMFIS
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telf1.Text) & "', " 'MAEPRV_TELEF1
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telf2.Text) & "', " 'MAEPRV_TELEF2
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telf3.Text) & "', " 'MAEPRV_TELEF3
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Email.Text) & "', " 'MAEPRV_CORREO
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaDetrac.Text) & "', " 'MAEPRV_CTADET
   g_str_Parame = g_str_Parame & CStr(cmb_TipCny.ItemData(cmb_TipCny.ListIndex)) & ", " 'MAEPRV_TIPCNT
   g_str_Parame = g_str_Parame & CStr(cmb_Condicion.ItemData(cmb_Condicion.ListIndex)) & ", " 'MAEPRV_CONDIC
   g_str_Parame = g_str_Parame & CStr(cmb_Padron1.ItemData(cmb_Padron1.ListIndex)) & ", " 'MAEPRV_PADRN1
   g_str_Parame = g_str_Parame & CStr(cmb_Padron2.ItemData(cmb_Padron2.ListIndex)) & ", " 'MAEPRV_PADRN2
   If (cmb_Banco_MN1.ListIndex = -1) Then 'MAEPRV_CODBNC_MN1
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_MN1.ItemData(cmb_Banco_MN1.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_MN1.Text) & "', " 'MAEPRV_CTACRR_MN1
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_MN1.Text) & "', " 'MAEPRV_NROCCI_MN1
   
   If (cmb_Banco_MN2.ListIndex = -1) Then 'MAEPRV_CODBNC_MN2
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_MN2.ItemData(cmb_Banco_MN2.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_MN2.Text) & "', " 'MAEPRV_CTACRR_MN2
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_MN2.Text) & "', " 'MAEPRV_NROCCI_MN2
   
   If (cmb_Banco_MN3.ListIndex = -1) Then 'MAEPRV_CODBNC_MN3
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_MN3.ItemData(cmb_Banco_MN3.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_MN3.Text) & "', " 'MAEPRV_CTACRR_MN3
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_MN3.Text) & "', " 'MAEPRV_NROCCI_MN3
   
   If (cmb_Banco_DL1.ListIndex = -1) Then 'MAEPRV_CODBNC_DL1
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_DL1.ItemData(cmb_Banco_DL1.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_DL1.Text) & "', " 'MAEPRV_CTACRR_DL1
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_DL1.Text) & "', " 'MAEPRV_NROCCI_DL1
   
   If (cmb_Banco_DL2.ListIndex = -1) Then 'MAEPRV_CODBNC_DL2
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_DL2.ItemData(cmb_Banco_DL2.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_DL2.Text) & "', " 'MAEPRV_CTACRR_DL2
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_DL2.Text) & "', " 'MAEPRV_NROCCI_DL2
   
   If (cmb_Banco_DL3.ListIndex = -1) Then 'MAEPRV_CODBNC_DL3
       g_str_Parame = g_str_Parame & " 0 , "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_Banco_DL3.ItemData(cmb_Banco_DL3.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CtaCorr_DL3.Text) & "', " 'MAEPRV_CTACRR_DL3
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CCI_DL3.Text) & "', " 'MAEPRV_NROCCI_DL3
   
   g_str_Parame = g_str_Parame & " 1 , " 'MAEPRV_SITUAC
   
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (g_rst_Genera!RESUL = 0) Then
       MsgBox "El numero de ruc ya fue registrado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   ElseIf (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
       frm_Ctb_RegCom_01.fs_BuscarProv
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       frm_Ctb_RegCom_01.fs_BuscarProv
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Function fs_ValTipoDoc() As Boolean
   fs_ValTipoDoc = True
   
   If (cmb_TipDoc.ListIndex > -1) Then
       If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
           If (Len(Trim(txt_NumDoc.Text)) <> 8) Then  'DNI 8
               MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               fs_ValTipoDoc = False
           End If
       ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
           If Not gf_Valida_RUC(Trim(txt_NumDoc.Text), Mid(Trim(txt_NumDoc.Text), Len(Trim(txt_NumDoc.Text)), 1)) Then
              MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(txt_NumDoc)
              fs_ValTipoDoc = False
           End If
           '----------------------------------
       End If
   End If
End Function

Private Sub ipp_FecIng_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_TarCre)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub txt_CodSic_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(ipp_FecIng)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub cmb_TarCre_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_DomFisc)
    End If
End Sub

Private Sub txt_DomFisc_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_DptDir)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
    End If
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_RazSoc)
        Call fs_ValTipoDoc
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub txt_NumDoc_LostFocus()
'   Call fs_ValTipoDoc
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_TipPer)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
    End If
End Sub

Private Sub txt_Telf1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_Telf2)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub txt_Telf2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_Telf3)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub txt_Telf3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_Email)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub txt_Email_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CtaDetrac)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_@.")
    End If
End Sub

Private Sub txt_CtaDetrac_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_TipCny)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
    End If
End Sub

Private Sub cmb_TipCny_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Condicion)
    End If
End Sub

Private Sub cmb_Condicion_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Padron1)
    End If
End Sub

Private Sub cmb_Padron1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Padron2)
    End If
End Sub

Private Sub cmb_Padron2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_MN1)
    End If
End Sub

Private Sub cmb_Banco_MN1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_MN1.ListIndex = -1) Then
            Call gs_SetFocus(cmb_Banco_DL1)
        Else
            Call gs_SetFocus(txt_CtaCorr_MN1)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_MN1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_MN1)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_MN1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_MN2)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub cmb_Banco_MN2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_MN2.ListIndex = -1) Then
            Call gs_SetFocus(cmb_Banco_DL1)
        Else
            Call gs_SetFocus(txt_CtaCorr_MN2)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_MN2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_MN2)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_MN2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_MN3)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub cmb_Banco_MN3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_MN3.ListIndex = -1) Then
            Call gs_SetFocus(cmb_Banco_DL1)
        Else
            Call gs_SetFocus(txt_CtaCorr_MN3)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_MN3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_MN3)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_MN3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_DL1)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_MN4_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmd_Grabar)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
    End If
End Sub

Private Sub cmb_Banco_DL1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_DL1.ListIndex = -1) Then
            Call gs_SetFocus(cmd_Grabar)
        Else
            Call gs_SetFocus(txt_CtaCorr_DL1)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_DL1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_DL1)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_DL1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_DL2)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub cmb_Banco_DL2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_DL2.ListIndex = -1) Then
            Call gs_SetFocus(cmd_Grabar)
        Else
            Call gs_SetFocus(txt_CtaCorr_DL2)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_DL2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_DL2)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_DL2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmb_Banco_DL3)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub cmb_Banco_DL3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (cmb_Banco_DL3.ListIndex = -1) Then
            Call gs_SetFocus(cmd_Grabar)
        Else
            Call gs_SetFocus(txt_CtaCorr_DL3)
        End If
    End If
End Sub

Private Sub txt_CtaCorr_DL3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(txt_CCI_DL3)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub

Private Sub txt_CCI_DL3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call gs_SetFocus(cmd_Grabar)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO) '& "-"
    End If
End Sub



