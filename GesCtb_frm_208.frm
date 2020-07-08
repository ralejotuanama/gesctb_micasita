VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GesCtb_frm_208.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7980
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   14265
      _Version        =   65536
      _ExtentX        =   25162
      _ExtentY        =   14076
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
         Height          =   650
         Left            =   40
         TabIndex        =   12
         Top             =   60
         Width           =   14130
         _Version        =   65536
         _ExtentX        =   24924
         _ExtentY        =   1147
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.32
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
            TabIndex        =   13
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Compensación - Acción"
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
            Picture         =   "GesCtb_frm_208.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   40
         TabIndex        =   14
         Top             =   750
         Width           =   14130
         _Version        =   65536
         _ExtentX        =   24924
         _ExtentY        =   1147
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.32
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Cheque 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   3630
            Picture         =   "GesCtb_frm_208.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Impresión de Cheque"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Nuevo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_208.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_208.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_Consulta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_208.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3030
            Picture         =   "GesCtb_frm_208.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Reversa 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_208.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_208.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   13500
            Picture         =   "GesCtb_frm_208.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2400
         Left            =   40
         TabIndex        =   15
         Top             =   1440
         Width           =   14130
         _Version        =   65536
         _ExtentX        =   24924
         _ExtentY        =   4233
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
         Begin VB.ComboBox cmb_Bancos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1305
            Width           =   3345
         End
         Begin VB.ComboBox cmb_Moneda 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1650
            Width           =   3345
         End
         Begin VB.TextBox txt_Referen 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            MaxLength       =   60
            TabIndex        =   8
            Top             =   1980
            Width           =   2500
         End
         Begin VB.ComboBox cmb_TipPag 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   975
            Width           =   3345
         End
         Begin Threed.SSPanel pnl_CtaNom 
            Height          =   315
            Left            =   9660
            TabIndex        =   7
            Top             =   1650
            Width           =   4125
            _Version        =   65536
            _ExtentX        =   7267
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
         Begin Threed.SSPanel pnl_CtaCod 
            Height          =   315
            Left            =   6270
            TabIndex        =   6
            Top             =   1650
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
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
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   315
            Left            =   6270
            TabIndex        =   2
            Top             =   645
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1350
            TabIndex        =   0
            Top             =   300
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
         Begin EditLib.fpDateTime ipp_FchCtb 
            Height          =   315
            Left            =   1350
            TabIndex        =   1
            Top             =   645
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   1410
            Width           =   510
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "T.Cambio SBS:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5130
            TabIndex        =   31
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pago:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Cta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8370
            TabIndex        =   20
            Top             =   1740
            Width           =   1170
         End
         Begin VB.Label Label25 
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
            Left            =   240
            TabIndex        =   19
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Referencia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   2070
            Width           =   825
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   1740
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Cargo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5130
            TabIndex        =   16
            Top             =   1740
            Width           =   1020
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4020
         Left            =   40
         TabIndex        =   21
         Top             =   3900
         Width           =   14130
         _Version        =   65536
         _ExtentX        =   24924
         _ExtentY        =   7091
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
         Begin Threed.SSPanel pnl_PagNet_Sol 
            Height          =   315
            Left            =   12540
            TabIndex        =   47
            Top             =   3300
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin Threed.SSPanel pnl_PagNet_Dol 
            Height          =   315
            Left            =   12540
            TabIndex        =   49
            Top             =   3630
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin Threed.SSPanel pnl_ImpTot_Sol 
            Height          =   315
            Left            =   9600
            TabIndex        =   43
            Top             =   3300
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin Threed.SSPanel pnl_ImpTot_Dol 
            Height          =   315
            Left            =   9600
            TabIndex        =   45
            Top             =   3630
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2910
            Left            =   30
            TabIndex        =   22
            Top             =   360
            Width           =   14085
            _ExtentX        =   24844
            _ExtentY        =   5133
            _Version        =   393216
            Rows            =   24
            Cols            =   23
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   2610
            TabIndex        =   23
            Top             =   60
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Documento"
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
            Left            =   3840
            TabIndex        =   24
            Top             =   60
            Width           =   2730
            _Version        =   65536
            _ExtentX        =   4815
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proveedor"
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
            Left            =   9870
            TabIndex        =   25
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   9180
            TabIndex        =   26
            Top             =   60
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1235
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
            Left            =   7245
            TabIndex        =   27
            Top             =   60
            Width           =   1950
            _Version        =   65536
            _ExtentX        =   3440
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Corriente"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   1140
            TabIndex        =   33
            Top             =   60
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Proceso"
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
            Left            =   10950
            TabIndex        =   36
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Aplicación"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   11820
            TabIndex        =   37
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Dscto."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   12705
            TabIndex        =   38
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Pago Neto"
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
            Left            =   6540
            TabIndex        =   51
            Top             =   60
            Width           =   730
            _Version        =   65536
            _ExtentX        =   1288
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descrip."
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Pago Neto ME:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11340
            TabIndex        =   50
            Top             =   3705
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Pago Neto MN:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11340
            TabIndex        =   48
            Top             =   3375
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Totales  ==>  Importe ME:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7650
            TabIndex        =   46
            Top             =   3705
            Width           =   1830
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Totales  ==>  Importe MN:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7650
            TabIndex        =   44
            Top             =   3375
            Width           =   1845
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_int_Contar        As Integer
Dim l_int_TopNiv        As Integer


Private Sub cmb_Moneda_Click()
   pnl_CtaCod.Caption = ""
   pnl_CtaNom.Caption = ""
      
   If cmb_Moneda.ListIndex > -1 Then
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
         'cuenta contable soles - BBVA
         pnl_CtaCod.Caption = "111301060102"
         pnl_CtaNom.Caption = moddat_gf_Consulta_NomCtaCtb(moddat_g_str_CodEmp, Trim(pnl_CtaCod.Caption))
      End If
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 Then
         'cuenta contable dolares - BBVA
         pnl_CtaCod.Caption = "112301060102"
         pnl_CtaNom.Caption = moddat_gf_Consulta_NomCtaCtb(moddat_g_str_CodEmp, Trim(pnl_CtaCod.Caption))
      End If
   End If
End Sub

Private Sub cmb_Moneda_LostFocus()
   Call cmb_Moneda_Click
End Sub

Private Sub cmb_TipPag_Click()
   If moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2 Then
      'si es adicion o modificar
      If cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 4 Then
         'detraccion siempre es en soles
         cmb_Moneda.ListIndex = 0
         cmb_Moneda.Enabled = False
      Else
         cmb_Moneda.ListIndex = 0
         cmb_Moneda.Enabled = True
      End If
   End If
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Row = -1 Then
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows = 1 Then
      grd_Listad.Rows = 0
   Else
      grd_Listad.RemoveItem (grd_Listad.Row)
   End If
   Call fs_SumTotal(8)
End Sub

Private Sub cmd_Cheque_Click()
Dim r_str_CadAux   As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 3 Then
      'CHEQUE
      r_str_CadAux = ""
      
      frm_Ctb_PagCom_08.ipp_FecChq.Text = date
      frm_Ctb_PagCom_08.txt_NomDe.Text = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 4)) 'PROVEEDOR
      If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
         frm_Ctb_PagCom_08.pnl_Import.Caption = pnl_PagNet_Sol.Caption & " "
      Else
         frm_Ctb_PagCom_08.pnl_Import.Caption = pnl_PagNet_Dol.Caption & " "
      End If
      frm_Ctb_PagCom_08.pnl_Moneda.Caption = Trim(cmb_Moneda.Text)
      frm_Ctb_PagCom_08.txt_CodOrigen.Text = "MODULO_COMPENSACION"
      frm_Ctb_PagCom_08.txt_CodOrigen.Tag = Trim(pnl_Codigo.Caption)
      frm_Ctb_PagCom_08.fs_NumeroLetra
      frm_Ctb_PagCom_08.Show 1
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
Dim r_bol_Estado   As Boolean
Dim r_dbl_TCSunat  As Double
Dim r_str_CadAux   As String
Dim r_str_NumDoc   As String
Dim r_str_NomPrv   As String
Dim r_int_Fila     As Integer

    r_dbl_TCSunat = 0
    r_str_CadAux = ""
    
   'VALIDA TIPO DE CAMBIO SBS
   If CDbl(pnl_TipCambio.Caption) = 0 Then
      MsgBox "Tiene que registrar el tipo de cambio sbs del día.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FchCtb)
      Exit Sub
   End If

   'VALIDA TIPO CAMBIO SUNAT-VENTA
   'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
   'TipTip = 1 - Venta / 2 - Compra
   r_dbl_TCSunat = moddat_gf_ObtieneTipCamDia(3, 2, Format(ipp_FchCtb.Text, "yyyymmdd"), 1)
   If r_dbl_TCSunat = 0 Then
      MsgBox "Tiene que registrar el tipo de cambio sunat-venta del día.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FchCtb)
      Exit Sub
   End If
   
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Tiene que selecconar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
      
   If cmb_TipPag.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar el tipo de pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPag)
      Exit Sub
   End If
   
   If Len(Trim(pnl_CtaCod.Caption)) = 0 Then
      MsgBox "La cuenta de Cargo no puede estar vacío.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   
   If Len(Trim(pnl_CtaNom.Caption)) = 0 Then
      MsgBox "La descripción de la cuenta no puede estar vacío.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   
   If Len(Trim(txt_Referen.Text)) = 0 Then
      MsgBox "Tiene que ingresar una referencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Referen)
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      MsgBox "Tiene que adicionar almenos un registro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Nuevo)
      Exit Sub
   End If
      
'   If (Format(ipp_FchCtb.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FchCtb.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FchCtb)
'       Exit Sub
'   End If

   If fs_ValidaPeriodo(ipp_FchCtb.Text) = False Then
      Exit Sub
   End If
         
   'VALIDACION TIPO DE PAGO
   r_str_NumDoc = Trim(grd_Listad.TextMatrix(0, 3))
   r_str_NomPrv = Trim(grd_Listad.TextMatrix(0, 4))
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 1 Or cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 6 Or cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 8 Then
       'TRANSFERENCIA, PAGO PROVEEDORES, HABERES
          If Trim(grd_Listad.TextMatrix(r_int_Fila, 6) & "") = "" Then
             MsgBox "El tipo pago seleccionado obliga a que la columna cuenta corriente sea obligatorio.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          'TODO MENOS DETRACCION
          If Mid(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 1, 2) = "06" And CInt(grd_Listad.TextMatrix(r_int_Fila, 20)) = 2 Then
             MsgBox "No se permite registros de tipo detracción, por el tipo de pago seleccionado.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          
         If cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 8 Then
              'HABERES (8)
              If Trim(grd_Listad.TextMatrix(r_int_Fila, 21)) = "" Then
                 MsgBox "El registro " & grd_Listad.TextMatrix(r_int_Fila, 0) & " no tiene código de planilla.", vbExclamation, modgen_g_str_NomPlt
                 Exit Sub
              End If
              If Trim(grd_Listad.TextMatrix(r_int_Fila, 22) & "") <> "2" Then
                 MsgBox "El registro " & grd_Listad.TextMatrix(r_int_Fila, 0) & " no es un personal interno.", vbExclamation, modgen_g_str_NomPlt
                 Exit Sub
              End If
              If CLng(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) <> CLng(grd_Listad.TextMatrix(r_int_Fila, 16)) Then
                 MsgBox "El tipo moneda seleccionada obliga a que la columna moneda sea " & Trim(cmb_Moneda.Text), vbExclamation, modgen_g_str_NomPlt
                 Exit Sub
              End If
         End If
                     
       ElseIf cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 4 Then
       'DETRACCION
          If Trim(grd_Listad.TextMatrix(r_int_Fila, 6) & "") = "" Then
             MsgBox "El tipo pago seleccionado obliga a que la columna cuenta corriente sea obligatorio.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          If Left(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 2) <> "06" Then
             MsgBox "Para el pago de detracción, el origen de los registros deben de ser: registro de compras", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          If CInt(grd_Listad.TextMatrix(r_int_Fila, 20)) <> 2 Then
             MsgBox "El tipo pago seleccionado obliga a que los registros sean solo detracciones.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          If CInt(grd_Listad.TextMatrix(r_int_Fila, 16)) <> 1 Then
             MsgBox "Solo se admiten registros en soles para los pagos en detracción.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
       Else 'CHEQUE O CARTA
          If CLng(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) <> CLng(grd_Listad.TextMatrix(r_int_Fila, 16)) Then
             MsgBox "El tipo moneda seleccionada obliga a que la columna moneda sea " & Trim(cmb_Moneda.Text), vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          'TODO MENOS DETRACCION
          If Mid(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 1, 2) = "06" And CInt(grd_Listad.TextMatrix(r_int_Fila, 20)) = 2 Then
             MsgBox "No se permite registros de tipo detracción, por el tipo de pago seleccionado.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          'mismo proveedor
          If r_str_NumDoc <> Trim(grd_Listad.TextMatrix(r_int_Fila, 3)) Then
             MsgBox "Los registros adicionados deben ser de un mismo proveedor:" & vbCrLf & _
                    r_str_NumDoc & " - " & r_str_NomPrv, vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
       End If
   Next
      
   r_str_CadAux = ""
   If fs_ValImp_Neg(r_str_CadAux) = False Then
       MsgBox "El pago del proveedor con nro documento " & r_str_CadAux & " no puede ser negativo.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Grabar(r_dbl_TCSunat)
   Screen.MousePointer = 0
End Sub

Private Function fs_ValImp_Neg(ByRef p_Codigo As String) As Boolean
Dim r_int_NumCab   As Integer
Dim r_int_NumDet   As Integer
Dim r_dbl_ImpBrt   As Double
Dim r_dbl_ImpNet   As Double

Dim r_str_NumDoc   As String
'Dim r_str_Estado   As Boolean

   fs_ValImp_Neg = True
   p_Codigo = ""

   For r_int_NumCab = 0 To grd_Listad.Rows - 1
       r_dbl_ImpBrt = 0: r_dbl_ImpNet = 0
       r_str_NumDoc = ""
       r_str_NumDoc = Trim(CStr(grd_Listad.TextMatrix(r_int_NumCab, 3)))
       If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = CInt(grd_Listad.TextMatrix(r_int_NumCab, 16)) Then
          r_dbl_ImpNet = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 11))
          r_dbl_ImpBrt = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 8))
       Else
          If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
             r_dbl_ImpNet = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 11)) * CDbl(pnl_TipCambio.Caption)
             r_dbl_ImpBrt = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 8)) * CDbl(pnl_TipCambio.Caption)
          Else
             r_dbl_ImpNet = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 11)) / CDbl(pnl_TipCambio.Caption)
             r_dbl_ImpBrt = CDbl(grd_Listad.TextMatrix(r_int_NumCab, 8)) / CDbl(pnl_TipCambio.Caption)
          End If
       End If
       'BUSCANDO EL MISMO PROVEEDOR
       For r_int_NumDet = 0 To grd_Listad.Rows - 1
           If r_int_NumCab <> r_int_NumDet Then
              If Trim(grd_Listad.TextMatrix(r_int_NumDet, 3)) = r_str_NumDoc Then
                 'CONVERSION MONEDA
                 If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = CInt(grd_Listad.TextMatrix(r_int_NumDet, 16)) Then
                    r_dbl_ImpBrt = r_dbl_ImpBrt + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 8))
                    r_dbl_ImpNet = r_dbl_ImpNet + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 11))
                 Else
                    If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
                       r_dbl_ImpNet = r_dbl_ImpNet + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 11)) * CDbl(pnl_TipCambio.Caption)
                       r_dbl_ImpBrt = r_dbl_ImpBrt + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 8)) * CDbl(pnl_TipCambio.Caption)
                    Else
                       r_dbl_ImpNet = r_dbl_ImpNet + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 11)) / CDbl(pnl_TipCambio.Caption)
                       r_dbl_ImpBrt = r_dbl_ImpBrt + CDbl(grd_Listad.TextMatrix(r_int_NumDet, 8)) / CDbl(pnl_TipCambio.Caption)
                    End If
                 End If
              End If
           End If
       Next
       If r_dbl_ImpBrt < 0 Then
          fs_ValImp_Neg = False
          p_Codigo = r_str_NumDoc
          Exit For
       End If
       If r_dbl_ImpNet < 0 Then
          fs_ValImp_Neg = False
          p_Codigo = r_str_NumDoc
          Exit For
       End If
   Next
End Function

Private Sub cmd_Nuevo_Click()
   frm_Ctb_PagCom_03.Show 1
   Call fs_SumTotal(8)
End Sub

Private Sub cmb_Consulta_Click()
Dim r_str_CodAux   As String
Dim r_str_FlgAux   As Integer

   r_str_CodAux = ""
   r_str_FlgAux = 0
   moddat_g_str_NumOpe = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   Select Case Left(grd_Listad.TextMatrix(grd_Listad.Row, 0), 2)
          Case "01" 'CUENTAS X PAGAR OPETRA
               moddat_g_str_NumOpe = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
               frm_Ctb_PagCom_04.Show 1
          Case "12" 'CUENTAS X PAGAR GESCTB
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CtaPag_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "07" 'GESTION PERSONAL
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_int_TipRec = 1 'GESTION DE PAGOS
                  frm_Ctb_GesPer_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "08" 'CARGA DEL ARCHIVO RECAUDO
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CarArc_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "06" 'REGISTRO DE COMPRAS
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 14))
                  moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 15))
                  moddat_g_int_InsAct = 0 'tipo registro compra
                  frm_Ctb_RegCom_04.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "05" 'ENTREGAS A RENDIR
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodMod = grd_Listad.TextMatrix(grd_Listad.Row, 16)
                  moddat_g_int_FlgGrb = 0
                  If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 20) & "") = "1" Then
                     frm_Ctb_EntRen_02.Show 1 'form principal
                  ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 20) & "") = "2" Then
                     frm_Ctb_EntRen_04.Show 1 'reembolso
                  End If
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
                  
          Case Else
               Exit Sub
   End Select
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Reversa_Click()
Dim r_bol_Estado As Boolean
   
   If MsgBox("¿Esta seguro que desea realizar esta operación de reversa?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMPAG_REVERSA ( "
   g_str_Parame = g_str_Parame & " " & CLng(Trim(pnl_Codigo.Caption)) & ", " 'COMPAG_CODCOM
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
       Call frm_Ctb_PagCom_01.fs_Buscar
       Unload Me
   ElseIf g_rst_Genera!as_resul = 2 Or g_rst_Genera!as_resul = 3 Then
       'procesado cuentas x pagar
       MsgBox "El registro ya fue revertido.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_PagCom_01.fs_Buscar
       Unload Me
   Else
       MsgBox "Favor de verificar la operación de reversa.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   cmd_Grabar.Visible = False
   cmd_Reversa.Visible = False
   cmd_Nuevo.Visible = False
   cmd_Borrar.Visible = False
   cmb_Consulta.Visible = False
   cmd_ExpExc.Visible = False
   
   If moddat_g_int_FlgGrb = 0 Then
      pnl_Titulo.Caption = "Compensación - Consultar"
      Call fs_Cargar_Datos
      Call fs_Desabilitar
      cmb_Consulta.Visible = True
      cmd_ExpExc.Visible = True
      cmb_Consulta.Left = 30
      cmd_ExpExc.Left = 630
      
      cmd_Cheque.Visible = False
      If cmb_TipPag.ItemData(cmb_TipPag.ListIndex) = 3 Then
         'CHEQUE
         cmd_Cheque.Visible = True
         cmd_Cheque.Left = 1230
      End If
   ElseIf moddat_g_int_FlgGrb = 1 Then
      pnl_Titulo.Caption = "Compensación - Adicionar"
      cmd_Grabar.Visible = True
      cmd_Nuevo.Visible = True
      cmd_Borrar.Visible = True
      cmb_Consulta.Visible = True
      cmd_ExpExc.Visible = True
      cmd_Nuevo.Left = 630
      cmd_Borrar.Left = 1230
      cmb_Consulta.Left = 1830
      cmd_ExpExc.Left = 2430
   ElseIf moddat_g_int_FlgGrb = 2 Then
      pnl_Titulo.Caption = "Compensación - Modificar"
      Call fs_Cargar_Datos
      cmd_Grabar.Visible = True
      cmd_Nuevo.Visible = True
      cmd_Borrar.Visible = True
      cmb_Consulta.Visible = True
      cmd_ExpExc.Visible = True
      cmd_Nuevo.Left = 630
      cmd_Borrar.Left = 1230
      cmb_Consulta.Left = 1830
      cmd_ExpExc.Left = 2430
   ElseIf moddat_g_int_FlgGrb = 3 Then
      pnl_Titulo.Caption = "Compensación - Reversa"
      cmd_Reversa.Visible = True
      cmb_Consulta.Visible = True
      cmd_ExpExc.Visible = True
      cmd_Reversa.Left = 30
      cmb_Consulta.Left = 630
      cmd_ExpExc.Left = 1230
      Call fs_Cargar_Datos
      Call fs_Desabilitar
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1080 'CODIGO
   grd_Listad.ColWidth(1) = 1470 'TIPO PROCESO
   grd_Listad.ColWidth(2) = 0    'FECHA - 1050
   grd_Listad.ColWidth(3) = 1240 'NRO DOCUMENTO
   grd_Listad.ColWidth(4) = 2700 'PROVEEDOR
   grd_Listad.ColWidth(5) = 690 'DESCRIPCION
   grd_Listad.ColWidth(6) = 1940 'CUENTA CORRIENTE
   grd_Listad.ColWidth(7) = 690  'MONEDA
   grd_Listad.ColWidth(8) = 1080 'IMPORTE
   grd_Listad.ColWidth(9) = 890 'APLICACION
   grd_Listad.ColWidth(10) = 870 'DESCUENTO
   grd_Listad.ColWidth(11) = 1080 'PAGO NETO
   grd_Listad.ColWidth(12) = 0 'CODIGO APLICACION
   grd_Listad.ColWidth(13) = 0 'COMAUT_CODAUT
   grd_Listad.ColWidth(14) = 0 'COMAUT_TIPDOC
   grd_Listad.ColWidth(15) = 0 'COMAUT_NUMDOC
   grd_Listad.ColWidth(16) = 0 'COMAUT_CODMON
   grd_Listad.ColWidth(17) = 0 'COMAUT_CODBNC
   grd_Listad.ColWidth(18) = 0 'COMAUT_CTACTB
   grd_Listad.ColWidth(19) = 0 'COMAUT_DATCTB
   grd_Listad.ColWidth(20) = 0 'COMAUT_TIPOPE
   
   grd_Listad.ColWidth(21) = 0 'MAEPRV_CODSIC
   grd_Listad.ColWidth(22) = 0 'MAEPRV_TIPPER
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter 'CODIGO
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   'grd_Listad.ColAlignment(2) = flexAlignCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter 'PROVEEDOR
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_SumTotal(8)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPag, 1, "135")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Bancos, 1, "122")
   cmb_Bancos.Enabled = False
End Sub

Private Sub fs_Limpia()
   pnl_Codigo.Caption = ""
   ipp_FchCtb.Text = date
   pnl_TipCambio.Caption = "0.000000" & " "
   cmb_Moneda.ListIndex = 0
   
   cmb_TipPag.ListIndex = 0
   Call gs_BuscarCombo_Item(cmb_Bancos, 11)
   
   pnl_CtaCod.Caption = ""
   pnl_CtaNom.Caption = ""
   txt_Referen.Text = ""
   
   Call ipp_FchCtb_LostFocus
   Call cmb_Moneda_Click
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Grabar(p_TCSunat_Vent As Double)
   
   '--01-INSERTAR COMPENSACION---
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMPAG ( "
   If moddat_g_int_FlgGrb = 1 Then
      'insert
      g_str_Parame = g_str_Parame & "0, " 'COMPAG_CODCOM
   Else
      'Update
      g_str_Parame = g_str_Parame & CLng(pnl_Codigo.Caption) & ", " 'COMPAG_CODCOM
   End If
   g_str_Parame = g_str_Parame & Format(ipp_FchCtb.Text, "yyyymmdd") & ", " 'COMPAG_FECPAG
   g_str_Parame = g_str_Parame & CDbl(pnl_TipCambio.Caption) & ", "  'COMPAG_TIPCAM
   g_str_Parame = g_str_Parame & cmb_TipPag.ItemData(cmb_TipPag.ListIndex) & ", " 'COMPAG_TIPPAG
   g_str_Parame = g_str_Parame & cmb_Bancos.ItemData(cmb_Bancos.ListIndex) & ", " 'COMPAG_CODBNC
   g_str_Parame = g_str_Parame & CLng(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) & ", " 'COMPAG_CODMON
   g_str_Parame = g_str_Parame & "'" & CStr(Trim(pnl_CtaCod.Caption)) & "', " 'COMPAG_CTACTB
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Referen.Text) & "', " 'COMPAG_REFERE
   g_str_Parame = g_str_Parame & "1, " 'COMPAG_SITUAC
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      'MsgBox "Error al ejecutar el Procedimiento CNTBL_COMPAG_COMDET.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   '--02-ELIMINAR DETALLE---
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMDET_BORRAR (  "
   g_str_Parame = g_str_Parame & Trim(CStr(g_rst_Genera!CODIGO)) & ", " 'COMDET_CODCOM
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      'MsgBox "Error al ejecutar el Procedimiento CNTBL_COMPAG_COMDET.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   '--03-INSERTAR DETALLE---
   For l_int_Contar = 0 To grd_Listad.Rows - 1
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & " USP_CNTBL_COMDET ( "
       g_str_Parame = g_str_Parame & CLng(g_rst_Genera!CODIGO) & ", " 'COMDET_CODCOM
              
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 13)) & ", " 'COMDET_CODAUT
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 0)) & ", " 'COMDET_CODOPE
       g_str_Parame = g_str_Parame & Format(grd_Listad.TextMatrix(l_int_Contar, 2), "yyyymmdd") & ", " 'COMDET_FECOPE
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 14)) & ", " 'COMDET_TIPDOC
       g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(l_int_Contar, 15)) & "', " 'COMDET_NUMDOC
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 16)) & ", " 'COMDET_CODMON
       g_str_Parame = g_str_Parame & CDbl(grd_Listad.TextMatrix(l_int_Contar, 8)) & ", " 'COMDET_IMPPAG
       If Trim(grd_Listad.TextMatrix(l_int_Contar, 17) & "") = "" Then
          g_str_Parame = g_str_Parame & "NULL , "   'COMDET_CODBNC
       Else
          g_str_Parame = g_str_Parame & Trim(grd_Listad.TextMatrix(l_int_Contar, 17) & "") & ", "   'COMDET_CODBNC
       End If
       g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(l_int_Contar, 6) & "") & "', " 'COMDET_CTACRR
       g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(l_int_Contar, 18) & "") & "', " 'COMDET_CTACTB
       g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(l_int_Contar, 19) & "") & "', " 'COMDET_DATCTB
       g_str_Parame = g_str_Parame & "1, " 'COMDET_SITUAC
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 12)) & ", " 'APLICACION
       g_str_Parame = g_str_Parame & CDbl(grd_Listad.TextMatrix(l_int_Contar, 10)) & ", " 'DESCUENTO
       If CDbl(grd_Listad.TextMatrix(l_int_Contar, 10)) > 0 Then
          g_str_Parame = g_str_Parame & CDbl(p_TCSunat_Vent) & ", " 'TIPO CAMBIO SUNAT-VENTA
       Else
          g_str_Parame = g_str_Parame & CDbl(pnl_TipCambio.Caption) & ", " 'TIPO CAMBIO SBS
       End If
       g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(l_int_Contar, 5) & "") & "', "   'COMDET_DESCRP
       g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(l_int_Contar, 20)) & ", " 'COMDET_TIPOPE
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
       g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
       g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          'MsgBox "Error al ejecutar el Procedimiento CNTBL_COMPAG_COMDET.", vbExclamation, modgen_g_str_NomPlt
          Exit Sub
       End If
   Next

   If (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_PagCom_01.fs_Buscar
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_PagCom_01.fs_Buscar
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub fs_SumTotal(p_Column As Integer)
Dim r_dbl_ImpSol As Double
Dim r_dbl_ImpDol As Double
Dim r_dbl_NetSol As Double
Dim r_dbl_NetDol As Double

   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   r_dbl_NetSol = 0
   r_dbl_NetDol = 0
   
   For l_int_Contar = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(l_int_Contar, 16) = 1 Then 'SOLES
          r_dbl_ImpSol = r_dbl_ImpSol + CDbl(grd_Listad.TextMatrix(l_int_Contar, p_Column))
          r_dbl_NetSol = r_dbl_NetSol + CDbl(grd_Listad.TextMatrix(l_int_Contar, p_Column + 3))
       Else
          r_dbl_ImpDol = r_dbl_ImpDol + CDbl(grd_Listad.TextMatrix(l_int_Contar, p_Column))
          r_dbl_NetDol = r_dbl_NetDol + CDbl(grd_Listad.TextMatrix(l_int_Contar, p_Column + 3))
       End If
   Next
   
   pnl_ImpTot_Sol.Caption = Format(r_dbl_ImpSol, "###,###,##0.00") & " "
   pnl_ImpTot_Dol.Caption = Format(r_dbl_ImpDol, "###,###,##0.00") & " "
   pnl_PagNet_Sol.Caption = Format(r_dbl_NetSol, "###,###,##0.00") & " "
   pnl_PagNet_Dol.Caption = Format(r_dbl_NetDol, "###,###,##0.00") & " "
End Sub

Private Sub fs_Cargar_Datos()
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.COMPAG_CODCOM, A.COMPAG_FECPAG, A.COMPAG_TIPCAM, A.COMPAG_TIPPAG,  "
   g_str_Parame = g_str_Parame & "        A.COMPAG_CODMON , A.COMPAG_CTACTB, A.COMPAG_REFERE  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMPAG A  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMPAG_CODCOM = " & CLng(moddat_g_str_Codigo)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Codigo.Caption = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      ipp_FchCtb.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      pnl_TipCambio.Caption = Format(g_rst_Princi!COMPAG_TIPCAM, "###,###,##0.000000") & " "
      Call gs_BuscarCombo_Item(cmb_TipPag, g_rst_Princi!COMPAG_TIPPAG)
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!COMPAG_CODMON)
      pnl_CtaCod.Caption = Trim(g_rst_Princi!COMPAG_CTACTB & "")
      pnl_CtaNom.Caption = moddat_gf_Consulta_NomCtaCtb(moddat_g_str_CodEmp, g_rst_Princi!COMPAG_CTACTB)
      txt_Referen.Text = Trim(g_rst_Princi!COMPAG_REFERE & "")
      'pnl_ImpTot.Caption = Format(g_rst_Princi!COMPAG_IMPTOT, "###,###,##0.00") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.COMDET_CODCOM, A.COMDET_CODAUT, A.COMDET_CODOPE, A.COMDET_FECOPE,  "
   g_str_Parame = g_str_Parame & "        A.COMDET_TIPDOC, A.COMDET_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS PROVEEDOR,  "
   g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) AS MONEDA, A.COMDET_IMPPAG, A.COMDET_CTACRR,  "
   g_str_Parame = g_str_Parame & "        A.COMDET_CODAUT, A.COMDET_CODMON,  "
   g_str_Parame = g_str_Parame & "        A.COMDET_CODBNC , A.COMDET_CTACTB, A.COMDET_DATCTB, TRIM(D.PARDES_DESCRI) AS TIPOPROCESO,  "
   g_str_Parame = g_str_Parame & "        A.COMDET_TIPDST, A.COMDET_IMPDST, A.COMDET_TIPOPE, A.COMDET_DESCRP,  "
   g_str_Parame = g_str_Parame & "        B.MAEPRV_CODSIC , B.MAEPRV_TIPPER "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMDET A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMDET_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMDET_NUMDOC)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMDET_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMDET_NUMDOC) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.COMDET_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 136 AND TO_NUMBER(D.PARDES_CODITE) = TO_NUMBER(SUBSTR(LPAD(COMDET_CODOPE,10,0),1,2)) AND D.PARDES_CODITE <> 0  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMDET_CODCOM =  " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "  ORDER BY PROVEEDOR ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
         
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!COMDET_CODOPE, "0000000000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!TIPOPROCESO & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMDET_FECOPE & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_TIPDOC) & "-" & Trim(g_rst_Princi!COMDET_NUMDOC)
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!PROVEEDOR & " ")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_DESCRP & " ")
            
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_CTACRR & " ")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & " ")
      
      grd_Listad.Col = 8
      grd_Listad.Text = Format(g_rst_Princi!COMDET_IMPPAG, "###,###,##0.00") & " "
      
      grd_Listad.Col = 9
      If g_rst_Princi!COMDET_TIPDST = 1 Then
         grd_Listad.Text = "NINGUNO"
      ElseIf g_rst_Princi!COMDET_TIPDST = 2 Then
         grd_Listad.Text = "ITF"
      ElseIf g_rst_Princi!COMDET_TIPDST = 3 Then
         grd_Listad.Text = "4TA"
      End If
      
      grd_Listad.Col = 10
      grd_Listad.Text = Format(g_rst_Princi!COMDET_IMPDST, "###,###,##0.00")
      grd_Listad.Col = 11
      grd_Listad.Text = Format(g_rst_Princi!COMDET_IMPPAG - g_rst_Princi!COMDET_IMPDST, "###,###,##0.00")
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!COMDET_TIPDST
            
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_CODAUT & " ")
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_TIPDOC & " ")
      
      grd_Listad.Col = 15
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_NUMDOC & " ")
      
      grd_Listad.Col = 16
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_CODMON & " ")
      
      grd_Listad.Col = 17
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_CODBNC & " ")
      
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_CTACTB & " ")
      
      grd_Listad.Col = 19
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_DATCTB & " ")
      
      grd_Listad.Col = 20
      grd_Listad.Text = Trim(g_rst_Princi!COMDET_TIPOPE & " ")
      
      grd_Listad.Col = 21
      grd_Listad.Text = Trim(g_rst_Princi!MAEPRV_CODSIC & " ")
      
      grd_Listad.Col = 22
      grd_Listad.Text = Trim(g_rst_Princi!MAEPRV_TIPPER & " ")
      
      g_rst_Princi.MoveNext
   Loop

   Call fs_SumTotal(8)
 
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_Desabilitar()
   ipp_FchCtb.Enabled = False
   cmb_TipPag.Enabled = False
   cmb_Moneda.Enabled = False
   txt_Referen.Enabled = False
   cmd_Nuevo.Enabled = False
   cmd_Borrar.Enabled = False
End Sub

Private Sub ipp_FchCtb_LostFocus()
  'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
  'TipTip = 1 - Venta / 2 - Compra
   pnl_TipCambio.Caption = Format(moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FchCtb.Text, "yyyymmdd"), 1), "###,###,##0.000000") & " "
End Sub


Private Sub txt_Referen_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Nuevo)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub ipp_FchCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipPag)
   End If
End Sub

Private Sub cmb_TipPag_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Referen)
   End If
End Sub

Public Sub fs_CalcITF(p_ImpTot As Double, ByRef p_AplNom As String, ByRef p_AplCod As Integer, ByRef p_AplDsct As String, ByRef p_ImpNeto As String)
Dim r_dbl_Import   As Double
Dim r_dbl_ImpAux   As Double
Dim r_str_CadAux   As String
   
   r_dbl_Import = 0
   r_dbl_Import = p_ImpTot
   r_dbl_ImpAux = Int((r_dbl_Import * (0.005 / 100)) * 100) / 100 '=+TRUNCAR(D12*(0.005/100),2)  Int ( OriginalNumber * 100 ) / 100
   
   r_dbl_ImpAux = IIf(CStr(Right(r_dbl_ImpAux, 2)) = "01", Int(r_dbl_ImpAux), r_dbl_ImpAux) '=+SI(DERECHA(D13,2)="01",TRUNCAR(D13,0),D13)
   r_str_CadAux = Format(r_dbl_ImpAux, "###,###.00") '=+TEXTO(D14,"###,###.00")
   r_dbl_ImpAux = CDbl(Mid(r_str_CadAux, 1, InStr(1, r_str_CadAux, ".") + 1)) + IIf(Int(Right(r_str_CadAux, 1)) < 5, 0, "0.05") '=+MED(D15,1,ENCONTRAR(".",D15,1)+1)+SI(DERECHA(D15,1)<5,0,0.05)
   
   p_AplNom = "ITF"
   p_AplDsct = Format(r_dbl_ImpAux, "###,###,###,##0.00")
   p_ImpNeto = Format(p_ImpTot - r_dbl_ImpAux, "###,###,###,##0.00")
   p_AplCod = 2 'CODIGO
End Sub

Public Sub fs_Calc4TA(p_ImpTot As Double, ByRef p_AplNom As String, ByRef p_AplCod As Integer, ByRef p_AplDsct As String, ByRef p_ImpNeto As String)
Dim r_dbl_Import    As Double
   
   r_dbl_Import = 0
   'r_dbl_Import = Round(p_ImpTot * (8 / 100), 0)
   r_dbl_Import = p_ImpTot * (8 / 100)
   
   p_AplNom = "4TA"
   p_AplDsct = Format(r_dbl_Import, "###,###,###,##0.00")
   p_ImpNeto = Format(p_ImpTot - r_dbl_Import, "###,###,###,##0.00")
   p_AplCod = 3 'CODIGO
End Sub

Public Sub fs_CalcNinguno(p_ImpTot As Double, ByRef p_AplNom As String, ByRef p_AplCod As Integer, ByRef p_AplDsct As String, ByRef p_ImpNeto As String)
   p_AplNom = "NINGUNO"
   p_AplDsct = Format(0, "###,###,###,##0.00")
   p_ImpNeto = Format(p_ImpTot, "###,###,###,##0.00")
   p_AplCod = 1 'CODIGO
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(1, 10) = "FECHA"
      .Cells(1, 11) = Format(date, "dd/mm/yyyy")
      .Cells(2, 2) = UCase(Trim(pnl_Titulo.Caption))
      .Range(.Cells(2, 2), .Cells(2, 12)).Merge
      .Range(.Cells(2, 2), .Cells(2, 12)).Font.Bold = True

      .Cells(4, 2) = "DATOS":
      .Cells(4, 2).Font.Bold = True
      .Cells(5, 2) = "CODIGO:": .Cells(5, 3) = "'" & Trim(pnl_Codigo.Caption)
      .Cells(6, 2) = "FECHA:": .Cells(6, 3) = "'" & ipp_FchCtb.Text
      .Cells(7, 2) = "TIPO PAGO:": .Cells(7, 3) = cmb_TipPag.Text
      .Cells(8, 2) = "BANCO:": .Cells(8, 3) = cmb_Bancos.Text
      .Cells(9, 2) = "MONEDA:": .Cells(9, 3) = cmb_Moneda.Text
      .Cells(10, 2) = "REFERENCIA:": .Cells(10, 3) = "'" & Trim(txt_Referen.Text)
      .Cells(6, 5) = "T.CAMBIO SBS    : " & pnl_TipCambio.Caption
      .Cells(9, 5) = "CUENTA CARGO: " & pnl_CtaCod.Caption
      .Cells(9, 6) = "DESCRIPCION CTA.:": .Cells(9, 7) = pnl_CtaNom.Caption

      .Cells(12, 2) = "CÓDIGO"
      .Cells(12, 3) = "TIPO PROCESO"
      .Cells(12, 4) = "NRO DOCUMENTO"
      .Cells(12, 5) = "PROVEEDOR"
      .Cells(12, 6) = "DESCRIPCIÓN"
      .Cells(12, 7) = "CUENTA CORRIENTE"
      .Cells(12, 8) = "MONEDA"
      .Cells(12, 9) = "IMPORTE"
      .Cells(12, 10) = "APLICACION"
      .Cells(12, 11) = "DESCUENTO"
      .Cells(12, 12) = "PAGO NETO"
         
      .Range(.Cells(12, 2), .Cells(12, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 2), .Cells(12, 12)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'CÓDIGO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22 'TIPO PROCESO
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 17 'NRO DOCUMENTO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 45 'PROVEEDOR
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 20 'DESCRIPCION
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 24 'CUENTA CORRIENTE
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 18 'MONEDA
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 15 'IMPORTE
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 13 'APLICACION
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 16 'DESCUENTO
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 15 'PAGO NETO
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 12)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 12)).Font.Size = 11
      
      r_int_NumFil = 13
      For l_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(l_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = grd_Listad.TextMatrix(l_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(l_int_Contar, 3)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(l_int_Contar, 4)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(l_int_Contar, 5)
         .Cells(r_int_NumFil, 7) = "'" & grd_Listad.TextMatrix(l_int_Contar, 6)
         .Cells(r_int_NumFil, 8) = grd_Listad.TextMatrix(l_int_Contar, 7)
         .Cells(r_int_NumFil, 9) = grd_Listad.TextMatrix(l_int_Contar, 8)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(l_int_Contar, 9)
         .Cells(r_int_NumFil, 11) = grd_Listad.TextMatrix(l_int_Contar, 10)
         .Cells(r_int_NumFil, 12) = grd_Listad.TextMatrix(l_int_Contar, 11)
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Cells(2, 2).HorizontalAlignment = xlHAlignCenter 'titulo principal
      .Cells(9, 7).HorizontalAlignment = xlHAlignLeft 'descripcion cta
      .Range(.Cells(4, 2), .Cells(10, 2)).HorizontalAlignment = xlHAlignLeft 'cabecera datos
      .Range(.Cells(12, 2), .Cells(12, 12)).HorizontalAlignment = xlHAlignCenter 'titulo detalle
      
      .Cells(r_int_NumFil + 1, 7) = "TOTALES PAGO ==> "
      .Cells(r_int_NumFil + 2, 7) = "TOTALES PAGO ==> "
      .Cells(r_int_NumFil + 1, 8) = "IMPORTE MN:"
      .Cells(r_int_NumFil + 2, 8) = "IMPORTE ME:"
      .Cells(r_int_NumFil + 1, 9) = pnl_ImpTot_Sol.Caption
      .Cells(r_int_NumFil + 2, 9) = pnl_ImpTot_Dol.Caption
      
      .Range(.Cells(13, 8), .Cells(r_int_NumFil + 2, 12)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(13, 10), .Cells(r_int_NumFil + 2, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 8), .Cells(r_int_NumFil, 8)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NumFil + 1, 11) = "PAGO NETO MN:"
      .Cells(r_int_NumFil + 2, 11) = "PAGO NETO ME:"
      .Cells(r_int_NumFil + 1, 12) = pnl_PagNet_Sol.Caption
      .Cells(r_int_NumFil + 2, 12) = pnl_PagNet_Dol.Caption
      .Range(.Cells(r_int_NumFil + 1, 7), .Cells(r_int_NumFil + 2, 12)).Font.Bold = True
      
      .Cells(1, 12).HorizontalAlignment = xlHAlignCenter 'Fecha
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


