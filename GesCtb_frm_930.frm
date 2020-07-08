VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Frm_Ctb_FacEle_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   Icon            =   "GesCtb_frm_930.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10155
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   14445
      _Version        =   65536
      _ExtentX        =   25479
      _ExtentY        =   17912
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
         Left            =   45
         TabIndex        =   21
         Top             =   45
         Width           =   14355
         _Version        =   65536
         _ExtentX        =   25321
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
            TabIndex        =   22
            Top             =   180
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Documentos Electrónicos"
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
            Picture         =   "GesCtb_frm_930.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   45
         TabIndex        =   23
         Top             =   750
         Width           =   14355
         _Version        =   65536
         _ExtentX        =   25321
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
         Begin VB.CommandButton cmd_ImpExc 
            Height          =   585
            Left            =   1200
            Picture         =   "GesCtb_frm_930.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Importar archivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_930.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modificar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   12600
            Picture         =   "GesCtb_frm_930.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   13185
            Picture         =   "GesCtb_frm_930.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_930.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13755
            Picture         =   "GesCtb_frm_930.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2520
         Left            =   45
         TabIndex        =   24
         Top             =   6015
         Width           =   14355
         _Version        =   65536
         _ExtentX        =   25321
         _ExtentY        =   4445
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
         Begin VB.TextBox txt_Refer 
            Height          =   315
            Left            =   2130
            MaxLength       =   200
            TabIndex        =   7
            Top             =   2060
            Width           =   10200
         End
         Begin VB.ComboBox cmb_TipCom 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2910
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   2910
         End
         Begin VB.TextBox txt_Glosa 
            Height          =   315
            Left            =   2130
            MaxLength       =   200
            TabIndex        =   6
            Top             =   1725
            Width           =   10200
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   5325
         End
         Begin EditLib.fpDateTime ipp_FecEmi 
            Height          =   315
            Left            =   10110
            TabIndex        =   2
            Top             =   720
            Width           =   2250
            _Version        =   196608
            _ExtentX        =   3969
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
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   10110
            TabIndex        =   4
            Top             =   1050
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         Begin EditLib.fpDoubleSingle ipp_ValVta 
            Height          =   315
            Left            =   2130
            TabIndex        =   5
            Top             =   1395
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
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   270
            TabIndex        =   39
            Top             =   2115
            Width           =   825
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta:"
            Height          =   195
            Left            =   270
            TabIndex        =   35
            Top             =   1455
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Comprobante:"
            Height          =   195
            Left            =   270
            TabIndex        =   34
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   270
            TabIndex        =   30
            Top             =   1110
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos del Documento"
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
            Left            =   270
            TabIndex        =   29
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión:"
            Height          =   195
            Left            =   8400
            TabIndex        =   28
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            Height          =   195
            Left            =   270
            TabIndex        =   27
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   8400
            TabIndex        =   26
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Proceso:"
            Height          =   195
            Left            =   270
            TabIndex        =   25
            Top             =   780
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1515
         Left            =   45
         TabIndex        =   31
         Top             =   8565
         Width           =   14355
         _Version        =   65536
         _ExtentX        =   25321
         _ExtentY        =   2672
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   10110
            MaxLength       =   11
            TabIndex        =   9
            Top             =   360
            Width           =   2235
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   5325
         End
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   2130
            TabIndex        =   10
            Top             =   705
            Width           =   11880
            _Version        =   65536
            _ExtentX        =   20955
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
         Begin Threed.SSPanel pnl_Direcc 
            Height          =   315
            Left            =   2130
            TabIndex        =   11
            Top             =   1080
            Width           =   5280
            _Version        =   65536
            _ExtentX        =   9313
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
         Begin Threed.SSPanel pnl_Correo 
            Height          =   315
            Left            =   9120
            TabIndex        =   12
            Top             =   1080
            Width           =   4890
            _Version        =   65536
            _ExtentX        =   8625
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
         Begin Threed.SSPanel pnl_Distri 
            Height          =   315
            Left            =   12720
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
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
         Begin Threed.SSPanel pnl_Provin 
            Height          =   315
            Left            =   13200
            TabIndex        =   50
            Top             =   120
            Visible         =   0   'False
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
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
         Begin Threed.SSPanel pnl_Depart 
            Height          =   315
            Left            =   13680
            TabIndex        =   51
            Top             =   120
            Visible         =   0   'False
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correo:"
            Height          =   195
            Left            =   8400
            TabIndex        =   41
            Top             =   1140
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   270
            TabIndex        =   40
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl_NumDoc 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Documento:"
            Height          =   195
            Left            =   8400
            TabIndex        =   38
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lbl_RazSoc 
            AutoSize        =   -1  'True
            Caption         =   "Receptor:"
            Height          =   195
            Left            =   270
            TabIndex        =   37
            Top             =   795
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Datos del Receptor"
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
            Left            =   270
            TabIndex        =   33
            Top             =   90
            Width           =   1665
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   420
            Width           =   1230
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   4550
         Left            =   45
         TabIndex        =   36
         Top             =   1425
         Width           =   14355
         _Version        =   65536
         _ExtentX        =   25321
         _ExtentY        =   8026
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
            Height          =   4125
            Left            =   30
            TabIndex        =   13
            Top             =   300
            Width           =   14265
            _ExtentX        =   25162
            _ExtentY        =   7276
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin Threed.SSPanel pnl_TipCom 
            Height          =   285
            Left            =   30
            TabIndex        =   42
            Top             =   0
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Comp."
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
         Begin Threed.SSPanel pnl_TipPro 
            Height          =   285
            Left            =   690
            TabIndex        =   43
            Top             =   0
            Width           =   3075
            _Version        =   65536
            _ExtentX        =   5424
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Proceso"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   11370
            TabIndex        =   44
            Top             =   0
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel pnl_Receptor 
            Height          =   285
            Left            =   6150
            TabIndex        =   45
            Top             =   0
            Width           =   5295
            _Version        =   65536
            _ExtentX        =   9340
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Receptor"
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   285
            Left            =   4830
            TabIndex        =   46
            Top             =   0
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
         Begin Threed.SSPanel pnl_FecEmi 
            Height          =   285
            Left            =   3750
            TabIndex        =   47
            Top             =   0
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Emisión"
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   285
            Left            =   12450
            TabIndex        =   48
            Top             =   0
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
      End
   End
End
Attribute VB_Name = "Frm_Ctb_FacEle_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type g_tpo_CarDocEle
   CarDocEle_Col1       As String      'NumRef
   CarDocEle_ColErr1    As String
   CarDocEle_Col2       As String      'TipCom
   CarDocEle_ColErr2    As String
   CarDocEle_Col3       As String      'TipPro
   CarDocEle_ColErr3    As String
   CarDocEle_Col4       As String      'FecEmi
   CarDocEle_ColErr4    As String
   CarDocEle_Col5       As String      'Moneda
   CarDocEle_ColErr5    As String
   CarDocEle_Col6       As String      'TipCam
   CarDocEle_ColErr6    As String
   CarDocEle_Col7       As String      'TipDoc
   CarDocEle_ColErr7    As String
   CarDocEle_Col8       As String      'NumDoc
   CarDocEle_ColErr8    As String
   CarDocEle_Col9       As String      'Direcc
   CarDocEle_ColErr9    As String
   CarDocEle_Col10      As String      'Distri
   CarDocEle_ColErr10   As String
   CarDocEle_Col11      As String      'Provin
   CarDocEle_ColErr11   As String
   CarDocEle_Col12      As String      'Depart
   CarDocEle_ColErr12   As String
   CarDocEle_Col13      As String      'Correo
   CarDocEle_ColErr13   As String
   CarDocEle_Col14      As String      'Cantid
   CarDocEle_ColErr14   As String
   CarDocEle_Col15      As String      'Codigo
   CarDocEle_ColErr15   As String
   CarDocEle_Col16      As String      'UniMed
   CarDocEle_ColErr16   As String
   CarDocEle_Col17      As String      'Glosa
   CarDocEle_ColErr17   As String
   CarDocEle_Col18      As String      'ValUni
   CarDocEle_ColErr18   As String
   CarDocEle_Col19      As String      'VtaTot
   CarDocEle_ColErr19   As String
   CarDocEle_Col20      As String      'RazSoc
   CarDocEle_ColErr20   As String
   CarDocEle_Col21      As String      'Observación
   CarDocEle_ColErr21   As String
End Type

'Dim l_arr_LogPro()      As modprc_g_tpo_LogPro
'Dim l_bol_FlgErr        As Boolean
'Dim l_int_NumReg        As Integer

'Dim l_str_MsjRef        As String
'Dim l_int_RegTot        As Integer
'Dim l_int_RegPro        As Integer
'Dim l_int_RegErr        As Integer
'Dim l_int_RegSPr        As Integer

Dim l_str_RutaLg        As String
Dim l_fsobj             As Scripting.FileSystemObject
Dim l_txtStr            As TextStream

Dim l_lng_Codigo        As Long
'Dim l_int_InsUpd        As Integer

Private Sub cmb_TipCom_Click()
   If cmb_TipPro.ListIndex <> -1 Then
      cmb_TipPro_Click
   End If
End Sub

Private Sub cmb_TipPro_Click()
   txt_Glosa.Text = cmb_TipPro.Text
End Sub

Private Function fs_Obtener_Codigo(ByRef p_CodIte As Long)
Dim r_str_Parame           As String
Dim r_rst_Codigo           As ADODB.Recordset
   
   p_CodIte = 0
     
   'Código Máximo de CNTBL_DOCELE
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT MAX(DOCELETMP_CODIGO) AS CODIGO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_DOCELETMP "
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      r_rst_Codigo.Close
      Set r_rst_Codigo = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontro ningún registro en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
      If IsNull(r_rst_Codigo!CODIGO) Then
         p_CodIte = 0
      Else
         p_CodIte = r_rst_Codigo!CODIGO
      End If
   End If
   
   p_CodIte = p_CodIte + 1
   
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing

End Function
Private Sub fs_Escribir_Linea(p_ruta As String, p_texto As String)
   On Error GoTo MyError

   'Escribir en archivo según se ubique
   Set l_fsobj = New FileSystemObject
   Set l_txtStr = l_fsobj.OpenTextFile(p_ruta, ForAppending, False)
   l_txtStr.WriteLine (p_texto)
   l_txtStr.Close
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Escribir_Linea")
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1 'insert
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_TipCom)
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpiar
   Call fs_Activa(False)
   Call fs_Buscar
End Sub

Private Sub cmd_Editar_Click()
  
   If grd_Listad.Rows = 0 Then Exit Sub
  'CODIGO
   moddat_g_int_FlgGrb = 2 'update
  
   grd_Listad.Col = 8
   moddat_g_str_CodIte = CStr(grd_Listad.Text)
   
   'Buscando Información de DocEleTmp
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELETMP_IDE_TIPDOC, DOCELETMP_TIPPRO    , DOCELETMP_IDE_FECEMI, DOCELETMP_IDE_TIPMON, DOCELETMP_TIPCAM    , DOCELETMP_CAB_TOTVTA_OPEINA, "
   g_str_Parame = g_str_Parame & "        DOCELETMP_DET_DESPRD, DOCELETMP_REFER     , DOCELETMP_REC_TIPDOC, DOCELETMP_REC_NUMDOC, DOCELETMP_REC_DENOMI, DOCELETMP_REC_DIRCOM, "
   g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DISTRI, DOCELETMP_REC_PROVIN, DOCELETMP_REC_DEPART, DOCELETMP_REC_CORREC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELETMP "
   g_str_Parame = g_str_Parame & "  WHERE DOCELETMP_CODIGO = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not g_rst_Princi.EOF Then
      If Not IsNull((g_rst_Princi!DOCELETMP_IDE_TIPDOC)) Then
         Call gs_BuscarCombo_Item(cmb_TipCom, g_rst_Princi!DOCELETMP_IDE_TIPDOC)
      End If
      If Not IsNull((g_rst_Princi!DOCELETMP_TIPPRO)) Then
         cmb_TipPro.Text = moddat_gf_Consulta_ParDes("539", g_rst_Princi!DOCELETMP_TIPPRO)
      End If
      
      ipp_FecEmi.Text = Format(gf_FormatoFecha(Format(CStr(g_rst_Princi!DOCELETMP_IDE_FECEMI), "YYYYMMDD")), "dd/mm/yyyy")
      
      If Not IsNull(g_rst_Princi!DOCELETMP_IDE_TIPMON) Then
         cmb_Moneda.Text = fs_Obtener_Moneda(g_rst_Princi!DOCELETMP_IDE_TIPMON, 2)
      End If
      
      pnl_TipCam.Caption = Format(g_rst_Princi!DOCELETMP_TIPCAM, "###,###,###,##0.00") & "  "
      ipp_ValVta.Text = Format(g_rst_Princi!DOCELETMP_CAB_TOTVTA_OPEINA, "###,###,###,##0.00") & "  "
      
      If Not IsNull(g_rst_Princi!DOCELETMP_DET_DESPRD) Then
         txt_Glosa.Text = Trim(g_rst_Princi!DOCELETMP_DET_DESPRD)
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REFER) Then
         txt_Refer.Text = Trim(g_rst_Princi!DOCELETMP_REFER)
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_TIPDOC) Then
         cmb_TipDoc.Text = moddat_gf_Consulta_ParDes("118", g_rst_Princi!DOCELETMP_REC_TIPDOC)
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_NUMDOC) Then
         txt_NumDoc.Text = g_rst_Princi!DOCELETMP_REC_NUMDOC
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DENOMI) Then
         pnl_RazSoc.Caption = Trim(g_rst_Princi!DOCELETMP_REC_DENOMI) & "  "
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DISTRI) And Not IsNull(g_rst_Princi!DOCELETMP_REC_PROVIN) And Not IsNull(g_rst_Princi!DOCELETMP_REC_DEPART) Then
         pnl_Direcc.Caption = Trim(g_rst_Princi!DOCELETMP_REC_DISTRI) & " - " & Trim(g_rst_Princi!DOCELETMP_REC_PROVIN) & " - " & Trim(g_rst_Princi!DOCELETMP_REC_DEPART) & "  "
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_CORREC) Then
         pnl_Correo.Caption = Trim(g_rst_Princi!DOCELETMP_REC_CORREC) & "  "
      End If
      
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DISTRI) Then
         pnl_Distri.Caption = Trim(g_rst_Princi!DOCELETMP_REC_DISTRI)
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_PROVIN) Then
         pnl_Provin.Caption = Trim(g_rst_Princi!DOCELETMP_REC_PROVIN)
      End If
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DEPART) Then
         pnl_Depart.Caption = Trim(g_rst_Princi!DOCELETMP_REC_DEPART)
      End If
   End If
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_TipCom)
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_NumCom        As String
Dim r_int_TipPro        As Integer
Dim r_int_SerFac        As Integer
Dim r_str_NumSer        As String
Dim r_lng_NumFac        As Long
Dim r_str_NumFac        As String
Dim r_str_Mensaj        As String
Dim r_arr_Matriz()      As g_tpo_CarDocEle

   ReDim r_arr_Matriz(0)
   
   If cmb_TipCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Comprobante.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCom)
      Exit Sub
   End If
   If cmb_TipPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Proceso.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPro)
      Exit Sub
   End If
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Debe seleccionar Moneda.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   If ipp_ValVta.Value = 0 Then
      MsgBox "El valor venta no debe ser cero.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValVta)
      Exit Sub
   End If
   If txt_Glosa.Text = "" Then
      MsgBox "Debe ingresar glosa.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Glosa)
      Exit Sub
   End If
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo documento del Receptor.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
         
   If txt_NumDoc.Text = "" Then
      MsgBox "Debe ingresar Número de documento del Receptor.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex <> -1 And txt_NumDoc.Text <> "" Then
      Call fs_Validar_NumDoc(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
   End If
   
   If pnl_RazSoc.Caption = "" Then
      MsgBox "Debe especificar información del Receptor.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      r_str_Mensaj = "¿Seguro que desea ingresar el documento electrónico?"
   Else
      r_str_Mensaj = "¿Seguro que desea modificar el documento electrónico?"
   End If
   
   If MsgBox(r_str_Mensaj, vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col1 = 1                                              'l_lng_Codigo
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col2 = cmb_TipCom.ItemData(cmb_TipCom.ListIndex)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col3 = cmb_TipPro.ItemData(cmb_TipPro.ListIndex)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col4 = Format(ipp_FecEmi.Value, "yyyy-mm-dd")
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col5 = fs_Obtener_Moneda(Trim(cmb_Moneda.Text), 1)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col6 = pnl_TipCam.Caption
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col7 = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col8 = Trim(txt_NumDoc.Text)
'   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col9 = pnl_Direcc.Caption
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col10 = Trim(pnl_Distri.Caption)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col11 = Trim(pnl_Provin.Caption)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col12 = Trim(pnl_Depart.Caption)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col13 = Trim(pnl_Correo.Caption)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col14 = 1
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col15 = "--"
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col16 = "UND"
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col17 = Trim(txt_Glosa.Text)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col18 = ipp_ValVta.Value
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col19 = ipp_ValVta.Value
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col20 = Trim(pnl_RazSoc.Caption)
   r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col21 = Trim(txt_Refer.Text)
   
   If moddat_g_int_FlgGrb = 1 Then
      If fs_Agregar_DocEleTmp(r_arr_Matriz()) = True Then
         MsgBox "Ingreso realizado satisfactoriamente.", vbInformation, modgen_g_con_PltPar
      Else
         MsgBox "Ocurrió un error en fs_Agregar_DocEle.", vbInformation, modgen_g_con_PltPar
      End If
   Else
      If fs_Actualiza_DocEleTmp(r_arr_Matriz()) = True Then
         MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
      Else
         MsgBox "Ocurrió un error en fs_Actualiza_DocEleTmp.", vbInformation, modgen_g_con_PltPar
      End If

   End If
   
   Screen.MousePointer = 0
   Call fs_Inicia
   Call fs_Limpiar
   Call fs_Buscar
   Call fs_Activa(False)
End Sub

Private Function fs_Agregar_DocEleTmp(ByRef p_Array() As g_tpo_CarDocEle) As Boolean
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_str_NumSer     As String
Dim r_str_NumFac     As String
Dim r_lng_NumFac     As Long
Dim r_str_TipCom     As String

    fs_Agregar_DocEleTmp = False
          
    For r_lng_Contad = 1 To UBound(p_Array)
       
       Call fs_Obtener_Codigo(l_lng_Codigo)
       r_str_TipCom = IIf(Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") = 1, "F", "B")
       
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELETMP (      "
       g_str_Parame = g_str_Parame & " DOCELETMP_CODIGO                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_FECPRO                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_FECAPR                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_FECEMI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_HOREMI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_TIPDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_TIPMON             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_NUMORC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_IDE_FECVCT             , "
       
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_TIPDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_NUMDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_NOMCOM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DENOMI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_UBIGEO             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DIRCOM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_URBANI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_PROVIN             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DEPART             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DISTRI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_CODPAI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_TELEMI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_EMI_COREMI             , "
       
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_TIPDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_NUMDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_DENOMI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_DIRCOM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_DISTRI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_PROVIN             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_DEPART             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_CODPAI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_TELREC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REC_CORREC             , "
       
       g_str_Parame = g_str_Parame & " DOCELETMP_DRF_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DRF_TIPDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DRF_NUMDOC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DRF_CODMOT             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DRF_DESMOT             , "
       
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEGRV      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEGRV      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEINA      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEINA      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEEXO      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEEXO      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEGRA      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEGRA      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEEXP      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEEXP      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_PERCEP      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_REGPER      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_BASIMP_PERCEP      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOPER             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOTOT_PERCEP      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIMP             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOIMP             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_OTRCAR             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_TOTDSC      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTDSC             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_IMPTOT_DOCUME      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_DSCGLO             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_INFPPG             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTANT             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TIPOPE             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_CAB_LEYEND             , "
       
       g_str_Parame = g_str_Parame & " DOCELETMP_ADI_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_ADI_TITADI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_ADI_VALADI             , "
       
       'DETALLE
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_SERNUM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_NUMITE             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODPRD             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_DESPRD             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CANTID             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_UNIDAD             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALUNI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_PUNVTA             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIMP             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_MTOIMP             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_TIPAFE             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALVTA             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALREF             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_DSTITE             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_NUMPLA             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODSUN             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODCON             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_NROCON             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_FECOTO      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_FECOTO             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_TIPPRE      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_TIPPRE             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_PARREG      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_PARREG             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_PRIVIV      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_PRIVIV             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_DIRCOM      , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_DIRCOM             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODUBI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_UBIGEO             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODURB             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_URBANI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODDPT             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_DEPART             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODPRV             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_PROVIN             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODDIS             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_DET_DISTRI             , "
       g_str_Parame = g_str_Parame & " DOCELETMP_TIPCAM                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_SITUAC                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_TIPPRO                 , "
       g_str_Parame = g_str_Parame & " DOCELETMP_REFER                  , "
       g_str_Parame = g_str_Parame & " SEGUSUCRE                        , "
       g_str_Parame = g_str_Parame & " SEGFECCRE                        , "
       g_str_Parame = g_str_Parame & " SEGHORCRE                        , "
       g_str_Parame = g_str_Parame & " SEGPLTCRE                        , "
       g_str_Parame = g_str_Parame & " SEGTERCRE                        , "
       g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
       
       g_str_Parame = g_str_Parame & " VALUES ( "
       g_str_Parame = g_str_Parame & "" & l_lng_Codigo & "                                                   , "         'DOCELETMP_CODIGO
       g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "                      , "         'DOCELETMP_FECPRO
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_FECAPR
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'                                                 , "         'DOCELETMP_IDE_SERNUM ----  r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col4 & "'                         , "         'DOCELETMP_IDE_FECEMI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_HOREMI
       g_str_Parame = g_str_Parame & "'" & Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") & "'           , "         'DOCELETMP_IDE_TIPDOC
       g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col5 & "'                         , "         'DOCELETMP_IDE_TIPMON
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_NUMORC
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_FECVCT
       
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_EMI_SERNUM  ---- r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & "'6'                                                                    , "         'DOCELETMP_EMI_TIPDOC
       g_str_Parame = g_str_Parame & "'20511904162'                                                          , "         'DOCELETMP_EMI_NUMDOC
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_NOMCOM
       g_str_Parame = g_str_Parame & "'EDPYME MICASITA SA'                                                   , "         'DOCELETMP_EMI_DENOMI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_UBIGEO
       g_str_Parame = g_str_Parame & "'AV RIVERA NAVARRETE 645'                                              , "         'DOCELETMP_EMI_DIRCOM
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_URBANI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_PROVIN
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_DEPART
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_DISTRI
       g_str_Parame = g_str_Parame & "'PE'                                                                   , "         'DOCELETMP_EMI_CODPAI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_TELEMI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_COREMI
       
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_REC_SERNUM ----   r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col7 & "'                         , "         'DOCELETMP_REC_TIPDOC
       g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & "'                   , "         'DOCELETMP_REC_NUMDOC
       g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col20) & "'                  , "         'DOCELETMP_REC_DENOMI
       If Len(p_Array(r_lng_Contad).CarDocEle_Col9) = 0 Then
          g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DIRCOM
       Else
          g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col9) & "'                , "
       End If
       If Len(p_Array(r_lng_Contad).CarDocEle_Col10) = 0 Then
          g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DISTRI
       Else
          g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col10) & "'               , "
       End If
       If Len(p_Array(r_lng_Contad).CarDocEle_Col11) = 0 Then
          g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_PROVIN
       Else
          g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col11) & "'               , "
       End If
       If Len(p_Array(r_lng_Contad).CarDocEle_Col12) = 0 Then
          g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DEPART
       Else
          g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col12) & "'               , "
       End If
       g_str_Parame = g_str_Parame & "'PE'                                                                   , "         'DOCELETMP_REC_CODPAI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_REC_TELREC
       If Len(p_Array(r_lng_Contad).CarDocEle_Col13) = 0 Then
          g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_CORREC
       Else
          g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col13) & "'               , "
       End If
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_DRF_SERNUM  ---- r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_TIPDOC
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_NUMDOC
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_CODMOT
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_DESMOT
       
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "          'DOCELETMP_CAB_SERNUM  ---- & r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & "'1001'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEGRV
       g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEGRV
       g_str_Parame = g_str_Parame & "'1002'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEINA
       g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col18 & "                          , "         'DOCELETMP_CAB_TOTVTA_OPEINA
       g_str_Parame = g_str_Parame & "'1003'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEEXO
       g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEEXO
       g_str_Parame = g_str_Parame & "'1004'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEGRA
       g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEGRA
       g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEEXP
       g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEEXP
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_PERCEP
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_REGPER
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_BASIMP_PERCEP
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_MTOPER
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_MTOTOT_PERCEP
       g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_CAB_CODIMP
       g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_MTOIMP
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_OTRCAR
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_TOTDSC
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_TOTDSC
       g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col19 & "                          , "         'DOCELETMP_CAB_IMPTOT_DOCUME
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_DSCGLO
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_INFPPG
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_TOTANT
       g_str_Parame = g_str_Parame & "'13'                                                                   , "         'DOCELETMP_CAB_TIPOPE  -- Verificar CATALOGO 17 -- Gasto Deducible Persona Natural
       g_str_Parame = g_str_Parame & "'[1000'                                                                , "         'DOCELETMP_CAB_LEYEND
       
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "          'DOCELETMP_ADI_SERNUM ---- r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_ADI_TITADI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_ADI_VALADI
       
       'DETALLE
       g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "         'DOCELETMP_DET_SERNUM ---- r_str_NumSer & "-" & r_str_NumFac &
       g_str_Parame = g_str_Parame & "'001'                                                                  , "         'DOCELETMP_DET_NUMITE
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODPRD
       g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col17) & "'                  , "         'DOCELETMP_DET_DESPRD
       g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col14 & "                          , "         'DOCELETMP_DET_CANTID
       g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col16) & "'                  , "         'DOCELETMP_DET_UNIDAD
       g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col18) & "                    , "         'DOCELETMP_DET_VALUNI
       g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "                    , "         'DOCELETMP_DET_PUNVTA
       g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_DET_CODIMP
       g_str_Parame = g_str_Parame & "0.00                                                                   , "         'DOCELETMP_DET_MTOIMP
       g_str_Parame = g_str_Parame & "'30'                                                                   , "         'DOCELETMP_DET_TIPAFE -   'Inafecto operación onerosa
       g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "                    , "         'DOCELETMP_DET_VALVTA
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_VALREF
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DSTITE
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_NUMPLA
       
       g_str_Parame = g_str_Parame & "'84121501'                                                             , "         'DOCELETMP_DET_CODSUN
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODCON
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_NROCON
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_FECOTO
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_FECOTO
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_TIPPRE
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_TIPPRE
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_PARREG
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PARREG
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_PRIVIV
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PRIVIV
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_DIRCOM
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DIRCOM
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODUBI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_UBIGEO
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODURB
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_URBANI
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODDPT
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DEPART
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODPRV
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PROVIN
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODDIS
       g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DISTRI
       
       g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col6) & "                     , "         'DOCELETMP_TIPCAM
       g_str_Parame = g_str_Parame & "" & 2 & "                                                              , "         'DOCELETMP_SITUAC
       g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col3 & "                           , "         'DOCELETMP_TIPPRO
       g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col21 & "'                        , "         'DOCELETMP_REFER

       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                                          , "         'SEGUSUCRE
       g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "                , "         'SEGFECCRE
       g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                                        , "         'SEGHORCRE
       g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                                           , "         'SEGPLTCRE
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                                          , "         'SEGTERCRE
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                             
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
          Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELETMP, Nro Doc:" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & ", Nro. Cod: " & l_lng_Codigo & ", procedimiento: fs_Agregar_DocEleTmp")
          Exit Function
       End If
       DoEvents: DoEvents: DoEvents
         
       Set g_rst_Genera = Nothing
   Next
   fs_Agregar_DocEleTmp = True
End Function
Private Function fs_Actualiza_DocEleTmp(ByRef p_Array() As g_tpo_CarDocEle) As Boolean
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_str_NumSer     As String
Dim r_str_NumFac     As String
Dim r_lng_NumFac     As Long
Dim r_str_TipCom     As String

      fs_Actualiza_DocEleTmp = False
            
      For r_lng_Contad = 1 To UBound(p_Array)
         
         r_str_TipCom = IIf(Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") = 1, "F", "B")
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " UPDATE CNTBL_DOCELETMP SET     "
         g_str_Parame = g_str_Parame & "        DOCELETMP_IDE_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_IDE_FECEMI = '" & p_Array(r_lng_Contad).CarDocEle_Col4 & "'                , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_IDE_TIPDOC = '" & Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") & "'  , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_IDE_TIPMON = '" & p_Array(r_lng_Contad).CarDocEle_Col5 & "'                , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_EMI_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_TIPDOC = '" & p_Array(r_lng_Contad).CarDocEle_Col7 & "'                , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_NUMDOC = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & "'          , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DENOMI = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col20) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DIRCOM = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col9) & "'          , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DISTRI = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col10) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_PROVIN = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col11) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DEPART = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col12) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REC_CORREC = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col13) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DRF_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_CAB_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_CAB_TOTVTA_OPEINA = " & p_Array(r_lng_Contad).CarDocEle_Col18 & "          , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_CAB_IMPTOT_DOCUME = " & p_Array(r_lng_Contad).CarDocEle_Col19 & "          , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_ADI_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_SERNUM = '" & r_str_TipCom & "'                                        , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_DESPRD = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col17) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_CANTID = " & p_Array(r_lng_Contad).CarDocEle_Col14 & "                 , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_UNIDAD = '" & CStr(p_Array(r_lng_Contad).CarDocEle_Col16) & "'         , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_VALUNI = " & CDbl(p_Array(r_lng_Contad).CarDocEle_Col18) & "           , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_PUNVTA = " & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "           , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_DET_VALVTA = " & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "           , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_TIPCAM = " & CDbl(p_Array(r_lng_Contad).CarDocEle_Col6) & "                , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_TIPPRO = " & p_Array(r_lng_Contad).CarDocEle_Col3 & "                      , "
         g_str_Parame = g_str_Parame & "        DOCELETMP_REFER = '" & p_Array(r_lng_Contad).CarDocEle_Col21 & "'                    , "
         g_str_Parame = g_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "'                                            , "
         g_str_Parame = g_str_Parame & "        SEGFECACT = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "                   , "
         g_str_Parame = g_str_Parame & "        SEGHORACT = " & Format(Time, "HHMMSS") & "                                           , "
         g_str_Parame = g_str_Parame & "        SEGPLTACT = '" & UCase(App.EXEName) & "'                                             , "
         g_str_Parame = g_str_Parame & "        SEGTERACT = '" & modgen_g_str_NombPC & "'                                            , "
         g_str_Parame = g_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "'                                            "
         
         g_str_Parame = g_str_Parame & "  WHERE DOCELETMP_CODIGO = " & moddat_g_str_CodIte & "                                       "
         g_str_Parame = g_str_Parame & "    AND DOCELETMP_SITUAC = 2                                                                 "
                                        
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
           ' Call fs_Escribir_Linea("", "ERR   No se puede insertar en la tabla CNTBL_DOCELETMP, Nro Doc:" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & ", Nro. Cod: " & 0 & ", procedimiento: fs_Agregar_DocEleTmp")
            Exit Function
         End If
         DoEvents: DoEvents: DoEvents
         
       Set g_rst_Genera = Nothing
   Next
   fs_Actualiza_DocEleTmp = True
End Function
Private Function fs_Obtener_Moneda(ByVal p_Moneda As String, p_Opcion As Integer) As String
Dim r_str_Parame  As String

   fs_Obtener_Moneda = ""
   r_str_Parame = ""
   If p_Opcion = 1 Then
      r_str_Parame = r_str_Parame & " SELECT CATSUN_CODIGO "
   Else
      r_str_Parame = r_str_Parame & " SELECT CATSUN_DESCRI "
   End If
      r_str_Parame = r_str_Parame & "   FROM CNTBL_CATSUN "
   If p_Opcion = 1 Then
      r_str_Parame = r_str_Parame & "  WHERE CATSUN_DESCRI = '" & Trim(p_Moneda) & "' "
   Else
      r_str_Parame = r_str_Parame & "  WHERE CATSUN_CODIGO = '" & Trim(p_Moneda) & "' "
   End If
      r_str_Parame = r_str_Parame & "    AND CATSUN_NROCAT = 2 "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
       g_rst_Genera.Close
       Set g_rst_Genera = Nothing
       Exit Function
   Else
      If p_Opcion = 1 Then
         fs_Obtener_Moneda = g_rst_Genera!CATSUN_CODIGO
      Else
         fs_Obtener_Moneda = g_rst_Genera!CATSUN_DESCRI
      End If
   End If
End Function

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)

   cmb_TipCom.ListIndex = -1
   cmb_TipPro.ListIndex = -1
   ipp_FecEmi.Text = moddat_g_str_FecSis
   cmb_Moneda.ListIndex = -1
   pnl_TipCam.Caption = "0.000000" & " "
   txt_Glosa.Text = ""
   txt_Refer.Text = ""
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   pnl_Direcc.Caption = Empty
   pnl_Distri.Caption = Empty
   pnl_Provin.Caption = Empty
   pnl_Depart.Caption = Empty
   pnl_Correo.Caption = Empty
   pnl_RazSoc.Caption = Empty
   
   ipp_ValVta.Text = "0.00"
   
End Sub
Private Function fs_Validar_NumDoc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Boolean
   fs_Validar_NumDoc = True
   
   If (p_TipDoc = 1) Then 'DNI - 8
      If Len(Trim(p_NumDoc)) <> 8 Then
         MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         fs_Validar_NumDoc = False
      End If
   ElseIf (p_TipDoc = 6) Then 'RUC - 11
      If Not gf_Valida_RUC(Trim(p_NumDoc), Mid(Trim(p_NumDoc), Len(Trim(p_NumDoc)), 1)) Then
         MsgBox "El Número de RUC no es válido " & p_NumDoc, vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         fs_Validar_NumDoc = False
      End If
   Else 'OTROS
      If Len(Trim(p_NumDoc)) = 0 Then
         MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         fs_Validar_NumDoc = False
      End If
   End If
End Function

Private Sub cmd_ImpExc_Click()
   Frm_Ctb_FacEle_04.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   Call fs_Buscar
   Call fs_Activa(False)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call fs_Cargar_MntPardes(cmb_TipDoc, "118", 2)  'Tipo de Documento
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPro, 1, "539")
   
   cmb_TipCom.Clear
'   cmb_TipCom.AddItem "- TODOS -"
'   cmb_TipCom.ItemData(cmb_TipCom.NewIndex) = 0
   cmb_TipCom.AddItem "FACTURAS"
   cmb_TipCom.ItemData(cmb_TipCom.NewIndex) = "01"
   cmb_TipCom.AddItem "BOLETAS"
   cmb_TipCom.ItemData(cmb_TipCom.NewIndex) = "03"
'   cmb_TipCom.AddItem "NOTAS DE CREDITO"
'   cmb_TipCom.ItemData(cmb_TipCom.NewIndex) = "07"
'   cmb_TipCom.AddItem "PREPAGOS"
'   cmb_TipCom.ItemData(cmb_TipCom.NewIndex) = "05"
   cmb_TipCom.ListIndex = -1
   
   'Datos del Credito
   grd_Listad.ColWidth(0) = 735     'Tipo Comprobante
   grd_Listad.ColWidth(1) = 3000    'Tipo Proceso
   grd_Listad.ColWidth(2) = 1080    'Fecha Emisión
   grd_Listad.ColWidth(3) = 1300    'Documento
   grd_Listad.ColWidth(4) = 5240    'Receptor
   grd_Listad.ColWidth(5) = 1125    'Moneda
   grd_Listad.ColWidth(6) = 1450    'Importe
   grd_Listad.ColWidth(7) = 0       'Fecha Emisión sin Formato
   grd_Listad.ColWidth(8) = 0       'Código
   grd_Listad.ColWidth(9) = 0       'Situac
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   
   grd_Listad.Rows = 0

End Sub

Private Sub fs_Activa(ByVal p_Estado As Boolean)
   cmb_TipCom.Enabled = p_Estado
   cmb_TipPro.Enabled = p_Estado
   ipp_FecEmi.Enabled = p_Estado
   cmb_Moneda.Enabled = p_Estado
   ipp_ValVta.Enabled = p_Estado
   cmb_TipDoc.Enabled = p_Estado
   txt_Glosa.Enabled = p_Estado
   cmd_Grabar.Enabled = p_Estado
   cmd_Cancel.Enabled = p_Estado
   txt_NumDoc.Enabled = p_Estado
   txt_Refer.Enabled = p_Estado
   
   cmd_Agrega.Enabled = Not p_Estado
   cmd_Editar.Enabled = Not p_Estado
   cmd_ImpExc.Enabled = Not p_Estado
End Sub
Public Sub fs_Buscar()
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando Información de DocEleTmp
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELETMP_IDE_TIPDOC, DOCELETMP_TIPPRO    , DOCELETMP_IDE_FECEMI       , DOCELETMP_REC_TIPDOC, DOCELETMP_REC_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        DOCELETMP_REC_DENOMI, DOCELETMP_IDE_TIPMON, DOCELETMP_CAB_TOTVTA_OPEINA, DOCELETMP_CODIGO    , DOCELETMP_SITUAC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELETMP "
   g_str_Parame = g_str_Parame & "  WHERE DOCELETMP_FECAPR IS NULL "
   g_str_Parame = g_str_Parame & "    AND DOCELETMP_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "  ORDER BY DOCELETMP_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Tipo de Comprobante
      grd_Listad.Col = 0
      If CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "01" Then
         grd_Listad.Text = "F"
      ElseIf CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "03" Then
         grd_Listad.Text = "B"
      ElseIf CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "07" Then
         grd_Listad.Text = "NC"
      End If
      
      'Tipo de Proceso
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("539", CStr(g_rst_Princi!DOCELETMP_TIPPRO))
      
      'Fecha de Emisión
      grd_Listad.Col = 2
      grd_Listad.Text = gf_FormatoFecha(Format(CStr(g_rst_Princi!DOCELETMP_IDE_FECEMI), "YYYYMMDD"))
      
      'Documento
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_TIPDOC) & "-" & CStr(g_rst_Princi!DOCELETMP_REC_NUMDOC)
            
      'Receptor
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_DENOMI)
      
      'Moneda
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_IDE_TIPMON)
     
     'Importe
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!DOCELETMP_CAB_TOTVTA_OPEINA, "###,###,###,##0.00")
      
      'Fecha Emisión (sin formato)
      grd_Listad.Col = 7
      If IsNull(g_rst_Princi!DOCELETMP_IDE_FECEMI) Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = Format(CStr(g_rst_Princi!DOCELETMP_IDE_FECEMI), "YYYYMMDD")
      End If
      
      'Código
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_CODIGO)
      
      'Estado
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
End Sub
Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub fs_Cargar_MntPardes(p_Combo As ComboBox, ByVal p_CodGrp As String, p_TipPer As Integer)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES A "
   g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   If p_TipPer = 1 Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000007') "
   Else
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000006','000007') "
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

Private Sub ipp_FecEmi_LostFocus()
   'TIPO CAMBIO SBS(2) - VENTA(1)
   pnl_TipCam.Caption = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecEmi.Text, "yyyymmdd"), 1)
   pnl_TipCam.Caption = Format(pnl_TipCam.Caption, "###,###,##0.000000") & " "
End Sub

Private Sub ipp_FecEmi_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ValVta)
   End If
End Sub

Private Sub ipp_ValVta_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Glosa)
   End If
End Sub
Private Sub txt_Glosa_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Refer)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       cmb_TipDoc_Click
   End If
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FecEmi)
   End If
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
Dim r_str_RazSoc        As String
Dim r_str_Direcc        As String
Dim r_str_Depart        As String
Dim r_str_Provin        As String
Dim r_str_Distri        As String
Dim r_str_Correo        As String

   If KeyAscii = 13 Then
   
      If cmb_TipDoc.ListIndex <> -1 And Len(Trim(txt_NumDoc.Text)) > 0 Then
         
         pnl_RazSoc.Caption = Empty
         pnl_Direcc.Caption = Empty
         pnl_Correo.Caption = Empty
         
         If fs_Validar_NumDoc(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text) = True Then
            
            Call fs_Buscar_Receptor(cmb_TipPro.ItemData(cmb_TipPro.ListIndex), cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, r_str_RazSoc, r_str_Direcc, r_str_Distri, r_str_Provin, r_str_Depart, r_str_Correo)
         
            pnl_RazSoc.Caption = Trim(r_str_RazSoc)
            If r_str_Depart <> "" And r_str_Depart <> "" And r_str_Depart <> "" Then
               pnl_Direcc.Caption = Trim(r_str_Depart) & " - " & Trim(r_str_Provin) & " - " & Trim(r_str_Distri)
               
               pnl_Distri.Caption = Trim(r_str_Distri)
               pnl_Provin.Caption = Trim(r_str_Provin)
               pnl_Depart.Caption = Trim(r_str_Depart)

            Else
               pnl_Direcc.Caption = Empty
            End If
            pnl_Correo.Caption = Trim(r_str_Correo)
            
         End If
         If cmb_TipDoc.ListIndex <> -1 Then
            Call gs_SetFocus(cmd_Grabar)
         End If
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub
Private Sub cmb_TipCom_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Function fs_Buscar_Receptor(ByVal p_TipPro As Integer, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByRef p_RazSoc As String, ByRef p_Direcc As String, ByRef p_Distri As String, ByRef p_Provin As String, ByRef p_Depart As String, ByRef p_Correo As String)
Dim r_int_TipPro     As Integer

   p_RazSoc = ""
   p_Direcc = ""
   p_Distri = ""
   p_Provin = ""
   p_Correo = ""
   p_Depart = ""
   
   g_str_Parame = ""
   
   If p_TipPro = 3 Or p_TipPro = 4 Or p_TipPro = 6 Then
      
      g_str_Parame = g_str_Parame & " SELECT MAEETE_TIPDOC, MAEETE_NUMDOC, MAEPRV_RAZSOC AS RECEPTOR, TRIM(MAEETE_DIRREP) AS DIRECCION, TRIM(C.PARDES_DESCRI) AS DEPARTAMENTO, "
      g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS PROVINCIA, TRIM(E.PARDES_DESCRI) AS DISTRITO, TRIM(MAEPRV_CORREO) AS CORREO "
      g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE A  "
      g_str_Parame = g_str_Parame & "        INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = MAEETE_TIPDOC AND B.MAEPRV_NUMDOC = MAEETE_NUMDOC "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 101 AND C.PARDES_CODITE = SUBSTR(A.MAEETE_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 101 AND D.PARDES_CODITE = SUBSTR(A.MAEETE_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES E ON E.PARDES_CODGRP = 101 AND E.PARDES_CODITE = A.MAEETE_UBIGEO "
      g_str_Parame = g_str_Parame & "  WHERE MAEETE_SITUAC = 1  "
      If p_TipDoc > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEETE_TIPDOC = " & p_TipDoc & "  "
      End If
      If Len(Trim(p_NumDoc)) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEETE_NUMDOC = '" & Trim(p_NumDoc) & "' "
      End If
      
   Else
      g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC AS RECEPTOR, TRIM(MAEPRV_DOMFIS) AS DIRECCION, TRIM(B.PARDES_DESCRI) AS DEPARTAMENTO, "
      g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) AS PROVINCIA, TRIM(D.PARDES_DESCRI) AS DISTRITO, TRIM(MAEPRV_CORREO) AS CORREO "
      g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 101 AND B.PARDES_CODITE = SUBSTR(A.MAEPRV_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 101 AND C.PARDES_CODITE = SUBSTR(A.MAEPRV_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 101 AND D.PARDES_CODITE = A.MAEPRV_UBIGEO "
      g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
      If p_TipDoc > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_TIPDOC = " & p_TipDoc & "  "
      End If
      If Len(Trim(p_NumDoc)) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & Trim(p_NumDoc) & "' "
      End If
      
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not g_rst_Princi.EOF Then
      If Not IsNull(g_rst_Princi!RECEPTOR) Then
         p_RazSoc = Trim(g_rst_Princi!RECEPTOR)
      End If
      If Not IsNull(g_rst_Princi!Direccion) Then
         p_Direcc = Trim(g_rst_Princi!Direccion)
      End If
      If Not IsNull(g_rst_Princi!DEPARTAMENTO) Then
         p_Depart = Trim(g_rst_Princi!DEPARTAMENTO)
      End If
      If Not IsNull(g_rst_Princi!PROVINCIA) Then
         p_Provin = Trim(g_rst_Princi!PROVINCIA)
      End If
      If Not IsNull(g_rst_Princi!DISTRITO) Then
         p_Distri = Trim(g_rst_Princi!DISTRITO)
      End If
      If Not IsNull(g_rst_Princi!CORREO) Then
         p_Correo = Trim(g_rst_Princi!CORREO)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Function

Private Sub txt_Refer_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipDoc)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub
