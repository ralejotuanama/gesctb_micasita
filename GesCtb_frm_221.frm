VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_TarCre_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15165
   Icon            =   "GesCtb_frm_221.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8595
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15165
      _Version        =   65536
      _ExtentX        =   26749
      _ExtentY        =   15161
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
         TabIndex        =   17
         Top             =   60
         Width           =   15040
         _Version        =   65536
         _ExtentX        =   26529
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   18
            Top             =   60
            Width           =   5025
            _Version        =   65536
            _ExtentX        =   8864
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registros de Tarjeta Cr�dito Corporativo"
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
            Left            =   30
            Picture         =   "GesCtb_frm_221.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   19
         Top             =   760
         Width           =   15040
         _Version        =   65536
         _ExtentX        =   26529
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
         Begin VB.CommandButton cmd_Cheque 
            Height          =   585
            Left            =   4860
            Picture         =   "GesCtb_frm_221.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Impresi�n Cheque"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Detalle 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_221.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   4260
            Picture         =   "GesCtb_frm_221.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_221.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_221.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_221.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14430
            Picture         =   "GesCtb_frm_221.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_221.frx":1ACC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_221.frx":1DD6
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_221.frx":20E0
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Procesar 
            Height          =   585
            Left            =   5460
            Picture         =   "GesCtb_frm_221.frx":23EA
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Procesar Registros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   21
         Top             =   1440
         Width           =   15040
         _Version        =   65536
         _ExtentX        =   26529
         _ExtentY        =   1455
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6780
            TabIndex        =   2
            Top             =   420
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8160
            TabIndex        =   3
            Top             =   420
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
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   24
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   450
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Rendici�n:"
            Height          =   195
            Left            =   5460
            TabIndex        =   22
            Top             =   450
            Width           =   1260
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   630
         Left            =   60
         TabIndex        =   25
         Top             =   7875
         Width           =   15040
         _Version        =   65536
         _ExtentX        =   26529
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
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   2595
         End
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   15
            Top             =   180
            Width           =   4425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   27
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5535
         Left            =   60
         TabIndex        =   28
         Top             =   2295
         Width           =   15040
         _Version        =   65536
         _ExtentX        =   26529
         _ExtentY        =   9763
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   13440
            TabIndex        =   40
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2205
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "  Seleccionar"
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   1000
               TabIndex        =   41
               Top             =   20
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Proces 
            Height          =   285
            Left            =   12600
            TabIndex        =   29
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1517
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Procesado"
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   50
            TabIndex        =   30
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Rendido"
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
         Begin Threed.SSPanel pnl_Respon 
            Height          =   285
            Left            =   2805
            TabIndex        =   31
            Top             =   60
            Width           =   2820
            _Version        =   65536
            _ExtentX        =   4974
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Asignaci�n"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5115
            Left            =   30
            TabIndex        =   32
            Top             =   360
            Width           =   14985
            _ExtentX        =   26432
            _ExtentY        =   9022
            _Version        =   393216
            Rows            =   24
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   8385
            TabIndex        =   33
            Top             =   60
            Width           =   730
            _Version        =   65536
            _ExtentX        =   1288
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
         Begin Threed.SSPanel pnl_MtoAsig 
            Height          =   285
            Left            =   9105
            TabIndex        =   34
            Top             =   60
            Width           =   1180
            _Version        =   65536
            _ExtentX        =   2081
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Asignado"
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
         Begin Threed.SSPanel pnl_MtoRen 
            Height          =   285
            Left            =   10275
            TabIndex        =   35
            Top             =   60
            Width           =   1180
            _Version        =   65536
            _ExtentX        =   2081
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Gastado"
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
         Begin Threed.SSPanel pnl_FecEnt 
            Height          =   285
            Left            =   2010
            TabIndex        =   36
            Top             =   60
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo"
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
         Begin Threed.SSPanel pnl_Benefi 
            Height          =   285
            Left            =   5610
            TabIndex        =   37
            Top             =   60
            Width           =   2790
            _Version        =   65536
            _ExtentX        =   4921
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Responsable"
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
         Begin Threed.SSPanel pnl_ano 
            Height          =   285
            Left            =   1230
            TabIndex        =   38
            Top             =   60
            Width           =   800
            _Version        =   65536
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "A�o"
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
         Begin Threed.SSPanel pnl_Saldo 
            Height          =   285
            Left            =   11445
            TabIndex        =   39
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Saldo"
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
Attribute VB_Name = "frm_Ctb_TarCre_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 11) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 11) = "X"
             End If
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Cheque_Click()
Dim r_str_CadAux   As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If grd_Listad.TextMatrix(grd_Listad.Row, 10) = "SI" Then
      'CHEQUE
      r_str_CadAux = ""
      
      frm_Ctb_PagCom_08.ipp_FecChq.Text = date
      frm_Ctb_PagCom_08.txt_NomDe.Text = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 5)) 'RESPONSABLE
      frm_Ctb_PagCom_08.pnl_Import.Caption = grd_Listad.TextMatrix(grd_Listad.Row, 8) & " " 'MONTO GASTADO
      
      frm_Ctb_PagCom_08.pnl_Moneda.Caption = grd_Listad.TextMatrix(grd_Listad.Row, 6) 'MONEDA
      frm_Ctb_PagCom_08.txt_CodOrigen.Text = "MODULO_TARJETA_CREDITO"
      frm_Ctb_PagCom_08.txt_CodOrigen.Tag = grd_Listad.TextMatrix(grd_Listad.Row, 0) 'CODIGO
      frm_Ctb_PagCom_08.fs_NumeroLetra
      frm_Ctb_PagCom_08.Show 1
   Else
      MsgBox "Solo se emite a registros que hayan sido procesados.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Procesar_Click()
Dim r_int_Contad   As Integer
Dim r_bol_Estado   As Boolean
Dim r_str_CajPrc   As String

   'PROCESADO
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_bol_Estado = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contad, 11) = "X" Then
             r_bol_Estado = True
             Exit For
          End If
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No se han seleccionados registros para procesar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_bol_Estado = True Then
      If MsgBox("�Seguro que desea procesar los registros seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      For r_int_Contad = 0 To grd_Listad.Rows - 1
          If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" Then
             If grd_Listad.TextMatrix(r_int_Contad, 11) = "X" Then
                'PROCESANDO REGISTROS
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_GEN ( "
                g_str_Parame = g_str_Parame & "'" & CLng(grd_Listad.TextMatrix(r_int_Contad, 0)) & "', "  'CAJDET_CODCAJ
                g_str_Parame = g_str_Parame & "6, " 'tarjeta de  credito
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "

                If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                   MsgBox "La rendici�n " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede procesar.", vbExclamation, modgen_g_str_NomPlt
                   Screen.MousePointer = 0
                   Exit Sub
                End If
                If (g_rst_Genera!RESUL = 1) Then 'insertado
                    r_str_CajPrc = r_str_CajPrc & "- " & grd_Listad.TextMatrix(r_int_Contad, 0)
                End If
                If (g_rst_Genera!RESUL = 2) Then 'saldo negativos
                    MsgBox "La rendici�n " & grd_Listad.TextMatrix(r_int_Contad, 0) & " tiene saldo negativo, no se proceso." & vbCrLf & _
                            "solo se procesaron:" & Trim(r_str_CajPrc), vbExclamation, modgen_g_str_NomPlt
                    Screen.MousePointer = 0
                    Exit Sub
                ElseIf (g_rst_Genera!RESUL = 3) Then 'no tiene detalle
                    MsgBox "La rendici�n " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no tiene detalle, no se proceso." & vbCrLf & _
                           "solo se procesaron:" & Trim(r_str_CajPrc), vbExclamation, modgen_g_str_NomPlt
                    Screen.MousePointer = 0
                    Exit Sub
                ElseIf (g_rst_Genera!RESUL = 4) Then 'moneda diferente
                    MsgBox "La rendici�n " & grd_Listad.TextMatrix(r_int_Contad, 0) & " tiene monedas distintas, no se proceso." & vbCrLf & _
                           "solo se procesaron:" & Trim(r_str_CajPrc), vbExclamation, modgen_g_str_NomPlt
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
             End If
          End If
      Next
        
      MsgBox "Se culmin� el proceso de registros seleccionados." & _
             vbCrLf & "Los registros procesados son: " & Trim(r_str_CajPrc), vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Call fs_BuscarCaja
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   cmd_Editar.Enabled = False
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   Call moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Buscar.ListIndex = 0
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "RESPONSABLE"
   cmb_Buscar.AddItem "PROCESADO"
   
   grd_Listad.ColWidth(0) = 1190 'Nro Caja
   grd_Listad.ColWidth(1) = 0 'Fecha caja
   grd_Listad.ColWidth(2) = 770 'a�o
   grd_Listad.ColWidth(3) = 790 'mes
   grd_Listad.ColWidth(4) = 2800 'Asignacion
   grd_Listad.ColWidth(5) = 2770 'Responsable
   grd_Listad.ColWidth(6) = 720  'Moneda
   grd_Listad.ColWidth(7) = 1170 'Mto Asignado
   grd_Listad.ColWidth(8) = 1150 'Mto Gastado
   grd_Listad.ColWidth(9) = 1170 'Mto Saldo
   grd_Listad.ColWidth(10) = 840 'procesado
   grd_Listad.ColWidth(11) = 1230 'selecconar
   grd_Listad.ColWidth(12) = 0 'Flag Proceso
   grd_Listad.ColWidth(13) = 0 'codigo moneda
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter 'nro rendido
   'grd_Listad.ColAlignment(1) = flexAlignCenterCenter 'fecha caja
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter 'a�o
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter 'mes
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter 'asignacion
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter 'responsable
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter 'moneda
   grd_Listad.ColAlignment(7) = flexAlignRightCenter 'mto asignado
   grd_Listad.ColAlignment(8) = flexAlignRightCenter 'mto gastado
   grd_Listad.ColAlignment(9) = flexAlignRightCenter 'mto saldo
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter 'procesado
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter 'seleccionado
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1 'insert
   frm_Ctb_TarCre_01.Show 1
End Sub

Private Sub cmd_Borrar_Click()
Dim r_rst_Genera     As ADODB.Recordset

   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 10
   If UCase(Trim(grd_Listad.Text)) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar el registro, se encuentra procesado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("�Seguro que desea eliminar el registro seleccionado?" & vbCrLf & _
             "Recuerde que debe eliminar el asiento contable manual.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & "6, " 'CAJCHC_TIPTAB
   g_str_Parame = g_str_Parame & "NULL, " 'CAJCHC_NUMERO
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminaci�n de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro de caja chica se elimino, recuerde que debe eliminar el asiento contable manual.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarCaja
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_BuscarCaja
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
End Sub

Private Sub cmd_Consul_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
      
   moddat_g_int_FlgGrb = 0 'consultar
   
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_TarCre_01.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Detalle_Click()
Dim r_int_Fila   As Integer

   moddat_g_str_Codigo = "" 'nro caja
   moddat_g_str_FecIng = "" 'ano y mes
   moddat_g_str_NomPrd = "" 'asignacion
   moddat_g_str_Descri = "" 'responsable
   moddat_g_str_DesMod = "" 'moneda
   moddat_g_dbl_MtoPre = 0 'importe
   moddat_g_int_Situac = 0
   moddat_g_str_CodMod = "" 'codigo moneda
   r_int_Fila = 0
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   r_int_Fila = grd_Listad.Row
   moddat_g_str_Codigo = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))  'nro caja
   moddat_g_str_FecIng = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 2)) & "-" & CStr(grd_Listad.TextMatrix(grd_Listad.Row, 3)) 'ano y mes
   moddat_g_str_NomPrd = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 4)) 'asignacion
   moddat_g_str_Descri = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 5)) 'responsable
   moddat_g_str_CodMod = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 13)) 'codigo moneda
   moddat_g_str_DesMod = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 6)) 'moneda
   moddat_g_dbl_MtoPre = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 7)) 'importe
   moddat_g_int_Situac = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 12)) 'Flag Proceso
   moddat_g_int_TipEva = 6 'tipo de tabla(tarjeta de credito)
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
   
   frm_Ctb_CajChc_03.Show 1
   Call fs_BuscarCaja
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_Fila   As Integer
    r_int_Fila = 0

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 12 'Flag Proceso
   If (grd_Listad.Text = 1) Then
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "No se pudo editar el registro, se encuentra procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   r_int_Fila = grd_Listad.Row
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
      
   moddat_g_int_FlgGrb = 2 'editar
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
   frm_Ctb_TarCre_01.Show 1
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Empres.Enabled = True
   cmb_Sucurs.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Public Sub fs_BuscarCaja()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String

   ReDim l_arr_CajChc(0)

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_PERANO, A.CAJCHC_PERMES,  "
   g_str_Parame = g_str_Parame & "         A.CAJCHC_CODMON, TRIM(C.PARDES_DESCRI) MONEDA, A.CAJCHC_IMPORT,  "
'   g_str_Parame = g_str_Parame & "         (NVL((SELECT SUM(NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0))  "
'   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X  "
'   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 6  "
'   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND CAJDET_TIPCPB NOT IN (7,88)),0) -  "
'   g_str_Parame = g_str_Parame & "         NVL((SELECT SUM(NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0))  "
'   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X  "
'   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 6  "
'   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND CAJDET_TIPCPB IN (7,88)),0)) MTOGASTADO,  "
   g_str_Parame = g_str_Parame & "        (NVL((SELECT SUM(CASE WHEN A.CAJCHC_CODMON = X.CAJDET_CODMON THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 1 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) / Y.TIPCAM_VENTAS "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 2 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) * Y.TIPCAM_VENTAS END) "
   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X "
   g_str_Parame = g_str_Parame & "                LEFT JOIN OPE_TIPCAM Y ON Y.TIPCAM_CODIGO = 3 AND Y.TIPCAM_TIPMON = 2 AND Y.TIPCAM_FECDIA = X.CAJDET_FECEMI "
   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 6 "
   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND X.CAJDET_TIPCPB NOT IN (7,88)),0) "
   g_str_Parame = g_str_Parame & "                  - "
   g_str_Parame = g_str_Parame & "         NVL((SELECT SUM(CASE WHEN A.CAJCHC_CODMON = X.CAJDET_CODMON THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 1 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) / Y.TIPCAM_VENTAS "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 2 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) * Y.TIPCAM_VENTAS END) "
   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X "
   g_str_Parame = g_str_Parame & "                LEFT JOIN OPE_TIPCAM Y ON Y.TIPCAM_CODIGO = 3 AND Y.TIPCAM_TIPMON = 2 AND Y.TIPCAM_FECDIA = X.CAJDET_FECEMI "
   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 6 "
   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND X.CAJDET_TIPCPB IN (7,88)),0)) "
   g_str_Parame = g_str_Parame & "        AS MTOGASTADO, "
   g_str_Parame = g_str_Parame & "         A.CAJCHC_TIPDOC || '-' || A.CAJCHC_NUMDOC NUMDOC_ASGINA, TRIM(B.MAEPRV_RAZSOC) NOM_ASIGNA,  "
   g_str_Parame = g_str_Parame & "         A.CAJCHC_RESPON,  "
   g_str_Parame = g_str_Parame & "         TRIM(EJECMC_APEPAT) ||' '|| TRIM(EJECMC_APEMAT) ||' '|| TRIM(EJECMC_NOMBRE) AS NOM_RESPONSABLE,  "
   g_str_Parame = g_str_Parame & "         A.CAJCHC_FLGPRC  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "   INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CAJCHC_CODMON  "
   g_str_Parame = g_str_Parame & "   INNER JOIN CRE_EJECMC D ON TRIM(D.EJECMC_CODEJE) = TRIM(A.CAJCHC_RESPON)  "
   g_str_Parame = g_str_Parame & "   WHERE A.CAJCHC_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "     AND A.CAJCHC_TIPTAB = 6  "
   g_str_Parame = g_str_Parame & "     AND TO_NUMBER(A.CAJCHC_PERANO || LPAD(A.CAJCHC_PERMES,2,'0')||'01') BETWEEN  " & r_str_FecIni & " AND " & r_str_FecFin
  
   If (cmb_Buscar.ListIndex = 1) Then 'responsable
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND UPPER(TRIM(EJECMC_APEPAT) ||' '|| TRIM(EJECMC_APEMAT) ||' '|| TRIM(EJECMC_NOMBRE)) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'procesado
       r_str_Cadena = ""
       Select Case UCase(Trim(txt_Buscar.Text))
              Case "S", "SI", "I": r_str_Cadena = "1"
              Case "N", "NO", "O": r_str_Cadena = "0"
       End Select
       If (Len(Trim(r_str_Cadena)) > 0) Then
           g_str_Parame = g_str_Parame & "   AND CAJCHC_FLGPRC = " & r_str_Cadena
       End If
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY CAJCHC_CODCAJ ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'No se ha encontrado ning�n registro.
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If

   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   ReDim l_arr_CajChc(0)

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!CajChc_CodCaj, "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)
      
      grd_Listad.Col = 2
      grd_Listad.Text = g_rst_Princi!CAJCHC_PERANO

      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!CAJCHC_PERMES, "00")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!NOM_ASIGNA & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!NOM_RESPONSABLE & "")
            
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")

      grd_Listad.Col = 7 'MTO ASIGNADO
      grd_Listad.Text = Format(g_rst_Princi!CajChc_Import, "###,###,###,##0.00")
      
      grd_Listad.Col = 8 'MTO GASTADO
      grd_Listad.Text = Format(g_rst_Princi!MTOGASTADO, "###,###,###,##0.00")
      
      grd_Listad.Col = 9 'MTO SALDO
      grd_Listad.Text = Format(g_rst_Princi!CajChc_Import - g_rst_Princi!MTOGASTADO, "###,###,###,##0.00")
      '------------------------------------------------------------------------------------------------
      grd_Listad.Col = 10
      grd_Listad.Text = IIf(g_rst_Princi!CAJCHC_FLGPRC = 1, "SI", "NO")
      
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!CAJCHC_FLGPRC
      
      grd_Listad.Col = 13
      grd_Listad.Text = g_rst_Princi!CAJCHC_CODMON
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contar        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE TARJETA DE CREDITO"
      .Range(.Cells(2, 2), .Cells(2, 11)).Merge
      .Range(.Cells(2, 2), .Cells(2, 11)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 11)).HorizontalAlignment = xlHAlignCenter

      r_int_NumFil = 4
      .Cells(r_int_NumFil, 2) = "NRO RENDIDO"
      .Cells(r_int_NumFil, 3) = "A�O"
      .Cells(r_int_NumFil, 4) = "PERIODO"
      .Cells(r_int_NumFil, 5) = "ASIGNACION"
      .Cells(r_int_NumFil, 6) = "RESPONSABLE"
      .Cells(r_int_NumFil, 7) = "MONEDA"
      .Cells(r_int_NumFil, 8) = "MONTO ASIGNADO"
      .Cells(r_int_NumFil, 9) = "MONTO GASTADO"
      .Cells(r_int_NumFil, 10) = "MONTO SALDO"
      .Cells(r_int_NumFil, 11) = "PROCESADO"
         
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 11)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 14 'NRO RENDIDO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 7 'A�O
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10 'PERIODO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40 'ASIGNACION
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 40 'RESPONSABLE
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 9 'MONEDA
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 19 'MONTO ASIGNADO
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 18 'MONTO GASTADO
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 16 'MONTO SALDO
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 13 'PROCESADO
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(10, 11)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 11)).Font.Size = 11
      
      r_int_NumFil = 5
      For r_int_Contar = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(r_int_Contar, 0)
          .Cells(r_int_NumFil, 3) = "'" & grd_Listad.TextMatrix(r_int_Contar, 2)
          .Cells(r_int_NumFil, 4) = "'" & grd_Listad.TextMatrix(r_int_Contar, 3)
          .Cells(r_int_NumFil, 5) = "'" & grd_Listad.TextMatrix(r_int_Contar, 4)
          .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(r_int_Contar, 5)
          .Cells(r_int_NumFil, 7) = "'" & grd_Listad.TextMatrix(r_int_Contar, 6)
          .Cells(r_int_NumFil, 8) = "'" & grd_Listad.TextMatrix(r_int_Contar, 7)
          .Cells(r_int_NumFil, 9) = "'" & grd_Listad.TextMatrix(r_int_Contar, 8)
          .Cells(r_int_NumFil, 10) = "'" & grd_Listad.TextMatrix(r_int_Contar, 9)
          .Cells(r_int_NumFil, 11) = "'" & grd_Listad.TextMatrix(r_int_Contar, 10)
          
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(4, 3), .Cells(4, 11)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 10
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 11
         If grd_Listad.Text = "X" Then
             grd_Listad.Text = ""
         Else
              grd_Listad.Text = "X"
         End If
      End If
      Call gs_RefrescaGrid(grd_Listad)
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

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarCaja
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?�)(/&%$�#@_.,;:")
   End If
End Sub

Private Sub cmb_Buscar_Click()
    If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
        txt_Buscar.Enabled = False
        Call gs_SetFocus(cmd_Buscar)
    Else
        txt_Buscar.Enabled = True
        Call gs_SetFocus(txt_Buscar)
    End If
    txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      moddat_g_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
      moddat_g_str_RazSoc = cmb_Empres.Text
      
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
      cmb_Sucurs.ListIndex = 0
      Call gs_SetFocus(cmb_Sucurs)
      Screen.MousePointer = 0
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub
