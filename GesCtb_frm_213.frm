VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_CarArc_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16440
   Icon            =   "GesCtb_frm_213.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16480
      _Version        =   65536
      _ExtentX        =   29069
      _ExtentY        =   13996
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
         TabIndex        =   1
         Top             =   60
         Width           =   16350
         _Version        =   65536
         _ExtentX        =   28840
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
            TabIndex        =   2
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Carga del Archivo de Recaudo"
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
            Picture         =   "GesCtb_frm_213.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   16350
         _Version        =   65536
         _ExtentX        =   28840
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
            Picture         =   "GesCtb_frm_213.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_213.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15720
            Picture         =   "GesCtb_frm_213.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_213.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_213.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_213.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_213.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   11
         Top             =   1470
         Width           =   16350
         _Version        =   65536
         _ExtentX        =   28840
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
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   420
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   90
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6780
            TabIndex        =   14
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
            TabIndex        =   15
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   6780
            TabIndex        =   16
            Top             =   90
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Proceso:"
            Height          =   195
            Left            =   5310
            TabIndex        =   20
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   450
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   120
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   17
            Top             =   120
            Width           =   1200
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5430
         Left            =   60
         TabIndex        =   21
         Top             =   2340
         Width           =   16350
         _Version        =   65536
         _ExtentX        =   28840
         _ExtentY        =   9578
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   10980
            TabIndex        =   32
            Top             =   300
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Top             =   300
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Proc."
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   2190
            TabIndex        =   23
            Top             =   300
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4785
            Left            =   30
            TabIndex        =   24
            Top             =   600
            Width           =   16320
            _ExtentX        =   28787
            _ExtentY        =   8440
            _Version        =   393216
            Rows            =   30
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   6420
            TabIndex        =   25
            Top             =   300
            Width           =   895
            _Version        =   65536
            _ExtentX        =   1579
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   300
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2046
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   13920
            TabIndex        =   27
            Top             =   300
            Width           =   1060
            _Version        =   65536
            _ExtentX        =   1870
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   14970
            TabIndex        =   28
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   3510
            TabIndex        =   29
            Top             =   300
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5151
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   7290
            TabIndex        =   30
            Top             =   300
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Banco"
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
            Left            =   9930
            TabIndex        =   31
            Top             =   300
            Width           =   1060
            _Version        =   65536
            _ExtentX        =   1870
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   11760
            TabIndex        =   33
            Top             =   300
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Filas"
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
            Left            =   12630
            TabIndex        =   34
            Top             =   300
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Pago Total"
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   9930
            TabIndex        =   35
            Top             =   30
            Width           =   4000
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Archivo Recaudo"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   13920
            TabIndex        =   36
            Top             =   30
            Width           =   2100
            _Version        =   65536
            _ExtentX        =   3704
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Columnas de Pago"
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
Attribute VB_Name = "frm_Ctb_CarArc_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()  As moddat_tpo_Genera
Dim l_arr_Sucurs()  As moddat_tpo_Genera

Private Sub cmd_Borrar_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   moddat_g_str_Codigo = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   
   '--procesado por Compensasion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL((SELECT COMAUT_CODEST FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "              Where A.COMAUT_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODEST IN (1,2,4,5)  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODOPE = " & CLng(moddat_g_str_Codigo) & ")  "
   g_str_Parame = g_str_Parame & "           ,0) AS CODEST  "
   g_str_Parame = g_str_Parame & "   FROM DUAL  "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   If g_rst_Princi!CODEST <> 0 Then
      Select Case g_rst_Princi!CODEST
             Case 1: MsgBox "El registro se encuentra como pendiente en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
             Case 2: MsgBox "El registro se encuentra como aprobado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
             Case 4: MsgBox "El registro se encuentra como aplicado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
             Case 5: MsgBox "El registro se encuentra como pagado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
      End Select
      Exit Sub
   End If
   '----------------------------------------
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_PROPAG_ARH_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'PROPAG_CODPAG
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_Buscar
   Call gs_SetFocus(grd_Listad)
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

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ctb_CarArc_02.Show 1
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_Buscar
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_int_FlgGrb = 0
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 0
   frm_Ctb_CarArc_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Empres.Enabled = True
   cmb_Sucurs.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   grd_Listad.ColWidth(0) = 1140  'CODIGO
   grd_Listad.ColWidth(1) = 990  'FECHA PROCESO
   grd_Listad.ColWidth(2) = 1380  'NRO-DOCUMENTO
   grd_Listad.ColWidth(3) = 2850  'PROVEEDOR
   grd_Listad.ColWidth(4) = 880   'MONEDA
   grd_Listad.ColWidth(5) = 2630  'BANCO
   grd_Listad.ColWidth(6) = 0   'NRO-CUENTA
   grd_Listad.ColWidth(7) = 1040   'FECHA ARCHIVO
   grd_Listad.ColWidth(8) = 790  'MONEDA ARCHIVO
   grd_Listad.ColWidth(9) = 890   'NRO-FILAS
   grd_Listad.ColWidth(10) = 1280 'PAGO TOTAL
   grd_Listad.ColWidth(11) = 1050 'FECHA PAGO
   grd_Listad.ColWidth(12) = 1040 'CODIGO PAGO
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter  'CODIGO
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter  'FECHA PROCESO
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter  'NRO-DOCUMENTO
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter    'PROVEEDOR
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter    'MONEDA
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter    'BANCO
   'grd_Listad.ColAlignment(6) = flexAlignLEFTCenter   'NRO-CUENTA
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter  'FECHA ARCHIVO
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter  'MONEDA ARCHIVO
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter   'NRO-FILAS
   grd_Listad.ColAlignment(10) = flexAlignRightCenter 'PAGO TOTAL
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter 'FECHA PAGO
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter 'CODIGO PAGO
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Public Sub fs_Buscar()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.PROPAG_CODPAG,A.PROPAG_TIPDOC, A.PROPAG_NUMDOC, TRIM(B.MAEPRV_RAZSOC) AS MAEPRV_RAZSOC,  "
   g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) AS NOM_MONEDA, TRIM(D.PARDES_DESCRI) AS NOM_BANCO, A.PROPAG_CTACRR,  "
   g_str_Parame = g_str_Parame & "        E.PAGCAB_FECREC, DECODE(E.PAGCAB_MONEDA,1,'PEN','USD') REC_MONEDA, E.PAGCAB_NUMREGFIL,  "
   g_str_Parame = g_str_Parame & "        E.PAGCAB_TOTPAGFIL, E.PAGCAB_FECPRO, A.PROPAG_FECREG, "
   g_str_Parame = g_str_Parame & "        G.COMPAG_FECPAG, G.COMPAG_CODCOM "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_PROPAG_ARH A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.PROPAG_TIPDOC AND B.MAEPRV_NUMDOC = A.PROPAG_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.PROPAG_CODMON  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 122 AND D.PARDES_CODITE = A.PROPAG_CODBCO  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PROPAGCAB E ON E.PAGCAB_FECPRO = A.PROPAG_FECPRO AND E.PAGCAB_NUMPRO = A.PROPAG_NUMPRO  "
   
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET F ON F.COMDET_CODOPE = A.PROPAG_CODPAG AND F.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG G ON G.COMPAG_CODCOM = F.COMDET_CODCOM AND G.COMPAG_SITUAC = 1 AND G.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "  WHERE A.PROPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND A.PROPAG_FECREG BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   g_str_Parame = g_str_Parame & "  ORDER BY A.PROPAG_CODPAG ASC  "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
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
      grd_Listad.Text = Format(g_rst_Princi!PROPAG_CODPAG, "0000000000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!PROPAG_FECREG)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!PROPAG_TIPDOC) & "-" & Trim(g_rst_Princi!PROPAG_NUMDOC)
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!NOM_MONEDA & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!NOM_BANCO & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!PROPAG_CTACRR & "")
      
      grd_Listad.Col = 7
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!PAGCAB_FECREC)
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!REC_MONEDA & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!PAGCAB_NUMREGFIL & "")
      
      grd_Listad.Col = 10
      grd_Listad.Text = Format(g_rst_Princi!PAGCAB_TOTPAGFIL, "###,###,###,##0.00")
                  
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 11
         grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_Listad.Col = 12
         grd_Listad.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
                  
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
Dim r_int_Contad        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE CARGAS DEL ARCHIVO DE RECAUDO"
      .Range(.Cells(2, 2), .Cells(2, 13)).Merge
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 13)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CODIGO"
      .Cells(3, 3) = "FECHA PROCESO"
      .Cells(3, 4) = "NRO DOCUMENTO"
      .Cells(3, 5) = "PROVEEDOR"
      .Cells(3, 6) = "MONEDA"
      .Cells(3, 7) = "BANCO"
      .Cells(3, 8) = "FECHA TXT"
      .Cells(3, 9) = "MONEDA TXT"
      .Cells(3, 10) = "NRO FILAS TXT"
      .Cells(3, 11) = "PAGO TOTAL TXT"
      .Cells(3, 12) = "FECHA PAGO"
      .Cells(3, 13) = "CODIGO PAGO"
         
      .Range(.Cells(3, 2), .Cells(3, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 13)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'CODIGO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15 'FECHA PROCESO
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18 'NRO DOCUMENTO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 41 'PROVEEDOR
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 10 'MONEDA
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 30 'BANCO
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 12 'FECHA ARCHIVO
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 13 'MONEDA ARCHIVO
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 13 'NRO FILAS ARCHIVO
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 17 'PAGO TOTAL ARCHIVO
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 12 'FECHA PAGO
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 14 'CODIGO PAGO
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_Contad, 0) 'CODIGO
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_Listad.TextMatrix(r_int_Contad, 1) 'FECHA PROCESO
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_Listad.TextMatrix(r_int_Contad, 2) 'NRO DOCUMENTO
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_Listad.TextMatrix(r_int_Contad, 3) 'PROVEEDOR
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_Listad.TextMatrix(r_int_Contad, 4) 'MONEDA
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_Listad.TextMatrix(r_int_Contad, 5) 'BANCO
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_Listad.TextMatrix(r_int_Contad, 7) 'FECHA ARCHIVO
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_Listad.TextMatrix(r_int_Contad, 8) 'MONEDA ARCHIVO
          .Cells(r_int_NumFil + 2, 10) = grd_Listad.TextMatrix(r_int_Contad, 9)  'NRO FILAS ARCHIVO
          .Cells(r_int_NumFil + 2, 11) = grd_Listad.TextMatrix(r_int_Contad, 10) 'PAGO TOTAL ARCHIVO
          .Cells(r_int_NumFil + 2, 12) = "'" & grd_Listad.TextMatrix(r_int_Contad, 11) 'FECHA PAGO
          .Cells(r_int_NumFil + 2, 13) = "'" & grd_Listad.TextMatrix(r_int_Contad, 12) 'CODIGO PAGO
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 13)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Call gs_SetFocus(cmb_Sucurs)
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
