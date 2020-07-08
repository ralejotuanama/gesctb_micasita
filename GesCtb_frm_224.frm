VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_GesPer_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16005
   Icon            =   "GesCtb_frm_224.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   16095
      _Version        =   65536
      _ExtentX        =   28390
      _ExtentY        =   14746
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
         TabIndex        =   11
         Top             =   60
         Width           =   15885
         _Version        =   65536
         _ExtentX        =   28019
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
            TabIndex        =   12
            Top             =   150
            Width           =   4125
            _Version        =   65536
            _ExtentX        =   7276
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registros de vacaciones"
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
            Picture         =   "GesCtb_frm_224.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   650
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   15885
         _Version        =   65536
         _ExtentX        =   28019
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_224.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_224.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   645
         End
         Begin VB.CommandButton cmd_Consultar 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_224.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpiar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_224.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_224.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_224.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_224.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   15270
            Picture         =   "GesCtb_frm_224.frx":185C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   585
         Left            =   60
         TabIndex        =   15
         Top             =   2160
         Width           =   15885
         _Version        =   65536
         _ExtentX        =   28019
         _ExtentY        =   1032
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
         Begin Threed.SSPanel pnl_Trabajador 
            Height          =   315
            Left            =   1155
            TabIndex        =   25
            Top             =   150
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
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
         Begin Threed.SSPanel pnl_FechIng 
            Height          =   315
            Left            =   12840
            TabIndex        =   26
            Top             =   150
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin Threed.SSPanel pnl_CodPla 
            Height          =   315
            Left            =   9990
            TabIndex        =   31
            Top             =   150
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código Planilla:"
            Height          =   195
            Left            =   8790
            TabIndex        =   32
            Top             =   210
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso:"
            Height          =   195
            Left            =   11640
            TabIndex        =   28
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   210
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   5490
         Left            =   60
         TabIndex        =   16
         Top             =   2790
         Width           =   15885
         _Version        =   65536
         _ExtentX        =   28019
         _ExtentY        =   9684
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   285
            Left            =   1335
            TabIndex        =   17
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Operación"
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
         Begin MSFlexGridLib.MSFlexGrid grd_ListVac 
            Height          =   5100
            Left            =   30
            TabIndex        =   18
            Top             =   360
            Width           =   15840
            _ExtentX        =   27940
            _ExtentY        =   8996
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   285
            Left            =   2790
            TabIndex        =   19
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Operación"
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
         Begin Threed.SSPanel SSPanel25 
            Height          =   285
            Left            =   5520
            TabIndex        =   20
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Desde"
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
         Begin Threed.SSPanel SSPanel26 
            Height          =   285
            Left            =   6900
            TabIndex        =   21
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Hasta"
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
         Begin Threed.SSPanel SSPanel29 
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Interno"
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
         Begin Threed.SSPanel SSPanel30 
            Height          =   285
            Left            =   8280
            TabIndex        =   23
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Solicitados"
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
         Begin Threed.SSPanel SSPanel31 
            Height          =   285
            Left            =   9570
            TabIndex        =   24
            Top             =   60
            Width           =   4690
            _Version        =   65536
            _ExtentX        =   8273
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Comentario"
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
            Left            =   14250
            TabIndex        =   33
            Top             =   60
            Width           =   1280
            _Version        =   65536
            _ExtentX        =   2258
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Left            =   60
         TabIndex        =   29
         Top             =   1470
         Width           =   15885
         _Version        =   65536
         _ExtentX        =   28019
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   5700
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   210
            Width           =   2830
         End
         Begin EditLib.fpDateTime ipp_FecIniVac 
            Height          =   315
            Left            =   1155
            TabIndex        =   0
            Top             =   210
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
         Begin EditLib.fpDateTime ipp_FecFinVac 
            Height          =   315
            Left            =   2550
            TabIndex        =   1
            Top             =   210
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Situación:"
            Height          =   195
            Left            =   4650
            TabIndex        =   34
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Oper:"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   270
            Width           =   885
         End
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "MnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu smnu 
         Caption         =   "Exportar a Excel Listado"
         Index           =   0
      End
      Begin VB.Menu smnu 
         Caption         =   "Exportar a Excel Solicitud"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_Ctb_GesPer_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Borrar_Click()
Dim r_int_Fila   As Integer
Dim r_str_CodGrb As String
Dim r_bol_Estado As Boolean

   If grd_ListVac.Rows = 0 Then
      Exit Sub
   End If
                                 
   If CLng(grd_ListVac.TextMatrix(grd_ListVac.Row, 8)) <> 1 Then
      'Pendiente
      MsgBox "Solo se pueden eliminar los registro con situación pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_ListVac)
   If MsgBox("¿Seguro que desea eliminar lo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_GESPER_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & CLng(grd_ListVac.TextMatrix(grd_ListVac.Row, 0)) & "', " 'GESPER_CODGES
   g_str_Parame = g_str_Parame & "2, " 'TIPO TABLA
   g_str_Parame = g_str_Parame & "0, " 'ESTADO ELIMINAR
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   If g_rst_Genera!RESUL = 1 Then
      MsgBox "Registro eliminado correctamente.", vbInformation, modgen_g_str_NomPlt
      Call frm_Ctb_GesPer_01.fs_BuscarVac
      Call fs_BuscarReg
      Call gs_SetFocus(grd_ListVac)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_BuscarReg
   ipp_FecIniVac.Enabled = False
   ipp_FecFinVac.Enabled = False
   cmb_Situac.Enabled = False
End Sub

Private Sub cmd_Consultar_Click()
   If grd_ListVac.Rows = 0 Then
      Exit Sub
   End If
   
   grd_ListVac.Col = 0
   moddat_g_str_Codigo = CLng(grd_ListVac.Text)
      
   moddat_g_int_FlgGrb = 0 'consultar
   
   Call gs_RefrescaGrid(grd_ListVac)
   frm_Ctb_GesPer_02.Show 1
   
   Call gs_SetFocus(grd_ListVac)
End Sub

Private Sub cmd_Editar_Click()
   If grd_ListVac.Rows = 0 Then
      Exit Sub
   End If
   
   If CLng(grd_ListVac.TextMatrix(grd_ListVac.Row, 8)) <> 1 Then
      MsgBox "Solo se pueden editar registros con situación pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_ListVac.Col = 0
   moddat_g_str_Codigo = CLng(grd_ListVac.Text)
      
   moddat_g_int_FlgGrb = 2 'consultar
   
   Call gs_RefrescaGrid(grd_ListVac)
   frm_Ctb_GesPer_02.Show 1
   
   Call gs_SetFocus(grd_ListVac)
End Sub

Private Sub cmd_ExpExc_Click()
   Me.PopupMenu MnuPopUp
End Sub

Private Sub cmd_Limpiar_Click()
   cmb_Situac.ListIndex = 3
   Call gs_LimpiaGrid(grd_ListVac)
   
   ipp_FecIniVac.Enabled = True
   ipp_FecFinVac.Enabled = True
   cmb_Situac.Enabled = True
   Call gs_SetFocus(ipp_FecIniVac)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Trabajador.Caption = moddat_g_int_TipDoc & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_CodPla.Caption = Trim(moddat_g_str_CodGen)
   pnl_FechIng.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call cmd_Buscar_Click
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset

   Call moddat_gs_FecSis
   cmb_Situac.Clear
   cmb_Situac.AddItem "PENDIENTES"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 1
   cmb_Situac.AddItem "APROBADOS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 2
   cmb_Situac.AddItem "RECHAZADOS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 3
   cmb_Situac.AddItem "<<TODOS>>"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 0
   cmb_Situac.ListIndex = -1
   Call gs_BuscarCombo_Item(cmb_Situac, 0)
   
   grd_ListVac.ColWidth(0) = 1290 'CODIGO INTERNO
   grd_ListVac.ColWidth(1) = 1455 'FECHA OPERACION
   grd_ListVac.ColWidth(2) = 2720 'TIPO OPERACION
   grd_ListVac.ColWidth(3) = 1380 'FECHA DESDE
   grd_ListVac.ColWidth(4) = 1380 'FECHA HASTA
   grd_ListVac.ColWidth(5) = 1290 'DIAS SOLICITADOS
   grd_ListVac.ColWidth(6) = 4680 'COMENTARIO
   grd_ListVac.ColWidth(7) = 1260 'SITUACION
   grd_ListVac.ColWidth(8) = 0 'SITUACION
   grd_ListVac.ColWidth(9) = 0 'TIPO OPERACION
   
   grd_ListVac.ColAlignment(0) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(1) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(2) = flexAlignLeftCenter
   grd_ListVac.ColAlignment(3) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(4) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(5) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(6) = flexAlignLeftCenter
   grd_ListVac.ColAlignment(7) = flexAlignCenterCenter
   
   ipp_FecIniVac.Text = DateAdd("m", -5, moddat_g_str_FecSis)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT MIN(A.GESPER_FECOPE) AS FECHA_INICIO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPTAB = 2 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_NUMDOC = '" & moddat_g_str_NumDoc & "'"
   r_str_Parame = r_str_Parame & "    AND A.GESPER_SITUAC != 0 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      If Trim(r_rst_Genera!FECHA_INICIO & "") <> "" Then
         ipp_FecIniVac.Text = DateAdd("m", -2, gf_FormatoFecha(r_rst_Genera!FECHA_INICIO))
      End If
   End If
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
      
   ipp_FecFinVac.Text = DateAdd("m", 12, moddat_g_str_FecSis)
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ctb_GesPer_02.Show 1
End Sub

Public Sub fs_BuscarReg()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_ListVac)
   r_str_FecIni = Format(ipp_FecIniVac.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFinVac.Text, "yyyymmdd")
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.GESPER_CODGES, A.GESPER_FECOPE, A.GESPER_FECHA1, A.GESPER_FECHA2, A.GESPER_IMPORT, A.GESPER_TIPOPE, "
   r_str_Parame = r_str_Parame & "        TRIM(A.GESPER_DESCRI) AS GESPER_DESCRI, TRIM(B.PARDES_DESCRI) AS TIPO_OPERACION, A.GESPER_SITUAC "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 140 AND B.PARDES_CODITE = A.GESPER_TIPOPE "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPTAB = 2 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_NUMDOC = '" & moddat_g_str_NumDoc & "'"
   r_str_Parame = r_str_Parame & "    AND A.GESPER_SITUAC != 0 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_FECOPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   If cmb_Situac.ListIndex <> -1 Then
      If cmb_Situac.ItemData(cmb_Situac.ListIndex) <> 0 Then
         r_str_Parame = r_str_Parame & "  AND A.GESPER_SITUAC =" & cmb_Situac.ItemData(cmb_Situac.ListIndex)
      End If
   End If
   r_str_Parame = r_str_Parame & " ORDER BY A.GESPER_FECOPE ASC, A.GESPER_CODGES ASC "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_ListVac.Redraw = False
   r_rst_Princi.MoveFirst
   
   Do While Not r_rst_Princi.EOF
      grd_ListVac.Rows = grd_ListVac.Rows + 1
      grd_ListVac.Row = grd_ListVac.Rows - 1
      
      grd_ListVac.Col = 0
      grd_ListVac.Text = Format(r_rst_Princi!GESPER_CODGES, "0000000000")
      
      grd_ListVac.Col = 1
      grd_ListVac.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECOPE & "")
      
      grd_ListVac.Col = 2
      grd_ListVac.Text = Trim(r_rst_Princi!TIPO_OPERACION & "")
      
      grd_ListVac.Col = 3
      grd_ListVac.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA1 & "")
      
      grd_ListVac.Col = 4
      grd_ListVac.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA2 & "")
      
      grd_ListVac.Col = 5
      grd_ListVac.Text = r_rst_Princi!GESPER_IMPORT & " "

      grd_ListVac.Col = 6
      grd_ListVac.Text = Trim(r_rst_Princi!GESPER_DESCRI & "")
      
      grd_ListVac.Col = 7
      Select Case r_rst_Princi!GESPER_SITUAC
             Case 0: grd_ListVac.Text = "ELIMINADO"
             Case 1: grd_ListVac.Text = "PENDIENTE"
             Case 2: grd_ListVac.Text = "APROBADO"
             Case 3: grd_ListVac.Text = "RECHAZADO"
      End Select
                        
      grd_ListVac.Col = 8
      grd_ListVac.Text = r_rst_Princi!GESPER_SITUAC
      
      grd_ListVac.Col = 9
      grd_ListVac.Text = r_rst_Princi!GESPER_TIPOPE
 
      r_rst_Princi.MoveNext
   Loop
   
   grd_ListVac.Redraw = True
   Call gs_UbiIniGrid(grd_ListVac)
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
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
      .Cells(2, 2) = "REGISTRO DE VACACIONES"
      .Range(.Cells(2, 2), .Cells(2, 9)).Merge
      .Range(.Cells(2, 2), .Cells(2, 9)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 9)).HorizontalAlignment = xlHAlignCenter

      r_int_NumFil = 7
      .Cells(r_int_NumFil, 2) = "CÓDIGO INTERNO"
      .Cells(r_int_NumFil, 3) = "FECHA OPERACION"
      .Cells(r_int_NumFil, 4) = "TIPO OPERACION"
      .Cells(r_int_NumFil, 5) = "FECHA DESDE"
      .Cells(r_int_NumFil, 6) = "FECHA HASTA"
      .Cells(r_int_NumFil, 7) = "DIAS SOLICITADOS"
      .Cells(r_int_NumFil, 8) = "COMENTARIO"
      .Cells(r_int_NumFil, 9) = "SITUACION"
               
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 9)).Font.Bold = True
             
      .Cells(4, 2) = "FECHA:"
      .Cells(4, 3) = "'" & Trim(ipp_FecIniVac.Text) & "    Hasta    " & Trim(ipp_FecFinVac.Text)
      .Cells(4, 6) = "SITUACION:"
      .Cells(4, 7) = "'" & Trim(cmb_Situac.Text)
      .Cells(5, 2) = "CODIGO PLANILLA:"
      .Cells(5, 3) = "'" & Trim(pnl_CodPla.Caption)
      .Cells(5, 4) = "FECHA INGRESO:"
      .Cells(5, 5) = "'" & Trim(pnl_FechIng.Caption)
      .Cells(5, 6) = "TRABAJADOR:"
      .Cells(5, 7) = "'" & Trim(pnl_Trabajador.Caption)
      
      .Cells(4, 2).Font.Bold = True
      .Cells(4, 6).Font.Bold = True
      .Cells(5, 2).Font.Bold = True
      .Cells(5, 4).Font.Bold = True
      .Cells(5, 6).Font.Bold = True
         
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 17 'CÓDIGO INTERNO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 17 'FECHA OPERACION
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 16 'TIPO OPERACION
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 12 'FECHA DESDE
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 13 'FECHA HASTA
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 17 'DIAS SOLICITADOS
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 50 'COMENTARIO
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 14 'SITUACION
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      
      .Range(.Cells(1, 1), .Cells(9, 9)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(9, 9)).Font.Size = 11
      
      r_int_NumFil = 6
      For r_int_Contad = 0 To grd_ListVac.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 0) 'CÓDIGO INTERNO
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 1) 'FECHA OPERACION
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 2) 'TIPO OPERACION
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 3) 'FECHA DESDE
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 4) 'FECHA HASTA
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 5) 'DIAS SOLICITADO
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 6) 'COMENTARIO
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 7) 'SITUACION
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(7, 3), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(5, 9)).HorizontalAlignment = xlHAlignLeft
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Sol()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer
Dim r_int_ConTot        As Integer
Dim r_str_Parame        As String
Dim r_rst_Princi        As ADODB.Recordset
Dim r_rst_Genera        As ADODB.Recordset
Dim r_str_PerAux        As String
Dim r_str_FecIni        As String
Dim r_str_FecFin        As String
Dim r_fsobj             As Scripting.FileSystemObject
Dim r_str_rutarz        As String
Dim r_str_FecAux        As String

   Set r_fsobj = New FileSystemObject
   r_str_rutarz = moddat_g_str_RutLoc & "\TEMP"
   If r_fsobj.FolderExists(r_str_rutarz) = False Then
      r_fsobj.CreateFolder (r_str_rutarz)
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT SUBSTR(A.GESPER_FECHA2,1,4) AS PERIODO_FIN "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "   LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 140 AND B.PARDES_CODITE = A.GESPER_TIPOPE "
   r_str_Parame = r_str_Parame & "   LEFT JOIN CNTBL_GESPER C ON C.GESPER_CODGES = A.GESPER_CODGES AND C.GESPER_TIPTAB = 2 "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPTAB = 4 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND TRIM(A.GESPER_NUMDOC) = '" & moddat_g_str_NumDoc & "'"
   r_str_Parame = r_str_Parame & "    AND A.GESPER_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_CODGES = " & grd_ListVac.TextMatrix(grd_ListVac.Row, 0)
   r_str_Parame = r_str_Parame & " GROUP BY SUBSTR(A.GESPER_FECHA2,1,4)"
   r_str_Parame = r_str_Parame & " ORDER BY PERIODO_FIN ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         '----------------------------------------------
         Set r_obj_Excel = New Excel.Application
         r_obj_Excel.SheetsInNewWorkbook = 1
         r_obj_Excel.Workbooks.Add
         With r_obj_Excel.ActiveSheet
              r_int_Contad = r_int_Contad + 1
              .Columns(1).ColumnWidth = 7
              r_str_FecAux = ""
              .Range(.Cells(3, 2), .Cells(3, 7)).Merge
              .Cells(3, 2).HorizontalAlignment = xlHAlignCenter
              .Cells(3, 2).Font.Bold = True
              .Cells(3, 2) = "SOLICITUD DE FRACCIONAMIENTO DEL DESCANSO VACACIONAL"
            
              .Range(.Cells(6, 4), .Cells(6, 7)).Merge
              .Cells(6, 4).HorizontalAlignment = xlHAlignCenter
              .Cells(6, 4) = "SOLICITO: FRACCIONAR EL DESCANSO VACACIONAL"
            
              .Cells(8, 2).Font.Bold = True
              .Cells(8, 2) = "Señores"
              .Cells(10, 2).Font.Bold = True
              .Cells(10, 2) = "EDPYME MICASITA S.A."
            
              .Cells(13, 2) = "Presente.-"
              .Cells(15, 2) = "De mi mayor consideración:"
            
              .Range(.Cells(17, 2), .Cells(20, 7)).Merge
              .Cells(17, 2).HorizontalAlignment = 6
              .Cells(17, 2).VerticalAlignment = xlHAlignCenter
            
              .Cells(17, 2) = "Por medio de la presente, en virtud del artículo 3° del D. Legislativo N° 1405, "
              .Cells(17, 2) = .Cells(17, 2) & "solicito a usted autorizar el goce fraccionado de mi descanso  vacacional, "
              .Cells(17, 2) = .Cells(17, 2) & "correspondiente al periodo " & Mid(r_rst_Genera!PERIODO_FIN, 1, 4) & ", en los periodos y la forma "
              .Cells(17, 2) = .Cells(17, 2) & "que indico a continuación:"
                                                            
              r_str_Parame = ""
              r_str_Parame = r_str_Parame & "SELECT A.GESPER_CODGES, A.GESPER_FECHA1 AS PERIODO_INI, A.GESPER_FECHA2 AS PERIODO_FIN, A.GESPER_IMPORT, "
              r_str_Parame = r_str_Parame & "       TRIM(B.PARDES_DESCRI) AS TIPO_OPERACION, C.GESPER_FECHA1 AS FEC_GOCE_INI, C.GESPER_FECHA2 AS FEC_GOCE_FIN "
              r_str_Parame = r_str_Parame & "  FROM CNTBL_GESPER A "
              r_str_Parame = r_str_Parame & "  LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 140 AND B.PARDES_CODITE = A.GESPER_TIPOPE "
              r_str_Parame = r_str_Parame & "  LEFT JOIN CNTBL_GESPER C ON C.GESPER_CODGES = A.GESPER_CODGES AND C.GESPER_TIPTAB = 2 "
              r_str_Parame = r_str_Parame & " WHERE A.GESPER_TIPTAB = 4 "
              r_str_Parame = r_str_Parame & "   AND A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
              r_str_Parame = r_str_Parame & "   AND TRIM(A.GESPER_NUMDOC) = '" & moddat_g_str_NumDoc & "'"
              r_str_Parame = r_str_Parame & "   AND A.GESPER_SITUAC = 1 "
              r_str_Parame = r_str_Parame & "   AND A.GESPER_CODGES = " & grd_ListVac.TextMatrix(grd_ListVac.Row, 0)
              r_str_Parame = r_str_Parame & "   AND SUBSTR(A.GESPER_FECHA2,1,4) = " & Mid(r_rst_Genera!PERIODO_FIN, 1, 4)
              r_str_Parame = r_str_Parame & " ORDER BY A.GESPER_FECHA1, A.GESPER_NUMERO ASC "
               
              If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
                 Exit Sub
              End If
               
              r_int_Contad = 23
              If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
                 r_rst_Princi.MoveFirst
                 Do While Not r_rst_Princi.EOF
                    If r_str_FecIni = "" Then
                       r_str_FecIni = gf_FormatoFecha(r_rst_Princi!FEC_GOCE_INI)
                       r_str_FecFin = DateAdd("d", r_rst_Princi!GESPER_IMPORT - 1, r_str_FecIni)
                    Else
                       r_str_FecIni = DateAdd("d", 1, r_str_FecFin)
                       r_str_FecFin = DateAdd("d", r_rst_Princi!GESPER_IMPORT - 1, r_str_FecIni)
                    End If
                    .Cells(r_int_Contad, 2) = "* " & r_rst_Princi!GESPER_IMPORT & "   Dias naturales del   " & r_str_FecIni & "   Al   " & r_str_FecFin & ";"
                    r_int_Contad = r_int_Contad + 1
                    r_rst_Princi.MoveNext
                 Loop
              End If
              r_rst_Princi.Close
              Set r_rst_Princi = Nothing
      
              .Cells(29, 2) = "Sin otro particular, quedo de usted."
            
              .Range(.Cells(33, 2), .Cells(33, 3)).Merge
              .Cells(33, 2).HorizontalAlignment = 2
              .Cells(33, 2) = "Atentamente,"
                  
              .Range(.Cells(40, 2), .Cells(40, 4)).Merge
              .Range(.Cells(41, 2), .Cells(41, 4)).Merge
              .Range(.Cells(40, 2), .Cells(40, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
              .Cells(40, 2).HorizontalAlignment = 3
              .Cells(41, 2).HorizontalAlignment = 3
              .Cells(40, 2) = Trim(moddat_g_str_NomCli)
              
              Select Case moddat_g_int_TipDoc
                     Case 1: .Cells(41, 2) = "(DNI:" & moddat_g_str_NumDoc & " )"
                     Case 4: .Cells(41, 2) = "(CE:" & moddat_g_str_NumDoc & " )"
                     Case 6: .Cells(41, 2) = "(RUC:" & moddat_g_str_NumDoc & " )"
                     Case 7: .Cells(41, 2) = "(PAS:" & moddat_g_str_NumDoc & " )"
                     Case 9999: .Cells(41, 2) = "(OTROS:" & moddat_g_str_NumDoc & " )"
              End Select
            
              .Range(.Cells(44, 5), .Cells(44, 7)).Merge
              .Cells(44, 5).HorizontalAlignment = 4
              .Cells(44, 5) = Format(moddat_g_str_FecSis, "Li""m""a, DD ""de"" mmmm ""del"" yyyy")
                  
              .Range(.Cells(1, 1), .Cells(42, 7)).Font.Size = 12
         End With
         
         'g_str_RutLog & "\" & "SOLICITUD_FRACC_" & TRIM(moddat_g_str_NumDoc) & "_" & Mid(r_rst_Genera!PERIODO_FIN, 1, 4) & ".PDF", Quality:=xlQualityStandard, IncludeDocProperties:=True,
         r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
         r_str_rutarz & "\SOLICITUD_FRACC_" & Trim(moddat_g_str_NumDoc) & "_" & Mid(r_rst_Genera!PERIODO_FIN, 1, 4) & ".PDF", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
         IgnorePrintAreas:=False, OpenAfterPublish:=True
         r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False
                        
         r_obj_Excel.Visible = False
         Set r_obj_Excel = Nothing
         '----------------------------------------------
         'Kill (g_str_RutLog & "\" & "SOLICITUD_FRACC_" & TRIM(moddat_g_str_NumDoc) & "_" & Mid(r_rst_Genera!PERIODO_FIN, 1, 4) & ".PDF")
         
         r_rst_Genera.MoveNext
      Loop
   End If
 
   Screen.MousePointer = 0
End Sub

Private Sub smnu_Click(Index As Integer)
    Select Case Index
        Case 0:
               If grd_ListVac.Rows = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               End If
               
               Screen.MousePointer = 11
               Call fs_GenExc
               Screen.MousePointer = 0
        Case 1:
               If grd_ListVac.Rows = 0 Then
                  Exit Sub
               End If
                        
               If grd_ListVac.TextMatrix(grd_ListVac.Row, 9) = 2 Then
                  'TIPO DE OPERACION A CUENTA
                  If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     Exit Sub
                  End If
                  
                  Screen.MousePointer = 11
                  Call fs_GenExc_Sol
                  Screen.MousePointer = 0
               Else
                   MsgBox "Solo registros de tipo operación a cuenta (de 7 dias a mas).", vbExclamation, modgen_g_str_NomPlt
                   Exit Sub
               End If
    End Select
End Sub

