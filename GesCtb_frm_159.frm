VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_AsiCtb_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9570
   ClientLeft      =   3165
   ClientTop       =   3615
   ClientWidth     =   16725
   Icon            =   "GesCtb_frm_159.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   16725
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   16725
      _Version        =   65536
      _ExtentX        =   29501
      _ExtentY        =   16960
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   6260
         Left            =   30
         TabIndex        =   16
         Top             =   2610
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
         _ExtentY        =   11042
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
         Begin Threed.SSPanel pnl_NroAsi 
            Height          =   285
            Left            =   1530
            TabIndex        =   17
            Top             =   60
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Asiento"
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
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   285
            Left            =   2460
            TabIndex        =   18
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Contable"
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
         Begin Threed.SSPanel pnl_Glosa 
            Height          =   285
            Left            =   3660
            TabIndex        =   19
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Glosa"
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
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   9090
            TabIndex        =   20
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe (MN)"
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
         Begin Threed.SSPanel pnl_HabMN 
            Height          =   285
            Left            =   10290
            TabIndex        =   21
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber (MN)"
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
         Begin Threed.SSPanel pnl_DebME 
            Height          =   285
            Left            =   12690
            TabIndex        =   22
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe (ME)"
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
         Begin Threed.SSPanel pnl_HabME 
            Height          =   285
            Left            =   13890
            TabIndex        =   23
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber (ME)"
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
            Height          =   5865
            Left            =   30
            TabIndex        =   24
            Top             =   360
            Width           =   16605
            _ExtentX        =   29289
            _ExtentY        =   10345
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
         Begin Threed.SSPanel pnl_DifMN 
            Height          =   285
            Left            =   11490
            TabIndex        =   33
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Difer. (MN)"
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
         Begin Threed.SSPanel pnl_DifME 
            Height          =   285
            Left            =   15090
            TabIndex        =   34
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Difer. (ME)"
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
         Begin Threed.SSPanel pnl_NroLib 
            Height          =   285
            Left            =   690
            TabIndex        =   37
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Libro"
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
         Begin Threed.SSPanel pnl_id 
            Height          =   285
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro."
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   25
         Top             =   60
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   26
            Top             =   60
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Asientos Contables"
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
            Picture         =   "GesCtb_frm_159.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   27
         Top             =   780
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_159.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_159.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_159.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_159.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16020
            Picture         =   "GesCtb_frm_159.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_159.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_159.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_159.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1095
         Left            =   30
         TabIndex        =   28
         Top             =   1470
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
         _ExtentY        =   1931
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   6285
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   390
            Width           =   6285
         End
         Begin VB.ComboBox cmb_LibCon 
            Height          =   315
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   6285
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   720
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
            Left            =   2910
            TabIndex        =   2
            Top             =   720
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
            Left            =   10200
            TabIndex        =   36
            Top             =   60
            Width           =   6285
            _Version        =   65536
            _ExtentX        =   11086
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
         Begin VB.Label lbl_NomEti 
            Caption         =   "Período Vigente:"
            Height          =   255
            Index           =   2
            Left            =   8730
            TabIndex        =   35
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   60
            TabIndex        =   31
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Libro Contable:"
            Height          =   255
            Index           =   1
            Left            =   8730
            TabIndex        =   30
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Asiento:"
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   720
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   630
         Left            =   30
         TabIndex        =   39
         Top             =   8910
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
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
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   4
            Top             =   180
            Width           =   4425
         End
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   40
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera
Dim r_str_Origen        As String
Dim l_var_ColAnt        As Variant

Private Sub cmb_Buscar_Click()
   If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
      txt_Buscar.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   Else
      txt_Buscar.Enabled = True
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

Private Sub cmb_Filtro_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Filtro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Filtro_Click
   End If
End Sub

Private Sub cmb_LibCon_GotFocus()
    cmb_LibCon.BackColor = modgen_g_con_ColAma
End Sub

Private Sub cmb_LibCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_LibCon_LostFocus()
   cmb_LibCon.BackColor = modgen_g_con_ColBla
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(cmb_LibCon)
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   frm_Ctb_AsiCtb_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_int_Situac     As Integer
   Dim r_str_NumAsi     As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 2
   r_str_NumAsi = grd_Listad.Text
   r_str_Origen = "LM"
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   modctb_int_PerAno = grd_Listad.TextMatrix(grd_Listad.Row, 11)
   modctb_int_PerMes = grd_Listad.TextMatrix(grd_Listad.Row, 12)
   grd_Listad.Redraw = False
    
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INGRESO_CNTBL_ASI_DET_1 ("
      g_str_Parame = g_str_Parame & "'" & r_str_Origen & "', "
      g_str_Parame = g_str_Parame & CStr(modctb_int_PerAno) & ", "
      g_str_Parame = g_str_Parame & CStr(modctb_int_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(r_str_NumAsi) & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      
      'Datos de Linea
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ","
      g_str_Parame = g_str_Parame & CInt(3) & ") "
      
'         'Datos de Auditoria
'         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
'         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
'         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
'         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INGRESO_CNTBL_ASI_DET_1. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INGRESO_CNTBL_ASIENTO_1 ("
      g_str_Parame = g_str_Parame & "'" & r_str_Origen & "', "
      g_str_Parame = g_str_Parame & CStr(modctb_int_PerAno) & ", "
      g_str_Parame = g_str_Parame & CStr(modctb_int_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(r_str_NumAsi) & ", "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & "'0', "
      g_str_Parame = g_str_Parame & CInt(3) & ") "
      
      'Datos de Auditoria
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
'      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INGRESO_CNTBL_ASIENTO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   grd_Listad.Redraw = True
  
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Consul_Click()
Dim r_int_LisInd As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'Nro de Libro
   grd_Listad.Col = 1
   modctb_int_CodLib = CLng(grd_Listad.Text)
   
   'Nombre de Libro
   If Me.cmb_LibCon.Text = "<<TODOS>>" Then
      r_int_LisInd = cmb_LibCon.ListIndex
      cmb_LibCon.ListIndex = modctb_int_CodLib - 1
      modctb_str_NomLib = Me.cmb_LibCon.Text
      cmb_LibCon.ListIndex = r_int_LisInd
   End If
   
   'Nro de Asiento
   grd_Listad.Col = 2
   modctb_lng_NumAsi = CLng(grd_Listad.Text)
   
   'grd_Listad.Col = 3
   'modctb_int_PerAno = Year(CDate(grd_Listad.Text))
   'modctb_int_PerMes = Month(CDate(grd_Listad.Text))
   modctb_int_PerAno = grd_Listad.TextMatrix(grd_Listad.Row, 11)
   modctb_int_PerMes = grd_Listad.TextMatrix(grd_Listad.Row, 12)
    
   Call gs_RefrescaGrid(grd_Listad)

   frm_Ctb_AsiCtb_03_1.Show 1
End Sub

Private Sub cmd_Editar_Click()
   Dim r_int_Situac     As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 2
   modctb_lng_NumAsi = CLng(grd_Listad.Text)

   Call gs_RefrescaGrid(grd_Listad)

   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   modctb_int_PerAno = grd_Listad.TextMatrix(grd_Listad.Row, 11)
   modctb_int_PerMes = grd_Listad.TextMatrix(grd_Listad.Row, 12)
   
   frm_Ctb_AsiCtb_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ExpExc_Click()
        
    If grd_Listad.Rows = 0 Then
        MsgBox "No existen datos a Exportar.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_LibCon)
        Exit Sub
    End If

    If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
     
    Screen.MousePointer = 11
    Call fs_GenExc
    Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NoFlLi        As Integer
Dim r_int_NoFlGr        As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
    
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "REPORTE DE ASIENTOS CONTABLES"
      .Range(.Cells(1, 1), .Cells(1, 11)).Merge
      .Range(.Cells(1, 1), .Cells(1, 11)).Font.Bold = True
      .Cells(2, 1) = "Del " & ipp_FecIni.Text & " Al " & ipp_FecFin.Text
      .Range(.Cells(2, 1), .Cells(2, 11)).Merge
      .Range(.Cells(2, 1), .Cells(2, 11)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 11)).HorizontalAlignment = xlCenter
      .Range(.Cells(2, 1), .Cells(2, 11)).HorizontalAlignment = xlCenter
      
      r_int_NoFlLi = 4
      
      .Cells(r_int_NoFlLi, 1) = "Nro."
      .Cells(r_int_NoFlLi, 2) = "Nro. Libro"
      .Cells(r_int_NoFlLi, 3) = "Nro. Asiento"
      .Cells(r_int_NoFlLi, 4) = "F. Contable"
      .Cells(r_int_NoFlLi, 5) = "Glosa"
      .Cells(r_int_NoFlLi, 6) = "Debe (MN)"
      .Cells(r_int_NoFlLi, 7) = "Haber (MN)"
      .Cells(r_int_NoFlLi, 8) = "Difer. (MN)"
      .Cells(r_int_NoFlLi, 9) = "Debe (ME)"
      .Cells(r_int_NoFlLi, 10) = "Haber (ME)"
      .Cells(r_int_NoFlLi, 11) = "Difer. (ME)"
      
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Font.Name = "Calibri"
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Font.Bold = True
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_NoFlLi, 1), .Cells(r_int_NoFlLi, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 10
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 12
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 60
      .Columns("F").ColumnWidth = 12
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 12
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 12
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 12
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
   
      r_int_NoFlLi = 5
         
      For r_int_NoFlGr = 0 To grd_Listad.Rows - 1
         For r_int_Contad = 0 To grd_Listad.Cols - 1
            If r_int_Contad = 3 Then
               .Cells(r_int_NoFlLi, r_int_Contad + 1) = CDate(grd_Listad.TextMatrix(r_int_NoFlGr, r_int_Contad))
            Else
               .Cells(r_int_NoFlLi, r_int_Contad + 1) = grd_Listad.TextMatrix(r_int_NoFlGr, r_int_Contad)
            End If
         Next r_int_Contad
         r_int_NoFlLi = r_int_NoFlLi + 1
      Next r_int_NoFlGr
         
   End With
   
   r_obj_Excel.Cells(5, 1).Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   r_obj_Excel.Visible = True
End Sub

Private Sub grd_Listad_DblClick()
   cmd_Consul_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecFin_InvalidData(NextWnd As Long)
   If CDate(ipp_FecFin.Text) < CDate(modctb_str_FecIni) Then
      ipp_FecFin.Text = modctb_str_FecIni
   ElseIf CDate(ipp_FecFin.Text) > CDate(modctb_str_FecFin) Then
      ipp_FecFin.Text = modctb_str_FecFin
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Call gs_SetFocus(cmd_Buscar)
      Call gs_SetFocus(cmb_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_InvalidData(NextWnd As Long)
   If CDate(ipp_FecIni.Text) < CDate(modctb_str_FecIni) Then
      ipp_FecIni.Text = modctb_str_FecIni
   ElseIf CDate(ipp_FecIni.Text) > CDate(modctb_str_FecFin) Then
      ipp_FecIni.Text = modctb_str_FecFin
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LibCon(cmb_LibCon) 'moddat_gs_Carga_LibCtb
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "NRO ASIENTO"
   cmb_Buscar.AddItem "GLOSA"
   
   cmb_Buscar.ListIndex = 0
'   txt_Buscar.Enabled = False
   Call cmb_Buscar_Click
   
   grd_Listad.ColWidth(0) = 635
   grd_Listad.ColWidth(1) = 835
   grd_Listad.ColWidth(2) = 945
   grd_Listad.ColWidth(3) = 1215
   grd_Listad.ColWidth(4) = 5420
   grd_Listad.ColWidth(5) = 1195
   grd_Listad.ColWidth(6) = 1195
   grd_Listad.ColWidth(7) = 1195
   grd_Listad.ColWidth(8) = 1195
   grd_Listad.ColWidth(9) = 1195
   grd_Listad.ColWidth(10) = 1195
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
     
   cmb_LibCon.AddItem "<<TODOS>>"
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0

   Call moddat_gs_FecSis
   cmb_Empres.ListIndex = 0
   
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   '---------------------------------------------------------
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   '---------------------------------------------------------
   
   ipp_FecIni.Text = Format(moddat_g_str_FecSis, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(moddat_g_str_FecSis, "dd/mm/yyyy")
   
   If moddat_g_str_CodIte = 1 Then
      ipp_FecIni.DateMin = Format(CDate(modctb_str_FecIni), "yyyymmdd")
      ipp_FecIni.DateMax = Format(CDate(modctb_str_FecFin), "yyyymmdd")
      
      ipp_FecFin.DateMin = Format(CDate(modctb_str_FecIni), "yyyymmdd")
      ipp_FecFin.DateMax = Format(CDate(modctb_str_FecFin), "yyyymmdd")
   End If
   cmb_Sucurs.ListIndex = 0
   cmb_LibCon.ListIndex = -1
   
'   cmb_Buscar.ListIndex = 0
   txt_Buscar.Text = ""
   txt_Buscar.Enabled = True
   cmb_Buscar.Enabled = True
   cmb_Buscar.ListIndex = 0
   Call cmb_Buscar_Click
   
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Empres.Enabled = p_Activa
   cmb_Sucurs.Enabled = p_Activa
   cmb_LibCon.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   'cmb_Buscar.Enabled = Not p_Activa
   'txt_Buscar.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_Consul.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
  
   grd_Listad.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   If moddat_g_str_CodIte = 1 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_Consul.Enabled = False
      cmd_ExpExc.Enabled = False
      grd_Listad.Enabled = False
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NRO_LIBRO, NRO_ASIENTO, FECHA_CNTBL, DESC_GLOSA, TOT_SOLDEB, "
   g_str_Parame = g_str_Parame & "        TOT_SOLHAB, TOT_DOLDEB, TOT_DOLHAB, ANO, MES "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_ASIENTO "
   g_str_Parame = g_str_Parame & "  WHERE FECHA_CNTBL >= TO_DATE ('" & Me.ipp_FecIni.Text & "','dd/mm/yyyy') "
   g_str_Parame = g_str_Parame & "    AND FECHA_CNTBL <= TO_DATE ('" & Me.ipp_FecFin.Text & "', 'dd/mm/yyyy') "
   If cmb_Buscar.ListIndex > -1 Then
      If cmb_Buscar.ListIndex = 1 Then
         g_str_Parame = g_str_Parame & "  AND NRO_ASIENTO LIKE '%" & txt_Buscar.Text & "%'"
      Else
         g_str_Parame = g_str_Parame & "  AND UPPER(TRIM(DESC_GLOSA)) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
      End If
   End If
   If Me.cmb_LibCon.Text <> "<<TODOS>>" Then
      g_str_Parame = g_str_Parame & " AND NRO_LIBRO = '" & Me.cmb_LibCon.ItemData(cmb_LibCon.ListIndex) & "' "
      g_str_Parame = g_str_Parame & "  ORDER BY NRO_ASIENTO ASC "
   Else
      g_str_Parame = g_str_Parame & "  ORDER BY NRO_LIBRO, NRO_ASIENTO ASC "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = grd_Listad.Row + 1
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!NRO_LIBRO)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!NRO_ASIENTO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!FECHA_CNTBL
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!DESC_GLOSA & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_SOLDEB), 0, g_rst_Princi!TOT_SOLDEB), "###,###,##0.00")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_SOLHAB), 0, g_rst_Princi!TOT_SOLHAB), "###,###,##0.00")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_SOLDEB), 0, g_rst_Princi!TOT_SOLDEB) - IIf(IsNull(g_rst_Princi!TOT_SOLHAB), 0, g_rst_Princi!TOT_SOLHAB), "###,###,##0.00")
      
      grd_Listad.Col = 8
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_DOLDEB), 0, g_rst_Princi!TOT_DOLDEB), "###,###,##0.00")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_DOLHAB), 0, g_rst_Princi!TOT_DOLHAB), "###,###,##0.00")
      
      grd_Listad.Col = 10
      grd_Listad.Text = Format(IIf(IsNull(g_rst_Princi!TOT_DOLDEB), 0, g_rst_Princi!TOT_DOLDEB) - IIf(IsNull(g_rst_Princi!TOT_DOLHAB), 0, g_rst_Princi!TOT_DOLHAB), "###,###,##0.00")
      
      grd_Listad.Col = 11
      grd_Listad.Text = g_rst_Princi!ANO
      
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!Mes
     
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   If moddat_g_str_CodIte = 1 Then
      If grd_Listad.Rows > 0 Then
         cmd_Editar.Enabled = True
         cmd_Borrar.Enabled = True
         cmd_Consul.Enabled = True
         cmd_ExpExc.Enabled = True
         grd_Listad.Enabled = True
      End If
   End If
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_LibCon)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   If cmb_Sucurs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Sucurs)
      Exit Sub
   End If
   If cmb_LibCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Libro Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_LibCon)
      Exit Sub
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   If DateDiff("m", CDate(ipp_FecIni.Text), CDate(ipp_FecFin.Text)) > 3 Then
      MsgBox "Consulta demasiada información, favor verifique el rango de fechas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   modctb_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
   modctb_str_NomEmp = cmb_Empres.Text
   
   modctb_str_CodSuc = l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo
   modctb_str_NomSuc = cmb_Sucurs.Text
   
   modctb_int_CodLib = cmb_LibCon.ItemData(cmb_LibCon.ListIndex)
   modctb_str_NomLib = cmb_LibCon.Text
   
   If moddat_g_str_CodIte = 1 Then
      Call fs_Activa(False)
   Else
      cmd_Buscar.Enabled = False
      cmd_Consul.Enabled = True
      cmd_ExpExc.Enabled = True
      grd_Listad.Enabled = True
      
      cmb_Empres.Enabled = False
      cmb_Sucurs.Enabled = False
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
      cmb_LibCon.Enabled = False
   End If
   'cmb_Buscar.Enabled = False
   'txt_Buscar.Enabled = False
      
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub pnl_id_Click()
 If Len(Trim(pnl_id.Tag)) = 0 Or pnl_id.Tag = "D" Then
      pnl_id.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "N")
   Else
      pnl_id.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
   End If
End Sub

Private Sub pnl_NroLib_Click()
   If Len(Trim(pnl_NroLib.Tag)) = 0 Or pnl_NroLib.Tag = "D" Then
      pnl_NroLib.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "N")
   Else
      pnl_NroLib.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "N-")
   End If
End Sub

Private Sub pnl_NroAsi_Click()
   If Len(Trim(pnl_NroAsi.Tag)) = 0 Or pnl_NroAsi.Tag = "D" Then
      pnl_NroAsi.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "N")
   Else
      pnl_NroAsi.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "N-")
   End If
End Sub

Private Sub pnl_FecCtb_Click()
   If Len(Trim(pnl_FecCtb.Tag)) = 0 Or pnl_FecCtb.Tag = "D" Then
      pnl_FecCtb.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_FecCtb.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Glosa_Click()
   If Len(Trim(pnl_Glosa.Tag)) = 0 Or pnl_Glosa.Tag = "D" Then
      pnl_Glosa.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Glosa.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_DebMN_Click()
   If Len(Trim(pnl_DebMN.Tag)) = 0 Or pnl_DebMN.Tag = "D" Then
      pnl_DebMN.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "N")
   Else
      pnl_DebMN.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "N-")
   End If
End Sub

Private Sub pnl_HabMN_Click()
   If Len(Trim(pnl_HabMN.Tag)) = 0 Or pnl_HabMN.Tag = "D" Then
      pnl_HabMN.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_HabMN.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_DifMN_Click()
   If Len(Trim(pnl_DifMN.Tag)) = 0 Or pnl_DifMN.Tag = "D" Then
      pnl_DifMN.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_DifMN.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_DebME_Click()
   If Len(Trim(pnl_DebME.Tag)) = 0 Or pnl_DebME.Tag = "D" Then
      pnl_DebME.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_DebME.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_HabME_Click()
   If Len(Trim(pnl_HabME.Tag)) = 0 Or pnl_HabME.Tag = "D" Then
      pnl_HabME.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 9, "N")
   Else
      pnl_HabME.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 9, "N-")
   End If
End Sub

Private Sub pnl_DifME_Click()
   If Len(Trim(pnl_DifME.Tag)) = 0 Or pnl_DifME.Tag = "D" Then
      pnl_DifME.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 10, "N")
   Else
      pnl_DifME.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 10, "N-")
   End If
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If (cmb_Buscar.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub
