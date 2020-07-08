VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   9060
   ClientTop       =   4140
   ClientWidth     =   10815
   Icon            =   "GesCtb_frm_008.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8805
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   15531
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
         Left            =   30
         TabIndex        =   26
         Top             =   30
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
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
            Height          =   270
            Left            =   630
            TabIndex        =   27
            Top             =   150
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de ITF"
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
            Picture         =   "GesCtb_frm_008.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   28
         Top             =   750
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   5490
            Picture         =   "GesCtb_frm_008.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CanGra 
            Height          =   585
            Left            =   7350
            Picture         =   "GesCtb_frm_008.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_008.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_008.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   4890
            Picture         =   "GesCtb_frm_008.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   3690
            Picture         =   "GesCtb_frm_008.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Nueva"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   4290
            Picture         =   "GesCtb_frm_008.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_IngMan 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_008.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ingreso Manual"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_008.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10140
            Picture         =   "GesCtb_frm_008.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_008.frx":23EA
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_008.frx":26F4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   9090
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   9540
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   29
         Top             =   1440
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
         _ExtentY        =   1402
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
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3800
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3800
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   6750
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   2000
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   6750
            TabIndex        =   3
            Top             =   390
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label13 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   90
            TabIndex        =   54
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   90
            TabIndex        =   32
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   5760
            TabIndex        =   31
            Top             =   450
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   5760
            TabIndex        =   30
            Top             =   120
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4215
         Left            =   30
         TabIndex        =   33
         Top             =   2250
         Width           =   10785
         _Version        =   65536
         _ExtentX        =   19024
         _ExtentY        =   7435
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisOpe 
            Height          =   3765
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   6641
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   90
            Width           =   700
            _Version        =   65536
            _ExtentX        =   1235
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   2280
            TabIndex        =   36
            Top             =   90
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Declarante"
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
            Left            =   8070
            TabIndex        =   37
            Top             =   90
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Soles"
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
            Left            =   3780
            TabIndex        =   38
            Top             =   90
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Mov."
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
            Left            =   4770
            TabIndex        =   39
            Top             =   90
            Width           =   3300
            _Version        =   65536
            _ExtentX        =   5821
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Movimiento"
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
            Left            =   9210
            TabIndex        =   40
            Top             =   90
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ITF Soles"
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
            Left            =   780
            TabIndex        =   41
            Top             =   90
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
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
      Begin Threed.SSPanel SSPanel12 
         Height          =   825
         Left            =   30
         TabIndex        =   42
         Top             =   6480
         Width           =   10785
         _Version        =   65536
         _ExtentX        =   19024
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   14
            Top             =   420
            Width           =   2400
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   90
            Width           =   2400
         End
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   6060
            MaxLength       =   6
            TabIndex        =   15
            Top             =   90
            Width           =   1800
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   45
            Top             =   450
            Width           =   1725
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   44
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   4590
            TabIndex        =   43
            Top             =   120
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   795
         Left            =   30
         TabIndex        =   46
         Top             =   7320
         Width           =   10785
         _Version        =   65536
         _ExtentX        =   19024
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_TipDec 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   60
            Width           =   2400
         End
         Begin VB.ComboBox cmb_TipOpe 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   390
            Width           =   2400
         End
         Begin VB.ComboBox cmb_TipMov 
            Height          =   315
            Left            =   6030
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   60
            Width           =   3600
         End
         Begin EditLib.fpDateTime ipp_FecDep 
            Height          =   315
            Left            =   6030
            TabIndex        =   19
            Top             =   390
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
         Begin VB.Label Label5 
            Caption         =   "Tipo Declarante:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Deposito:"
            Height          =   315
            Left            =   4590
            TabIndex        =   49
            Top             =   450
            Width           =   1395
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Movimiento:"
            Height          =   315
            Left            =   4590
            TabIndex        =   47
            Top             =   90
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   465
         Left            =   30
         TabIndex        =   51
         Top             =   8130
         Width           =   10785
         _Version        =   65536
         _ExtentX        =   19024
         _ExtentY        =   820
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
         Begin EditLib.fpDoubleSingle ipp_MtoItf 
            Height          =   315
            Left            =   6030
            TabIndex        =   21
            Top             =   60
            Width           =   1800
            _Version        =   196608
            _ExtentX        =   3175
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
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
         Begin EditLib.fpDoubleSingle ipp_MtoSol 
            Height          =   315
            Left            =   1860
            TabIndex        =   20
            Top             =   60
            Width           =   1800
            _Version        =   196608
            _ExtentX        =   3175
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
         Begin VB.Label Label11 
            Caption         =   "Monto Soles:"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   90
            Width           =   1725
         End
         Begin VB.Label Label10 
            Caption         =   "ITF Soles:"
            Height          =   285
            Left            =   4620
            TabIndex        =   52
            Top             =   90
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_FecIni       As String
Dim r_str_FecFin       As String
Dim r_arr_PriItf()     As String
Dim r_int_Contad       As Integer
Dim l_arr_Empres()     As moddat_tpo_Genera
Dim l_arr_Sucurs()     As moddat_tpo_Genera
Dim r_int_TipRep       As Integer

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
      Call gs_SetFocus(cmb_Sucurs)
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_PerMes_Click()
   If cmb_PerMes.ListIndex > -1 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmb_Sucurs_Click()
   If cmb_Sucurs.ListIndex > -1 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   Call gs_Activa_2(False)
   Call gs_Activa_3(True)
   Call gs_SetFocus(cmb_TipDoc)
   
   ipp_FecDep.Text = CDate("01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000"))
   ipp_FecDep.DateMin = Format(CDate(CDate("01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000"))), "yyyymmdd")
   ipp_FecDep.DateMax = Format(CDate(CDate(Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text))) & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000"))), "yyyymmdd")
End Sub

Private Sub cmd_Borrar_Click()
   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 1
   moddat_g_str_TipDoc = Left(Trim(grd_LisOpe.Text), 1)
   moddat_g_str_NumDoc = Mid(Trim(grd_LisOpe.Text), 3, Len(Trim(grd_LisOpe.Text)) - 2)
   Screen.MousePointer = 11
     
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_DETITF WHERE DETITF_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND DETITF_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & moddat_g_str_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & moddat_g_str_NumDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_CODEMP = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "DETITF_CODSUC = '" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   Screen.MousePointer = 0
   Call gs_RefrescaGrid(grd_LisOpe)
   MsgBox "Registro eliminado.", vbInformation, modgen_g_str_NomPlt
   Call cmd_BusOpe_Click
End Sub

Private Sub cmd_Cancel_Click()
   Call gs_Activa_2(False)
   Call gs_Activa_1(True)
   Call gs_LimpiaGrid(grd_LisOpe)
End Sub

Private Sub cmd_CanIng_Click()
   txt_NumCom.Text = ""
   txt_NumDoc.Text = ""
   ipp_MtoSol.Text = 0
   ipp_MtoItf.Text = 0
   ipp_FecDep.Text = Format(Now, "DD/MM/YYYY")
   cmb_TipDoc.ListIndex = -1
   cmb_TipDec.ListIndex = -1
   cmb_TipOpe.ListIndex = -1
   cmb_TipMov.ListIndex = -1
End Sub

Private Sub cmd_CanGra_Click()
   Call gs_Activa_3(False)
   Call gs_Activa_2(True)
   cmd_CanIng_Click
End Sub

Private Sub cmd_Editar_Click()
   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   Call gs_Activa_2(False)
   Call gs_Activa_3(True)
   
   modsec_g_str_Period = Trim(grd_LisOpe.Text)
   grd_LisOpe.Col = 1
   moddat_g_str_TipDoc = Left(Trim(grd_LisOpe.Text), 1)
   moddat_g_str_NumDoc = Mid(Trim(grd_LisOpe.Text), 3, Len(Trim(grd_LisOpe.Text)) - 2)
   
   grd_LisOpe.Col = 5
   modsec_g_dbl_MtoSol = Trim(grd_LisOpe.Text)
   
   grd_LisOpe.Col = 6
   modsec_g_dbl_ITFSol = Trim(grd_LisOpe.Text)
   
   grd_LisOpe.Col = 7
   moddat_g_int_TipCli = Trim(grd_LisOpe.Text)
   
   Call gs_RefrescaGrid(grd_LisOpe)
   ipp_FecDep.DateMin = Format(CDate(CDate("01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000"))), "yyyymmdd")
   ipp_FecDep.DateMax = Format(CDate(CDate(Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text))) & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000"))), "yyyymmdd")
   Call ff_Buscar
   
   cmb_TipDoc.Enabled = False
   txt_NumCom.Enabled = False
   txt_NumDoc.Enabled = False
   Call gs_SetFocus(cmb_TipDec)
End Sub

Private Sub cmd_ExpExc_Click()
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
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   r_int_TipRep = MsgBoxExText("Elija el Tipo de Reporte", vbInformation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt, YES, "1ra Quincena", NO, "2da Quincena")
   
   'Confirmacion
   If MsgBox("¿Está seguro que deseas exportar el ITF de la " & IIf(r_int_TipRep = 6, "Primera Quincena", "Segunda Quincena") & " ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
      
   'Call fs_Genera_Cajmov(r_str_FecIni, r_str_FecFin)
   Call fs_GenExc
End Sub

Private Sub cmd_Export_Click()
Dim r_rst_PerMes        As ADODB.Recordset
Dim r_str_CtbFin        As String
   
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
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
         
   If MsgBox("¿Está seguro de generar los archivos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
'   'COMPARACION CON TABLE CTB_PERMES
'   g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
'   g_str_Parame = g_str_Parame & "PERMES_CODEMP = 00001 AND "
'   g_str_Parame = g_str_Parame & "PERMES_TIPPER = 1 AND "
'   g_str_Parame = g_str_Parame & "PERMES_CODANO = " & ipp_PerAno.Text & " AND "
'   g_str_Parame = g_str_Parame & "PERMES_CODMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & " "
'   'g_str_Parame = g_str_Parame & "PERMES_SITUAC = 1 "
'   g_str_Parame = g_str_Parame & "ORDER BY PERMES_CODANO, PERMES_CODMES ASC"
'
'   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerMes, 3) Then
'      Exit Sub
'   End If
'
'   If Not (r_rst_PerMes.BOF And r_rst_PerMes.EOF) Then
'      r_rst_PerMes.MoveFirst
'      Do While Not r_rst_PerMes.EOF
'         r_str_FecIni = r_rst_PerMes!PERMES_FECINI
'         r_str_FecFin = r_rst_PerMes!PERMES_FECFIN
'         r_rst_PerMes.MoveNext
'      Loop
'   End If
'
'   r_rst_PerMes.Close
'   Set r_rst_PerMes = Nothing
   
'   If cmb_TipRep.ListIndex = 0 Then
'      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15"
'      If ff_BusPer(Left(r_str_FecIni, 6) & "01", Left(r_str_FecIni, 6) & "15") = 0 Then
'         MsgBox "1ra Quincena No Procesada. Procesar la 1ra Quincena.", vbInformation, modgen_g_str_NomPlt
'         Screen.MousePointer = 0
'         Exit Sub
'      End If
'   ElseIf cmb_TipRep.ListIndex = 1 Then
'      If ff_BusPer(Left(r_str_FecIni, 6) & "01", Left(r_str_FecIni, 6) & "15") = 0 Then
'         MsgBox "No se a procesado la 1ra Quincena. Necesita procesarla antes de continuar con la 2da Quincena.", vbInformation, modgen_g_str_NomPlt
'         Screen.MousePointer = 0
'         Exit Sub
'      End If
'      If ff_BusPer(Left(r_str_FecIni, 6) & "16", r_str_FecFin) = 0 Then
'         MsgBox "2da Quincena No Procesada. Procesar la 2da Quincena.", vbInformation, modgen_g_str_NomPlt
'         Screen.MousePointer = 0
'         Exit Sub
'      End If
'   End If
   
   Screen.MousePointer = 11
   If ff_BusITF(ipp_PerAno.Text, cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "DELETE FROM CTB_CABITF "
      g_str_Parame = g_str_Parame & " WHERE CABITF_PERANO = " & ipp_PerAno.Text & " "
      g_str_Parame = g_str_Parame & "   AND CABITF_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Sub
      End If
   End If
   
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   Call fs_Genera_MovAcum(r_str_FecIni, r_str_FecFin)
   Call fs_Genera_MovAcu(dlg_Guarda.FileName, r_str_FecIni, r_str_FecFin)
   Call fs_Genera_AsiErr(dlg_Guarda.FileName, r_str_FecIni, r_str_FecFin)
   
   Screen.MousePointer = 0
   MsgBox "Archivos Creados.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_TipCam        As Double
   Dim r_dbl_Porcen        As Double
   Dim r_int_NroCom        As Long
   Dim r_str_Numero        As String

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:
            If Len(Trim(txt_NumDoc.Text)) <> 8 Then
               MsgBox "Debe ingresar un Número de Documento de 8 dígitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
         Case 6:
            If Len(Trim(txt_NumDoc.Text)) <> 11 Then
               MsgBox "Debe ingresar un Número de RUC de 11 dígitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
            
         Case Else:
            If Len(Trim(txt_NumDoc.Text)) < 8 Then
               MsgBox "Debe ingresar un Número de Documento de 12 dígitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
      End Select
   End If
   
   If Trim(txt_NumDoc.Text) = "" Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
      If Trim(txt_NumCom.Text) = "" Then
         MsgBox "Debe ingresar el Número del Comprobante.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumCom)
         Exit Sub
      End If
   End If
   If cmb_TipDec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Declarante.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDec)
      Exit Sub
   End If
   If cmb_TipOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOpe)
      Exit Sub
   End If
   'If cmb_TipMov.ListIndex = -1 Then
   '   MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(cmb_TipMov)
   '   Exit Sub
   'End If
   If ipp_MtoSol.Text = 0 Or ipp_MtoSol.Text = "" Then
      MsgBox "Debe ingresar el Monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoSol)
      Exit Sub
   End If
   If ipp_MtoItf.Text = "" Then
      MsgBox "Debe ingresar el ITF.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoItf)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      If MsgBox("¿Está seguro que desea realizar el ingreso manual del ITF?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
         r_int_NroCom = modsec_gf_BusMov(Format(Mid(ipp_FecDep.Text, 4, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000"))
      Else
         r_int_NroCom = Trim(txt_NumCom.Text)
      End If
      
      r_dbl_TipCam = ff_TipCam(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      r_dbl_Porcen = ff_Porcen(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      
      If r_dbl_TipCam = 0 Then
         Screen.MousePointer = 0
         MsgBox "El tipo de cambio no esta registrado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         r_str_Numero = ff_BuscarNumero(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex), cmb_TipMov.ItemData(cmb_TipMov.ListIndex))
         If Len(Trim(r_str_Numero)) = 0 Then
            Screen.MousePointer = 0
            MsgBox "No se pudo ubicar el numero de operacion y/o solicitud de referencia.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
      Else
         r_str_Numero = ""
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 1 Then
            If Len(r_str_Numero) < 12 Then
               Screen.MousePointer = 0
               MsgBox "El documento ingresado no presenta Número de Solicitud .", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         Else
            If Len(r_str_Numero) < 10 Then
               Screen.MousePointer = 0
               MsgBox "El documento ingresado no presenta Número de Operación .", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         End If
      End If
               
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_DETITF("
      g_str_Parame = g_str_Parame & "DETITF_PERMES, "
      g_str_Parame = g_str_Parame & "DETITF_PERANO, "
      g_str_Parame = g_str_Parame & "DETITF_TIPDOC, "
      g_str_Parame = g_str_Parame & "DETITF_TIPDEC, "
      g_str_Parame = g_str_Parame & "DETITF_FECMOV, "
      g_str_Parame = g_str_Parame & "DETITF_NUMDOC, "
      g_str_Parame = g_str_Parame & "DETITF_TIPMOV, "
      g_str_Parame = g_str_Parame & "DETITF_NROCOM, "
      g_str_Parame = g_str_Parame & "DETITF_TIPCOD, "
      g_str_Parame = g_str_Parame & "DETITF_ITFPOR, "
      g_str_Parame = g_str_Parame & "DETITF_MTOORG, "
      g_str_Parame = g_str_Parame & "DETITF_ITFORG, "
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL, "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL, "
      g_str_Parame = g_str_Parame & "DETITF_MTODOL, "
      g_str_Parame = g_str_Parame & "DETITF_ITFDOL, "
      g_str_Parame = g_str_Parame & "DETITF_OPEREF, "
      g_str_Parame = g_str_Parame & "DETITF_TIPMON, "
      g_str_Parame = g_str_Parame & "DETITF_TIPCAM, "
      g_str_Parame = g_str_Parame & "DETITF_MANUAL, "
      g_str_Parame = g_str_Parame & "DETITF_CODEMP, "
      g_str_Parame = g_str_Parame & "DETITF_CODSUC) "
                                 
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(ipp_PerAno.Text, "0000") & ", "
      g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
      g_str_Parame = g_str_Parame & cmb_TipDec.ItemData(cmb_TipDec.ListIndex) & ", "                             ' 1 - DECLARANTE / 2 - EXTORNO
      g_str_Parame = g_str_Parame & Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00") & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_TipMov.Text & "', "
      g_str_Parame = g_str_Parame & "'" & r_int_NroCom & "', "
      g_str_Parame = g_str_Parame & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "                              '1 - SOLICITUD / 2 - OPERACION
      g_str_Parame = g_str_Parame & r_dbl_Porcen & ","
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ","
      Else
         g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text, "###########0.00") & ","
      End If
      
      g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text, "###########0.00") & ","
      g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ","
      g_str_Parame = g_str_Parame & "'" & r_str_Numero & "', "
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & 2 & ", "
      Else
         g_str_Parame = g_str_Parame & 1 & ", "
      End If
      
      g_str_Parame = g_str_Parame & r_dbl_TipCam & ","
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "'" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      MsgBox "Se realizó el ingreso manual.", vbInformation, modgen_g_str_NomPlt
   
   ElseIf moddat_g_int_FlgGrb = 2 Then
   
      If MsgBox("¿Está seguro que desea modificar el ingreso manual del ITF?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      r_dbl_TipCam = ff_TipCam(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      r_dbl_Porcen = ff_Porcen(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         r_str_Numero = ff_BuscarNumero(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex), cmb_TipMov.ItemData(cmb_TipMov.ListIndex))
      Else
         r_str_Numero = ""
      End If
               
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CTB_DETITF SET "
      g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & cmb_TipDec.ItemData(cmb_TipDec.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "DETITF_FECMOV = " & Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_TIPMOV = '" & cmb_TipMov.Text & "', "
      g_str_Parame = g_str_Parame & "DETITF_NROCOM = '" & Trim(txt_NumCom.Text) & "', "
      g_str_Parame = g_str_Parame & "DETITF_TIPCOD = " & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFPOR = " & r_dbl_Porcen & ", "
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & "DETITF_MTOORG = " & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & "DETITF_ITFORG = " & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ", "
      Else
         g_str_Parame = g_str_Parame & "DETITF_MTOORG = " & Format(ipp_MtoSol.Text, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & "DETITF_ITFORG = " & Format(ipp_MtoItf.Text, "###########0.00") & ", "
      End If
      
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & Format(ipp_MtoSol.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & Format(ipp_MtoItf.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_MTODOL = " & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFDOL = " & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_OPEREF = '" & r_str_Numero & "', "
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & "DETITF_TIPMON = " & 2 & ", "
      Else
         g_str_Parame = g_str_Parame & "DETITF_TIPMON = " & 1 & ", "
      End If
      g_str_Parame = g_str_Parame & "DETITF_TIPCAM = " & r_dbl_TipCam & " "
      
      g_str_Parame = g_str_Parame & "WHERE "
      g_str_Parame = g_str_Parame & "DETITF_PERMES = " & Mid(modsec_g_str_Period, 6, 2) & " AND "
      g_str_Parame = g_str_Parame & "DETITF_PERANO = " & Mid(modsec_g_str_Period, 1, 4) & " AND "
      g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & moddat_g_str_TipDoc & " AND "
      g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & moddat_g_str_NumDoc & " AND "
      g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & moddat_g_int_TipCli & " AND "
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & modsec_g_dbl_MtoSol & " AND "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & modsec_g_dbl_ITFSol & " AND "
      g_str_Parame = g_str_Parame & "DETITF_CODEMP = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "DETITF_CODSUC = '" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "' "
                                                   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      MsgBox "Se modificó el ingreso manual.", vbInformation, modgen_g_str_NomPlt
      Call cmd_CanGra_Click
   End If
   
   Call cmd_BusOpe_Click
   Call cmd_CanIng_Click
End Sub

Private Sub cmd_IngMan_Click()
   Call cmd_BusOpe_Click
End Sub

Private Sub gs_Activa_1(ByVal p_Activa As Boolean)
   cmd_Proces.Enabled = p_Activa
   cmd_ExpExc.Enabled = p_Activa
   cmd_Export.Enabled = p_Activa
   cmd_IngMan.Enabled = p_Activa
   cmd_Limpia.Enabled = p_Activa
   cmb_Empres.Enabled = p_Activa
   cmb_Sucurs.Enabled = p_Activa
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
End Sub

Private Sub gs_Activa_2(ByVal p_Activa As Boolean)
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
End Sub

Private Sub gs_Activa_3(ByVal p_Activa As Boolean)
   cmd_Grabar.Enabled = p_Activa
   cmd_CanGra.Enabled = p_Activa
   cmb_TipDoc.Enabled = p_Activa
   txt_NumCom.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   cmb_TipDec.Enabled = p_Activa
   cmb_TipOpe.Enabled = p_Activa
   cmb_TipMov.Enabled = p_Activa
   ipp_FecDep.Enabled = p_Activa
   ipp_MtoSol.Enabled = p_Activa
   ipp_MtoItf.Enabled = p_Activa
End Sub

Private Sub cmd_Limpia_Click()
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = 0
   cmb_Empres.ListIndex = -1
   cmb_Sucurs.Clear
End Sub

Private Sub cmd_Proces_Click()
Dim r_rst_PerMes        As ADODB.Recordset
Dim r_str_CtbFin        As String
   
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
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   r_int_TipRep = MsgBoxExText("Elija el Tipo de Reporte", vbInformation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt, YES, "1ra Quincena", NO, "2da Quincena")
   
   If MsgBox("¿Está seguro que desea realizar el proceso de ITF de la " & IIf(r_int_TipRep = 6, "Primera Quincena", "Segunda Quincena") & " ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If r_int_TipRep = 6 Then
      r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15"
      r_str_CtbFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15"
      
      If ff_BusPer(Left(r_str_FecIni, 6) & "01", Left(r_str_FecIni, 6) & "15") = 0 Then
         Call fs_Genera_CajMov(r_str_FecIni, r_str_FecFin, r_str_CtbFin)
      Else
         MsgBox "1ra Quincena Procesada. No se puede volver a Procesar.", vbInformation, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
   
   ElseIf r_int_TipRep = 7 Then
      r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "16"
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "01"
      r_str_CtbFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
            
      If ff_BusPer(Left(r_str_FecIni, 6) & "01", Left(r_str_FecIni, 6) & "15") = 0 Then
         MsgBox "No se a procesado la 1ra Quincena. Necesita procesarla antes continuar con la 2da Quincena.", vbInformation, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      If ff_BusPer(r_str_FecIni, r_str_CtbFin) = 0 Then
         Call ff_PriItf(Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15", Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15")
         
         'COMPARACION CON TABLE CTB_PERMES
         g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
         g_str_Parame = g_str_Parame & "PERMES_CODEMP = 00001 AND "
         g_str_Parame = g_str_Parame & "PERMES_TIPPER = 1 AND "
         g_str_Parame = g_str_Parame & "PERMES_CODANO = " & ipp_PerAno.Text & " AND "
         g_str_Parame = g_str_Parame & "PERMES_CODMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & " "
         g_str_Parame = g_str_Parame & "ORDER BY PERMES_CODANO, PERMES_CODMES ASC"
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerMes, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_PerMes.BOF And r_rst_PerMes.EOF) Then
            r_rst_PerMes.MoveFirst
            
            Do While Not r_rst_PerMes.EOF
               r_str_FecFin = r_rst_PerMes!PERMES_FECFIN
               r_rst_PerMes.MoveNext
            Loop
         End If
         
         r_rst_PerMes.Close
         Set r_rst_PerMes = Nothing
         Call fs_Genera_CajMov(r_str_FecIni, r_str_FecFin, r_str_CtbFin)
      Else
         MsgBox "2da Quincena Procesada. No se puede volver a Procesar.", vbInformation, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = 0
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'PROCESO ITF
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
   'LISTADO INGRESOS MANUALES
   grd_LisOpe.ColWidth(0) = 700
   grd_LisOpe.ColWidth(1) = 1500
   grd_LisOpe.ColWidth(2) = 1450
   grd_LisOpe.ColWidth(3) = 1000
   grd_LisOpe.ColWidth(4) = 3300
   grd_LisOpe.ColWidth(5) = 1140
   grd_LisOpe.ColWidth(6) = 1140
   grd_LisOpe.ColWidth(7) = 0
      
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(2) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(3) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(4) = flexAlignLeftCenter
   grd_LisOpe.ColAlignment(5) = flexAlignRightCenter
   grd_LisOpe.ColAlignment(6) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'INGRESO MANUAL DEL ITF
   ipp_FecDep.Text = Format(Now, "DD/MM/YYYY")
   ipp_FecDep.BackColor = modgen_g_con_ColAma
   txt_NumCom.Enabled = False
   
   cmb_TipDoc.Clear
   cmb_TipDoc.AddItem "DNI"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(1)
   cmb_TipDoc.AddItem "CARNE DE EXTRANJERIA"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(2)
   cmb_TipDoc.AddItem "PASAPORTE"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(5)
   cmb_TipDoc.AddItem "RUC"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(6)
   
   cmb_TipDec.Clear
   cmb_TipDec.AddItem "DECLARANTE"
   cmb_TipDec.ItemData(cmb_TipDec.NewIndex) = CInt(1)
   cmb_TipDec.AddItem "EXTORNO"
   cmb_TipDec.ItemData(cmb_TipDec.NewIndex) = CInt(2)
   
   cmb_TipOpe.Clear
   cmb_TipOpe.AddItem "SOLICITUD"
   cmb_TipOpe.ItemData(cmb_TipOpe.NewIndex) = CInt(1)
   cmb_TipOpe.AddItem "OPERACION"
   cmb_TipOpe.ItemData(cmb_TipOpe.NewIndex) = CInt(2)
   cmb_TipOpe.AddItem "PLAN AHORRO"
   cmb_TipOpe.ItemData(cmb_TipOpe.NewIndex) = CInt(3)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMov, 1, "301")
   ipp_FecDep.BackColor = modgen_g_con_ColAma
   
   Call cmd_Limpia_Click
   Call cmd_Cancel_Click
   Call cmd_CanIng_Click
   Call gs_Activa_1(True)
   Call gs_Activa_2(False)
   Call gs_Activa_3(False)
End Sub

Private Sub grd_LisOpe_SelChange()
   If grd_LisOpe.Rows > 2 Then
      grd_LisOpe.RowSel = grd_LisOpe.Row
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

'Llenado de la Tabla CTB_DETITF con los datos de la OPE_CAJMOV
Private Sub fs_Genera_CajMov(ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_CtbFin As String)
Dim r_str_TipMov As String
Dim r_str_Bancos As String
Dim r_dbl_TipCam As Double
Dim r_dbl_Porcen As Double
Dim r_str_FecPag As String
Dim r_int_FlgRep As Integer
Dim r_str_DesCof As String
         
   'Leyendo Tabla de Movimientos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_FECMOV >= " & p_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & p_FecFin & " "
   If r_int_TipRep = 6 Then
      g_str_Parame = g_str_Parame & "   AND CAJMOV_FECDEP >= " & p_FecIni & " "
   ElseIf r_int_TipRep = 7 Then
      g_str_Parame = g_str_Parame & "   AND CAJMOV_FECDEP <= " & p_CtbFin & " "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_FECMOV ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_int_FlgRep = 0
         
         If CStr(Trim(g_rst_Princi!CAJMOV_TIPMOV)) = "1103" Then
            If CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "001" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "002" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "006" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "011" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "012" Then
               r_str_FecPag = ff_BusCof(Trim(g_rst_Princi!CAJMOV_TIPDOC), Trim(g_rst_Princi!CAJMOV_NUMDOC), p_CtbFin)
               If r_str_FecPag <= p_CtbFin Then
                  r_int_FlgRep = 0
               Else
                  r_int_FlgRep = 1
               End If
            Else
               If CDate(gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))) > CDate(gf_FormatoFecha(CStr(p_CtbFin))) Then
                  r_str_FecPag = p_CtbFin
               Else
                  r_str_FecPag = g_rst_Princi!CAJMOV_FECMOV
               End If
            End If
         Else
            If CDate(gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))) > CDate(gf_FormatoFecha(CStr(p_CtbFin))) Then
               r_str_FecPag = p_CtbFin
            Else
               r_str_FecPag = g_rst_Princi!CAJMOV_FECMOV
            End If
         End If
         
         'Obtenemos los campos correspondientes de las tablas a relacionar
         r_str_TipMov = moddat_gf_Consulta_ParDes("301", CStr(g_rst_Princi!CAJMOV_TIPMOV))
         r_str_Bancos = moddat_gf_Consulta_ParDes("505", CStr(g_rst_Princi!CAJMOV_CODBAN))
         r_dbl_TipCam = ff_TipCam(g_rst_Princi!CAJMOV_FECDEP, p_CtbFin)
         r_dbl_Porcen = ff_Porcen(p_FecIni, p_FecFin)
         
         If r_dbl_TipCam = 0 Then
            MsgBox "El tipo de cambio no esta registrado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmd_Proces)
            Exit Sub
         End If
         
         If r_int_TipRep = 6 Then
            If g_rst_Princi!CAJMOV_FECMOV > p_CtbFin Then
               If CStr(Trim(g_rst_Princi!CAJMOV_TIPMOV)) = "1103" Then
                  If CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "001" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "002" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "006" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "011" And CStr(Left(Trim(g_rst_Princi!CAJMOV_NUMOPE), 3)) <> "012" Then
                     r_str_DesCof = ff_BusCof(Trim(g_rst_Princi!CAJMOV_TIPDOC), Trim(g_rst_Princi!CAJMOV_NUMDOC), p_CtbFin)
                  Else
                     If CDate(gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))) > CDate(gf_FormatoFecha(CStr(p_CtbFin))) Then
                        r_str_DesCof = g_rst_Princi!CAJMOV_FECMOV
                     End If
                  End If
                  
                  If r_str_DesCof <= p_CtbFin Then
                     r_str_FecPag = r_str_DesCof
                     r_int_FlgRep = 0
                  Else
                     r_int_FlgRep = 1
                  End If
               Else
                  r_int_FlgRep = 1
               End If
            Else
               r_int_FlgRep = 0
            End If
         End If
         
         If r_int_TipRep = 7 Then
            If UBound(r_arr_PriItf) > 0 Then
               For r_int_Contad = 0 To UBound(r_arr_PriItf) Step 7
                  If r_arr_PriItf(r_int_Contad + 0) = cmb_PerMes.ItemData(cmb_PerMes.ListIndex) And _
                     r_arr_PriItf(r_int_Contad + 1) = ipp_PerAno.Text And _
                     r_arr_PriItf(r_int_Contad + 2) = Trim(g_rst_Princi!CAJMOV_TIPDOC) And _
                     r_arr_PriItf(r_int_Contad + 3) = Trim(g_rst_Princi!CAJMOV_NUMDOC) And _
                     r_arr_PriItf(r_int_Contad + 4) = IIf(g_rst_Princi!CAJMOV_TIPMOV = 2101, 2, 1) And _
                     r_arr_PriItf(r_int_Contad + 5) = Trim(g_rst_Princi!CAJMOV_FECMOV) And _
                     r_arr_PriItf(r_int_Contad + 6) = Trim(r_str_TipMov) Then
                     r_int_FlgRep = 1
                     Exit For
                  End If
               Next
            End If
         End If
                  
         If r_int_FlgRep = 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO CTB_DETITF("
            g_str_Parame = g_str_Parame & "DETITF_PERMES, "
            g_str_Parame = g_str_Parame & "DETITF_PERANO, "
            g_str_Parame = g_str_Parame & "DETITF_TIPDOC, "
            g_str_Parame = g_str_Parame & "DETITF_NUMDOC, "
            g_str_Parame = g_str_Parame & "DETITF_TIPDEC, "
            g_str_Parame = g_str_Parame & "DETITF_FECMOV, "
            g_str_Parame = g_str_Parame & "DETITF_TIPMOV, "
            g_str_Parame = g_str_Parame & "DETITF_NROCOM, "
            g_str_Parame = g_str_Parame & "DETITF_TIPCOD, "
            g_str_Parame = g_str_Parame & "DETITF_OPEREF, "
            g_str_Parame = g_str_Parame & "DETITF_TIPMON, "
            g_str_Parame = g_str_Parame & "DETITF_MTOORG, "
            g_str_Parame = g_str_Parame & "DETITF_ITFORG, "
            g_str_Parame = g_str_Parame & "DETITF_ITFPOR, "
            g_str_Parame = g_str_Parame & "DETITF_MTOSOL, "
            g_str_Parame = g_str_Parame & "DETITF_ITFSOL, "
            g_str_Parame = g_str_Parame & "DETITF_MTODOL, "
            g_str_Parame = g_str_Parame & "DETITF_ITFDOL, "
            g_str_Parame = g_str_Parame & "DETITF_TIPCAM, "
            g_str_Parame = g_str_Parame & "DETITF_CODBAN, "
            g_str_Parame = g_str_Parame & "DETITF_NUMCTA, "
            g_str_Parame = g_str_Parame & "DETITF_MANUAL, "
            g_str_Parame = g_str_Parame & "DETITF_CODEMP, "
            g_str_Parame = g_str_Parame & "DETITF_CODSUC) "
                                       
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & ", "
            g_str_Parame = g_str_Parame & Format(ipp_PerAno.Text, "0000") & ", "
            g_str_Parame = g_str_Parame & g_rst_Princi!CAJMOV_TIPDOC & ", "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAJMOV_NUMDOC & "', "
            
            If g_rst_Princi!CAJMOV_TIPMOV = 2101 Then
               g_str_Parame = g_str_Parame & 2 & ", "                             ' 1 - DECLARANTE / 2 - EXTORNO
            Else
               g_str_Parame = g_str_Parame & 1 & ", "
            End If
            
            g_str_Parame = g_str_Parame & r_str_FecPag & ", "
            g_str_Parame = g_str_Parame & "'" & r_str_TipMov & "', "
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!CAJMOV_NUMMOV) & "', "
            
            If Len(Trim(g_rst_Princi!CAJMOV_NUMOPE)) = 10 Then                            '1 - SOLICITUD / 2 - OPERACION
               g_str_Parame = g_str_Parame & 2 & ", "
            Else
               g_str_Parame = g_str_Parame & 1 & ", "
            End If
            
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!CAJMOV_NUMOPE) & "', "
            g_str_Parame = g_str_Parame & g_rst_Princi!CAJMOV_MONPAG & ", "
            
            'BASE IMPONIBLE EN MONEDA ORIGINAL
            g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_IMPPAG) & ", "
            
            'IMPORTE ITF EN MONEDA ORIGINAL
            g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_ITFIMP) & ", "
            
            'PORCENTAJE DE ITF
            g_str_Parame = g_str_Parame & r_dbl_Porcen & ", "
            
            'BASE IMPONIBLE EN SOLES
            If g_rst_Princi!CAJMOV_MONPAG = 1 Then
               g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_IMPPAG) & ", "
            ElseIf g_rst_Princi!CAJMOV_MONPAG = 2 Then
               g_str_Parame = g_str_Parame & gf_Truncar_Numero(g_rst_Princi!CAJMOV_IMPPAG * r_dbl_TipCam, 2) & ", "
            End If
         
            'IMPORTE ITF EN SOLES
            If g_rst_Princi!CAJMOV_MONPAG = 1 Then
               g_str_Parame = g_str_Parame & g_rst_Princi!CAJMOV_ITFIMP & ", "
            ElseIf g_rst_Princi!CAJMOV_MONPAG = 2 Then
               If g_rst_Princi!CAJMOV_ITFIMP <> 0 Then
                  'g_str_Parame = g_str_Parame & TRUNC((r_dbl_Porcen / 100) * (g_rst_Princi!CAJMOV_IMPPAG * r_dbl_TipCam), 2) & ", "
                  g_str_Parame = g_str_Parame & gf_Truncar_Numero(g_rst_Princi!CAJMOV_ITFIMP * r_dbl_TipCam, 2) & ", "
               Else
                  g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_ITFIMP) & ", "
               End If
            End If
                              
            'BASE IMPONIBLE EN DOLARES
            If g_rst_Princi!CAJMOV_MONPAG = 1 Then
               g_str_Parame = g_str_Parame & gf_Truncar_Numero((g_rst_Princi!CAJMOV_IMPPAG / r_dbl_TipCam), 2) & ", "
            ElseIf g_rst_Princi!CAJMOV_MONPAG = 2 Then
               g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_IMPPAG) & ", "
            End If
         
            'IMPORTE ITF EN DOLARES
            If g_rst_Princi!CAJMOV_MONPAG = 1 Then
               If g_rst_Princi!CAJMOV_TIPMOV <> 1102 Then
                  'g_str_Parame = g_str_Parame & TRUNC((r_dbl_Porcen / 100) * (g_rst_Princi!CAJMOV_IMPPAG / r_dbl_TipCam), 2) & ", "
                  g_str_Parame = g_str_Parame & gf_Truncar_Numero((g_rst_Princi!CAJMOV_ITFIMP / r_dbl_TipCam), 2) & ", "
               Else
                  g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_ITFIMP) & ", "
               End If
            ElseIf g_rst_Princi!CAJMOV_MONPAG = 2 Then
               g_str_Parame = g_str_Parame & CStr(g_rst_Princi!CAJMOV_ITFIMP) & ", "
            End If
                           
            g_str_Parame = g_str_Parame & r_dbl_TipCam & ", "
            g_str_Parame = g_str_Parame & "'" & r_str_Bancos & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAJMOV_NUMCTA & "', "
            g_str_Parame = g_str_Parame & "2, "
            g_str_Parame = g_str_Parame & "'" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & "'" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "')"
                  
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
   Else
      Screen.MousePointer = 0
      MsgBox "No se encontraron Pagos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Genera_MovAcum(ByVal p_FecIni As String, ByVal p_FecFin As String)
   Dim r_str_TipMov As String
   Dim r_str_Bancos As String
   Dim r_dbl_TipCam As Double
   Dim r_dbl_Porcen As Double
   Dim r_dbl_SumMon As Double
   Dim r_dbl_MtoSol As Double
   Dim r_dbl_ItfSol As Double
   Dim r_dbl_MtoDol As Double
   Dim r_dbl_ItfDol As Double
   Dim r_str_PaiCli As String
   
   'Leyendo Tabla de Movimientos
   g_str_Parame = "SELECT DISTINCT DETITF_TIPDOC, DETITF_NUMDOC, DETITF_TIPDEC FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV >= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV <= " & p_FecFin & " "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_TIPDOC, DETITF_NUMDOC ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF

         'Obtenemos los campos correspondientes de las tablas a relacionar
         r_dbl_SumMon = ff_SumMon(g_rst_Princi!DETITF_TIPDOC, Trim(g_rst_Princi!DETITF_NUMDOC), g_rst_Princi!DETITF_TIPDEC, r_dbl_MtoSol, r_dbl_ItfSol, r_dbl_MtoDol, r_dbl_ItfDol, Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00"), Format(ipp_PerAno.Text, "0000"))
         r_str_PaiCli = modsec_gf_BusPai(Trim(g_rst_Princi!DETITF_TIPDOC), Trim(g_rst_Princi!DETITF_NUMDOC))

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_CABITF("
         g_str_Parame = g_str_Parame & "CABITF_PERMES, "
         g_str_Parame = g_str_Parame & "CABITF_PERANO, "
         g_str_Parame = g_str_Parame & "CABITF_TIPDOC, "
         g_str_Parame = g_str_Parame & "CABITF_NUMDOC, "
         g_str_Parame = g_str_Parame & "CABITF_TIPDEC, "
         g_str_Parame = g_str_Parame & "CABITF_MTOSOL, "
         g_str_Parame = g_str_Parame & "CABITF_ITFSOL, "
         g_str_Parame = g_str_Parame & "CABITF_MTODOL, "
         g_str_Parame = g_str_Parame & "CABITF_ITFDOL, "
         g_str_Parame = g_str_Parame & "CABITF_PAICLI, "
         g_str_Parame = g_str_Parame & "CABITF_TIPOPE, "
         g_str_Parame = g_str_Parame & "CABITF_TIPMOV, "
         If g_rst_Princi!DETITF_TIPDEC = 1 Then
            g_str_Parame = g_str_Parame & "CABITF_CODOPE) "
         Else
            g_str_Parame = g_str_Parame & "CABITF_CODOPE, "
            g_str_Parame = g_str_Parame & "CABITF_PERDEV, "
            g_str_Parame = g_str_Parame & "CABITF_MODOPE) "
         End If

         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & ", "
         g_str_Parame = g_str_Parame & Format(ipp_PerAno.Text, "0000") & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!DETITF_TIPDOC & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DETITF_NUMDOC & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!DETITF_TIPDEC & ", "
         g_str_Parame = g_str_Parame & Round(r_dbl_MtoSol, 2) & ", "
         g_str_Parame = g_str_Parame & Round(r_dbl_ItfSol, 2) & ", "
         g_str_Parame = g_str_Parame & Round(r_dbl_MtoDol, 2) & ", "
         g_str_Parame = g_str_Parame & Round(r_dbl_ItfDol, 2) & ", "
         g_str_Parame = g_str_Parame & IIf(r_str_PaiCli = "4028", "''", "'" & r_str_PaiCli & "'") & ", "
         g_str_Parame = g_str_Parame & "01" & ", "
         g_str_Parame = g_str_Parame & "03" & ", "

         If g_rst_Princi!DETITF_TIPDEC = 1 Then
            g_str_Parame = g_str_Parame & "02" & ") "
         Else
            g_str_Parame = g_str_Parame & "02" & ", "
            g_str_Parame = g_str_Parame & ff_BusFec(g_rst_Princi!DETITF_TIPDOC, g_rst_Princi!DETITF_NUMDOC) & ", "
            g_str_Parame = g_str_Parame & "1" & ") "
         End If

         r_dbl_MtoSol = 0
         r_dbl_ItfSol = 0
         r_dbl_MtoDol = 0
         r_dbl_ItfDol = 0

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If

         g_rst_Princi.MoveNext
      Loop

      g_rst_Princi.Close
      Set g_rst_Princi = Nothing

   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Pagos registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
End Sub

Private Sub fs_Genera_MovAcu(ByVal p_NomFil As String, ByVal p_FecIni As String, ByVal p_FecFin As String)
   Dim r_int_NumRes        As Integer
   Dim r_int_Contad        As Integer
   Dim r_str_NomRes        As String
   Dim r_str_FecIni_bas    As String
   Dim r_str_FecFin_bas    As String
   Dim ff_MonItf           As Double
   Dim ff_MonSol           As Double
   Dim r_str_AcuMto        As String
   
   'Para obtener nombres de Archivo
   For r_int_Contad = Len(p_NomFil) To 1 Step -1
      If Mid(p_NomFil, r_int_Contad, 1) = "\" Then
         Exit For
      End If
   Next r_int_Contad
   
   r_str_NomRes = "C:\SUNATPDT\0695\" & Mid(p_NomFil, 1, r_int_Contad) & "0695" & Mid(p_FecIni, 1, 4) & Mid(p_FecIni, 5, 2) & "20511904162" & ".mov"
   
   'Creando Archivo de Resumen
   r_int_NumRes = FreeFile
   
   Open r_str_NomRes For Output As r_int_NumRes
   
   'Ejecutando Query
   g_str_Parame = "SELECT * FROM CTB_CABITF WHERE"
   g_str_Parame = g_str_Parame & " CABITF_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   g_str_Parame = g_str_Parame & " AND CABITF_PERANO = " & ipp_PerAno.Text
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         r_str_FecIni_bas = Format(g_rst_Princi!CABITF_PERANO, "0000") & Format(g_rst_Princi!CABITF_PERMES, "00") & "01"
         r_str_FecFin_bas = Format(g_rst_Princi!CABITF_PERANO, "0000") & Format(g_rst_Princi!CABITF_PERMES, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
            
         If g_rst_Princi!CABITF_TIPDEC = 1 And g_rst_Princi!CABITF_ITFSOL <> 0 Then
            Print #r_int_NumRes, Format(IIf(Trim(g_rst_Princi!CABITF_TIPDOC) = 2, 4, Trim(g_rst_Princi!CABITF_TIPDOC)), "00") & "|" & Trim(g_rst_Princi!CABITF_NUMDOC) & "|" & IIf(Trim(g_rst_Princi!CABITF_TIPDOC) = 2, Trim(g_rst_Princi!CABITF_PAICLI), "") & "|" & "0" + Trim(g_rst_Princi!CABITF_TIPOPE) & "|" & _
                                 "0" + Trim(g_rst_Princi!CABITF_TIPMOV) & "|" & "0" + Trim(g_rst_Princi!CABITF_CODOPE) & "||" & Trim(Format(g_rst_Princi!CABITF_MTOSOL, "0.00")) & "|" & _
                                 Format(gf_NueImp_Numero(Trim(g_rst_Princi!CABITF_ITFSOL)), "0.00") & "|" & Trim(Format(g_rst_Princi!CABITF_ITFSOL, "0.00")) & "|" & "||||||"
         End If
            
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
End Sub

Private Sub fs_Genera_AsiErr(ByVal p_NomFil As String, ByVal p_FecIni As String, ByVal p_FecFin As String)
   Dim r_int_NumRes        As Integer
   Dim r_int_Contad        As Integer
   Dim r_str_NomRes        As String
   Dim r_str_FecIni_bas    As String
   Dim r_str_FecFin_bas    As String
      
   'Para obtener nombres de Archivo
   For r_int_Contad = Len(p_NomFil) To 1 Step -1
      If Mid(p_NomFil, r_int_Contad, 1) = "\" Then
         Exit For
      End If
   Next r_int_Contad

   r_str_NomRes = "C:\SUNATPDT\0695\" & Mid(p_NomFil, 1, r_int_Contad) & "0695" & Mid(p_FecIni, 1, 4) & Mid(p_FecIni, 5, 2) & "20511904162" & ".ext"
   
   'Creando Archivo de Resumen
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   'Ejecutando Query para obtener Créditos
   g_str_Parame = "SELECT * FROM CTB_CABITF WHERE"
   g_str_Parame = g_str_Parame & " CABITF_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   g_str_Parame = g_str_Parame & " AND CABITF_PERANO = " & ipp_PerAno.Text
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         r_str_FecIni_bas = Format(g_rst_Princi!CABITF_PERANO, "0000") & Format(g_rst_Princi!CABITF_PERMES, "00") & "01"
         r_str_FecFin_bas = Format(g_rst_Princi!CABITF_PERANO, "0000") & Format(g_rst_Princi!CABITF_PERMES, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
         
         If g_rst_Princi!CABITF_TIPDEC = 2 Then
            Print #r_int_NumRes, Format(IIf(Trim(g_rst_Princi!CABITF_TIPDOC) = 2, 4, Trim(g_rst_Princi!CABITF_TIPDOC)), "00") & "|" & Trim(g_rst_Princi!CABITF_NUMDOC) & "||" & Trim(g_rst_Princi!CABITF_PERDEV) & "|" & _
                                 "0" + Trim(g_rst_Princi!CABITF_TIPOPE) & "|" & "0" + Trim(g_rst_Princi!CABITF_TIPMOV) & "|" & _
                                 "0" + Trim(g_rst_Princi!CABITF_CODOPE) & "||" & _
                                 Trim(g_rst_Princi!CABITF_MODOPE) & "|" & Trim(Format(g_rst_Princi!CABITF_MTOSOL, "0.00")) & "|" & _
                                 Format(gf_NueImp_Numero(Trim(g_rst_Princi!CABITF_ITFSOL)), "0.00") & "|" & Trim(Format(g_rst_Princi!CABITF_ITFSOL, "0.00")) & "|" & "||||||"
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
End Sub

Private Function ff_BusFec(ByVal p_TipDoc As String, ByVal p_NumDoc As String) As String
   Dim r_str_Parame As String
      
   ff_BusFec = 0
   r_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   r_str_Parame = r_str_Parame & "DETITF_TIPDOC = " & p_TipDoc & " AND "
   r_str_Parame = r_str_Parame & "DETITF_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "DETITF_TIPDEC = 1 "
   r_str_Parame = r_str_Parame & "ORDER BY DETITF_PERANO DESC, DETITF_PERMES DESC, DETITF_FECMOV DESC"
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      ff_BusFec = Trim(g_rst_Listas!DETITF_PERANO) & Format(g_rst_Listas!DETITF_PERMES, "00")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_BusCof(ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_FecDes As String) As String
   Dim r_str_Parame As String
   Dim r_str_NumSol As String
   
   ff_BusCof = 0
   r_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   r_str_Parame = r_str_Parame & "HIPMAE_TDOCLI = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "HIPMAE_NDOCLI = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      r_str_NumSol = Trim(g_rst_Listas!HIPMAE_NUMSOL)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   r_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   r_str_Parame = r_str_Parame & "EVACOF_NUMSOL = '" & r_str_NumSol & "' "
   r_str_Parame = r_str_Parame & "ORDER BY EVACOF_FECDES DESC "
   'r_str_Parame = r_str_Parame & "EVACOF_FECDES <= " & p_FecDes & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      ff_BusCof = Trim(g_rst_Listas!EVACOF_FECDES)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'Metodo para obtener el tipo de cambio
'Private Function ff_TipCam_01(ByVal p_FecPag As Double) As Double
'   g_str_Parame = "SELECT TIPCAM_FECDIA, TIPCAM_COMPRA FROM OPE_TIPCAM WHERE "
'   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = " & 3 & " AND "
'   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA = '" & p_FecPag & "'"

'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'      Exit Function
'   End If
   
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'      Do While Not g_rst_Listas.EOF
'         If g_rst_Listas!TIPCAM_FECDIA = p_FecPag Then
'            ff_TipCam = g_rst_Listas!TIPCAM_COMPRA
'            g_rst_Listas.MoveNext
'         Else
'            ff_TipCam = 0
'            g_rst_Listas.MoveNext
'         End If
'      Loop
'   End If
   
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function

'Metodo para obtener el tipo de cambio
Private Function ff_TipCam(ByVal p_FecPag As String, ByVal p_CtbFin As String) As Double
   Dim r_str_FecPag As String
            
   If CDate(gf_FormatoFecha(CStr(p_FecPag))) > CDate(gf_FormatoFecha(CStr(p_CtbFin))) Then
      r_str_FecPag = CDate(gf_FormatoFecha(CStr(p_CtbFin)))
   Else
      r_str_FecPag = CDate(gf_FormatoFecha(CStr(p_FecPag)))
   End If
                        
   g_str_Parame = "SELECT FECHA, VTA_DOL_PROM, CMP_DOL_PROM FROM CALENDARIO WHERE "
   g_str_Parame = g_str_Parame & "FECHA = to_date ('" & r_str_FecPag & "','DD/MM/YYYY')"
   g_str_Parame = g_str_Parame & "ORDER BY FECHA DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         If Trim(g_rst_Listas!fecha) = Format(r_str_FecPag, "dd/mm/yyyy") Then
            ff_TipCam = g_rst_Listas!CMP_DOL_PROM
            g_rst_Listas.MoveNext
         Else
            ff_TipCam = 0
            g_rst_Listas.MoveNext
         End If
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_PerMes     As String
Dim r_str_TipMon     As String
Dim r_str_Nombre     As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CTB_DETITF "
   g_str_Parame = g_str_Parame & " WHERE DETITF_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "   AND DETITF_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   If r_int_TipRep = 6 Then
      g_str_Parame = g_str_Parame & "   AND DETITF_FECMOV <= " & Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "15" & " "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_FECMOV ASC"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "SUCURSAL"
      .Cells(1, 3) = "PERIODO"
      .Cells(1, 4) = "AÑO"
      .Cells(1, 5) = "F. MOVIMIENTO"
      .Cells(1, 6) = "DOC. IDENTIDAD"
      .Cells(1, 7) = "NOMBRE DEL CLIENTE"
      .Cells(1, 8) = "OPE. REFERENCIA"
      .Cells(1, 9) = "ENTIDAD FINANCIERA"
      .Cells(1, 10) = "NRO CUENTA DEPOSITO"
      .Cells(1, 11) = "TIPO COMPROBANTE"
      .Cells(1, 12) = "TIPO DE DECLARANTE"
      .Cells(1, 13) = "NRO COMPROBANTE"
      .Cells(1, 14) = "TIPO DE CAMBIO (COMPRA)"
      .Cells(1, 15) = "TIPO DE MONEDA"
      .Cells(1, 16) = "BASE IMPONIBLE MON. ORG."
      .Cells(1, 17) = "IMPUESTO MON. ORG."
      .Cells(1, 18) = "BASE IMPONIBLE S/."
      .Cells(1, 19) = "IMPUESTO S/."
      .Cells(1, 20) = "BASE IMPONIBLE US$."
      .Cells(1, 21) = "IMPUESTO US$."
      .Cells(1, 22) = "ING. MAN."
   
      .Range(.Cells(1, 1), .Cells(1, 22)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 22)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 4.57
      .Columns("B").ColumnWidth = 11.86
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 17.14
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 5.71
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17.43
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 15.14
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 44
      .Columns("H").ColumnWidth = 19
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 25.71
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 22.71
      .Columns("J").NumberFormat = "@"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 36
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 19.57
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 19.29
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 27
      .Columns("N").NumberFormat = "##0.000"
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 21.43
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 27.57
      .Columns("P").NumberFormat = "###,###,##0.00"
      .Columns("Q").ColumnWidth = 22.29
      .Columns("Q").NumberFormat = "###,###,##0.00"
      .Columns("R").ColumnWidth = 18.43
      .Columns("R").NumberFormat = "###,###,##0.00"
      .Columns("S").ColumnWidth = 14.29
      .Columns("S").NumberFormat = "###,###,##0.00"
      .Columns("T").ColumnWidth = 20.57
      .Columns("T").NumberFormat = "###,###,##0.00"
      .Columns("U").ColumnWidth = 15.86
      .Columns("U").NumberFormat = "###,###,##0.00"
      .Columns("V").ColumnWidth = 13
      .Columns("V").HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos
      r_str_PerMes = moddat_gf_Consulta_ParDes("033", CStr(g_rst_Princi!DETITF_PERMES))
      r_str_TipMon = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!DETITF_TIPMON))
         
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "PRINCIPAL"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_str_PerMes
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Format(ipp_PerAno.Text, "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!DETITF_FECMOV)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(g_rst_Princi!DETITF_TIPDOC) & "-" & Trim(g_rst_Princi!DETITF_NUMDOC)
      
      r_str_Nombre = ""
      If Len(Trim(g_rst_Princi!DETITF_NUMDOC)) > 8 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = ""
      Else
         r_str_Nombre = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!DETITF_TIPDOC), Trim(g_rst_Princi!DETITF_NUMDOC))
         If Len(Trim(r_str_Nombre)) = 0 Then
            r_str_Nombre = fs_BuscarCliente_PlanAhorro(CStr(g_rst_Princi!DETITF_TIPDOC), Trim(g_rst_Princi!DETITF_NUMDOC))
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = r_str_Nombre
      End If
      
      If Len(Trim(g_rst_Princi!DETITF_NUMDOC)) > 8 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = ""
      Else
         If Len(Trim(g_rst_Princi!DETITF_OPEREF)) > 14 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Mid(Trim(g_rst_Princi!DETITF_OPEREF), 1, 4) & "-" & Mid(Trim(g_rst_Princi!DETITF_OPEREF), 5, 8) & "-" & Mid(Trim(g_rst_Princi!DETITF_OPEREF), 13, 3)
         ElseIf Len(Trim(g_rst_Princi!DETITF_OPEREF)) > 10 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CStr(gf_Formato_NumSol(Trim(g_rst_Princi!DETITF_OPEREF)))
         Else
            If Not IsNull(g_rst_Princi!DETITF_OPEREF) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CStr(gf_Formato_NumOpe(Trim(g_rst_Princi!DETITF_OPEREF)))
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = ""
            End If
         End If
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!DETITF_CODBAN)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!DETITF_NUMCTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!DETITF_TIPMOV)
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "DECLARANTE"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "EXTORNO"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Trim(g_rst_Princi!DETITF_NROCOM))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = CDbl(g_rst_Princi!DETITF_TIPCAM)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = r_str_TipMon
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CDbl(Format(g_rst_Princi!DETITF_MTOORG, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CDbl(Format(g_rst_Princi!DETITF_MTOORG, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CDbl(Format(g_rst_Princi!DETITF_ITFORG, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CDbl(Format(g_rst_Princi!DETITF_ITFORG, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDbl(Format(g_rst_Princi!DETITF_MTOSOL, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDbl(Format(g_rst_Princi!DETITF_MTOSOL, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CDbl(Format(g_rst_Princi!DETITF_ITFSOL, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CDbl(Format(g_rst_Princi!DETITF_ITFSOL, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = CDbl(Format(g_rst_Princi!DETITF_MTODOL, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = CDbl(Format(g_rst_Princi!DETITF_MTODOL, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_TIPDEC = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDbl(Format(g_rst_Princi!DETITF_ITFDOL, "###,###,##0.00"))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDbl(Format(g_rst_Princi!DETITF_ITFDOL, "###,###,##0.00")) * -1
      End If
      
      If g_rst_Princi!DETITF_MANUAL = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = "SI"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = "NO"
      End If
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22)).Font.Name = "Arial"
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22)).Font.Size = 8
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

'Metodo para obtener Movimientos Acumulados en la Tabla CTB_CABITF
Private Function ff_SumMon(ByVal p_TipDoc As Double, ByVal p_NumDoc As Double, ByVal p_TipDec As Integer, Optional ByRef ff_Mtosol As Double, Optional ByRef ff_itfsol As Double, Optional ByRef ff_mtodol As Double, Optional ByRef ff_itfdol As Double, Optional ByRef p_PerMes As Integer, Optional ByRef p_PerAno As Integer) As Double
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & p_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & p_NumDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & p_TipDec & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERMES = " & p_PerMes & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO = " & p_PerAno & " AND "
   g_str_Parame = g_str_Parame & "DETITF_ITFSOL <> 0 "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_TIPDOC, DETITF_NUMDOC ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         
         ff_Mtosol = ff_Mtosol + g_rst_Listas!DETITF_MTOSOL
         ff_itfsol = ff_itfsol + g_rst_Listas!DETITF_ITFSOL
         ff_mtodol = ff_mtodol + g_rst_Listas!DETITF_MTODOL
         ff_itfdol = ff_itfdol + g_rst_Listas!DETITF_ITFDOL
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_BusPer(ByVal p_FecIni As String, ByVal p_FecFin As String) As Integer
   ff_BusPer = 0
   
   'COMPARACION CON TABLE CTB_DETITF
   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV >= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV <= " & p_FecFin & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERMES <= " & Mid(p_FecIni, 5, 2) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO <= " & Left(p_FecIni, 4) & " "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_TIPDOC, DETITF_NUMDOC ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
                                   
         ff_BusPer = g_rst_Listas!TOTREG
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub ff_PriItf(ByVal p_FecIni As String, ByVal p_FecFin As String)
   r_int_Contad = 6
   ReDim r_arr_PriItf(0)
   
   'COMPARACION CON TABLE CTB_DETITF
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV >= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "DETITF_FECMOV <= " & p_FecFin & " "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_TIPDOC, DETITF_NUMDOC ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
      
         ReDim Preserve r_arr_PriItf(UBound(r_arr_PriItf) + IIf(UBound(r_arr_PriItf) = 0, 6, 7))
         'r_arr_PriItf(r_int_Contad + 0) = Trim(g_rst_Listas!DETITF_PERMES)
         'r_arr_PriItf(r_int_Contad + 1) = Trim(g_rst_Listas!DETITF_PERANO)
         'r_arr_PriItf(r_int_Contad + 2) = Trim(g_rst_Listas!DETITF_TIPDOC)
         'r_arr_PriItf(r_int_Contad + 3) = Trim(g_rst_Listas!DETITF_NUMDOC)
         'r_arr_PriItf(r_int_Contad + 4) = Trim(g_rst_Listas!DETITF_TIPDEC)
         'r_arr_PriItf(r_int_Contad + 5) = Trim(g_rst_Listas!DETITF_FECMOV)
         'r_arr_PriItf(r_int_Contad + 6) = Trim(g_rst_Listas!DETITF_TIPMOV)
         
         r_arr_PriItf(UBound(r_arr_PriItf) - 6) = Trim(g_rst_Listas!DETITF_PERMES)
         r_arr_PriItf(UBound(r_arr_PriItf) - 5) = Trim(g_rst_Listas!DETITF_PERANO)
         r_arr_PriItf(UBound(r_arr_PriItf) - 4) = Trim(g_rst_Listas!DETITF_TIPDOC)
         r_arr_PriItf(UBound(r_arr_PriItf) - 3) = Trim(g_rst_Listas!DETITF_NUMDOC)
         r_arr_PriItf(UBound(r_arr_PriItf) - 2) = Trim(g_rst_Listas!DETITF_TIPDEC)
         r_arr_PriItf(UBound(r_arr_PriItf) - 1) = Trim(g_rst_Listas!DETITF_FECMOV)
         r_arr_PriItf(UBound(r_arr_PriItf) - 0) = Trim(g_rst_Listas!DETITF_TIPMOV)
         
'         r_arr_PriItf(r_int_Contad + 7) = g_rst_Listas!DETITF_NROCOM
'         r_arr_PriItf(r_int_Contad + 8) = g_rst_Listas!DETITF_TIPCOD
'         r_arr_PriItf(r_int_Contad + 9) = g_rst_Listas!DETITF_OPEREF
'         r_arr_PriItf(r_int_Contad + 10) = g_rst_Listas!DETITF_TIPMON
'         r_arr_PriItf(r_int_Contad + 11) = g_rst_Listas!DETITF_MTOORG
'         r_arr_PriItf(r_int_Contad + 12) = g_rst_Listas!DETITF_ITFORG
'         r_arr_PriItf(r_int_Contad + 13) = g_rst_Listas!DETITF_ITFPOR
'         r_arr_PriItf(r_int_Contad + 14) = g_rst_Listas!DETITF_MTOSOL
'         r_arr_PriItf(r_int_Contad + 15) = g_rst_Listas!DETITF_ITFSOL
'         r_arr_PriItf(r_int_Contad + 16) = g_rst_Listas!DETITF_MTODOL
'         r_arr_PriItf(r_int_Contad + 17) = g_rst_Listas!DETITF_ITFDOL
'         r_arr_PriItf(r_int_Contad + 18) = g_rst_Listas!DETITF_TIPCAM
'         r_arr_PriItf(r_int_Contad + 19) = g_rst_Listas!DETITF_CODBAN
'         r_arr_PriItf(r_int_Contad + 20) = g_rst_Listas!DETITF_NUMCTA
         
         'r_int_Contad = r_int_Contad + 1
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Function ff_BusITF(ByVal p_PerAno As String, ByVal p_PerMes As String) As Integer
   ff_BusITF = 0
   
   'COMPARACION CON TABLE CTB_DETITF
   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM CTB_CABITF WHERE "
   g_str_Parame = g_str_Parame & "CABITF_PERANO = " & p_PerAno & " AND "
   g_str_Parame = g_str_Parame & "CABITF_PERMES = " & p_PerMes & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      ff_BusITF = g_rst_Listas!TOTREG
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub cmd_BusOpe_Click()
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
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If cmd_Grabar.Enabled = False Then
      Call gs_Activa_1(False)
      Call gs_Activa_2(True)
   End If
      
   grd_LisOpe.Redraw = False
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'BUSCANDO ITF
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO = '" & Trim(ipp_PerAno.Text) & "' AND DETITF_MANUAL = 1 AND "
   g_str_Parame = g_str_Parame & "DETITF_CODEMP = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "DETITF_CODSUC = '" & l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_PERANO DESC, DETITF_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_PERANO) & "-" & Format(Trim(g_rst_Princi!DETITF_PERMES), "00")
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_TIPDOC) & "-" & Trim(g_rst_Princi!DETITF_NUMDOC)
         
         grd_LisOpe.Col = 2
         grd_LisOpe.Text = IIf(Trim(g_rst_Princi!DETITF_TIPDEC) = 1, "DECLARANTE", "EXTORNO")
         
         grd_LisOpe.Col = 3
         grd_LisOpe.Text = gf_FormatoFecha(g_rst_Princi!DETITF_FECMOV)
         
         grd_LisOpe.Col = 4
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_TIPMOV)
         
         grd_LisOpe.Col = 5
         grd_LisOpe.Text = Format(g_rst_Princi!DETITF_MTOSOL, "###,###,###,##0.00")
         
         grd_LisOpe.Col = 6
         grd_LisOpe.Text = Format(g_rst_Princi!DETITF_ITFSOL, "###,###,###,##0.00")
         
         grd_LisOpe.Col = 7
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_TIPDEC)
                  
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   grd_LisOpe.Redraw = True
   If grd_LisOpe.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisOpe)
   Else
      MsgBox "No se encontraron registros manuales del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub ff_Buscar()
   Dim r_int_Contad As Integer
         
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_PERMES = " & Mid(modsec_g_str_Period, 6, 2) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO = " & Mid(modsec_g_str_Period, 1, 4) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & moddat_g_str_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & moddat_g_str_NumDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & moddat_g_int_TipCli & " AND "
   g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & modsec_g_dbl_MtoSol & " AND "
   g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & modsec_g_dbl_ITFSol & "  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      'cmb_TipDoc.ListIndex = g_rst_Listas!DETITF_TIPDOC
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Listas!DETITF_TIPDOC)
      txt_NumDoc.Text = g_rst_Listas!DETITF_NUMDOC
      txt_NumCom.Text = g_rst_Listas!DETITF_NROCOM
      'cmb_TipDec.ListIndex = g_rst_Listas!DETITF_TIPDEC
      Call gs_BuscarCombo_Item(cmb_TipDec, g_rst_Listas!DETITF_TIPDEC)
      'cmb_TipOpe.ListIndex = g_rst_Listas!DETITF_TIPCOD
      Call gs_BuscarCombo_Item(cmb_TipOpe, g_rst_Listas!DETITF_TIPCOD)
      'cmb_TipMov.ListIndex = g_rst_Listas!DETITF_TIPMOV
      
      For r_int_Contad = 1 To cmb_TipMov.ListCount Step 1
         cmb_TipMov.ListIndex = r_int_Contad
         If cmb_TipMov.Text = Trim(g_rst_Listas!DETITF_TIPMOV) Then
            Exit For
         End If
      Next
      
      'Call gs_BuscarCombo_Item(cmb_TipMov, g_rst_Listas!DETITF_TIPMOV)
      ipp_FecDep.Text = Mid(g_rst_Listas!DETITF_FECMOV, 7, 2) & "/" & Mid(g_rst_Listas!DETITF_FECMOV, 5, 2) & "/" & Mid(g_rst_Listas!DETITF_FECMOV, 1, 4)
      ipp_MtoSol.Text = g_rst_Listas!DETITF_MTOSOL
      ipp_MtoItf.Text = g_rst_Listas!DETITF_ITFSOL
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Function ff_Porcen(ByVal p_FecIni As String, ByVal p_FecFin As String) As Double
   ff_Porcen = 0
   g_str_Parame = "SELECT * FROM OPE_TABITF WHERE "
   g_str_Parame = g_str_Parame & "TABITF_FECINI <= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "TABITF_FECFIN >= " & p_FecFin & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_Porcen = g_rst_Listas!TABITF_PORCEN
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function fs_BuscarCliente_PlanAhorro(ByVal p_TdoCli As String, ByVal p_ndocli As String) As String
   fs_BuscarCliente_PlanAhorro = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(AHOCLI_APEPAT) ||' '|| TRIM(AHOCLI_APEMAT) ||' '|| TRIM(AHOCLI_NOMBRE) AS NOMBRE"
   g_str_Parame = g_str_Parame & "  FROM CRE_AHOCLI "
   g_str_Parame = g_str_Parame & " WHERE AHOCLI_TIPDOC = " & p_TdoCli & " "
   g_str_Parame = g_str_Parame & "   AND AHOCLI_NUMDOC = '" & p_ndocli & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      fs_BuscarCliente_PlanAhorro = Trim(g_rst_Listas!NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_BuscarNumero(ByVal p_TdoCli As String, ByVal p_ndocli As String, ByVal p_TipPag As Integer, ByVal p_TipMov As Integer) As String
   ff_BuscarNumero = ""
   
   If p_TipPag = 3 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT AHOMAE_NUMERO "
      g_str_Parame = g_str_Parame & "  FROM CRE_AHOMAE "
      g_str_Parame = g_str_Parame & " WHERE AHOMAE_TIPDOC = " & p_TdoCli & " "
      g_str_Parame = g_str_Parame & "   AND AHOMAE_NUMDOC = '" & p_ndocli & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!AHOMAE_NUMERO)
      End If
   
   ElseIf p_TipPag = 2 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPMAE_NUMOPE "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCLI = " & p_TdoCli & " "
      g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCLI = '" & p_ndocli & "' "
      g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 9)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!HIPMAE_NUMOPE)
      End If
      
   ElseIf p_TipPag = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT SOLMAE_NUMERO "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
      g_str_Parame = g_str_Parame & " WHERE SOLMAE_TITTDO = " & p_TdoCli & " "
      g_str_Parame = g_str_Parame & "   AND SOLMAE_TITNDO = " & p_ndocli & "  "
      If Left(p_TipMov, 2) <> 21 Then
         g_str_Parame = g_str_Parame & "AND SOLMAE_SITUAC = 1 "
      Else
         g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!SOLMAE_NUMERO)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub cmb_TipDec_Click()
   Call gs_SetFocus(cmb_TipOpe)
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumCom.MaxLength = 6
         Case Else:  txt_NumCom.MaxLength = 6
      End Select
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
         txt_NumCom.Enabled = False
      Else
         txt_NumCom.Enabled = True
      End If
      
      Call gs_SetFocus(txt_NumDoc)
      txt_NumCom.Text = ""
      txt_NumDoc.Text = ""
   End If
End Sub

Private Sub cmb_TipMov_Click()
   Call gs_SetFocus(ipp_FecDep)
End Sub

Private Sub cmb_TipOpe_Click()
   Call gs_SetFocus(cmb_TipMov)
End Sub

Private Sub ipp_FecDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoSol)
   End If
End Sub

Private Sub ipp_MtoItf_GotFocus()
   Call gs_SelecTodo(ipp_MtoItf)
End Sub

Private Sub ipp_MtoItf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_MtoSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoItf)
   End If
End Sub

Private Sub ipp_MtoSol_GotFocus()
   Call gs_SelecTodo(ipp_MtoSol)
End Sub

Private Sub txt_NumCom_GotFocus()
   Call gs_SelecTodo(txt_NumCom)
End Sub

Private Sub txt_NumCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
         Call gs_SetFocus(cmb_TipDec)
      Else
         Call gs_SetFocus(txt_NumCom)
      End If
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 5:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 6:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

