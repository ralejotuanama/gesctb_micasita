VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_TrnCta_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15600
   Icon            =   "GesCtb_frm_206.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   15600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7395
      Left            =   -90
      TabIndex        =   10
      Top             =   0
      Width           =   15735
      _Version        =   65536
      _ExtentX        =   27755
      _ExtentY        =   13044
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
         Left            =   150
         TabIndex        =   11
         Top             =   60
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27340
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
            TabIndex        =   12
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registros de Movimientos Bancarios"
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
            Picture         =   "GesCtb_frm_206.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   150
         TabIndex        =   13
         Top             =   780
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27340
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_206.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_206.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_206.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Reversa 
            Height          =   585
            Left            =   2490
            Picture         =   "GesCtb_frm_206.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14880
            Picture         =   "GesCtb_frm_206.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3120
            Picture         =   "GesCtb_frm_206.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consulta 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_206.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5070
         Left            =   150
         TabIndex        =   14
         Top             =   2190
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27340
         _ExtentY        =   8943
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
            Height          =   4620
            Left            =   45
            TabIndex        =   15
            Top             =   390
            Width           =   15390
            _ExtentX        =   27146
            _ExtentY        =   8149
            _Version        =   393216
            Rows            =   24
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   285
            Left            =   1240
            TabIndex        =   16
            Top             =   90
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
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
         Begin Threed.SSPanel pnl_Glosa 
            Height          =   285
            Left            =   2340
            TabIndex        =   17
            Top             =   90
            Width           =   1120
            _Version        =   65536
            _ExtentX        =   1976
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Cambio"
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
            Left            =   13290
            TabIndex        =   18
            Top             =   90
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Referencia"
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
            Left            =   6405
            TabIndex        =   19
            Top             =   90
            Width           =   2200
            _Version        =   65536
            _ExtentX        =   3881
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Movimiento"
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
            Left            =   60
            TabIndex        =   20
            Top             =   90
            Width           =   1220
            _Version        =   65536
            _ExtentX        =   2152
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Mov"
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
            Left            =   11910
            TabIndex        =   21
            Top             =   90
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   3420
            TabIndex        =   25
            Top             =   90
            Width           =   3010
            _Version        =   65536
            _ExtentX        =   5309
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   8600
            TabIndex        =   26
            Top             =   90
            Width           =   2430
            _Version        =   65536
            _ExtentX        =   4286
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Cuenta"
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
            Left            =   11010
            TabIndex        =   27
            Top             =   90
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1605
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
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   675
         Left            =   150
         TabIndex        =   22
         Top             =   1470
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27340
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
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6510
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   7890
            TabIndex        =   2
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
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   24
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Movimiento:"
            Height          =   195
            Left            =   5040
            TabIndex        =   23
            Top             =   240
            Width           =   1350
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_TrnCta_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Buscar_Click()
   Call fs_BuscarMov
   cmb_Banco.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
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

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Banco.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   Call gs_SetFocus(cmb_Banco)
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1170 'NRO MOVIMIENTO
   grd_Listad.ColWidth(1) = 1115 'FECHA
   grd_Listad.ColWidth(2) = 1100 'TIPO CAMBIO
   grd_Listad.ColWidth(3) = 2950 'BANCO
   grd_Listad.ColWidth(4) = 2190 'TIPO MOVIMIENTO
   grd_Listad.ColWidth(5) = 2420 'NRO CUENTA
   grd_Listad.ColWidth(6) = 880  'MONEDA
   grd_Listad.ColWidth(7) = 1390 'IMPORTE
   grd_Listad.ColWidth(8) = 1770 'NRO REFERENCIA

   grd_Listad.ColAlignment(0) = flexAlignCenterCenter 'NRO MOVIMIENTO
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter 'FECHA
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter 'TIPO CAMBIO
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter   'BANCO
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter   'TIPO MOVIMIENTO
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter   'NRO CUENTA
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter   'MONEDA
   grd_Listad.ColAlignment(7) = flexAlignRightCenter  'IMPORTE
   grd_Listad.ColAlignment(8) = flexAlignLeftCenter   'NRO REFERENCIA
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Banco, 1, "516")
   cmb_Banco.AddItem Trim$("<<TODOS>>")
   cmb_Banco.ItemData(cmb_Banco.NewIndex) = CLng(0)
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   Call gs_LimpiaGrid(grd_Listad)
   
   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   r_str_CadAux = ""
   
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   Call gs_BuscarCombo_Item(cmb_Banco, CLng(0))
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Public Sub fs_BuscarMov()

   Call gs_LimpiaGrid(grd_Listad)
   
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT LPAD(A.MOVCTA_NUMMOV,8,'0') AS MOVCTA_NUMMOV, TRIM(B.PARDES_DESCRI) AS TIPOMOVIMIENTO,  "
'   g_str_Parame = g_str_Parame & "        A.MOVCTA_FECOPE, A.MOVCTA_TIPCAM, A.MOVCTA_NUMREF, A.MOVCTA_IMPORT, A.MOVCTA_TIPMOV,  "
'   g_str_Parame = g_str_Parame & "        CASE A.MOVCTA_TIPMOV  "
'   g_str_Parame = g_str_Parame & "          WHEN 1 THEN (SELECT TRIM(Z.PARDES_DESCRI) ||' - '||TRIM(X.MOVCTA_NUMCTA)  "
'   g_str_Parame = g_str_Parame & "                         FROM CNTBL_MOVCTA X  "
'   g_str_Parame = g_str_Parame & "                        INNER JOIN MNT_PARDES Z ON Z.PARDES_CODGRP = 516 AND Z.PARDES_CODITE = X.MOVCTA_CODBAN  "
'   g_str_Parame = g_str_Parame & "                        WHERE X.MOVCTA_NUMMOV = A.MOVCTA_NUMMOV AND X.MOVCTA_TIPMOV = 2)  "
'   g_str_Parame = g_str_Parame & "          WHEN 2 THEN (SELECT TRIM(Z.PARDES_DESCRI) ||' - '||TRIM(X.MOVCTA_NUMCTA)  "
'   g_str_Parame = g_str_Parame & "                         FROM CNTBL_MOVCTA X  "
'   g_str_Parame = g_str_Parame & "                        INNER JOIN MNT_PARDES Z ON Z.PARDES_CODGRP = 516 AND Z.PARDES_CODITE = X.MOVCTA_CODBAN  "
'   g_str_Parame = g_str_Parame & "                        WHERE X.MOVCTA_NUMMOV = A.MOVCTA_NUMMOV AND X.MOVCTA_TIPMOV = 1)  "
'   g_str_Parame = g_str_Parame & "          ELSE 'GLOSA DESCRIPTIVA'  "
'   g_str_Parame = g_str_Parame & "        END AS GLOSA_TEMP  "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_MOVCTA A  "
'   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 133 AND A.MOVCTA_TIPMOV = B.PARDES_CODITE  "
'   g_str_Parame = g_str_Parame & "  WHERE MOVCTA_CODBAN = '" & Trim(moddat_g_str_Codigo) & "'  "
'   'g_str_Parame = g_str_Parame & "    AND MOVCTA_NUMCTA = '" & Trim(moddat_g_str_CodIte) & "'  "
'   g_str_Parame = g_str_Parame & "    AND MOVCTA_SITUAC = 1  "
'   g_str_Parame = g_str_Parame & "  ORDER BY MOVCTA_FECOPE ASC, MOVCTA_NUMMOV ASC "

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT LPAD(A.MOVCTA_NUMMOV,8,'0') AS MOVCTA_NUMMOV, TRIM(B.PARDES_DESCRI) AS TIPOMOVIMIENTO,  "
   g_str_Parame = g_str_Parame & "        A.MOVCTA_FECOPE, A.MOVCTA_TIPCAM, A.MOVCTA_NUMREF, A.MOVCTA_IMPORT, A.MOVCTA_TIPMOV,  "
   g_str_Parame = g_str_Parame & "        A.MOVCTA_CODBAN, TRIM(C.PARDES_DESCRI) AS NOM_BANCO, A.MOVCTA_NUMCTA, TRIM(E.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "        Case A.MOVCTA_TIPMOV  "
   g_str_Parame = g_str_Parame & "          WHEN 1 THEN (SELECT TRIM(Z.PARDES_DESCRI) ||' - '||TRIM(X.MOVCTA_NUMCTA)  "
   g_str_Parame = g_str_Parame & "                         FROM CNTBL_MOVCTA X  "
   g_str_Parame = g_str_Parame & "                        INNER JOIN MNT_PARDES Z ON Z.PARDES_CODGRP = 516 AND Z.PARDES_CODITE = X.MOVCTA_CODBAN  "
   g_str_Parame = g_str_Parame & "                        WHERE X.MOVCTA_NUMMOV = A.MOVCTA_NUMMOV AND X.MOVCTA_TIPMOV = 2)  "
   g_str_Parame = g_str_Parame & "          WHEN 2 THEN (SELECT TRIM(Z.PARDES_DESCRI) ||' - '||TRIM(X.MOVCTA_NUMCTA)  "
   g_str_Parame = g_str_Parame & "                         FROM CNTBL_MOVCTA X  "
   g_str_Parame = g_str_Parame & "                        INNER JOIN MNT_PARDES Z ON Z.PARDES_CODGRP = 516 AND Z.PARDES_CODITE = X.MOVCTA_CODBAN  "
   g_str_Parame = g_str_Parame & "                        WHERE X.MOVCTA_NUMMOV = A.MOVCTA_NUMMOV AND X.MOVCTA_TIPMOV = 1)  "
   g_str_Parame = g_str_Parame & "          Else 'GLOSA DESCRIPTIVA'  "
   g_str_Parame = g_str_Parame & "        END As GLOSA_TEMP  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MOVCTA A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 133 AND A.MOVCTA_TIPMOV = B.PARDES_CODITE  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 516 AND C.PARDES_CODITE = A.MOVCTA_CODBAN  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_CTABAN D ON D.CTABAN_CODBAN = A.MOVCTA_CODBAN AND D.CTABAN_NUMCTA = A.MOVCTA_NUMCTA  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 204 AND E.PARDES_CODITE = D.CTABAN_TIPMON  "
   g_str_Parame = g_str_Parame & "  WHERE MOVCTA_FECOPE BETWEEN " & Format(ipp_FecIni.Text, "yyyymmdd") & " AND " & Format(ipp_FecFin.Text, "yyyymmdd")
   If cmb_Banco.ItemData(cmb_Banco.ListIndex) <> 0 Then
      g_str_Parame = g_str_Parame & "  AND MOVCTA_CODBAN = '" & Format(cmb_Banco.ItemData(cmb_Banco.ListIndex), "000000") & "'  "
   End If
   g_str_Parame = g_str_Parame & "    AND MOVCTA_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "  ORDER BY MOVCTA_FECOPE ASC, MOVCTA_NUMMOV ASC, MOVCTA_TIPMOV ASC  "
   
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(CStr(g_rst_Princi!MOVCTA_NUMMOV & ""))
   
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!MOVCTA_FECOPE)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!MOVCTA_TIPCAM, "###,###,##0.000000")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(CStr(g_rst_Princi!NOM_BANCO & ""))
   
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(CStr(g_rst_Princi!TIPOMOVIMIENTO & ""))
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(CStr(g_rst_Princi!MOVCTA_NUMCTA & ""))
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(CStr(g_rst_Princi!Moneda & ""))
      
      grd_Listad.Col = 7
      'grd_Listad.Text = IIf(g_rst_Princi!MOVCTA_TIPMOV = 1, "-", "+") & Format(g_rst_Princi!MOVCTA_IMPORT, "###,###,##0.00")
      grd_Listad.Text = IIf(g_rst_Princi!MOVCTA_TIPMOV = 1, "-", "") & Format(g_rst_Princi!MOVCTA_IMPORT, "###,###,##0.00") & " "
            
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(CStr(g_rst_Princi!MOVCTA_NUMREF & ""))

      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ctb_TrnCta_01.Show 1
End Sub

Private Sub cmd_Consulta_Click()
   moddat_g_str_NumOpe = "" 'NRO MOVIMIENTO

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = CLng(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 0
   frm_Ctb_TrnCta_01.Show 1
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Reversa_Click()
   moddat_g_str_NumOpe = "" 'NRO MOVIMIENTO

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = CLng(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   frm_Ctb_TrnCta_01.Show 1
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REGISTRO DE MOVIMIENTOS BANCARIOS"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True

      .Cells(5, 2) = "NRO MOV"
      .Cells(5, 3) = "FECHA"
      .Cells(5, 4) = "TIPO CAMBIO"
      .Cells(5, 5) = "BANCO"
      .Cells(5, 6) = "TIPO MOVIMIENTO"
      .Cells(5, 7) = "NRO CUENTA"
      .Cells(5, 8) = "MONEDA"
      .Cells(5, 9) = "IMPORTE"
      .Cells(5, 10) = "NRO REFERENCIA"
         
      .Cells(3, 2) = "Banco: " & Trim(cmb_Banco.Text)
      .Cells(3, 6) = "Fecha Inicio: " & ipp_FecIni.Text
      .Cells(3, 8) = "Fecha Fin: " & ipp_FecFin.Text
      
      .Range(.Cells(5, 2), .Cells(5, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'NRO MOV
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 12 'FECHA
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 13 'TIPO CAMBIO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 29 'BANCO
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 24 'TIPO MOVIMIENTO
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 27 'NRO CUENTA
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 8.14 'MONEDA
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 15 'IMPORTE -- .NumberFormat = "###,###,##0.00"
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 17 'NRO REFERENCIA
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      
      r_int_NumFil = 4
      For r_int_Contad = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_Contad, 0) 'NRO MOV
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_Listad.TextMatrix(r_int_Contad, 1) 'TIPO MOVIMIENTO
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_Listad.TextMatrix(r_int_Contad, 2) 'FECHA
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_Listad.TextMatrix(r_int_Contad, 3) 'TIPO CAMBIO
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_Listad.TextMatrix(r_int_Contad, 4) 'NRO REFERENCIA
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_Listad.TextMatrix(r_int_Contad, 5) 'BANCO
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_Listad.TextMatrix(r_int_Contad, 6) 'NRO CUENTA
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_Listad.TextMatrix(r_int_Contad, 7) 'MONEDA
          .Cells(r_int_NumFil + 2, 10) = grd_Listad.TextMatrix(r_int_Contad, 8) 'IMPORTE
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(3, 2), .Cells(3, 10)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(1, 1), .Cells(r_int_NumFil + 2, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_NumFil + 2, 10)).Font.Size = 11
   End With
      
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

