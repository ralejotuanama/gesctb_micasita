VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RegDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15045
   Icon            =   "GesCtb_frm_179.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7630
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15090
      _Version        =   65536
      _ExtentX        =   26617
      _ExtentY        =   13458
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CommandButton cmd_Seguimiento 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_179.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Seguimiento por Instancias"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "GesCtb_frm_179.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_179.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_179.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Buscar Operaciones"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SegSol 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_179.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Detalle de la Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_179.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   1440
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   1349
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
         Begin VB.CheckBox chk_Estado 
            Caption         =   "Todas las Instancias"
            Height          =   315
            Left            =   1110
            TabIndex        =   7
            Top             =   420
            Width           =   2685
         End
         Begin VB.ComboBox cmb_Estado 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   90
            Width           =   3975
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6780
            TabIndex        =   8
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8160
            TabIndex        =   9
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Registro:"
            Height          =   195
            Left            =   5550
            TabIndex        =   26
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Instancias:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            TabIndex        =   15
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Operaciones a Desembolsar al Promotor"
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
            Left            =   80
            Picture         =   "GesCtb_frm_179.frx":1940
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5300
         Left            =   30
         TabIndex        =   16
         Top             =   2250
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   9349
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4650
            Left            =   80
            TabIndex        =   17
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8202
            _Version        =   393216
            Rows            =   45
            Cols            =   22
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   13140
            TabIndex        =   18
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2611
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cons. Hipotecario"
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   3180
            TabIndex        =   19
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   4395
            TabIndex        =   20
            Top             =   60
            Width           =   3240
            _Version        =   65536
            _ExtentX        =   5715
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Tit_SitAct 
            Height          =   285
            Left            =   9750
            TabIndex        =   21
            Top             =   60
            Width           =   2020
            _Version        =   65536
            _ExtentX        =   3563
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia"
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
         Begin Threed.SSPanel pnl_Tit_FecReg 
            Height          =   285
            Left            =   7620
            TabIndex        =   22
            Top             =   60
            Width           =   1100
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Registro"
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
            Left            =   90
            TabIndex        =   23
            Top             =   60
            Width           =   3120
            _Version        =   65536
            _ExtentX        =   5503
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
         Begin Threed.SSPanel pnl_Tit_TotDsm 
            Height          =   285
            Left            =   11760
            TabIndex        =   24
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Desembolso"
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
         Begin Threed.SSPanel pnl_tit_Fec_aprob 
            Height          =   285
            Left            =   8700
            TabIndex        =   25
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Aprob Ope."
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
Attribute VB_Name = "frm_RegDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_Estado_Click()
   Call Estado_Ctrl
   If chk_Estado.Value = 1 Then
      cmb_Estado.ListIndex = -1
      cmb_Estado.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   ElseIf chk_Estado.Value = 0 Then
      cmb_Estado.Enabled = True
      Call gs_SetFocus(cmb_Estado)
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   moddat_g_str_NumOpe = ""
   moddat_g_str_NumSol = ""
   moddat_g_int_TipDoc = 0
   moddat_g_str_NumDoc = ""
   moddat_g_str_NomPrd = ""
   moddat_g_str_CodIte = ""
   moddat_g_int_CodIns = 0
   moddat_g_str_CodPrd = ""
   moddat_g_str_CodSub = ""
   moddat_g_int_TipMon = 0
   moddat_g_str_NomCli = ""
   moddat_g_str_FecRec = ""
   moddat_g_str_FecHip = ""
   moddat_g_str_Situac = ""
   
   cmb_Estado.ListIndex = 0
   chk_Estado.Value = 0
   
   ipp_FecIni.Text = DateAdd("M", -3, date)
   ipp_FecFin.Text = date
   
   Call gs_LimpiaGrid(grd_Listad)
   Call Estado_Ctrl
   
   Call gs_SetFocus(cmb_Estado)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   Dim r_str_CodIns As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'grd_Listad.Col = 14
   'r_str_CodIns =   Trim(grd_Listad.Text)
   r_str_CodIns = grd_Listad.TextMatrix(grd_Listad.Row, 14)
   'If r_str_CodIns <> "4" Then
   '   MsgBox "No se puede editar este registro.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   
   '4 = Operacion Aprobada
   If (r_str_CodIns = "4") Then
      moddat_g_int_TipRep = 1
      Else
      moddat_g_int_TipRep = 2
   End If

   'numero de operacion
   grd_Listad.Col = 12
   moddat_g_str_NumOpe = Trim(grd_Listad.Text)
   'numero de solicitud
   grd_Listad.Col = 9
   moddat_g_str_NumSol = Trim(grd_Listad.Text)
   'tipo de documento
   grd_Listad.Col = 10
   moddat_g_int_TipDoc = Trim(grd_Listad.Text)
   'numero de documento
   grd_Listad.Col = 11
   moddat_g_str_NumDoc = Trim(grd_Listad.Text)
   'Nombre producto
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = Trim(grd_Listad.Text)
   'Codigo de Item
   grd_Listad.Col = 14
   moddat_g_str_CodIte = Trim(grd_Listad.Text)
   'codigo de area
   grd_Listad.Col = 13
   moddat_g_int_CodIns = Trim(grd_Listad.Text)
   'HIPMAE_CODPRD
   grd_Listad.Col = 15
   moddat_g_str_CodPrd = Trim(grd_Listad.Text)
   'HIPMAE_CODSUB
   grd_Listad.Col = 16
   moddat_g_str_CodSub = Trim(grd_Listad.Text)
   'HIPMAE_MONEDA
   grd_Listad.Col = 17
   moddat_g_int_TipMon = Trim(grd_Listad.Text)
   'Nombre del Cliente
   grd_Listad.Col = 2
   moddat_g_str_NomCli = Trim(grd_Listad.Text)
   'Fecha Registro
   grd_Listad.Col = 18
   moddat_g_str_FecRec = Trim(grd_Listad.Text)
   'Hora Registro
   grd_Listad.Col = 19
   moddat_g_str_FecHip = Trim(grd_Listad.Text)
   'Estado Actual
   grd_Listad.Col = 5
   moddat_g_str_Situac = Trim(Trim(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_RegDes_02.Show 1
   
   Call Estado_Ctrl
End Sub

Private Sub cmd_Seguimiento_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   '1 = Guardar = 1, Eliminar = 2, (Aprobar = 3, Rechazar = 4)
   moddat_g_int_TipRep = 2

   'numero de operacion
   grd_Listad.Col = 12
   moddat_g_str_NumOpe = Trim(grd_Listad.Text)
   'numero de solicitud
   grd_Listad.Col = 9
   moddat_g_str_NumSol = Trim(grd_Listad.Text)
   'tipo de documento
   grd_Listad.Col = 10
   moddat_g_int_TipDoc = Trim(grd_Listad.Text)
   'numero de documento
   grd_Listad.Col = 11
   moddat_g_str_NumDoc = Trim(grd_Listad.Text)
   'Nombre producto
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = Trim(grd_Listad.Text)
   'Codigo de Item
   grd_Listad.Col = 14
   moddat_g_str_CodIte = Trim(grd_Listad.Text)
   '-------------------------------------------------------
   'HIPMAE_CODPRD
   grd_Listad.Col = 15
   moddat_g_str_CodPrd = Trim(grd_Listad.Text)
   'HIPMAE_CODSUB
   grd_Listad.Col = 16
   moddat_g_str_CodSub = Trim(grd_Listad.Text)
   'HIPMAE_MONEDA
   grd_Listad.Col = 17
   moddat_g_int_TipMon = Trim(grd_Listad.Text)
   'Nombre del Cliente
   grd_Listad.Col = 2
   moddat_g_str_NomCli = Trim(grd_Listad.Text)
   
   'Fecha Registro
   grd_Listad.Col = 18
   moddat_g_str_FecRec = Trim(grd_Listad.Text)
   'Hora Registro
   grd_Listad.Col = 19
   moddat_g_str_FecHip = Trim(grd_Listad.Text)
   'Estado Actual
   grd_Listad.Col = 5
   moddat_g_str_Situac = Trim(Trim(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_RegDes_03.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   chk_Estado.Value = 0
   Call chk_Estado_Click
   
   Call gs_SetFocus(cmd_Buscar)
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub Estado_Ctrl()
   If grd_Listad.Rows = 0 Then
      cmd_Seguimiento.Enabled = False
      cmd_ExpExc.Enabled = False
      cmd_SegSol.Enabled = False
      cmb_Estado.Enabled = True
      chk_Estado.Enabled = True
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = True
   Else
      cmd_Seguimiento.Enabled = True
      cmd_ExpExc.Enabled = True
      cmd_SegSol.Enabled = True
      cmb_Estado.Enabled = False
      chk_Estado.Enabled = False
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
   End If
End Sub

Private Sub fs_Inicia()
   cmb_Estado.Clear

   g_str_Parame = " SELECT to_number(PARDES_CODITE)||' - '||trim(PARDES_DESCRI) as glosa, PARDES_CODITE as codigo "
   g_str_Parame = g_str_Parame & " FROM MNT_PARDES WHERE PARDES_CODGRP = '374' "
   g_str_Parame = g_str_Parame & " and PARDES_CODITE <> '000000' AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " and PARDES_CODITE IN (3,5) "
   g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC "

   'g_str_Parame = " SELECT to_number(PARDES_CODITE)||' - '||trim(PARDES_DESCRI) as glosa, PARDES_CODITE as codigo "
   'g_str_Parame = g_str_Parame & " FROM MNT_PARDES WHERE PARDES_CODGRP = '375' "
   'g_str_Parame = g_str_Parame & " and PARDES_CODITE <> '000000' AND PARDES_SITUAC = 1 "
   'g_str_Parame = g_str_Parame & " and PARDES_CODITE = 4 "
   'g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC "
   
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
      cmb_Estado.AddItem Trim$(g_rst_Genera!GLOSA)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   '------------------------------------------------------
   grd_Listad.ColWidth(0) = 3070 'PRODUCTO
   grd_Listad.ColWidth(1) = 1250 'OPERACION
   grd_Listad.ColWidth(2) = 3200 'CLIENTE
   grd_Listad.ColWidth(3) = 1080 'FECHA_REGISTRO
   grd_Listad.ColWidth(4) = 1070 'FECHA_APROB_OPERA
   
   grd_Listad.ColWidth(5) = 2000 'ESTADO
   grd_Listad.ColWidth(6) = 1380 'IMPORTE
   grd_Listad.ColWidth(7) = 0 'OPERACION
   grd_Listad.ColWidth(8) = 0 'FECHA_REGISTRO
   grd_Listad.ColWidth(9) = 0 'HIPMAE_NUMSOL
   grd_Listad.ColWidth(10) = 0 'INSTANCIA
   grd_Listad.ColWidth(11) = 0 'HIPMAE_TDOCLI
   grd_Listad.ColWidth(12) = 0 'OPERACIONES
   grd_Listad.ColWidth(13) = 0 'DESCAB_CODAREA
   grd_Listad.ColWidth(14) = 0 'DESCAB_CODEST
   grd_Listad.ColWidth(15) = 0 'HIPMAE_CODPRD
   grd_Listad.ColWidth(16) = 0 'HIPMAE_CODSUB
   grd_Listad.ColWidth(17) = 0 'HIPMAE_MONEDA
   grd_Listad.ColWidth(18) = 0 'DESCAB_FECREG
   grd_Listad.ColWidth(19) = 0 'DESCAB_HORREG
   grd_Listad.ColWidth(20) = 1480 'CONSEJEROS
   grd_Listad.ColWidth(21) = 0 'fecha aprob ope.
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(20) = flexAlignLeftCenter
End Sub

Private Sub cmb_Estado_Click()
   Call Estado_Ctrl
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Estado_Click
   End If
End Sub

Public Sub cmd_Buscar_Click()
   If chk_Estado.Value = 0 Then
      If cmb_Estado.ListIndex = -1 Then
         MsgBox "Debe seleccionar una instancia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Estado)
         Exit Sub
      End If
   End If
      
   Screen.MousePointer = 11
   Call fs_Buscar_Creditos
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_SelChange()
'   Dim r_str_CodIns As String
'   If grd_Listad.Rows = 0 Then
'      Exit Sub
'   End If
'   grd_Listad.Col = 13
'   r_str_CodIns = Trim(grd_Listad.Text)
'
'   If grd_Listad.Rows > 2 Then
'      grd_Listad.RowSel = grd_Listad.Row
'   End If
   
   'Call gs_RefrescaGrid(grd_Listad)
End Sub

Public Sub fs_Buscar_Creditos()
Dim r_int_FlgIn1     As Integer
Dim r_int_FlgIn2     As Integer

   g_str_Parame = "  "
   g_str_Parame = g_str_Parame & "SELECT TRIM(E.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       TRIM(A.DESCAB_NUMOPE) AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       A.DESCAB_FECREG AS FECHA_REGISTRO, A.DESCAB_FECREG, A.DESCAB_HORREG, "
   g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS INSTANCIA, descab_codarea, descab_codest, DESCAB_CODEST, "
   g_str_Parame = g_str_Parame & "       TRIM(F.PARDES_DESCRI) AS ESTADO, B.HIPMAE_CODPRD, B.HIPMAE_CODSUB, B.HIPMAE_MONEDA, "
   g_str_Parame = g_str_Parame & "       B.hipmae_numsol, B.HIPMAE_TDOCLI, B.HIPMAE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_CONHIP) AS CONSEJERO, "
   g_str_Parame = g_str_Parame & "       (SELECT SUM(DT.DESDAT_IMPORT) FROM CRE_DESPRODAT DT WHERE DT.DESDAT_NUMOPE = A.DESCAB_NUMOPE  "
   g_str_Parame = g_str_Parame & "                                                             AND DT.DESDAT_FECREG = A.DESCAB_FECREG  "
   g_str_Parame = g_str_Parame & "                                                             AND DT.DESDAT_HORREG = A.DESCAB_HORREG) AS TOTAL,  "
   g_str_Parame = g_str_Parame & "       (SELECT DESDET_FECFIN FROM CRE_DESPRODET WHERE DESDET_NUMOPE = A.DESCAB_NUMOPE "
   g_str_Parame = g_str_Parame & "                                                 AND DESDET_FECREG = A.DESCAB_FECREG AND DESDET_HORREG = A.DESCAB_HORREG "
   g_str_Parame = g_str_Parame & "                                                 AND DESDET_CODAREA = 2 AND DESDET_CODEST = 4) AS FECHA_APROB_OPE "
   g_str_Parame = g_str_Parame & "  FROM CRE_DESPROCAB A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.DESCAB_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 374 AND D.PARDES_CODITE = A.DESCAB_CODAREA "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC E ON E.PRODUC_CODIGO = B.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 375 AND F.PARDES_CODITE = A.DESCAB_CODEST "
   
   If (cmb_Estado.ListIndex = -1 And chk_Estado.Value = False) Then
       Exit Sub
   End If
   
   If chk_Estado.Value = False Then
      g_str_Parame = g_str_Parame & " WHERE A.DESCAB_CODAREA = " & Trim(Mid(cmb_Estado.Text, 1, InStr(1, cmb_Estado.Text, "-") - 1)) & ""
      If Trim(Mid(cmb_Estado.Text, 1, InStr(1, cmb_Estado.Text, "-") - 1)) = 5 Then
         g_str_Parame = g_str_Parame & "   AND A.DESCAB_CODEST = 8"
      Else
         g_str_Parame = g_str_Parame & "   AND A.DESCAB_CODEST = 4"
      End If
   Else
      g_str_Parame = g_str_Parame & " WHERE A.DESCAB_CODEST IN (4,8)"
   End If
   g_str_Parame = g_str_Parame & " AND A.DESCAB_FECREG BETWEEN " & Format(ipp_FecIni.Value, "yyyymmdd") & " AND " & Format(ipp_FecFin.Value, "yyyymmdd")

   g_str_Parame = g_str_Parame & " ORDER BY NOMBRE_CLIENTE "
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Operaciones para esa selección.", vbExclamation, modgen_g_str_NomPlt
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
      grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!OPERACION))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
      
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FECHA_REGISTRO))
            
      If Trim(g_rst_Princi!FECHA_APROB_OPE & "") <> "" Then
         grd_Listad.Col = 4
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FECHA_APROB_OPE))
      End If
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!INSTANCIA)
      
      grd_Listad.Col = 6 'adicion nueva
      grd_Listad.Text = gf_FormatoNumero(g_rst_Princi!total, 12, 2) & " "
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!FECHA_REGISTRO)
      '-----------------
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMSOL)
      
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI)
      
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NDOCLI)
      
      grd_Listad.Col = 12
      grd_Listad.Text = Trim(g_rst_Princi!OPERACION)
      
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_CODAREA)
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_CODEST)
      '--------------------------------------------------------
      grd_Listad.Col = 15
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODPRD)
      grd_Listad.Col = 16
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODSUB)
      grd_Listad.Col = 17
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_MONEDA)
      
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_FECREG)
      grd_Listad.Col = 19
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_HORREG)
      
      grd_Listad.Col = 20
      grd_Listad.Text = CStr(g_rst_Princi!CONSEJERO)
      
      grd_Listad.Col = 21
      grd_Listad.Text = g_rst_Princi!FECHA_APROB_OPE
            
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
   
   Call Estado_Ctrl
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_nrofil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_nrofil, 1) = "PRODUCTO":              .Columns("A").ColumnWidth = 40
      .Cells(r_int_nrofil, 2) = "NRO OPERACION":         .Columns("B").ColumnWidth = 16
      .Cells(r_int_nrofil, 3) = "APELLIDOS Y NOMBRES":   .Columns("C").ColumnWidth = 40
      .Cells(r_int_nrofil, 4) = "FECHA REGISTRO":        .Columns("D").ColumnWidth = 16
      .Cells(r_int_nrofil, 5) = "FECHA APROBACION OPERACIONES": .Columns("E").ColumnWidth = 32
      .Cells(r_int_nrofil, 6) = "INSTANCIA":             .Columns("F").ColumnWidth = 25
      .Cells(r_int_nrofil, 7) = "TOTAL DESEMBOLSO":      .Columns("G").ColumnWidth = 18
      .Cells(r_int_nrofil, 8) = "CONSEJERO":             .Columns("H").ColumnWidth = 17
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_nrofil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_nrofil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 4) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 5) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_nrofil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 20)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_old()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_nrofil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_nrofil, 1) = "PRODUCTO":              .Columns("A").ColumnWidth = 40
      .Cells(r_int_nrofil, 2) = "NRO OPERACION":         .Columns("B").ColumnWidth = 16
      .Cells(r_int_nrofil, 3) = "APELLIDOS Y NOMBRES":   .Columns("C").ColumnWidth = 40
      .Cells(r_int_nrofil, 4) = "FECHA REGISTRO":        .Columns("D").ColumnWidth = 15
      .Cells(r_int_nrofil, 5) = "INSTANCIA ACTUAL":      .Columns("E").ColumnWidth = 25
      .Cells(r_int_nrofil, 6) = "TOTAL DESEMBOLSO":      .Columns("F").ColumnWidth = 18
      .Cells(r_int_nrofil, 7) = "CONSEJERO":             .Columns("G").ColumnWidth = 17
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 7)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_nrofil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_nrofil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 19)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
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

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 19, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 19, "C-")
   End If
End Sub

Private Sub pnl_tit_Fec_aprob_Click()
   If Len(Trim(pnl_tit_Fec_aprob.Tag)) = 0 Or pnl_tit_Fec_aprob.Tag = "D" Then
      pnl_tit_Fec_aprob.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 21, "N")
   Else
      pnl_tit_Fec_aprob.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 21, "N-")
   End If
End Sub

Private Sub pnl_Tit_FecReg_Click()
   If Len(Trim(pnl_Tit_FecReg.Tag)) = 0 Or pnl_Tit_FecReg.Tag = "D" Then
      pnl_Tit_FecReg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecReg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If

End Sub

Private Sub pnl_Tit_SitAct_Click()
   If Len(Trim(pnl_Tit_SitAct.Tag)) = 0 Or pnl_Tit_SitAct.Tag = "D" Then
      pnl_Tit_SitAct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_SitAct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub
 
Private Sub pnl_Tit_TotDsm_Click()
   If Len(Trim(pnl_Tit_TotDsm.Tag)) = 0 Or pnl_Tit_TotDsm.Tag = "D" Then
      pnl_Tit_TotDsm.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_Tit_TotDsm.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

