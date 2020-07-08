VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbPpg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13695
   Icon            =   "GesCtb_frm_912.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   285
      Left            =   11850
      TabIndex        =   25
      Top             =   2370
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
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
         Left            =   1030
         TabIndex        =   9
         Top             =   20
         Width           =   255
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   735
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   1296
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Left            =   660
         TabIndex        =   12
         Top             =   60
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Proceso"
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   315
         Left            =   660
         TabIndex        =   13
         Top             =   360
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Contabilización de Prepagos"
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
         Picture         =   "GesCtb_frm_912.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   675
      Left            =   0
      TabIndex        =   14
      Top             =   780
      Width           =   13680
      _Version        =   65536
      _ExtentX        =   24130
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
      Begin VB.CommandButton cmd_Detalle 
         Height          =   585
         Left            =   1890
         Picture         =   "GesCtb_frm_912.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ver Detalle"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Proces 
         Enabled         =   0   'False
         Height          =   585
         Left            =   2505
         Picture         =   "GesCtb_frm_912.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Generar asientos automaticos"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   1275
         Picture         =   "GesCtb_frm_912.frx":0A62
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar a Excel"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   13065
         Picture         =   "GesCtb_frm_912.frx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Limpia 
         Height          =   585
         Left            =   660
         Picture         =   "GesCtb_frm_912.frx":11AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpiar Datos"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Buscar 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_912.frx":14B8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Registros"
         Top             =   60
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   810
      Left            =   0
      TabIndex        =   15
      Top             =   1455
      Width           =   13680
      _Version        =   65536
      _ExtentX        =   24130
      _ExtentY        =   1429
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
      Begin VB.ComboBox cmb_TipPre 
         Height          =   315
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   2265
      End
      Begin VB.ComboBox cmb_PerMes 
         Height          =   315
         ItemData        =   "GesCtb_frm_912.frx":17C2
         Left            =   1080
         List            =   "GesCtb_frm_912.frx":17C4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2265
      End
      Begin EditLib.fpLongInteger ipp_PerAno 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   405
         Width           =   825
         _Version        =   196608
         _ExtentX        =   1455
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
      Begin VB.Label Label2 
         Caption         =   "Tipo de prepago:"
         Height          =   315
         Left            =   4920
         TabIndex        =   26
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   315
         Left            =   135
         TabIndex        =   17
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "Mes:"
         Height          =   315
         Left            =   135
         TabIndex        =   16
         Top             =   90
         Width           =   1365
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   5415
      Left            =   0
      TabIndex        =   18
      Top             =   2280
      Width           =   13680
      _Version        =   65536
      _ExtentX        =   24130
      _ExtentY        =   9551
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
      Begin Threed.SSPanel pnl_Tit_NumOpe 
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
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
      Begin Threed.SSPanel pnl_Tit_TipPpg 
         Height          =   285
         Left            =   5295
         TabIndex        =   20
         Top             =   90
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Tipo de Prepago"
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
      Begin Threed.SSPanel pnl_Tit_Import 
         Height          =   285
         Left            =   9990
         TabIndex        =   21
         Top             =   90
         Width           =   1860
         _Version        =   65536
         _ExtentX        =   3281
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
      Begin Threed.SSPanel pnl_Tit_FecPro 
         Height          =   285
         Left            =   8805
         TabIndex        =   22
         Top             =   90
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "F. Proceso"
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
         Left            =   1320
         TabIndex        =   23
         Top             =   90
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Nombre Cliente"
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
      Begin Threed.SSPanel pnl_Tit_FecPpg 
         Height          =   285
         Left            =   7620
         TabIndex        =   24
         Top             =   90
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "F. Prepago"
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
         Height          =   4995
         Left            =   80
         TabIndex        =   10
         Top             =   360
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   8811
         _Version        =   393216
         Rows            =   15
         Cols            =   30
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPpg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer
Dim l_rst_Prepagos      As ADODB.Recordset

Private Sub cmd_Buscar_Click()
   If Trim(cmb_PerMes.Text) = "" Then
      MsgBox "Debe seleccionar el tipo de mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If cmb_TipPre.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPre)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
   
   If (grd_Listad.Rows = 0) Then
       Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_Limpia_Click()
   grd_Listad.Rows = 0
   cmb_PerMes.ListIndex = -1
   cmb_TipPre.ListIndex = 0
   ipp_PerAno.Text = Year(date)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PerMes)
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

Private Sub cmd_Detalle_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'numero de operacion
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = Trim(grd_Listad.Text)
   'DNI
   grd_Listad.Col = 1
   moddat_g_str_NumDoc = Trim(grd_Listad.Text)
   'nombre del cliente
   grd_Listad.Col = 2
   moddat_g_str_NomCli = Trim(grd_Listad.Text)
   'fecha de prepago
   grd_Listad.Col = 11
   moddat_g_str_FecIng = Trim(grd_Listad.Text)
      
   Call gs_RefrescaGrid(grd_Listad)
      
   frm_Pro_CtbPpg_02.Show 1
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida seleccion
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionados registros para generar asientos automaticos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro de generar los asientos contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiento
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_BuscaPeriodo
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   
   cmb_TipPre.Clear
   cmb_TipPre.AddItem "- TODOS -"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 0
   cmb_TipPre.AddItem "PREPAGO PARCIAL"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 1
   cmb_TipPre.AddItem "PREPAGO TOTAL"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 2
   cmb_TipPre.ListIndex = 0
      
   grd_Listad.ColWidth(0) = 1222    'numero de oparcion
   grd_Listad.ColWidth(1) = 0       'tipo de documento
   grd_Listad.ColWidth(2) = 3972    'nombre del cliente
   grd_Listad.ColWidth(3) = 2322    'tipo de prepago
   grd_Listad.ColWidth(4) = 1182    'fecha de prepago
   grd_Listad.ColWidth(5) = 1182    'fecha de roceso
   grd_Listad.ColWidth(6) = 1857    'importe del prepago
   grd_Listad.ColWidth(7) = 1330    'Seleccionar
   grd_Listad.ColWidth(8) = 0       'tipo de prepago
   grd_Listad.ColWidth(9) = 0       'tipo de prepago (total o parcial)
   grd_Listad.ColWidth(10) = 0      'tipo de prepago (monto o tiempo)
   grd_Listad.ColWidth(11) = 0      'numero de operacion
   grd_Listad.ColWidth(12) = 0      'fecha de prepago
   grd_Listad.ColWidth(13) = 0      'fecha de proceso
   grd_Listad.ColWidth(14) = 0      '
   grd_Listad.ColWidth(15) = 0      'PPGCAB_MTODEP
   grd_Listad.ColWidth(16) = 0      'PPGCAB_MTOTOT
   grd_Listad.ColWidth(17) = 0      'PPGCAB_MTOPOR
   grd_Listad.ColWidth(18) = 0      'PPGCAB_INTCAL_TNC
   grd_Listad.ColWidth(19) = 0      'PPGCAB_INTCAL_TC
   grd_Listad.ColWidth(20) = 0      'PPGCAB_SEGDES
   grd_Listad.ColWidth(21) = 0      'PPGCAB_SEGINM
   grd_Listad.ColWidth(22) = 0      'PPGCAB_PBPPER
   grd_Listad.ColWidth(23) = 0      'PPGCAB_PBPINT
   grd_Listad.ColWidth(24) = 0      'PPGCAB_MTOITF
   grd_Listad.ColWidth(25) = 0      'PPGCAB_MTOAPL
   grd_Listad.ColWidth(26) = 0      'PPGCAB_TIPPPG
   grd_Listad.ColWidth(27) = 0      'HIPMAE_MONEDA
   grd_Listad.ColWidth(28) = 0      'PPGCAB_SLDACT_TNC
   grd_Listad.ColWidth(29) = 0      'PPGCAB_SLDACT_TC
   '*******************************************************
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.Rows = 0
End Sub

Private Sub fs_BuscaPeriodo()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PERMES_CODANO, PERMES_CODMES "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND PERMES_TIPPER = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se pudo determinar el período actual.", vbInformation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   l_int_PerMes = g_rst_Princi!PERMES_CODMES
   l_int_PerAno = g_rst_Princi!PERMES_CODANO
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   cmb_TipPre.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   cmd_Proces.Enabled = Not p_Activa
   cmd_Detalle.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
Dim swtFechaInicio      As Date
Dim swtFechaFin         As Date
Dim swtMes              As String
Dim swtNextFecha        As Date
Dim swtObect            As Object
Dim r_str_FchIni        As String
Dim r_str_FchFin        As String
   
   Call gs_LimpiaGrid(grd_Listad)
   
   swtMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   swtFechaInicio = Trim("01-" & swtMes & "-" & CStr(ipp_PerAno.Value))
   swtNextFecha = DateAdd("m", 1, swtFechaInicio)
   swtFechaFin = DateAdd("d", -1, swtNextFecha)
   r_str_FchIni = Left(swtFechaInicio, 2) & Mid(swtFechaInicio, 4, 2) & Right(swtFechaInicio, 4)
   r_str_FchFin = Left(swtFechaFin, 2) & Mid(swtFechaFin, 4, 2) & Right(swtFechaFin, 4)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PP.PPGCAB_NUMOPE, CH.HIPMAE_TDOCLI, CH.HIPMAE_NDOCLI, CH.HIPMAE_MONEDA, PP.PPGCAB_TIPPPG, "
   g_str_Parame = g_str_Parame & "       PP.PPGCAB_FECPRO, PP.PPGCAB_FECPPG, PP.PPGCAB_MTODEP, PP.PPGCAB_MTOTOT, PP.PPGCAB_TIPPPGPAR, "
   g_str_Parame = g_str_Parame & "       TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "       PP.PPGCAB_MTOAPL, PP.PPGCAB_SEGDES, PPGCAB_SEGINM, PPGCAB_INTCAL_TNC, "
   g_str_Parame = g_str_Parame & "       PPGCAB_INTCAL_TC, PPGCAB_PBPPER, PPGCAB_PBPINT, PPGCAB_MTOPOR, PPGCAB_MTOITF, "
   g_str_Parame = g_str_Parame & "       PPGCAB_SLDACT_TNC, PPGCAB_SLDACT_TC "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB PP  "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE   "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI  "
   g_str_Parame = g_str_Parame & " WHERE PP.PPGCAB_NUMOPE > 0 "
   g_str_Parame = g_str_Parame & "   AND (PP.PPGCAB_FECPPG >= " & Format(swtFechaInicio, "yyyymmdd") & ") "
   g_str_Parame = g_str_Parame & "   AND (PP.PPGCAB_FECPPG <= " & Format(swtFechaFin, "yyyymmdd") & ") "
   g_str_Parame = g_str_Parame & "   AND ((PP.PPGCAB_FLGCNT = 0) OR (PP.PPGCAB_FLGCNT IS NULL)) "
   If cmb_TipPre.ListIndex <> 0 Then
      g_str_Parame = g_str_Parame & "   AND (PP.PPGCAB_TIPPPG = " & cmb_TipPre.ItemData(cmb_TipPre.ListIndex) & " ) "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY PP.PPGCAB_NUMOPE ASC, PP.PPGCAB_FECPPG ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Prepagos, 3) Then
      Exit Sub
   End If

   If l_rst_Prepagos.BOF And l_rst_Prepagos.EOF Then
      l_rst_Prepagos.Close
      Set l_rst_Prepagos = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   l_rst_Prepagos.MoveFirst
   Do While Not l_rst_Prepagos.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'numero operacion (formateado)
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(Trim(l_rst_Prepagos!PPGCAB_NUMOPE & ""))
      
      'tipo de documento
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(l_rst_Prepagos!HIPMAE_TDOCLI) & "-" & Trim(l_rst_Prepagos!HIPMAE_NDOCLI & "")
      
      'nombre del cliente
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(l_rst_Prepagos!CLIENTE & "")
      
      'tipo de prepago
      grd_Listad.Col = 3
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         If l_rst_Prepagos!PPGCAB_TIPPPGPAR = 1 Then
            grd_Listad.Text = "PARCIAL - RED MONTO"
         Else
            grd_Listad.Text = "PARCIAL - RED PLAZO"
         End If
      Else
         grd_Listad.Text = "TOTAL"
      End If
      
      'fecha del prepago (formateado)
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPPG))
      
      'fecha de proceso (formateado)
      grd_Listad.Col = 5
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPRO))
      
      'importe del prepago (formateado)
      grd_Listad.Col = 6
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         If l_rst_Prepagos!HIPMAE_MONEDA = 1 Then
            grd_Listad.Text = "S/.   " & Format(l_rst_Prepagos!PPGCAB_MTODEP, "###,###,###,##0.00")
         Else
            grd_Listad.Text = "US$   " & Format(l_rst_Prepagos!PPGCAB_MTODEP, "###,###,###,##0.00")
         End If
      Else
         If l_rst_Prepagos!HIPMAE_MONEDA = 1 Then
            grd_Listad.Text = "S/.   " & Format(l_rst_Prepagos!PPGCAB_MTOTOT, "###,###,###,##0.00")
         Else
            grd_Listad.Text = "US$   " & Format(l_rst_Prepagos!PPGCAB_MTOTOT, "###,###,###,##0.00")
         End If
      End If
      
      'Tipo de prepago (parcial o total)
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPG)
      
      'Tipo de prepago Parcial (monto o tiempo)
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPGPAR & "")
      
      'numero de operacion
      grd_Listad.Col = 10
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_NUMOPE & "")
      
      'fecha de prepago
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPPG)
      
      'fecha de proceso
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPRO)
      
      'importe del prepago
      grd_Listad.Col = 13
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOTOT
      
      '***********************************************************************************
      'PPGCAB_MTODEP
      grd_Listad.Col = 15
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTODEP
      'PPGCAB_MTOTOT
      grd_Listad.Col = 16
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOTOT
      'PPGCAB_MTOPOR
      grd_Listad.Col = 17
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOPOR
      'PPGCAB_INTCAL_TNC
      grd_Listad.Col = 18
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_INTCAL_TNC
      'PPGCAB_INTCAL_TC
      grd_Listad.Col = 19
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_INTCAL_TC
      'PPGCAB_SEGDES
      grd_Listad.Col = 20
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_SEGDES
      'PPGCAB_SEGINM
      grd_Listad.Col = 21
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_SEGINM
      'PPGCAB_PBPPER
      grd_Listad.Col = 22
      grd_Listad.Text = IIf(IsNull(l_rst_Prepagos!PPGCAB_PBPPER) = True, 0, l_rst_Prepagos!PPGCAB_PBPPER)
      'PPGCAB_PBPINT
      grd_Listad.Col = 23
      grd_Listad.Text = IIf(IsNull(l_rst_Prepagos!PPGCAB_PBPINT) = True, 0, l_rst_Prepagos!PPGCAB_PBPINT)
      'PPGCAB_MTOITF
      grd_Listad.Col = 24
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOITF
      'PPGCAB_MTOAPL
      grd_Listad.Col = 25
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOAPL
      'PPGCAB_TIPPPG
      grd_Listad.Col = 26
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_TIPPPG
      'hipmae_moneda
      grd_Listad.Col = 27
      grd_Listad.Text = l_rst_Prepagos!HIPMAE_MONEDA
      'PPGCAB_SLDACT_TNC
      grd_Listad.Col = 28
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_SLDACT_TNC
      'PPGCAB_SLDACT_TC
      grd_Listad.Col = 29
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_SLDACT_TC
      '***********************************************************************************
      
      l_rst_Prepagos.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GeneraAsiento()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_dbl_TipSbs        As Double
Dim r_str_NroOpe        As String
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
Dim r_str_NumOpe        As String
Dim r_str_CtaCtb        As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double

Dim r_str_Portes        As String
Dim r_str_IntTNC        As String
Dim r_str_IntTC         As String
Dim r_str_SegDesg       As String
Dim r_str_SegInm        As String
Dim r_str_CapPBP        As String
Dim r_str_IntPBP        As String
Dim r_str_ITF           As String
Dim r_str_MtoApl        As String
Dim r_str_MtoDep        As String
Dim r_str_Codigo        As String
Dim r_int_Colum         As Integer
Dim r_int_Total         As Integer
Dim r_str_DebHab        As String
Dim r_int_TipPgo        As Integer
Dim r_str_Glosa         As String
Dim swtCadena           As String
Dim swtImporte          As Double

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      If grd_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
         '*************************************************
         'GENERACION DE ASIENTOS CONTABLES DEL PREPAGO
         '*************************************************
         
         'Inicializa variables
         r_int_NumAsi = 0
         r_str_FecPrPgoC = grd_Listad.TextMatrix(r_int_Contad, 11)
         r_str_FecPrPgoL = grd_Listad.TextMatrix(r_int_Contad, 4)
         r_str_NroOpe = grd_Listad.TextMatrix(r_int_Contad, 10)
         r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, 2, Format(r_str_FecPrPgoL, "yyyymmdd"), 1)
         
         'Tipo de prepago (parcial = 1 o total = 2)
         r_int_TipPgo = grd_Listad.TextMatrix(r_int_Contad, 8)
         If r_int_TipPgo = 1 Then
            If grd_Listad.TextMatrix(r_int_Contad, 9) = 1 Then
               swtCadena = "PPG PARCIAL - RED MONTO"
            Else
               swtCadena = "PPG PARCIAL - RED PLAZO"
            End If
         Else
            swtCadena = "PPG TOTAL"
         End If
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
         r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
         
         'Insertar en CABECERA
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(swtCadena) & " - " & Trim(r_str_NroOpe), r_str_FecPrPgoL, "1")
         
         'Inicializa
         r_int_NumIte = 0
         r_str_Codigo = Left(r_str_NroOpe, 3)
         r_str_Portes = "521229010109"
         r_str_SegDesg = "251602010103"
         r_str_ITF = "251705010107"
              
         Select Case r_str_Codigo
            'CREDITO CRC-PBP
            Case "001"
               r_str_IntTNC = "512401042401"
               r_str_IntTC = "512401042401"
               r_str_SegInm = "252602010104"
               r_str_CapPBP = "142104240101"
               r_str_IntPBP = "512401042401"
               r_str_MtoApl = "142104240101"
               r_str_MtoDep = "112301060202"
               
            'CREDITO MICASITA DOLARES
            Case "002"
               r_str_IntTNC = "512401040601"
               r_str_IntTC = "512401040601"
               r_str_SegInm = "252602010104"
               r_str_CapPBP = "142104060101"
               r_str_IntPBP = "512401040601"
               r_str_MtoApl = "142104060101"
               r_str_MtoDep = "112301060202"
               r_str_SegDesg = "252602010103"
            
            'CREDITO CME
            Case "003"
               r_str_IntTNC = "511401042501"
               r_str_IntTC = "511401042501"
               r_str_SegInm = "251602010103"
               r_str_CapPBP = "141104250101"
               r_str_IntPBP = "511401042501"
               r_str_MtoApl = "141104250101"
               r_str_MtoDep = "111301060201"
            
            'CREDITO PROYECTO MIHOGAR
            Case "004"
                 r_str_IntTNC = "511401042301"
                 r_str_IntTC = "511401042301"
                 r_str_SegInm = "251602010104"
                 r_str_CapPBP = "141104230101"
                 r_str_IntPBP = "511401042301"
                 r_str_MtoApl = "141104230101"
                 r_str_MtoDep = "111301060201"
            
            'CREDITO FMV UNION ANDINA
            Case "009"
               r_str_IntTNC = "511401042305"
               r_str_IntTC = "511401042305"
               r_str_SegInm = "251602010104"
               r_str_CapPBP = "141104230103"
               r_str_IntPBP = "511401042305"
               r_str_MtoApl = "141104230103"
               r_str_MtoDep = "111301060201"
            
            'CREDITO MICASITA
            Case "011", "006"
               r_str_IntTNC = "511401042302"
               r_str_IntTC = "511401042302"
               r_str_SegInm = "251602010104"
               r_str_CapPBP = "141104060101"
               r_str_IntPBP = "511401042302"
               r_str_MtoApl = "141104060101"
               r_str_MtoDep = "111301060201"
            
            'CREDITO MIVIVIENDA
            Case "007", "010", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023", "024", "025"
               r_str_IntTNC = "511401042302"
               r_str_IntTC = "511401042302"
               r_str_SegInm = "251602010104"
               r_str_CapPBP = "141104230102"
               r_str_IntPBP = "511401042302"
               r_str_MtoApl = "141104230102"
               r_str_MtoDep = "111301060201"
         End Select
         
         r_int_Colum = 15
         r_int_Total = grd_Listad.Cols
         
         Do While (r_int_Colum < r_int_Total)
            '26 = Tipo de prepago(PPGCAB_TIPPPG), 27 = Moneda, 28 y 29 (Saldo)
            If (r_int_Colum <> 26 And r_int_Colum <> 27 And r_int_Colum <> 29) Then
               
               If ((grd_Listad.TextMatrix(r_int_Contad, r_int_Colum) <> Null) Or (grd_Listad.TextMatrix(r_int_Contad, r_int_Colum) > 0)) Then
                  If r_int_TipPgo = 1 Then 'PPGCAB_TIPPPG
                     r_str_Glosa = "PPG PAR - " & Trim(r_str_NroOpe) & " - "
                  Else
                     r_str_Glosa = "PPG TOT - " & Trim(r_str_NroOpe) & " - "
                  End If
                  r_str_DebHab = "H"
                  
                  Select Case r_int_Colum
                     Case 15, 16       'Monto Depositado - Debe - PPGCAB_MTODEP, PPGCAB_MTOTOT
                          r_str_CtaCtb = r_str_MtoDep
                          r_str_DebHab = "D"
                          r_str_Glosa = r_str_Glosa & "MTO DEP"
                     Case 25, 28       'Monto de Prepago a Aplicar - Haber - PPGCAB_MTOAPL, (saldo)
                          r_str_CtaCtb = r_str_MtoApl
                          r_str_Glosa = r_str_Glosa & "MTO APL"
                     Case 20           'Seguro Desgravamen - Haber - PPGCAB_SEGDES
                          r_str_CtaCtb = r_str_SegDesg
                          r_str_Glosa = r_str_Glosa & "SEG DES"
                     Case 21           'Seguro Inmueble - Haber - PPGCAB_SEGINM
                          r_str_CtaCtb = r_str_SegInm
                          r_str_Glosa = r_str_Glosa & "SEG INM"
                     Case 18           'Intereses TNC a la fecha - Haber - PPGCAB_INTCAL_TNC
                          r_str_CtaCtb = r_str_IntTNC
                          r_str_Glosa = r_str_Glosa & "INT TNC"
                     Case 19           'Intereses TC a la fecha - Haber - PPGCAB_INTCAL_TC
                          r_str_CtaCtb = r_str_IntTC
                          r_str_Glosa = r_str_Glosa & "INT TC"
                     Case 22           'Capital PBP - Haber - PPGCAB_PBPPER
                          r_str_CtaCtb = r_str_CapPBP
                          r_str_Glosa = r_str_Glosa & "CAP PBP"
                     Case 23           'Interes PBP - Haber - PPGCAB_PBPINT
                          r_str_CtaCtb = r_str_IntPBP
                          r_str_Glosa = r_str_Glosa & "INT PBP"
                     Case 17           'Portes - Haber - PPGCAB_MTOPOR
                          r_str_CtaCtb = r_str_Portes
                          r_str_Glosa = r_str_Glosa & "MTO PORTES"
                     Case 24           'ITF - Haber - PPGCAB_MTOITF
                          r_str_CtaCtb = r_str_ITF
                          r_str_Glosa = r_str_Glosa & "ITF"
                  End Select
                  
                  '28 Saldo Actual, 25 Monto de Prepago a Aplicar
                  If (r_int_Colum = 28 And r_int_TipPgo = 2) Then
                     swtImporte = CDbl(grd_Listad.TextMatrix(r_int_Contad, 28)) + CDbl(grd_Listad.TextMatrix(r_int_Contad, 29))
                  ElseIf (r_int_Colum = 28 And r_int_TipPgo = 1) Then
                     swtImporte = 0
                  ElseIf (r_int_Colum = 25 And r_int_TipPgo = 2) Then
                     swtImporte = 0
                  Else
                     swtImporte = CDbl(grd_Listad.TextMatrix(r_int_Contad, r_int_Colum))
                  End If
               
                  If (swtImporte > 0) Then
                     r_int_NumIte = r_int_NumIte + 1
                     If (grd_Listad.TextMatrix(r_int_Contad, 27) = 1) Then
                        'SOLES
                        r_dbl_MtoSol = Format(swtImporte, "###,###,##0.00") 'Importe soles
                        r_dbl_MtoDol = Format(CDbl(r_dbl_MtoSol / r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
                     Else
                        'DOLARES
                        r_dbl_MtoDol = Format(swtImporte, "###,###,##0.00") 'Importe dolares
                        r_dbl_MtoSol = Format(CDbl(r_dbl_MtoDol * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
                     End If
                     
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecPrPgoL), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                  End If
               End If
            End If
            r_int_Colum = r_int_Colum + 1
         Loop
         
         '****************************************************************************************************************
         Set g_rst_Princi = Nothing
         
         'Actualiza flag de contabilizacion para las operacion con el numero de carta
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE CRE_PPGCAB "
         g_str_Parame = g_str_Parame & "   SET PPGCAB_FLGCNT = 1 "
         g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE  = '" & Trim(r_str_NroOpe) & "' "
         g_str_Parame = g_str_Parame & "   AND PPGCAB_FECPPG  = " & r_str_FecPrPgoC & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros seleccionados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE ASIENTOS DEL PREPAGO"
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(5, 2) = "NRO. DE OPERACION"
      .Cells(5, 3) = "NOMBRE DEL CLIENTE"
      .Cells(5, 4) = "TIPO DE PREPAGO"
      .Cells(5, 5) = "FECHA DE PREPAGO"
      .Cells(5, 6) = "FECHA DE PROCESO"
      .Cells(5, 7) = "IMPORTE"
      
      .Range(.Cells(5, 2), .Cells(5, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 7)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 7)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 20
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 45
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 20
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 20
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 20
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 20
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 7)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 7)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0))   'N°Operacion
         .Cells(r_int_NumFil + 3, 3) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 2))         'Nombre Cliente
         .Cells(r_int_NumFil + 3, 4) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 3)) 'Tipo Prepago
         .Cells(r_int_NumFil + 3, 5) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 4) 'Fecha de prepago
         .Cells(r_int_NumFil + 3, 6) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 5) 'Fecha de proceso
         .Cells(r_int_NumFil + 3, 7) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 6)) 'Importe Prepago
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 7
      If grd_Listad.Text = "X" Then
          grd_Listad.Text = ""
      Else
           grd_Listad.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 7) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 7) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

