VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Con_CreHip_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   5250
   ClientTop       =   765
   ClientWidth     =   10440
   Icon            =   "GesCtb_frm_108.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10425
      _Version        =   65536
      _ExtentX        =   18389
      _ExtentY        =   17806
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   3615
         Left            =   30
         TabIndex        =   9
         Top             =   6420
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
         _ExtentY        =   6376
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
         Begin Threed.SSPanel pnl_Cuo_TotSal 
            Height          =   315
            Left            =   8610
            TabIndex        =   10
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cuo_TotPag 
            Height          =   315
            Left            =   7320
            TabIndex        =   11
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cuo_TotDeu 
            Height          =   315
            Left            =   6030
            TabIndex        =   12
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Cuotas 
            Height          =   2565
            Left            =   60
            TabIndex        =   1
            Top             =   600
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   4524
            _Version        =   393216
            Rows            =   11
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   330
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
            Left            =   840
            TabIndex        =   14
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencim."
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
            Left            =   6030
            TabIndex        =   15
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Cuota"
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
            Left            =   2130
            TabIndex        =   16
            Top             =   330
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D. Atraso"
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
            Left            =   7320
            TabIndex        =   17
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Pagado"
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
            Left            =   4740
            TabIndex        =   18
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ult. Pago"
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
            Left            =   8610
            TabIndex        =   19
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Deudor"
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   3120
            TabIndex        =   20
            Top             =   330
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
         Begin VB.Label Label12 
            Caption         =   "Resumen de Cuotas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ==> US$ "
            Height          =   255
            Left            =   4350
            TabIndex        =   21
            Top             =   3270
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
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
            Height          =   585
            Left            =   660
            TabIndex        =   24
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   9630
            Top             =   150
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_108.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4935
         Left            =   30
         TabIndex        =   25
         Top             =   1440
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
         _ExtentY        =   8705
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
            Height          =   4575
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   8070
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Datos del Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   30
         TabIndex        =   27
         Top             =   750
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
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
         Begin VB.CommandButton cmd_PosOtr 
            Height          =   585
            Left            =   3030
            Picture         =   "GesCtb_frm_108.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Consulta de Posición del Cliente en otras Entidades Financieras"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_108.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Consulta de Datos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9720
            Picture         =   "GesCtb_frm_108.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatHip 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_108.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consulta de Datos de la Hipoteca"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_108.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consulta de Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_108.frx":1940
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Consulta de Cronogramas de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_108.frx":1C4A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Consulta de Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   1740
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_TipGar     As Integer

Private Sub cmd_DatCli_Click()
   frm_Con_CreHip_09.Show 1
End Sub

Private Sub cmd_DatHip_Click()
   frm_Con_CreHip_06.Show 1
End Sub

Private Sub cmd_DatInm_Click()
   frm_Con_CreHip_08.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   frm_Con_CreHip_07.Show 1
End Sub

Private Sub cmd_PosOtr_Click()
   frm_Con_CreHip_10.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerPag_Click()
   frm_Con_CreHip_05.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmd_DatCli)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2850
   grd_Listad.ColWidth(1) = 7000
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Cuotas
   grd_Cuotas.ColWidth(0) = 750
   grd_Cuotas.ColWidth(1) = 1295
   grd_Cuotas.ColWidth(2) = 1005
   grd_Cuotas.ColWidth(3) = 1625
   grd_Cuotas.ColWidth(4) = 1295
   grd_Cuotas.ColWidth(5) = 1295
   grd_Cuotas.ColWidth(6) = 1295
   grd_Cuotas.ColWidth(7) = 1295
   grd_Cuotas.ColAlignment(0) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(1) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(2) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(3) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(4) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(5) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(6) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_Cuotas)
   
   pnl_Cuo_TotDeu.Caption = "0.00 "
   pnl_Cuo_TotPag.Caption = "0.00 "
   pnl_Cuo_TotSal.Caption = "0.00 "
End Sub

Private Sub fs_Buscar()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_TDOCLI, HIPMAE_NDOCLI, HIPMAE_NUMSOL, HIPMAE_NUMOPE, HIPMAE_TDOCYG, HIPMAE_NDOCYG, HIPMAE_CODPRD, HIPMAE_CODSUB,"
   g_str_Parame = g_str_Parame & "       HIPMAE_CODMOD, HIPMAE_EJESEG, HIPMAE_CONHIP, HIPMAE_MONEDA, HIPMAE_MTOPRE, HIPMAE_CUOPEN, HIPMAE_NUMCUO, HIPMAE_SALCAP,"
   g_str_Parame = g_str_Parame & "       HIPMAE_FECDES, HIPMAE_TDOCLI, HIPMAE_NDOCLI, HIPMAE_NUMSOL, HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")

   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                  'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                  'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                  'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                  'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad) 'fs_Buscar
   
   lbl_Totale.Caption = "Totales ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
         
   'Buscando Cuotas
   Call fs_Buscar_Cuotas
   Call gs_SetFocus(grd_Cuotas)
End Sub

Private Sub fs_Buscar_Cuotas()
Dim r_dbl_Pag_TotCuo    As Double
Dim r_dbl_Pag_Capita    As Double
Dim r_dbl_Pag_Intere    As Double
Dim r_dbl_Pag_SegDes    As Double
Dim r_dbl_Pag_SegViv    As Double
Dim r_dbl_Pag_OtrCar    As Double
Dim r_dbl_Pag_IntMor    As Double
Dim r_dbl_Pag_IntCom    As Double
Dim r_dbl_Pag_GasCob    As Double
Dim r_dbl_Pag_OtrGas    As Double
Dim r_dbl_Deu_TotCuo    As Double
Dim r_dbl_Deu_Capita    As Double
Dim r_dbl_Deu_Intere    As Double
Dim r_dbl_Deu_SegDes    As Double
Dim r_dbl_Deu_SegViv    As Double
Dim r_dbl_Deu_OtrCar    As Double
Dim r_dbl_Deu_IntMor    As Double
Dim r_dbl_Deu_IntCom    As Double
Dim r_dbl_Deu_GasCob    As Double
Dim r_dbl_Deu_OtrGas    As Double
Dim r_dbl_Sal_TotCuo    As Double
Dim r_dbl_Sal_Capita    As Double
Dim r_dbl_Sal_Intere    As Double
Dim r_dbl_Sal_SegDes    As Double
Dim r_dbl_Sal_SegViv    As Double
Dim r_dbl_Sal_OtrCar    As Double
Dim r_dbl_Sal_IntMor    As Double
Dim r_dbl_Sal_IntCom    As Double
Dim r_dbl_Sal_GasCob    As Double
Dim r_dbl_Sal_OtrGas    As Double
Dim r_dbl_Gen_TotDeu    As Double
Dim r_dbl_Gen_TotPag    As Double
Dim r_dbl_Gen_TotSal    As Double

   r_dbl_Gen_TotDeu = 0
   r_dbl_Gen_TotPag = 0
   r_dbl_Gen_TotSal = 0
   
   'Cuotas Vencidas
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Cuotas.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'A Pagar
         r_dbl_Deu_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_CAPBBP, "###,###,##0.00"))
         r_dbl_Deu_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_INTBBP, "###,###,##0.00"))
         r_dbl_Deu_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_Deu_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_Deu_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_Deu_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_INTMOR, "###,###,##0.00"))
         r_dbl_Deu_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_INTCOM, "###,###,##0.00"))
         r_dbl_Deu_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GASCOB, "###,###,##0.00"))
         r_dbl_Deu_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTRGAS, "###,###,##0.00"))
         
         r_dbl_Deu_TotCuo = 0
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Capita
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Intere
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegDes
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegViv
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrCar
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntMor
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntCom
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_GasCob
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrGas
         
         'Pagado
         r_dbl_Pag_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPPAG + g_rst_Princi!HIPCUO_CBPPAG, "###,###,##0.00"))
         r_dbl_Pag_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTPAG + g_rst_Princi!HIPCUO_IBPPAG, "###,###,##0.00"))
         r_dbl_Pag_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00"))
         r_dbl_Pag_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00"))
         r_dbl_Pag_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00"))
         r_dbl_Pag_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00"))
         r_dbl_Pag_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00"))
         
         r_dbl_Pag_TotCuo = 0
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Capita
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Intere
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegDes
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegViv
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrCar
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntCom
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntMor
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_GasCob
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrGas
         
         'Saldo Pago
         r_dbl_Sal_Capita = r_dbl_Deu_Capita - r_dbl_Pag_Capita
         r_dbl_Sal_Intere = r_dbl_Deu_Intere - r_dbl_Pag_Intere
         r_dbl_Sal_IntCom = r_dbl_Deu_IntCom - r_dbl_Pag_IntCom
         r_dbl_Sal_IntMor = r_dbl_Deu_IntMor - r_dbl_Pag_IntMor
         r_dbl_Sal_GasCob = r_dbl_Deu_GasCob - r_dbl_Pag_GasCob
         r_dbl_Sal_OtrGas = r_dbl_Deu_OtrGas - r_dbl_Pag_OtrGas
         
         r_dbl_Sal_SegDes = r_dbl_Deu_SegDes - r_dbl_Pag_SegDes
         r_dbl_Sal_SegViv = r_dbl_Deu_SegViv - r_dbl_Pag_SegViv
         r_dbl_Sal_OtrCar = r_dbl_Deu_OtrCar - r_dbl_Pag_OtrCar
         
         'Total Cuota
         r_dbl_Sal_TotCuo = 0
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Capita
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Intere
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegDes
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegViv
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrCar
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntCom
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntMor
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_GasCob
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrGas
         
         grd_Cuotas.Rows = grd_Cuotas.Rows + 1
         grd_Cuotas.Row = grd_Cuotas.Rows - 1
         
         grd_Cuotas.Col = 0
         grd_Cuotas.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_Cuotas.Col = 1
         grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         'Si Situación es No-Pagado
         If g_rst_Princi!HIPCUO_SITUAC = 2 Then
            If moddat_g_int_Situac = 2 Then
               If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < CDate(moddat_g_str_FecSis) Then
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "VENCIDA"
               Else
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = "-"
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "POR VENCER"
               End If
            End If
         Else
            If CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))) > 0 Then
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
            Else
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = "-"
            End If
            
            grd_Cuotas.Col = 3
            grd_Cuotas.Text = "PAGADA"
         End If
         
         If g_rst_Princi!HIPCUO_FECPAG > 0 Then
            grd_Cuotas.Col = 4
            grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))
         End If
      
         'Valor Cuota
         grd_Cuotas.Col = 5
         grd_Cuotas.Text = Format(r_dbl_Deu_TotCuo, "###,###,##0.00")
      
         'Importe Pagado
         grd_Cuotas.Col = 6
         grd_Cuotas.Text = Format(r_dbl_Pag_TotCuo, "###,###,##0.00")
      
         'Saldo
         grd_Cuotas.Col = 7
         grd_Cuotas.Text = Format(r_dbl_Sal_TotCuo, "###,###,##0.00")
      
         'Sumando Totales
         r_dbl_Gen_TotDeu = r_dbl_Gen_TotDeu + r_dbl_Deu_TotCuo
         r_dbl_Gen_TotPag = r_dbl_Gen_TotPag + r_dbl_Pag_TotCuo
         r_dbl_Gen_TotSal = r_dbl_Gen_TotSal + r_dbl_Sal_TotCuo
      
         g_rst_Princi.MoveNext
      Loop
      
      pnl_Cuo_TotDeu.Caption = Format(r_dbl_Gen_TotDeu, "###,###,##0.00") & " "
      pnl_Cuo_TotPag.Caption = Format(r_dbl_Gen_TotPag, "###,###,##0.00") & " "
      pnl_Cuo_TotSal.Caption = Format(r_dbl_Gen_TotSal, "###,###,##0.00") & " "
      
      grd_Cuotas.Redraw = True
      
      Call gs_UbiIniGrid(grd_Cuotas)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub grd_Cuotas_DblClick()
   If grd_Cuotas.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Cuotas.Col = 0
   moddat_g_int_NumCuo = CInt(grd_Cuotas)
   
   Call gs_RefrescaGrid(grd_Cuotas)
   
   frm_Con_CreHip_03.Show 1
End Sub

Private Sub grd_Cuotas_SelChange()
   If grd_Cuotas.Rows > 2 Then
      grd_Cuotas.RowSel = grd_Cuotas.Row
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
