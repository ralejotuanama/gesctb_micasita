VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbPbp_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   Icon            =   "GesCtb_frm_920.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel11 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
      _ExtentY        =   1085
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
         Height          =   345
         Left            =   630
         TabIndex        =   1
         Top             =   150
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Consulta del detalle de Glosa"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
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
         Index           =   1
         Left            =   45
         Picture         =   "GesCtb_frm_920.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   4710
      Left            =   0
      TabIndex        =   2
      Top             =   2265
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
      _ExtentY        =   8308
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
      Begin Threed.SSPanel pnl_Cliente 
         Height          =   285
         Left            =   1290
         TabIndex        =   3
         Top             =   120
         Width           =   4580
         _Version        =   65536
         _ExtentX        =   8079
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Cliente"
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
      Begin Threed.SSPanel pnl_Operacion 
         Height          =   285
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Nro.Operacion"
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
      Begin Threed.SSPanel pnl_Monto 
         Height          =   315
         Left            =   9750
         TabIndex        =   11
         Top             =   4230
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
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
      Begin Threed.SSPanel pnl_Glosa 
         Height          =   285
         Left            =   5850
         TabIndex        =   12
         Top             =   120
         Width           =   3870
         _Version        =   65536
         _ExtentX        =   6826
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
      Begin Threed.SSPanel pnl_MontoDet 
         Height          =   285
         Left            =   9705
         TabIndex        =   4
         Top             =   120
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Monto"
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
         Height          =   3690
         Left            =   0
         TabIndex        =   6
         Top             =   450
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6509
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin Threed.SSPanel SSPanel22 
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   630
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
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
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   10530
         Picture         =   "GesCtb_frm_920.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_920.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel33 
      Height          =   920
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   1623
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin Threed.SSPanel pnl_NomProd 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   480
         Width           =   5625
         _Version        =   65536
         _ExtentX        =   9922
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
      Begin Threed.SSPanel pnl_Periodo 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   120
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
      Begin Threed.SSPanel pnl_NroCta 
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         Top             =   480
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
         Caption         =   "Nro.Cuenta :"
         Height          =   195
         Left            =   7920
         TabIndex        =   18
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   150
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   510
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPbp_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fs_IniciaGrid()
   grd_Listad.ColWidth(0) = 1250
   grd_Listad.ColWidth(1) = 4600
   grd_Listad.ColWidth(2) = 3850
   grd_Listad.ColWidth(3) = 980

   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   
   pnl_Monto.Caption = "0.00000"
   grd_Listad.Rows = 0
End Sub

Function MesEspanol(Mes As Integer) As String
   Select Case Mes
      Case 1:  MesEspanol = "Enero"
      Case 2:  MesEspanol = "Febrero"
      Case 3:  MesEspanol = "Marzo"
      Case 4:  MesEspanol = "Abril"
      Case 5:  MesEspanol = "Mayo"
      Case 6:  MesEspanol = "Junio"
      Case 7:  MesEspanol = "Julio"
      Case 8:  MesEspanol = "Agosto"
      Case 9:  MesEspanol = "Setiembre"
      Case 10: MesEspanol = "Octubre"
      Case 11: MesEspanol = "Noviembre"
      Case 12: MesEspanol = "Diciembre"
   End Select
End Function

Private Sub fs_Buscar_detalle()
Dim r_str_Codigo     As String

   r_str_Codigo = moddat_g_str_CodPrd
   pnl_NomProd.Caption = Trim(moddat_g_str_NomPrd)
   pnl_Periodo.Caption = UCase(MesEspanol(moddat_g_int_EdaMes)) & " - " & moddat_g_int_EdaAno
   pnl_NroCta.Caption = Trim(frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPMAE_CODPRD, DETPBP_NUMOPE, TRIM(DATGEN_NUMDOC) || ' - ' || TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE) CLIENTE, TRIM(HIPMAE_OPEMVI) HIPMAE_OPEMVI, DETPBP_CUOCON, DETPBP_CAPCLI, DETPBP_INTADE, DETPBP_CAPADE, A.DETPBP_INTCLI "
   g_str_Parame = g_str_Parame & "   FROM CRE_DETPBP A INNER JOIN CRE_HIPMAE B ON A.DETPBP_NUMOPE=B.HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "                     INNER JOIN CLI_DATGEN C ON B.HIPMAE_NDOCLI=C.DATGEN_NUMDOC AND B.HIPMAE_TDOCLI=C.DATGEN_TIPDOC "
   g_str_Parame = g_str_Parame & "  WHERE DETPBP_PERMES = " & moddat_g_int_EdaMes & " AND DETPBP_PERANO = " & moddat_g_int_EdaAno & " "
   g_str_Parame = g_str_Parame & "    AND DETPBP_NUMOPE=HIPMAE_NUMOPE AND HIPMAE_CODPRD IN (" & moddat_g_str_CodPrd & ") AND DETPBP_FLGPBP=1 "
   g_str_Parame = g_str_Parame & "  ORDER BY HIPMAE_CODPRD, DATGEN_APEPAT "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do Until g_rst_Princi.EOF
      r_str_Codigo = g_rst_Princi!HIPMAE_CODPRD
   
      Select Case r_str_Codigo
         Case "001", "003"
            If frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME") Then
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME") & " - CAP"
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,###,##0.00")
                           
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME") & " - INT"
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_INTADE, "###,###,##0.00")
            Else
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = frm_Pro_CtbPbp_02.grd_Listad.Text
            End If
            
         Case "004"
            If frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MIHOGAR" Then
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIHOGAR"
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,###,##0.00")
            Else
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = frm_Pro_CtbPbp_02.grd_Listad.Text
            End If
         Case "006"
            If frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MICASITA" Then
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MICASITA"
            Else
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
               grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
               grd_Listad.Col = 2: grd_Listad.Text = frm_Pro_CtbPbp_02.grd_Listad.Text
            End If
            
         Case "007", "009", "010", "012", "013", "014", "015", "016", "017", "018"
            If frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MIVIVIENDA" Then
               If r_str_Codigo = "007" Or r_str_Codigo = "012" Or r_str_Codigo = "013" Or r_str_Codigo = "014" Or r_str_Codigo = "015" Or r_str_Codigo = "016" Or r_str_Codigo = "017" Or r_str_Codigo = "018" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - CAP"
                  grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,###,##0.00")
               End If
               
               If r_str_Codigo = "009" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - PER EXT CAP"
                  grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,###,##0.00")
               End If
               
               If r_str_Codigo = "010" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - UNI AND CAP"
                  grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,###,##0.00")
               End If
            
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MIVIVIENDA - CAP" Then
               If r_str_Codigo = "007" Or r_str_Codigo = "012" Or r_str_Codigo = "013" Or r_str_Codigo = "014" Or r_str_Codigo = "015" Or r_str_Codigo = "016" Or r_str_Codigo = "017" Or r_str_Codigo = "018" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - CAP"
               End If
            
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MIVIVIENDA - PER EXT CAP" Then
               If r_str_Codigo = "010" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - PER EXT CAP"
               End If
               
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MIVIVIENDA - UNI AND CAP" Then
               If r_str_Codigo = "009" Then
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  grd_Listad.Col = 0: grd_Listad.Text = g_rst_Princi!DETPBP_NUMOPE
                  grd_Listad.Col = 1: grd_Listad.Text = g_rst_Princi!CLIENTE
                  grd_Listad.Col = 2: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - UNI AND CAP"
               End If
               
            End If
      End Select
      
      Select Case r_str_Codigo
         'CREDITO CRC-PBP
         Case "001"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "142104240101" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPCLI, "###,##0.00")
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "152719010105" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_INTADE, "###,##0.00")
            End If
         
         'CREDITO CME
         Case "003"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104250101" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPCLI, "###,##0.00")
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "151719010105" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_INTADE, "###,##0.00")
            End If
         
         'CREDITO PROYECTO MIHOGAR
         Case "004"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104230101" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPCLI, "###,##0.00")
            End If
         
         'CREDITO MICASITA SOLES
         Case "006"
            If frm_Pro_CtbPbp_02.grd_Listad.Text = "APLICACION PBP MICASITA" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPCLI + g_rst_Princi!DETPBP_INTCLI, "###,###,##0.00")
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104060101" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPCLI, "###,##0.00")
            ElseIf frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "511401040601" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_INTCLI, "###,##0.00")
            End If

         'CREDITO MIVIVIENDA
         Case "007", "012", "013", "014", "015", "016", "017", "018"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104230102" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,##0.00")
            End If
         
         'CREDITO MIVIVIENDA
         Case "009"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104230103" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,##0.00")
            End If
         
         'CREDITO MIVIVIENDA
         Case "010"
            If frm_Pro_CtbPbp_02.grd_Listad.TextMatrix(frm_Pro_CtbPbp_02.grd_Listad.Row, 1) = "141104230104" Then
               grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!DETPBP_CAPADE, "###,##0.00")
            End If
            
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Row = 0
   
   Call Sumar_Columnas
End Sub

Private Sub Sumar_Columnas()
Dim r_dbl_Total As Double
Dim r_int_Fila  As Integer
    
   For r_int_Fila = 0 To grd_Listad.Rows - 1
      r_dbl_Total = r_dbl_Total + grd_Listad.TextMatrix(r_int_Fila, 3)
   Next r_int_Fila
    
   pnl_Monto.Caption = Format(r_dbl_Total, "###,###,##0.00")
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NumFil     As Integer
Dim r_dbl_Monto      As Double

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONSULTA DE DETALLE DE GLOSA : " & Trim(pnl_Periodo.Caption)
      .Cells(3, 2) = "PRODUCTO : " & Trim(Me.pnl_NomProd.Caption) & Space(20) & "NRO. CUENTA : " & Trim(Me.pnl_NroCta.Caption)
      .Range(.Cells(2, 2), .Cells(2, 6)).Merge
      .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 6)).Font.Size = 12
      .Range(.Cells(2, 2), .Cells(2, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(3, 2), .Cells(3, 6)).Merge
      .Range(.Cells(3, 2), .Cells(3, 6)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 2), .Cells(3, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(2, 2), .Cells(3, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(3, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(3, 2), .Cells(3, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(3, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(3, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Cells(5, 2) = "N°"
      .Cells(5, 3) = "N°OPERACION"
      .Cells(5, 4) = "CLIENTE"
      .Cells(5, 5) = "GLOSA"
      .Cells(5, 6) = "MONTO (S/.)"
      .Range(.Cells(5, 2), .Cells(5, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Range(.Cells(5, 2), .Cells(5, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 6)).Font.Bold = True
      .Range(.Cells(5, 2), .Cells(5, 6)).HorizontalAlignment = xlHAlignCenter
             
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 14
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 50
      .Columns("E").ColumnWidth = 40
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 12
      .Columns("F").NumberFormat = "###,##0.00"
      
      .Range(.Cells(3, 1), .Cells(5, 6)).Font.Name = "Calibri"
      .Range(.Cells(3, 1), .Cells(5, 6)).Font.Size = 11
            
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
                  
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 6)).Font.Size = 10
         
         .Cells(r_int_NumFil + 3, 2) = r_int_NumFil - 2
         .Cells(r_int_NumFil + 3, 3) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0))
         .Cells(r_int_NumFil + 3, 4) = Space(2) & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 1))
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 3, 2)
         .Cells(r_int_NumFil + 3, 6) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         
         r_dbl_Monto = r_dbl_Monto + grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
            
      .Cells(r_int_NumFil + 3, 6).Interior.Color = RGB(146, 208, 80)
      .Cells(r_int_NumFil + 3, 6).Font.Bold = True
      .Cells(r_int_NumFil + 3, 6) = r_dbl_Monto
      .Cells(r_int_NumFil + 3, 6).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NumFil + 3, 5).Interior.Color = RGB(146, 208, 80)
      .Cells(r_int_NumFil + 3, 5).Font.Bold = True
      .Cells(r_int_NumFil + 3, 5).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NumFil + 3, 5) = "TOTAL : " & Space(5)
      
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
     
   Call fs_IniciaGrid
   Call fs_Buscar_detalle
   
   Call gs_CentraForm(Me)
   Call gs_RefrescaGrid(grd_Listad)
   Screen.MousePointer = 0
End Sub
