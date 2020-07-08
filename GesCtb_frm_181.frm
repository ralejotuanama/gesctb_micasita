VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RegDes_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11625
   Icon            =   "GesCtb_frm_181.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   15055
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
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   30
            Picture         =   "GesCtb_frm_181.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "GesCtb_frm_181.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            TabIndex        =   5
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Consulta de Desembolso a Promotor"
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
            Picture         =   "GesCtb_frm_181.frx":0758
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4125
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7276
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
         Begin MSFlexGridLib.MSFlexGrid grd_DatSol 
            Height          =   3735
            Left            =   60
            TabIndex        =   7
            Top             =   330
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6588
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
            Caption         =   "Datos Generales de Solicitud de Crédito"
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
            TabIndex        =   8
            Top             =   60
            Width           =   3945
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2895
         Left            =   30
         TabIndex        =   9
         Top             =   5610
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5106
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisIns 
            Height          =   2265
            Left            =   60
            TabIndex        =   10
            Top             =   630
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3995
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   4785
            TabIndex        =   11
            Top             =   330
            Width           =   1370
            _Version        =   65536
            _ExtentX        =   2417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Fin Eval."
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
            Left            =   7260
            TabIndex        =   12
            Top             =   330
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado Actual"
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
            Left            =   3390
            TabIndex        =   13
            Top             =   330
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inicio Eval."
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
            Left            =   90
            TabIndex        =   14
            Top             =   330
            Width           =   3300
            _Version        =   65536
            _ExtentX        =   5821
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   6150
            TabIndex        =   15
            Top             =   330
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Transc."
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
         Begin VB.Label Label1 
            Caption         =   "Seguimiento por Instancias"
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
            TabIndex        =   16
            Top             =   60
            Width           =   3165
         End
      End
   End
End
Attribute VB_Name = "frm_RegDes_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_numoper As String
Dim r_str_client As String
Dim r_str_tpDocCl As String
Dim r_str_NdocCl As String
Dim r_str_apoder As String
Dim r_str_tpDocApo As String
Dim r_str_NdocApo As String
Dim r_str_conyug As String
Dim r_str_produc As String
Dim r_str_priviv As String
Dim r_str_moneda As String
Dim r_str_modali As String
Dim r_str_dirInm As String
Dim r_str_Distri As String
Dim r_str_proyInm As String
Dim r_str_valorCV As String
Dim r_str_Aporte As String
Dim r_str_MtoPtmo As String
Dim r_str_TEA As String
Dim r_str_PLAZO As String
Dim r_str_NCuotas As String
Dim r_str_PGracia As String
Dim r_str_CtaExtra As String
Dim r_str_CompSeg As String
Dim r_str_Desgrav As String
Dim r_str_diaPago As String
Dim r_str_Situac As String
Dim r_str_FchIng As String
Dim r_str_consej As String
Dim r_str_Ejecut As String

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_str_FecRec As String
Dim r_str_FecHip As String

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   r_str_FecRec = moddat_g_str_FecRec
   r_str_FecHip = moddat_g_str_FecHip
   'Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol) 'fs_Buscar_DatGen
   Call fs_Buscar_DatGen
   moddat_g_str_FecRec = r_str_FecRec
   moddat_g_str_FecHip = r_str_FecHip
   Call fs_Buscar_Seguim
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos de la Solicitud
   grd_DatSol.ColWidth(0) = 2600
   grd_DatSol.ColWidth(1) = 8470
   grd_DatSol.ColAlignment(0) = flexAlignLeftCenter
   grd_DatSol.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Instancias
   grd_LisIns.ColWidth(0) = 3300
   grd_LisIns.ColWidth(1) = 1385
   grd_LisIns.ColWidth(2) = 1385
   grd_LisIns.ColWidth(3) = 1115
   grd_LisIns.ColWidth(4) = 3915
   grd_LisIns.ColWidth(5) = 0
   grd_LisIns.ColWidth(6) = 0
   grd_LisIns.ColAlignment(0) = flexAlignLeftCenter
   grd_LisIns.ColAlignment(1) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(2) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(3) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(4) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_DatGen()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   Call gs_LimpiaGrid(grd_DatSol)
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Cliente
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Cónyuge
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Producto
   moddat_g_str_CodPrd = g_rst_Princi!SOLMAE_CODPRD
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))
   moddat_g_str_CodSub = g_rst_Princi!SOLMAE_CODSUB
   
   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   'Modalidad
   moddat_g_str_CodMod = ""
   moddat_g_str_DesMod = ""
   
   If Len(Trim(g_rst_Princi!SOLMAE_CODMOD & "")) > 0 Then
      moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
      moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   End If
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJESEG))
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
   
   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   'Situación
   moddat_g_int_Situac = g_rst_Princi!SOLMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
   
   'Inmueble Identificado
   moddat_g_int_InmIde = g_rst_Princi!SOLMAE_INMIDE
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Número de Solicitud"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
   r_str_numoper = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Cliente"
   grd_DatSol.Col = 1:
   grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & " - " & Trim(g_rst_Princi!SOLMAE_TITNDO) & " / " & moddat_g_str_NomCli
   r_str_client = Trim(moddat_g_str_NomCli)
   r_str_tpDocCl = g_rst_Princi!SOLMAE_TITTDO
   r_str_NdocCl = Trim(g_rst_Princi!SOLMAE_TITNDO)
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Cónyuge"
      grd_DatSol.Col = 1:
      grd_DatSol.Text = CStr(moddat_g_int_CygTDo) & " - " & Trim(moddat_g_str_CygNDo) & " / " & moddat_g_str_CygNom
      r_str_conyug = CStr(moddat_g_int_CygTDo) & " - " & Trim(moddat_g_str_CygNDo) & " / " & moddat_g_str_CygNom
   End If
   
   'Apoderado
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera!DATGEN_APOTDO > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Apoderado"
      grd_DatSol.Col = 1:
      grd_DatSol.Text = CStr(g_rst_Genera!DATGEN_APOTDO) & " - " & Trim(g_rst_Genera!DATGEN_APONDO) & " / " & Trim(g_rst_Genera!DATGEN_APOAPP) & " " & Trim(g_rst_Genera!DATGEN_APOAPM) & " " & Trim(g_rst_Genera!DATGEN_APONOM)
      r_str_apoder = Trim(g_rst_Genera!DATGEN_APONOM)
      r_str_tpDocApo = g_rst_Genera!DATGEN_APOTDO
      r_str_NdocApo = Trim(g_rst_Genera!DATGEN_APONDO)
   End If
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Producto"
   grd_DatSol.Col = 1:
   r_str_produc = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   grd_DatSol.Text = r_str_produc
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Primera Vivienda"
   grd_DatSol.Col = 1:
   r_str_priviv = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
   grd_DatSol.Text = r_str_priviv
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Moneda Préstamo"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = moddat_g_str_Moneda
   r_str_moneda = moddat_g_str_Moneda
   
   If Len(Trim(moddat_g_str_Direcc)) > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Modalidad"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_DesMod
      r_str_modali = moddat_g_str_DesMod
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Dirección Inmueble"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_Direcc
      r_str_dirInm = moddat_g_str_Direcc
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Distrito"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_Distri
      r_str_Distri = moddat_g_str_Distri
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1: grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0:                    grd_DatSol.Text = "Proyecto Inmobiliario"
         grd_DatSol.Col = 1:
         r_str_proyInm = moddat_gf_Consulta_NomPry(r_str_CodPry)
         grd_DatSol.Text = r_str_proyInm
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1: grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0:                    grd_DatSol.Text = "Proyecto Inmobiliario"
         grd_DatSol.Col = 1:
         r_str_proyInm = r_str_NomPry & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
         grd_DatSol.Text = r_str_proyInm
         
      End If
   End If
   
   If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Valor Compra Venta"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8:
      r_str_valorCV = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2))
      grd_DatSol.Text = r_str_valorCV

      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Aporte Propio"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      If moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_DatSol.CellFontSize = 8:
         r_str_Aporte = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)) & " (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & ") "
         grd_DatSol.Text = r_str_Aporte
      Else
         grd_DatSol.CellFontSize = 8:
         r_str_Aporte = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2))
         grd_DatSol.Text = r_str_Aporte
      End If
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Monto Préstamo"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8:
      r_str_MtoPtmo = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_MPR, 12, 2)
      grd_DatSol.Text = r_str_MtoPtmo
   End If
      
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Tasa de Interés"
   grd_DatSol.Col = 1:
   r_str_TEA = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & "%"
   grd_DatSol.Text = r_str_TEA
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Plazo"
   grd_DatSol.Col = 1:
   r_str_PLAZO = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
   grd_DatSol.Text = r_str_PLAZO
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Número de Cuotas"
   grd_DatSol.Col = 1:
   r_str_NCuotas = CStr(g_rst_Princi!SOLMAE_PLAANO * 12)
   grd_DatSol.Text = r_str_NCuotas
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Período de Gracia"
   grd_DatSol.Col = 1:
   r_str_PGracia = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
   grd_DatSol.Text = r_str_PGracia

   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Cuotas Extraordinarias"
   grd_DatSol.Col = 1:
   r_str_CtaExtra = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
   grd_DatSol.Text = r_str_CtaExtra
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Compañía de Seguros"
   grd_DatSol.Col = 1:
   r_str_CompSeg = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
   grd_DatSol.Text = r_str_CompSeg
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Tipo de Seguro Desgravamen"
   grd_DatSol.Col = 1:
   r_str_Desgrav = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
   grd_DatSol.Text = r_str_Desgrav
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Día de Pago"
   grd_DatSol.Col = 1:
   r_str_diaPago = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   grd_DatSol.Text = r_str_diaPago
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Situación"
   grd_DatSol.Col = 1:
   r_str_Situac = moddat_g_str_Situac
   grd_DatSol.Text = moddat_g_str_Situac
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Fecha de Ingreso"
   grd_DatSol.Col = 1:
   r_str_FchIng = moddat_g_str_FecIng
   grd_DatSol.Text = moddat_g_str_FecIng
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Consejero Hipotecario"
   grd_DatSol.Col = 1:
   r_str_consej = moddat_g_str_NomConHip
   grd_DatSol.Text = moddat_g_str_NomConHip
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Ejecutivo Seguimiento"
   grd_DatSol.Col = 1:
   r_str_Ejecut = moddat_g_str_NomEjeSeg
   grd_DatSol.Text = moddat_g_str_NomEjeSeg
   
   '/******************************************************************************************/
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   
   If g_rst_Genera!DATGEN_TDOVIN > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Vinculado"
      grd_DatSol.Col = 1
      
      If g_rst_Genera!DATGEN_TIPVIN = 1 Then
         grd_DatSol.Text = "TRABAJADOR"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 2 Or g_rst_Genera!DATGEN_TIPVIN = 3 Then
         grd_DatSol.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 4 Then
         grd_DatSol.Text = "FUNCIONARIO"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 5 Then
         grd_DatSol.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      Else
         grd_DatSol.Text = ""
      End If
   End If
   
   If g_rst_Genera!DATGEN_TDOACC > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Accionista"
      grd_DatSol.Col = 1
      
      If g_rst_Genera!DATGEN_ACCVIN = 1 Then
         grd_DatSol.Text = "ACCIONISTA"
      ElseIf g_rst_Genera!DATGEN_ACCVIN = 2 Then
         grd_DatSol.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Genera!DATGEN_TDOACC, Trim(g_rst_Genera!DATGEN_NDOACC)) & ")"
      End If
   End If
   
   modmip_g_int_PaiRes = g_rst_Genera!DATGEN_PAIRES
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_DatSol)
End Sub

Private Sub fs_Buscar_Seguim()
   Dim r_int_DiaTra     As Integer
   Dim r_int_DiaTas     As Integer
   Dim r_int_DiaSeg     As Integer
   Dim r_int_DiaPol     As Integer
   Dim r_int_DiaMVi     As Integer
   
   Call gs_LimpiaGrid(grd_LisIns)
   r_int_DiaTra = 0
   r_int_DiaTas = 0
   r_int_DiaSeg = 0
   r_int_DiaPol = 0
   r_int_DiaMVi = 0
      
   g_str_Parame = "SELECT * FROM cre_desprodet WHERE desdet_numope = '" & moddat_g_str_NumOpe & "' and " & _
                  " desdet_fecreg =  '" & moddat_g_str_FecRec & "' and " & _
                  " desdet_horreg = '" & moddat_g_str_FecHip & "' "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   g_rst_Princi.MoveFirst
   grd_LisIns.Redraw = False
   
   Do While Not g_rst_Princi.EOF
      grd_LisIns.Rows = grd_LisIns.Rows + 1
      grd_LisIns.Row = grd_LisIns.Rows - 1
                  
      'Instancia
      grd_LisIns.Col = 0
      grd_LisIns.Text = moddat_gf_Consulta_ParDes("374", Format(g_rst_Princi!desdet_codarea, "000000"))
      
      grd_LisIns.Col = 5
      grd_LisIns.Text = g_rst_Princi!desdet_codarea
      
      'Fecha de Inicio
      grd_LisIns.Col = 1
      grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!desdet_fecini))
      
      'Fecha de Inicio
      grd_LisIns.Col = 1
      grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!desdet_fecini))
      
      If g_rst_Princi!desdet_fecfin > 0 Then
         'Fecha de Fin
         grd_LisIns.Col = 2
         grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!desdet_fecfin))
         'Días Transcurridos
         grd_LisIns.Col = 3
         grd_LisIns.Text = CStr(g_rst_Princi!desdet_diatra)
         'Situación
         grd_LisIns.Col = 4
         grd_LisIns.Text = moddat_gf_Consulta_ParDes("375", CStr(g_rst_Princi!desdet_codest))
      Else
         'Días Transcurridos
         r_int_DiaTra = r_int_DiaTra + CInt(date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!desdet_fecini))))
         grd_LisIns.Col = 3
         grd_LisIns.Text = CStr(r_int_DiaTra)
      End If
      
      grd_LisIns.Col = 6
      grd_LisIns.Text = CStr(g_rst_Princi!desdet_codest)
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   grd_LisIns.Redraw = True
   Call gs_UbiIniGrid(grd_LisIns)
End Sub

Private Function ff_ObsExp(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   ff_ObsExp = " "
      
   g_str_Parame = "SELECT SEGDET_OBSERV "
   g_str_Parame = g_str_Parame & "FROM (SELECT RANK() OVER (ORDER BY SEGFECCRE DESC, SEGHORCRE DESC) R,SEGDET_OBSERV FROM TRA_SEGDET"
   g_str_Parame = g_str_Parame & " WHERE  SEGDET_NUMSOL='" & p_NumSol & "' AND SEGDET_CODINS='" & p_CodIns & "' )"
   g_str_Parame = g_str_Parame & " WHERE R < 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   DoEvents
   g_rst_Genera.MoveFirst
   ff_ObsExp = Trim(g_rst_Genera!SEGDET_OBSERV & "")

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub fs_GenExc()
   Dim r_str_nrohij As String
   Dim r_str_clinom As String
   Dim r_str_edahij As String
   Dim r_str_comvta As String
   Dim r_str_apopro As String
   Dim r_str_mtopre As String
   Dim r_str_nompfs As String
   Dim r_str_nomcar As String
   Dim r_str_fecing As String
   Dim r_str_cuoapr As String
   Dim r_str_cuomen As String
   Dim r_str_ingliq As String
   Dim r_str_obs As String
   
   'r_str_CodPry
   Dim r_obj_Excel      As Excel.Application
                
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
     
   With r_obj_Excel.ActiveSheet
      'Unir celdas
      .Range("B6") = "SEGUIMIENTO POR INSTANCIAS DE DESEMBOLSO A PROMOTOR"
      .Range("B6:D5").Font.Underline = True
      .Range("B6:F6").Merge
      
      .Range("B45:F45").Merge
      .Range("E47:F47").Merge
      .Range("E48:F48").Merge
      .Range("E49:F49").Merge
      .Range("E50:F50").Merge
      .Range("E51:F51").Merge
      .Range("E52:F52").Merge
      .Range("E53:F53").Merge
      
'      'Bordes de las celdas
      .Range("B45:F45").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B45:F45").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B45:F45").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B45:F45").Borders(xlEdgeTop).Weight = xlThin
      .Range("B45:F45").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B45:F45").Borders(xlEdgeRight).Weight = xlThin
      .Range("B45:F45").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B45:F45").Borders(xlEdgeBottom).Weight = xlThin
      
      
      .Range("B47:F47").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B47:F47").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B47:F53").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B47:F53").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B47:F53").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B47:F53").Borders(xlEdgeTop).Weight = xlThin
      .Range("B47:F53").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B47:F53").Borders(xlEdgeRight).Weight = xlThin
      .Range("B47:F53").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B47:F53").Borders(xlEdgeBottom).Weight = xlThin
       
       Dim ndoc_Cli As String
       Dim ndoc_Apo As String
       ndoc_Cli = moddat_gf_Consulta_ParDes("230", CStr(r_str_tpDocCl))
       ndoc_Apo = moddat_gf_Consulta_ParDes("230", CStr(r_str_tpDocApo))

      .Cells(9, 2) = "NRO DE SOLICITUD:"
      .Cells(11, 2) = "CLIENTE:"
      .Cells(11, 4) = ndoc_Cli
      .Cells(13, 2) = "APODERADO:"
      .Cells(13, 4) = ndoc_Apo
      .Cells(15, 2) = "PRODUCTO"
      .Cells(17, 2) = "PRIMERA VIVIENDA:"
      .Cells(17, 4) = "MONEDA PRESTAMO:"
      .Cells(19, 2) = "DIRECCIÓN INMUEBLE:"
      .Cells(21, 2) = "DISTRITO:"
      .Cells(21, 4) = "PROYECTO INMOBILIARIO:"
      .Cells(23, 2) = "VALOR COMPRA VENTA:"
      .Cells(25, 2) = "APORTE PROPIO:"
      .Cells(27, 2) = "MONTO PRESTAMO:"
      .Cells(29, 2) = "TASA DE INTERES:"
      .Cells(29, 4) = "PLAZO:"
      .Cells(31, 2) = "NRO DE CUOTAS:"
      .Cells(31, 4) = "PERIODO DE GRACIA:"
      .Cells(33, 2) = "CUOTAS EXTRAORDINARIAS:"
      .Cells(35, 2) = "COMPAÑIA DE SEGUROS:"
      .Cells(37, 2) = "TIPO DE SEGURO DESGRAVAMEN:"
      .Cells(37, 4) = "DIA DE PAGO:"
      .Cells(39, 2) = "SITUACION:"
      .Cells(39, 4) = "FECHA DE INGRESO"
      .Cells(41, 2) = "CONSEJERO HIPOTECARIO:"
      .Cells(43, 2) = "EJECUTIVO SEGUIMIENTO:"
            
      .Cells(45, 2) = "SEGUIMIENTO POR INSTANCIAS"
           
      .Columns("A").ColumnWidth = 1.43
      .Columns("B").ColumnWidth = 26
      .Columns("C").ColumnWidth = 33.86
      .Columns("D").ColumnWidth = 20.29
      .Columns("E").ColumnWidth = 19.29
      
      .Rows("1:65536").RowHeight = 11.25
      
      .Cells(43, 2).Font.Bold = True
      .Columns("B").Font.Bold = True
      .Columns("D").Font.Bold = True
      .Range("B7").Font.Bold = True
            
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Range("B6").HorizontalAlignment = xlHAlignCenter
      .Range("B47:F53").HorizontalAlignment = xlHAlignLeft
      .Cells(2, 6).HorizontalAlignment = xlHAlignRight
      .Cells(3, 6).HorizontalAlignment = xlHAlignRight
                                    
      .Range("B47:F53").VerticalAlignment = xlVAlignTop
      .Range("B43:F53").WrapText = True
                                    
      .Range("A1:F85").Font.Name = "Arial"
      .Range("A1:F85").Font.Size = 8
   End With
'***************************************************************************
With r_obj_Excel.ActiveSheet
      .Cells(9, 3) = r_str_numoper
      .Cells(11, 3) = r_str_client
      .Cells(11, 5) = r_str_NdocCl
      .Cells(13, 3) = r_str_apoder
      .Cells(13, 5) = r_str_NdocApo
      .Cells(15, 3) = r_str_produc
      .Cells(17, 3) = r_str_priviv
      .Cells(17, 5) = r_str_moneda
      .Cells(19, 3) = r_str_dirInm
      .Cells(21, 3) = r_str_Distri
      .Cells(21, 5) = r_str_proyInm
      .Cells(23, 3) = r_str_valorCV
      .Cells(25, 3) = r_str_Aporte
      .Cells(27, 3) = r_str_MtoPtmo
      .Cells(29, 3) = r_str_TEA
      .Cells(29, 5) = r_str_PLAZO
      .Cells(31, 3) = r_str_NCuotas
      .Cells(31, 5) = r_str_PGracia
      .Cells(33, 3) = r_str_CtaExtra
      .Cells(35, 3) = r_str_CompSeg
      .Cells(37, 3) = r_str_Desgrav
      .Cells(37, 5) = r_str_diaPago
      .Cells(39, 3) = r_str_Situac
      .Cells(39, 5) = r_str_FchIng
      .Cells(41, 3) = r_str_consej
      .Cells(43, 3) = r_str_Ejecut
      
      .Cells(47, 2) = "INSTANCIAS"
      .Cells(47, 3) = "F.INICIO EVAL.       F.FIN EVAL."
      .Cells(47, 4) = "DIAS TRANSCURRIDOS"
      .Cells(47, 5) = "ESTADO ACTUAL"
      
      .Range("C19:F20").Merge
      .Cells(19, 3).HorizontalAlignment = xlHAlignLeft
      .Cells(19, 3).VerticalAlignment = xlHAlignJustify
      
      .Cells(47, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(47, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(47, 5).HorizontalAlignment = xlHAlignCenter
      
      .Range("B45: F45").Font.Bold = True
      
      .Cells(45, 3).HorizontalAlignment = xlHAlignCenter
      .Range("B47: E47").Font.Bold = True
      .Range("B48: E53").Font.Bold = False
      .Range("B47: E47").Interior.Color = RGB(146, 208, 80)
      
      Dim r_int_nrofil  As Integer
      Dim r_int_nroaux  As Integer
      r_int_nrofil = 48
      For r_int_nroaux = 0 To grd_LisIns.Rows - 1
         .Cells(r_int_nrofil, 2) = grd_LisIns.TextMatrix(r_int_nroaux, 0)
         If (Len(Trim(grd_LisIns.TextMatrix(r_int_nroaux, 2))) = 0) Then
             .Cells(r_int_nrofil, 3) = Trim(grd_LisIns.TextMatrix(r_int_nroaux, 1)) & "             " & "__/__/____"
         Else
             .Cells(r_int_nrofil, 3) = Trim(grd_LisIns.TextMatrix(r_int_nroaux, 1)) & "             " & Trim(grd_LisIns.TextMatrix(r_int_nroaux, 2))
         End If
         .Cells(r_int_nrofil, 4) = grd_LisIns.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 5) = grd_LisIns.TextMatrix(r_int_nroaux, 4)
         
         .Cells(r_int_nrofil, 3).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_nrofil, 4).HorizontalAlignment = xlHAlignCenter
         r_int_nrofil = r_int_nrofil + 1
      Next
      
End With

     
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


