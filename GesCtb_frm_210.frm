VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_PagCom_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "GesCtb_frm_210.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel7 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      _Version        =   65536
      _ExtentX        =   14790
      _ExtentY        =   10557
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   2110
         Left            =   60
         TabIndex        =   1
         Top             =   1500
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
         _ExtentY        =   3722
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.23
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            Top             =   330
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_FecCtb 
            Height          =   315
            Left            =   1590
            TabIndex        =   19
            Top             =   660
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_TipOper 
            Height          =   315
            Left            =   1590
            TabIndex        =   20
            Top             =   990
            Width           =   4395
            _Version        =   65536
            _ExtentX        =   7761
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1590
            TabIndex        =   21
            Top             =   1320
            Width           =   4395
            _Version        =   65536
            _ExtentX        =   7761
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
         Begin Threed.SSPanel pnl_ImpPag 
            Height          =   315
            Left            =   1590
            TabIndex        =   27
            Top             =   1650
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Importe a Pagar:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1710
            Width           =   1170
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operación"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operación"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1380
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   390
            Width           =   540
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   750
            TabIndex        =   9
            Top             =   180
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cuentas por Pagar - Consulta"
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
            Left            =   150
            Picture         =   "GesCtb_frm_210.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2130
         Left            =   60
         TabIndex        =   10
         Top             =   3660
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
         _ExtentY        =   3757
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1590
            TabIndex        =   22
            Top             =   330
            Width           =   6060
            _Version        =   65536
            _ExtentX        =   10689
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
         Begin Threed.SSPanel pnl_Proveedor 
            Height          =   315
            Left            =   1590
            TabIndex        =   23
            Top             =   660
            Width           =   6060
            _Version        =   65536
            _ExtentX        =   10689
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
         Begin Threed.SSPanel pnl_Descrip 
            Height          =   315
            Left            =   1590
            TabIndex        =   24
            Top             =   990
            Width           =   6060
            _Version        =   65536
            _ExtentX        =   10689
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
         Begin Threed.SSPanel pnl_Banco 
            Height          =   315
            Left            =   1590
            TabIndex        =   25
            Top             =   1320
            Width           =   4395
            _Version        =   65536
            _ExtentX        =   7761
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
         Begin Threed.SSPanel pnl_CtaCte 
            Height          =   315
            Left            =   1590
            TabIndex        =   26
            Top             =   1650
            Width           =   4395
            _Version        =   65536
            _ExtentX        =   7761
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   90
            Width           =   885
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   390
            Width           =   1230
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1380
            Width           =   510
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1710
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   17
         Top             =   780
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7560
            Picture         =   "GesCtb_frm_210.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   Call fs_CargarDatos
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_CargarDatos()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CTAPAG_CODPAG, A.CTAPAG_FECOPE, TRIM(B.PARDES_DESCRI) AS TIPOOPERACION,  "
   g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) AS NOM_MONEDA, TRIM(D.PARDES_DESCRI) AS TIPODOCUMENTO,  "
   g_str_Parame = g_str_Parame & "        A.CTAPAG_NUMDOC, E.MAEPRV_RAZSOC, A.CTAPAG_DESCRP, A.CTAPAG_CODBCO,  "
   g_str_Parame = g_str_Parame & "        TRIM(F.PARDES_DESCRI) AS NOM_BANCO, A.CTAPAG_CTACRR, A.CTAPAG_IMPPAG  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CTAPAG A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 134 AND B.PARDES_CODITE = A.CTAPAG_TIPOPE  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CTAPAG_CODMON  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 118 AND D.PARDES_CODITE = A.CTAPAG_TIPDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV E ON E.MAEPRV_TIPDOC = A.CTAPAG_TIPDOC AND E.MAEPRV_NUMDOC = A.CTAPAG_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 122 AND F.PARDES_CODITE = A.CTAPAG_CODBCO  "
   g_str_Parame = g_str_Parame & "  WHERE A.CTAPAG_CODPAG = " & CLng(moddat_g_str_NumOpe)
 
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Codigo.Caption = Format(g_rst_Princi!CTAPAG_CODPAG, "0000000000")
      pnl_FecCtb.Caption = gf_FormatoFecha(g_rst_Princi!CTAPAG_FECOPE)
      pnl_TipOper.Caption = Trim(g_rst_Princi!TIPOOPERACION)
      pnl_Moneda.Caption = Trim(g_rst_Princi!NOM_MONEDA)
      pnl_TipDoc.Caption = Trim(g_rst_Princi!TIPODOCUMENTO)
      pnl_Proveedor.Caption = Trim(g_rst_Princi!CTAPAG_NUMDOC) & " - " & Trim(g_rst_Princi!MaePrv_RazSoc)
      pnl_Descrip.Caption = Trim(g_rst_Princi!CTAPAG_DESCRP)
      
      If g_rst_Princi!CTAPAG_CODBCO = 11 Then
         lbl_Cuenta = "Cuenta:"
      Else
         lbl_Cuenta = "CCI:"
      End If
      
      pnl_Banco.Caption = Trim(g_rst_Princi!NOM_BANCO)
      pnl_CtaCte.Caption = Trim(g_rst_Princi!CTAPAG_CTACRR)
      pnl_ImpPag.Caption = Format(g_rst_Princi!CTAPAG_IMPPAG, "###,###,##0.00") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
