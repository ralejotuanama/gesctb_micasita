VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_PagCom_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   Icon            =   "GesCtb_frm_217.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   4850
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10605
      _Version        =   65536
      _ExtentX        =   18706
      _ExtentY        =   8555
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
            TabIndex        =   16
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Modificar Cuenta Corriente"
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
            Left            =   30
            Picture         =   "GesCtb_frm_217.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   60
         TabIndex        =   17
         Top             =   780
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_217.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9750
            Picture         =   "GesCtb_frm_217.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2205
         Left            =   60
         TabIndex        =   18
         Top             =   1500
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   3889
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1020
            TabIndex        =   0
            Top             =   420
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
         Begin Threed.SSPanel pnl_Evaluador 
            Height          =   315
            Left            =   7470
            TabIndex        =   3
            Top             =   720
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
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
         Begin Threed.SSPanel pnl_TipProceso 
            Height          =   315
            Left            =   7470
            TabIndex        =   1
            Top             =   390
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
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
            Left            =   1020
            TabIndex        =   4
            Top             =   1080
            Width           =   4620
            _Version        =   65536
            _ExtentX        =   8149
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
         Begin Threed.SSPanel pnl_Fecha 
            Height          =   315
            Left            =   1020
            TabIndex        =   2
            Top             =   750
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
            Left            =   7470
            TabIndex        =   5
            Top             =   1050
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
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
            Left            =   1020
            TabIndex        =   8
            Top             =   1740
            Width           =   4620
            _Version        =   65536
            _ExtentX        =   8149
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   315
            Left            =   7470
            TabIndex        =   9
            Top             =   1710
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Banco 
            Height          =   315
            Left            =   1020
            TabIndex        =   6
            Top             =   1410
            Width           =   4620
            _Version        =   65536
            _ExtentX        =   8149
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
            Left            =   7470
            TabIndex        =   7
            Top             =   1380
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
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
            TabIndex        =   32
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   6060
            TabIndex        =   27
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total a Pagar:"
            Height          =   195
            Left            =   6060
            TabIndex        =   26
            Top             =   1770
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1770
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   6060
            TabIndex        =   24
            Top             =   1110
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Evaluador:"
            Height          =   195
            Left            =   6060
            TabIndex        =   23
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1110
            Width           =   780
         End
         Begin VB.Label lbl_Fecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   450
            Width           =   540
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Proceso:"
            Height          =   195
            Left            =   6060
            TabIndex        =   19
            Top             =   450
            Width           =   990
         End
      End
      Begin Threed.SSPanel SSPanel17 
         Height          =   945
         Left            =   60
         TabIndex        =   29
         Top             =   3750
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   1667
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
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   7470
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   2520
         End
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   4620
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente"
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
            TabIndex        =   33
            Top             =   90
            Width           =   1440
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   6060
            TabIndex        =   31
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   540
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CtaCte()  As moddat_tpo_Genera
Dim l_int_CodAut    As Long
Dim l_str_CodOpe    As String
Dim l_int_TipOpe    As Integer
Dim l_int_TipDoc    As Integer
Dim l_str_NumDoc    As String

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_CargarDatos
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   With frm_Ctb_PagCom_06
        l_int_CodAut = CLng(.grd_Listad.TextMatrix(.grd_Listad.Row, 9))
        l_str_CodOpe = Trim(CStr(.grd_Listad.TextMatrix(.grd_Listad.Row, 0)))
        l_int_TipOpe = CLng(.grd_Listad.TextMatrix(.grd_Listad.Row, 16))
        l_int_TipDoc = CLng(.grd_Listad.TextMatrix(.grd_Listad.Row, 10))
        l_str_NumDoc = Trim(CStr(.grd_Listad.TextMatrix(.grd_Listad.Row, 11)))
   End With
   
   cmb_Banco.ListIndex = -1
   cmb_CtaCte.ListIndex = -1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Public Sub fs_CargarDatos()

   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COMAUT_CODAUT, COMAUT_CODOPE, COMAUT_FECOPE, COMAUT_TIPDOC, COMAUT_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        COMAUT_IMPPAG, COMAUT_CODBNC, COMAUT_CTACRR, TRIM(C.PARDES_DESCRI) AS MONEDA, A.COMAUT_CODMON,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS TIPOPROCESO, COMAUT_USUAPR, TRIM(F.PARDES_DESCRI) AS NOM_BANCO,  "
   g_str_Parame = g_str_Parame & "        DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE)  "
   g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS MaePrv_RazSoc, A.COMAUT_TIPOPE, A.COMAUT_DESCRP, MAEPRV_CTADET  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMAUT_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMAUT_NUMDOC)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMAUT_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMAUT_NUMDOC)  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.COMAUT_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 136 AND TO_NUMBER(D.PARDES_CODITE) = TO_NUMBER(SUBSTR(LPAD(COMAUT_CODOPE,10,0),1,2)) AND D.PARDES_CODITE <> 0  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 122 AND F.PARDES_CODITE = A.COMAUT_CODBNC  "
   g_str_Parame = g_str_Parame & "  WHERE A.COMAUT_CODAUT =  " & l_int_CodAut
   g_str_Parame = g_str_Parame & "    AND A.COMAUT_CODOPE =  " & l_str_CodOpe
   g_str_Parame = g_str_Parame & "    AND A.COMAUT_TIPOPE =  " & l_int_TipOpe
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   pnl_Codigo.Caption = Format(g_rst_Princi!COMAUT_CODOPE, "0000000000")
   pnl_TipProceso.Caption = Trim(g_rst_Princi!TIPOPROCESO)
   pnl_Fecha.Caption = gf_FormatoFecha(g_rst_Princi!COMAUT_FECOPE)
   pnl_Evaluador.Caption = Trim(g_rst_Princi!COMAUT_USUAPR)
   pnl_Proveedor.Caption = Trim(g_rst_Princi!MaePrv_RazSoc)
   pnl_Descrip.Caption = Trim(g_rst_Princi!COMAUT_DESCRP & "")
   pnl_Banco.Caption = Trim(g_rst_Princi!NOM_BANCO & "")
   pnl_CtaCte.Caption = Trim(g_rst_Princi!COMAUT_CTACRR & "")
   pnl_Moneda.Caption = Trim(g_rst_Princi!Moneda)
   pnl_Moneda.Tag = g_rst_Princi!COMAUT_CODMON
   pnl_TotPag.Caption = Format(g_rst_Princi!COMAUT_IMPPAG, "###,###,##0.00")

   ReDim l_arr_CtaCte(0)
   
   If Left(l_str_CodOpe, 2) = "06" And l_int_TipOpe = 2 Then
      'SI ES REG.COMPRAS Y DETRACCION
      cmb_Banco.AddItem Trim(Trim(moddat_gf_Consulta_ParDes("122", Format(18, "000000")))) 'BANCO NACION
      cmb_Banco.ItemData(cmb_Banco.NewIndex) = 18
                 
      ReDim Preserve l_arr_CtaCte(UBound(l_arr_CtaCte) + 1)
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Codigo = 18
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(18, "000000")))
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Prefij = Trim(CStr(g_rst_Princi!MaePrv_CtaDet & ""))
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_TipMon = 1
   Else
      Call fs_CargarCtaCte
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
End Sub

Private Sub fs_CargarCtaCte()
   ReDim l_arr_CtaCte(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT *  "
   g_str_Parame = g_str_Parame & "   FROM  (SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN1 AS COD_BANCO, DECODE(MAEPRV_CODBNC_MN1,11, MAEPRV_CTACRR_MN1, MAEPRV_NROCCI_MN1) AS NUM_CUENTA  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "         Union  "
   g_str_Parame = g_str_Parame & "          SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN2, DECODE(MAEPRV_CODBNC_MN2,11, MAEPRV_CTACRR_MN2, MAEPRV_NROCCI_MN2)  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "         Union  "
   g_str_Parame = g_str_Parame & "          SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN3, DECODE(MAEPRV_CODBNC_MN3,11, MAEPRV_CTACRR_MN3, MAEPRV_NROCCI_MN3)  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "         Union  "
   g_str_Parame = g_str_Parame & "          SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL1, DECODE(MAEPRV_CODBNC_DL1,11, MAEPRV_CTACRR_DL1, MAEPRV_NROCCI_DL1)  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "         Union  "
   g_str_Parame = g_str_Parame & "          SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL2, DECODE(MAEPRV_CODBNC_DL2,11, MAEPRV_CTACRR_DL2, MAEPRV_NROCCI_DL2)  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "         Union  "
   g_str_Parame = g_str_Parame & "          SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL3, DECODE(MAEPRV_CODBNC_DL3,11, MAEPRV_CTACRR_DL3, MAEPRV_NROCCI_DL3)  "
   g_str_Parame = g_str_Parame & "            FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "           WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "             AND TRIM(MAEPRV_NUMDOC) = '" & l_str_NumDoc & "') A  "
   g_str_Parame = g_str_Parame & "  WHERE A.COD_BANCO > 0  "
   g_str_Parame = g_str_Parame & "    AND A.COD_MONEDA = " & CLng(pnl_Moneda.Tag)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      Call gs_SetFocus(cmb_Banco)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If

   ReDim l_arr_CtaCte(0)
   
   g_rst_GenAux.MoveFirst
   Do While Not g_rst_GenAux.EOF
   
      ReDim Preserve l_arr_CtaCte(UBound(l_arr_CtaCte) + 1)
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Codigo = Trim(g_rst_GenAux!COD_BANCO)
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!COD_BANCO, "000000")))
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_Prefij = Trim(g_rst_GenAux!NUM_CUENTA & "")
      l_arr_CtaCte(UBound(l_arr_CtaCte)).Genera_TipMon = g_rst_GenAux!COD_MONEDA
      
      g_rst_GenAux.MoveNext
   Loop

   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   '================================================
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM (SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN1 AS COD_BANCO  "
   g_str_Parame = g_str_Parame & "                  FROM CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 WHERE MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                Union  "
   g_str_Parame = g_str_Parame & "                SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN2  "
   g_str_Parame = g_str_Parame & "                  From CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 Where MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                Union  "
   g_str_Parame = g_str_Parame & "                SELECT 1 AS COD_MONEDA, MAEPRV_CODBNC_MN3  "
   g_str_Parame = g_str_Parame & "                  From CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 Where MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                Union  "
   g_str_Parame = g_str_Parame & "                SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL1  "
   g_str_Parame = g_str_Parame & "                  From CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 Where MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                Union  "
   g_str_Parame = g_str_Parame & "                SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL2  "
   g_str_Parame = g_str_Parame & "                  From CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 Where MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  = '" & l_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                Union  "
   g_str_Parame = g_str_Parame & "                SELECT 2 AS COD_MONEDA, MAEPRV_CODBNC_DL3  "
   g_str_Parame = g_str_Parame & "                  From CNTBL_MAEPRV  "
   g_str_Parame = g_str_Parame & "                 Where MAEPRV_TIPDOC = " & l_int_TipDoc
   g_str_Parame = g_str_Parame & "                   AND TRIM(MAEPRV_NUMDOC)  =  '" & l_str_NumDoc & "') A  "
   g_str_Parame = g_str_Parame & "  WHERE A.COD_BANCO > 0  "
   g_str_Parame = g_str_Parame & "    AND A.COD_MONEDA = " & CLng(pnl_Moneda.Tag)
   g_str_Parame = g_str_Parame & "  GROUP BY COD_MONEDA, COD_BANCO  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      Call gs_SetFocus(cmb_Banco)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   g_rst_GenAux.MoveFirst
   Do While Not g_rst_GenAux.EOF
      cmb_Banco.AddItem Trim(Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!COD_BANCO, "000000"))))
      cmb_Banco.ItemData(cmb_Banco.NewIndex) = g_rst_GenAux!COD_BANCO

      g_rst_GenAux.MoveNext
   Loop

   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   cmb_Banco.ListIndex = -1
   cmb_CtaCte.ListIndex = -1
End Sub

Private Sub cmb_Banco_Click()
Dim r_int_Contar  As Integer
Dim r_str_Cadena  As String

   cmb_CtaCte.Clear
   For r_int_Contar = 1 To UBound(l_arr_CtaCte)
       r_str_Cadena = ""
       If cmb_Banco.ItemData(cmb_Banco.ListIndex) = CLng(l_arr_CtaCte(r_int_Contar).Genera_Codigo) And _
          CLng(l_arr_CtaCte(r_int_Contar).Genera_TipMon = CLng(pnl_Moneda.Tag)) Then
          
          If Left(l_str_CodOpe, 2) = "06" And l_int_TipOpe = 2 Then
             'REGISTRO DE COMPRAS
             lbl_Cuenta.Caption = "Cuenta Detractora:"
          Else
             If cmb_Banco.ItemData(cmb_Banco.ListIndex) = 11 Then 'Banco continental
                lbl_Cuenta.Caption = "Cuenta Corriente:"
             Else
                lbl_Cuenta.Caption = "CCI:"
             End If
          End If
          
          cmb_CtaCte.AddItem Trim(l_arr_CtaCte(r_int_Contar).Genera_Prefij)
       End If
   Next
End Sub

Private Sub cmd_Grabar_Click()
    If cmb_Banco.ListIndex = -1 Then
        MsgBox "Debe de seleccionar un banco", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Banco)
        Exit Sub
    End If
    
    If cmb_CtaCte.ListIndex = -1 Then
        MsgBox "Debe de seleccionar una cuenta corriente", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_CtaCte)
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar()

   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " UPDATE CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "    SET COMAUT_CODBNC = " & cmb_Banco.ItemData(cmb_CtaCte.ListIndex) & ",  "
   g_str_Parame = g_str_Parame & "        COMAUT_CTACRR =  '" & Trim(cmb_CtaCte.Text) & "', "
   g_str_Parame = g_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "        SEGFECACT = " & Format(date, "yyyymmdd") & ",  "
   g_str_Parame = g_str_Parame & "        SEGHORACT = " & Format(Time, "HHmmss") & ",  "
   g_str_Parame = g_str_Parame & "        SEGPLTACT = '" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "        SEGTERACT = '" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "' "
   g_str_Parame = g_str_Parame & "  WHERE A.COMAUT_CODAUT = " & l_int_CodAut
   g_str_Parame = g_str_Parame & "    AND A.COMAUT_CODOPE = " & l_str_CodOpe
   g_str_Parame = g_str_Parame & "    AND A.COMAUT_TIPOPE = " & l_int_TipOpe
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
   Call frm_Ctb_PagCom_06.fs_Buscar
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

