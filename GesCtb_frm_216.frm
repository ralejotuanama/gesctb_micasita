VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15615
   Icon            =   "GesCtb_frm_216.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   15615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7155
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15645
      _Version        =   65536
      _ExtentX        =   27596
      _ExtentY        =   12621
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
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27331
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
            TabIndex        =   7
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registros de Pagos Aprobados"
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
            Picture         =   "GesCtb_frm_216.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   670
         Left            =   60
         TabIndex        =   8
         Top             =   780
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27331
         _ExtentY        =   1182
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.19
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmb_Rechazar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_216.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Rechazar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_216.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14880
            Picture         =   "GesCtb_frm_216.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_216.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_Modificar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_216.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5475
         Left            =   60
         TabIndex        =   9
         Top             =   1500
         Width           =   15495
         _Version        =   65536
         _ExtentX        =   27331
         _ExtentY        =   9657
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1140
            TabIndex        =   17
            Top             =   60
            Width           =   2120
            _Version        =   65536
            _ExtentX        =   3739
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Proceso"
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4290
            TabIndex        =   10
            Top             =   60
            Width           =   1380
            _Version        =   65536
            _ExtentX        =   2434
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Evaluador"
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
            Height          =   5055
            Left            =   30
            TabIndex        =   12
            Top             =   360
            Width           =   15450
            _ExtentX        =   27252
            _ExtentY        =   8916
            _Version        =   393216
            Rows            =   30
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   5655
            TabIndex        =   13
            Top             =   60
            Width           =   3120
            _Version        =   65536
            _ExtentX        =   5503
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proveedor"
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
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         Begin Threed.SSPanel pnl_Tit_SitIns 
            Height          =   285
            Left            =   13905
            TabIndex        =   15
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Pagar"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   11100
            TabIndex        =   16
            Top             =   60
            Width           =   1950
            _Version        =   65536
            _ExtentX        =   3440
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Corriente"
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
            Left            =   8760
            TabIndex        =   18
            Top             =   60
            Width           =   2350
            _Version        =   65536
            _ExtentX        =   4145
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción"
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
         Begin Threed.SSPanel pnl_Tit_IngIns 
            Height          =   285
            Left            =   13035
            TabIndex        =   19
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1552
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   3240
            TabIndex        =   11
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1870
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
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_Contar   As Integer

Private Sub cmb_Modificar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   frm_Ctb_PagCom_07.Show 1
End Sub

Private Sub cmb_Rechazar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea rechazar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT_ESTADO ( "
   g_str_Parame = g_str_Parame & " " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 9)) & ", " 'COMAUT_CODAUT
   g_str_Parame = g_str_Parame & " " & 3 & ", " 'RECHAZAR
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   If g_rst_Genera!RESUL = 1 Then
      'COMAUT_CODOPE
      MsgBox "Registro rechazado correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 0
   Call fs_Buscar
   Call gs_SetFocus(grd_Listad)
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

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call gs_LimpiaGrid(grd_Listad)
   
   grd_Listad.ColWidth(0) = 1080 'CODIGO
   grd_Listad.ColWidth(1) = 2100 'TIPO PROCESO
   grd_Listad.ColWidth(2) = 1050 'FECHA
   grd_Listad.ColWidth(3) = 1360 'USU_APRUEBA
   grd_Listad.ColWidth(4) = 3100 'PROVEEDOR
   grd_Listad.ColWidth(5) = 2350 'DESCRIPCION
   grd_Listad.ColWidth(6) = 1920 'CUENTA CORRIENTE
   grd_Listad.ColWidth(7) = 870  'MONEDA
   grd_Listad.ColWidth(8) = 1210 'TOTAL
   grd_Listad.ColWidth(9) = 0 'COMAUT_CODAUT
   grd_Listad.ColWidth(10) = 0 'COMAUT_TIPDOC
   grd_Listad.ColWidth(11) = 0 'COMAUT_NUMDOC
   grd_Listad.ColWidth(12) = 0 'COMAUT_CODMON
   grd_Listad.ColWidth(13) = 0 'COMAUT_CODBNC
   grd_Listad.ColWidth(14) = 0 'COMAUT_CTACTB
   grd_Listad.ColWidth(15) = 0 'COMAUT_DATCTB
   grd_Listad.ColWidth(16) = 0 'COMAUT_TIPOPE
   grd_Listad.ColWidth(17) = 0 'NRO-DOCUMENTO
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter 'CODIGO
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter 'PROVEEDOR
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter 'TOTAL
End Sub

Public Sub fs_Buscar()
Dim r_str_Cadena  As String

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)

  '--------------------------------------
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COMAUT_CODAUT, COMAUT_CODOPE, COMAUT_FECOPE, COMAUT_TIPDOC, COMAUT_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        COMAUT_CODMON, COMAUT_IMPPAG, COMAUT_CODBNC, COMAUT_CTACRR, COMAUT_CTACTB,  "
   g_str_Parame = g_str_Parame & "        COMAUT_DATCTB , COMAUT_CODEST, TRIM(C.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS TIPOPROCESO, COMAUT_USUAPR, "
   g_str_Parame = g_str_Parame & "        DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS MaePrv_RazSoc, A.COMAUT_TIPOPE, A.COMAUT_DESCRP  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMAUT_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMAUT_NUMDOC)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMAUT_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMAUT_NUMDOC) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.COMAUT_CODMON  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 136 AND TO_NUMBER(D.PARDES_CODITE) = TO_NUMBER(SUBSTR(LPAD(COMAUT_CODOPE,10,0),1,2)) AND D.PARDES_CODITE <> 0   "
   g_str_Parame = g_str_Parame & "  WHERE COMAUT_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND COMAUT_CODEST = 2  "
   g_str_Parame = g_str_Parame & "  ORDER BY COMAUT_FECOPE ASC, A.COMAUT_CODAUT ASC  "
  '---------------------------------------
  
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

   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      grd_Listad.Col = 0
      grd_Listad.Text = Format(Trim(g_rst_Princi!COMAUT_CODOPE), "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!TIPOPROCESO & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMAUT_FECOPE)
   
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!COMAUT_USUAPR & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
                                    
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(CStr(g_rst_Princi!COMAUT_DESCRP & ""))
                  
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!COMAUT_CTACRR & "")
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 8 'TOTAL A PAGAR
      grd_Listad.Text = Format(g_rst_Princi!COMAUT_IMPPAG, "###,###,###,##0.00")
            
      grd_Listad.Col = 9
      grd_Listad.Text = g_rst_Princi!COMAUT_CODAUT
      
      grd_Listad.Col = 10
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC
      
      grd_Listad.Col = 11
      grd_Listad.Text = g_rst_Princi!COMAUT_NUMDOC
      
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!COMAUT_CODMON
      
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(CStr(g_rst_Princi!COMAUT_CODBNC & ""))
      
      grd_Listad.Col = 14
      grd_Listad.Text = g_rst_Princi!COMAUT_CTACTB
      
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!COMAUT_DATCTB
      
      grd_Listad.Col = 16
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPOPE
      
      grd_Listad.Col = 17
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC & "-" & Trim(g_rst_Princi!COMAUT_NUMDOC & "")
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Consul_Click()
Dim r_str_CodAux   As String
Dim r_str_FlgAux   As Integer

   r_str_CodAux = ""
   r_str_FlgAux = 0
   moddat_g_str_NumOpe = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   Select Case Left(grd_Listad.TextMatrix(grd_Listad.Row, 0), 2)
          Case "01" 'CUENTAS X PAGAR
               moddat_g_str_NumOpe = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
               frm_Ctb_PagCom_04.Show 1
          Case "12" 'CUENTAS X PAGAR GESCTB
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CtaPag_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "07" 'GESTION PERSONAL
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_int_TipRec = 1 'GESTION DE PAGOS
                  frm_Ctb_GesPer_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "08" 'CARGA DEL ARCHIVO RECAUDO
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CarArc_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "06" 'REGISTRO DE COMPRAS
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 10))
                  moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 11))
                  moddat_g_int_InsAct = 0 'tipo registro compra
                  frm_Ctb_RegCom_04.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "05" 'ENTREGAS A RENDIR
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodMod = grd_Listad.TextMatrix(grd_Listad.Row, 12)
                  moddat_g_int_FlgGrb = 0
                  If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16) & "") = "1" Then
                     frm_Ctb_EntRen_02.Show 1 'form principal
                  ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16) & "") = "2" Then
                     frm_Ctb_EntRen_04.Show 1 'reembolso
                  End If
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case Else
               Exit Sub
   End Select
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE PAGOS APROBADOS"
      .Range(.Cells(2, 2), .Cells(2, 11)).Merge
      .Range(.Cells(2, 2), .Cells(2, 11)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 11)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "TIPO PROCESO"
      .Cells(3, 4) = "FECHA"
      .Cells(3, 5) = "EVALUADOR"
      .Cells(3, 6) = "NRO DOCUMENTO"
      .Cells(3, 7) = "PROVEEDOR"
      .Cells(3, 8) = "DESCRIPCIÓN"
      .Cells(3, 9) = "CUENTA CORRIENTE"
      .Cells(3, 10) = "MONEDA"
      .Cells(3, 11) = "TOTAL PAGAR"
         
      .Range(.Cells(3, 2), .Cells(3, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 11)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22 'tipo proceso
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 13 'fecha
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17 'EVALUDOR
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 17 'nro documento
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 45 'proveedor
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 22 'descripcion
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 22 'cuenta corriente
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 22 'moneda
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      .Columns("K").ColumnWidth = 14 'total a pagar
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 11)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 11)).Font.Size = 11
      
      r_int_NumFil = 4
      For l_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(l_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = grd_Listad.TextMatrix(l_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = "'" & grd_Listad.TextMatrix(l_int_Contar, 2)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(l_int_Contar, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(l_int_Contar, 17)
         .Cells(r_int_NumFil, 7) = grd_Listad.TextMatrix(l_int_Contar, 4)
         .Cells(r_int_NumFil, 8) = "'" & grd_Listad.TextMatrix(l_int_Contar, 5)
         .Cells(r_int_NumFil, 9) = "'" & grd_Listad.TextMatrix(l_int_Contar, 6)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(l_int_Contar, 7)
         .Cells(r_int_NumFil, 11) = grd_Listad.TextMatrix(l_int_Contar, 8)
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(3, 3), .Cells(3, 11)).HorizontalAlignment = xlHAlignCenter
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

