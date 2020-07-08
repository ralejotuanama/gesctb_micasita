VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_EntRen_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14820
   Icon            =   "GesCtb_frm_218.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14955
      _Version        =   65536
      _ExtentX        =   26379
      _ExtentY        =   10663
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
         TabIndex        =   1
         Top             =   60
         Width           =   14720
         _Version        =   65536
         _ExtentX        =   25964
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
            TabIndex        =   2
            Top             =   60
            Width           =   3645
            _Version        =   65536
            _ExtentX        =   6429
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Tipos de Pagos Asociados"
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
            Picture         =   "GesCtb_frm_218.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   650
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   14720
         _Version        =   65536
         _ExtentX        =   25964
         _ExtentY        =   1147
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.32
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14100
            Picture         =   "GesCtb_frm_218.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_218.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_Adicionar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_218.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4435
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   7823
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
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_FecEnt 
            Height          =   285
            Left            =   1140
            TabIndex        =   13
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha de ER"
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
            Height          =   4050
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   7144
            _Version        =   393216
            Rows            =   24
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1100
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro ER"
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
         Begin Threed.SSPanel pnl_Respon 
            Height          =   285
            Left            =   3255
            TabIndex        =   10
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Responsable"
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
            Left            =   9405
            TabIndex        =   11
            Top             =   60
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
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
         Begin Threed.SSPanel pnl_MtoAsig 
            Height          =   285
            Left            =   10260
            TabIndex        =   12
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Asignado"
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
            Left            =   7605
            TabIndex        =   14
            Top             =   60
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
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
         Begin Threed.SSPanel pnl_Benefi 
            Height          =   285
            Left            =   5430
            TabIndex        =   15
            Top             =   60
            Width           =   2190
            _Version        =   65536
            _ExtentX        =   3863
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Benficiario"
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
         Begin Threed.SSPanel pnl_TipPag 
            Height          =   285
            Left            =   2160
            TabIndex        =   16
            Top             =   60
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Pago"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   11430
            TabIndex        =   17
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1499
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Procesado"
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
         Begin Threed.SSPanel pnl_FecPag 
            Height          =   285
            Left            =   12270
            TabIndex        =   18
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
         Begin Threed.SSPanel pnl_CodPag 
            Height          =   285
            Left            =   13335
            TabIndex        =   19
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Pago"
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
Attribute VB_Name = "frm_Ctb_EntRen_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Consul_Click()
Dim r_str_Codigo As String
Dim r_int_FlgGrb As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_Codigo = moddat_g_str_Codigo
   r_int_FlgGrb = moddat_g_int_FlgGrb

   moddat_g_str_Codigo = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   moddat_g_int_FlgGrb = 0 'consultar
   frm_Ctb_EntRen_02.Show 1
   moddat_g_str_Codigo = r_str_Codigo
   moddat_g_int_FlgGrb = r_int_FlgGrb
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1080 'Nro Caja
   grd_Listad.ColWidth(1) = 1030 'Fecha caja
   grd_Listad.ColWidth(2) = 1080 'tipo pago - 1100
   grd_Listad.ColWidth(3) = 2190 'Responsable
   grd_Listad.ColWidth(4) = 2160 'Beneficiario
   grd_Listad.ColWidth(5) = 1800 'Glosa
   grd_Listad.ColWidth(6) = 850 'Moneda
   grd_Listad.ColWidth(7) = 1180 'Mto Asignado
   grd_Listad.ColWidth(8) = 830 'Procesado
   grd_Listad.ColWidth(9) = 1060  'fecha pago
   grd_Listad.ColWidth(10) = 1000 'codigo pago
   grd_Listad.ColWidth(11) = 0  'fecha pago
   grd_Listad.ColWidth(12) = 0  'Mto Asignado
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   'g_str_Parame = ""
   'g_str_Parame = g_str_Parame & " SELECT CAJCHC_TIPDOC_2, CAJCHC_NUMDOC_2  "
   'g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   'g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_CODCAJ = " & CLng(moddat_g_str_Codigo)
   'g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB = 2 "

   'If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
   '   Screen.MousePointer = 0
   '   Exit Sub
   'End If
   
   'If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
   '   'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
   '   g_rst_GenAux.Close
   '   Set g_rst_GenAux = Nothing
   '   Screen.MousePointer = 0
   '   Exit Sub
   'End If
  '
   'If Trim(g_rst_GenAux!CAJCHC_NUMDOC_2 & "") = "" Then
   '   Exit Sub
   'End If

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_TIPDOC || '-' || A.CAJCHC_NUMDOC NUMDOC_RESPON, TRIM(B.MAEPRV_RAZSOC) NOM_RESPON,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_CODMON, TRIM(C.PARDES_DESCRI) MONEDA, A.CAJCHC_IMPORT, A.CAJCHC_FLGPRC,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_DESCRI, A.CAJCHC_NUMOPE,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.MAEPRV_RAZSOC) AS NOM_BENEFI, TRIM(F.PARDES_DESCRI) AS TIPO_PAGO, A.CAJCHC_TIPPAG,  "
   g_str_Parame = g_str_Parame & "        j.COMPAG_FECPAG , j.COMPAG_CODCOM  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CAJCHC_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV D ON D.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC_2 AND D.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC_2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 138 AND F.PARDES_CODITE = A.CAJCHC_TIPPAG  "  'tipo pago
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.CAJCHC_CODCAJ) AND H.COMAUT_TIPOPE = 1 AND H.COMAUT_CODEST NOT IN (3)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                                 AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_FLGPRC = 0 "  'sin procesar
   'g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPDOC_2 = " & g_rst_GenAux!CAJCHC_TIPDOC_2 'beneficiario
   'g_str_Parame = g_str_Parame & "    AND A.CAJCHC_NUMDOC_2 = '" & Trim(CStr(g_rst_GenAux!CAJCHC_NUMDOC_2)) & "' " 'beneficiario
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB = 2  " 'entregas a rendir
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODREF_1 IS NULL  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODREF_2 IS NULL  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODMON = " & CLng(moddat_g_str_CodMod)
   g_str_Parame = g_str_Parame & "    AND NVL(A.CAJCHC_TIPPAG,0) = 1  " 'SOLO ANTICIPOS
   g_str_Parame = g_str_Parame & "    AND J.COMPAG_CODCOM IS NOT NULL  " 'SOLO ANTCIPOS PAGADOS
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODCAJ NOT IN (" & CLng(moddat_g_str_Codigo) & ")"
   g_str_Parame = g_str_Parame & "    AND NVL((SELECT COUNT(*) FROM CNTBL_CAJCHC_DET B  "
   g_str_Parame = g_str_Parame & "             Where B.CajDet_CodCaj = A.CajChc_CodCaj  "
   g_str_Parame = g_str_Parame & "               AND B.CAJDET_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "               AND B.CAJDET_SITUAC = 1),0) = 0  " 'no tiene monto rendido(facturas)
   g_str_Parame = g_str_Parame & "    AND NVL((SELECT COUNT(*) FROM CNTBL_CAJCHC C"
   g_str_Parame = g_str_Parame & "              WHERE C.CajChc_CodCaj = A.CajChc_CodCaj"
   g_str_Parame = g_str_Parame & "                AND C.CAJCHC_TIPTAB IN (4,5)"
   g_str_Parame = g_str_Parame & "                AND C.CAJCHC_SITUAC = 1),0) = 0 " '--DEVOLUCION Y REEMBOLSO (no tiene)
   g_str_Parame = g_str_Parame & "  ORDER BY CAJCHC_CODCAJ ASC  "
      
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
      grd_Listad.Text = Format(CStr(g_rst_Princi!CajChc_CodCaj), "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)

      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!TIPO_PAGO & "")

      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!NOM_RESPON & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!NOM_BENEFI & "")
                  
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!CajChc_Descri & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
      
      grd_Listad.Col = 7 'MTO ASIGNADO
      grd_Listad.Text = Format(g_rst_Princi!CajChc_Import, "###,###,###,##0.00")
      
      '--------------------------------------------------------------------------------------------------
      grd_Listad.Col = 8
      grd_Listad.Text = IIf(g_rst_Princi!CAJCHC_FLGPRC = 1, "SI", "NO")
            
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 9
         grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_Listad.Col = 10
         grd_Listad.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
      
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 11
         grd_Listad.Text = g_rst_Princi!COMPAG_FECPAG
      End If
      
      grd_Listad.Col = 12 'MTO ASIGNADO
      grd_Listad.Text = g_rst_Princi!CajChc_Import
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_Listad)
   grd_Listad.Redraw = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub cmb_Adicionar_Click()
Dim r_int_Fila  As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
    If MsgBox("¿Esta seguro de adicionar el registro seleccionado?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
   
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_ASOC ( "
    g_str_Parame = g_str_Parame & CLng(moddat_g_str_Codigo) & ", "
    g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & ", "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
    g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
    g_str_Parame = g_str_Parame & CStr(1) & ") "
       
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
    End If

    MsgBox "El registro se grabo correctamente.", vbInformation, modgen_g_str_NomPlt
    Call frm_Ctb_EntRen_01.fs_BuscarCaja
    Call frm_Ctb_EntRen_03.fs_BuscarAsoc
    Unload Me
End Sub

Private Sub pnl_MtoAsig_Click()
   If pnl_MtoAsig.Tag = "" Then
      pnl_MtoAsig.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 12, "N")
   Else
      pnl_MtoAsig.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 12, "N-")
   End If
End Sub

Private Sub pnl_Respon_Click()
   If pnl_Respon.Tag = "" Then
      pnl_Respon.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Respon.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Benefi_Click()
   If pnl_Benefi.Tag = "" Then
      pnl_Benefi.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Benefi.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Glosa_Click()
   If pnl_Glosa.Tag = "" Then
      pnl_Glosa.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Glosa.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_FecPag_Click()
   If pnl_FecPag.Tag = "" Then
      pnl_FecPag.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 11, "N")
   Else
      pnl_FecPag.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 11, "N-")
   End If
End Sub

Private Sub pnl_CodPag_Click()
   If pnl_CodPag.Tag = "" Then
      pnl_CodPag.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 10, "N")
   Else
      pnl_CodPag.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 10, "N-")
   End If
End Sub

