VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_32 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   1620
   ClientTop       =   1620
   ClientWidth     =   14235
   Icon            =   "GesCtb_frm_820.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7605
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14250
      _Version        =   65536
      _ExtentX        =   25135
      _ExtentY        =   13414
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   45
         TabIndex        =   8
         Top             =   60
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   330
            Left            =   570
            TabIndex        =   9
            Top             =   180
            Width           =   2010
            _Version        =   65536
            _ExtentX        =   3545
            _ExtentY        =   582
            _StockProps     =   15
            Caption         =   "Reporte de Indicadores"
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
            Autosize        =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_820.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   45
         TabIndex        =   10
         Top             =   780
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   15
            Picture         =   "GesCtb_frm_820.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Procesar información"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   615
            Picture         =   "GesCtb_frm_820.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13530
            Picture         =   "GesCtb_frm_820.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1305
         Left            =   45
         TabIndex        =   11
         Top             =   1470
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
         _ExtentY        =   2302
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   150
            Width           =   8655
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   510
            Width           =   3855
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1680
            TabIndex        =   4
            Top             =   870
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reporte:"
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Periodo:"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   510
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año:"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   870
            Width           =   330
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4710
         Left            =   45
         TabIndex        =   14
         Top             =   2820
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listado 
            Height          =   4575
            Left            =   45
            TabIndex        =   6
            Top             =   60
            Width           =   14040
            _ExtentX        =   24765
            _ExtentY        =   8070
            _Version        =   393216
            Rows            =   16
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_PerMes     As Integer
Dim l_int_PerAno     As Integer
Dim l_str_FecLim     As String
Dim l_str_FecIni     As String
Dim l_str_FecFin     As String
Dim l_dbl_TipCam     As Double
Dim l_str_CieMes     As String
Dim l_str_CieAno     As String

Private Sub fs_Procesar_FactElectronica()
Dim r_str_Parame  As String
Dim r_lng_FecIni  As Long
Dim r_lng_FecFin  As Long
Dim r_lng_FecAux  As Long
Dim r_int_FilIni  As Integer
Dim r_int_FilFin  As Integer
Dim r_int_NumFil  As Integer
Dim r_int_filAux  As Integer
Dim r_rst_GenAux  As ADODB.Recordset

   r_lng_FecIni = l_int_PerAno & Format(l_int_PerMes, "00") & "01"
   r_lng_FecFin = l_int_PerAno & Format(l_int_PerMes, "00") & Format(ff_Ultimo_Dia_Mes(l_int_PerMes, l_int_PerAno), "00")
   r_lng_FecAux = l_int_PerAno & Format(l_int_PerMes, "00")
      
   grd_Listado.Rows = 1
   Call gs_LimpiaGrid(grd_Listado)
   Call fs_SeteaColumnas_FactElectronica
   DoEvents
   
   'Consulta informacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT *  "
   g_str_Parame = g_str_Parame & "   FROM (SELECT SUBSTR(A.HIPCUO_FECPAG,7,2)||'/'||SUBSTR(A.HIPCUO_FECPAG,5,2)||'/'||SUBSTR(A.HIPCUO_FECPAG,1,4) AS COL_01,  "
   g_str_Parame = g_str_Parame & "                '13' AS COL_02, '00000000000000000017' AS COL_03, lpad(TRIM(C.HIPPAG_NUMMOV),20,'0') AS COL_04,  "
   g_str_Parame = g_str_Parame & "                A.HIPCUO_INTERE AS COL_05, A.HIPCUO_IMPPAG AS COL_06, DECODE(B.HIPMAE_MONEDA, 1, 'PEN','USD') AS COL_07,  "
   g_str_Parame = g_str_Parame & "                TRIM(B.HIPMAE_TDOCLI) AS COL_08, TRIM(B.HIPMAE_NDOCLI) AS COL_09,  "
   g_str_Parame = g_str_Parame & "                TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS COL_10, '13' AS COL_11,  "
   g_str_Parame = g_str_Parame & "                '-' AS COL_12, '-' AS COL_13,  "
   g_str_Parame = g_str_Parame & "                SUBSTR(B.HIPMAE_FECDES,7,2)||'/'||SUBSTR(B.HIPMAE_FECDES,5,2)||'/'||SUBSTR(B.HIPMAE_FECDES,1,4) AS COL_14,  "
   g_str_Parame = g_str_Parame & "                A.HIPCUO_NUMOPE AS COL_15, '1' AS COL_16,  "
   g_str_Parame = g_str_Parame & "                CASE WHEN E.HIPGAR_PARFIC IS NULL THEN '-' ELSE TRIM(E.HIPGAR_PARFIC) END AS COL_17,  "
   g_str_Parame = g_str_Parame & "                TRIM(DECODE(G.SOLINM_TIPVIA, 12, '', TRIM(R.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMVIA)||' '||TRIM(G.SOLINM_NUMVIA)||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_INTDPT)), 0), 0, '', '('||TRIM(G.SOLINM_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_NOMZON)),0), 0, '', ' - '||DECODE(G.SOLINM_TIPZON, 12, '', TRIM(S.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMZON))) AS COL_18,  "
   g_str_Parame = g_str_Parame & "                CASE WHEN B.HIPMAE_FECDES > 20121231 THEN '3' ELSE '0' END AS COL_19  "
   g_str_Parame = g_str_Parame & "           FROM CRE_HIPCUO A  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE AND B.HIPMAE_SITUAC = 2 AND B.HIPMAE_TDOCLI = 1  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_HIPPAG C ON C.HIPPAG_NUMOPE = A.HIPCUO_NUMOPE AND C.HIPPAG_NUMCUO = A.HIPCUO_NUMCUO AND C.HIPPAG_NUMPAG = 1  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = B.HIPMAE_NDOCLI  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN CRE_HIPGAR E ON E.HIPGAR_NUMOPE = A.HIPCUO_NUMOPE  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLMAE F ON F.SOLMAE_NUMERO = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLINM G ON G.SOLINM_NUMSOL = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES R ON R.PARDES_CODGRP = 201 AND R.PARDES_CODITE = G.SOLINM_TIPVIA  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 202 AND S.PARDES_CODITE = G.SOLINM_TIPZON  "
   g_str_Parame = g_str_Parame & "          WHERE A.HIPCUO_TIPCRO = 1  "
   g_str_Parame = g_str_Parame & "            AND A.HIPCUO_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "            AND SUBSTR(A.HIPCUO_FECPAG,1,6) =  " & r_lng_FecAux
   'g_str_Parame = g_str_Parame & "            AND A.HIPCUO_FECPAG >=  " & r_lng_FecIni
   'g_str_Parame = g_str_Parame & "            AND A.HIPCUO_FECPAG <=  " & r_lng_FecFin
   'g_str_Parame = g_str_Parame & "            --UN PERIODO  "
   g_str_Parame = g_str_Parame & "         Union All  "
   g_str_Parame = g_str_Parame & "         SELECT SUBSTR(A.PPGCAB_FECPPG,7,2)||'/'||SUBSTR(A.PPGCAB_FECPPG,5,2)||'/'||SUBSTR(A.PPGCAB_FECPPG,1,4) AS COL_01,  "
   g_str_Parame = g_str_Parame & "                '13' AS COL_02, '00000000000000000017' AS COL_03, '00000000000000000000' AS COL_04,  "
   g_str_Parame = g_str_Parame & "                A.PPGCAB_INTCAL_TNC AS COL_05, A.PPGCAB_MTODEP AS COL_06, DECODE(B.HIPMAE_MONEDA, 1, 'PEN','USD') AS COL_07,  "
   g_str_Parame = g_str_Parame & "                TRIM(B.HIPMAE_TDOCLI) AS COL_08, TRIM(B.HIPMAE_NDOCLI) AS COL_09,  "
   g_str_Parame = g_str_Parame & "                TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS COL_10, '13' AS COL_11,  "
   g_str_Parame = g_str_Parame & "                '-' AS COL_12, '-' AS COL_13,  "
   g_str_Parame = g_str_Parame & "                SUBSTR(B.HIPMAE_FECDES,7,2)||'/'||SUBSTR(B.HIPMAE_FECDES,5,2)||'/'||SUBSTR(B.HIPMAE_FECDES,1,4) AS COL_14,  "
   g_str_Parame = g_str_Parame & "                A.PPGCAB_NUMOPE AS COL_15, '1' AS COL_16,  "
   g_str_Parame = g_str_Parame & "                CASE WHEN E.HIPGAR_PARFIC IS NULL THEN '-' ELSE TRIM(E.HIPGAR_PARFIC) END AS COL_17,  "
   g_str_Parame = g_str_Parame & "                TRIM(DECODE(G.SOLINM_TIPVIA, 12, '', TRIM(R.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMVIA)||' '||TRIM(G.SOLINM_NUMVIA)||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_INTDPT)), 0), 0, '', '('||TRIM(G.SOLINM_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_NOMZON)),0), 0, '', ' - '||DECODE(G.SOLINM_TIPZON, 12, '', TRIM(S.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMZON))) AS COL_18,  "
   g_str_Parame = g_str_Parame & "                CASE WHEN B.HIPMAE_FECDES > 20121231 THEN '3' ELSE '0' END AS COL_19  "
   g_str_Parame = g_str_Parame & "           FROM CRE_PPGCAB A  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.PPGCAB_NUMOPE AND B.HIPMAE_SITUAC = 2 AND B.HIPMAE_TDOCLI = 1  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = B.HIPMAE_NDOCLI  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN CRE_HIPGAR E ON E.HIPGAR_NUMOPE = A.PPGCAB_NUMOPE  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLMAE F ON F.SOLMAE_NUMERO = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLINM G ON G.SOLINM_NUMSOL = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES R ON R.PARDES_CODGRP = 201 AND R.PARDES_CODITE = G.SOLINM_TIPVIA  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 202 AND S.PARDES_CODITE = G.SOLINM_TIPZON  "
   g_str_Parame = g_str_Parame & "          WHERE SUBSTR(A.PPGCAB_FECPPG,1,6) =  " & r_lng_FecAux
   'g_str_Parame = g_str_Parame & "          WHERE A.PPGCAB_FECPPG >= " & r_lng_FecIni
   'g_str_Parame = g_str_Parame & "            AND A.PPGCAB_FECPPG <= " & r_lng_FecFin
   'g_str_Parame = g_str_Parame & "            --UN PERIODO  "
   g_str_Parame = g_str_Parame & "            AND A.PPGCAB_INTCAL_TNC > 0  "
   g_str_Parame = g_str_Parame & "         Union All  "
   g_str_Parame = g_str_Parame & "         SELECT SUBSTR(A.HIPCUO_FECPAG,7,2)||'/'||SUBSTR(A.HIPCUO_FECPAG,5,2)||'/'||SUBSTR(A.HIPCUO_FECPAG,1,4) AS COL_01,  "
   g_str_Parame = g_str_Parame & "                        '13' AS COL_02, '00000000000000000017' AS COL_03, lpad(TRIM(C.HIPPAG_NUMMOV),20,'0') AS COL_04,  "
   g_str_Parame = g_str_Parame & "                        A.HIPCUO_INTERE AS COL_05, A.HIPCUO_IMPPAG AS COL_06, DECODE(B.HIPMAE_MONEDA, 1, 'PEN','USD') AS COL_07,  "
   g_str_Parame = g_str_Parame & "                        TRIM(B.HIPMAE_TDOCLI) AS COL_08, TRIM(B.HIPMAE_NDOCLI) AS COL_09,  "
   g_str_Parame = g_str_Parame & "                        TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS COL_10, '13' AS COL_11,  "
   g_str_Parame = g_str_Parame & "                        '-' AS COL_12, '-' AS COL_13,  "
   g_str_Parame = g_str_Parame & "                        SUBSTR(B.HIPMAE_FECDES,7,2)||'/'||SUBSTR(B.HIPMAE_FECDES,5,2)||'/'||SUBSTR(B.HIPMAE_FECDES,1,4) AS COL_14,  "
   g_str_Parame = g_str_Parame & "                        A.HIPCUO_NUMOPE AS COL_15, '1' AS COL_16,  "
   g_str_Parame = g_str_Parame & "                        CASE WHEN E.HIPGAR_PARFIC IS NULL THEN '-' ELSE TRIM(E.HIPGAR_PARFIC) END AS COL_17,  "
   g_str_Parame = g_str_Parame & "                        TRIM(DECODE(G.SOLINM_TIPVIA, 12, '', TRIM(R.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMVIA)||' '||TRIM(G.SOLINM_NUMVIA)||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_INTDPT)), 0), 0, '', '('||TRIM(G.SOLINM_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(G.SOLINM_NOMZON)),0), 0, '', ' - '||DECODE(G.SOLINM_TIPZON, 12, '', TRIM(S.PARDES_DESCRI))||' '||TRIM(G.SOLINM_NOMZON))) AS COL_18,  "
   g_str_Parame = g_str_Parame & "                        CASE WHEN B.HIPMAE_FECDES > 20121231 THEN '3' ELSE '0' END AS COL_19  "
   g_str_Parame = g_str_Parame & "           FROM CRE_HIPCUO A  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE AND B.HIPMAE_SITUAC IN (6,9) AND B.HIPMAE_TDOCLI = 1  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_HIPPAG C ON C.HIPPAG_NUMOPE = A.HIPCUO_NUMOPE AND C.HIPPAG_NUMCUO = A.HIPCUO_NUMCUO AND C.HIPPAG_NUMPAG = 1  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = B.HIPMAE_NDOCLI  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN CRE_HIPGAR E ON E.HIPGAR_NUMOPE = A.HIPCUO_NUMOPE  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLMAE F ON F.SOLMAE_NUMERO = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLINM G ON G.SOLINM_NUMSOL = B.HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES R ON R.PARDES_CODGRP = 201 AND R.PARDES_CODITE = G.SOLINM_TIPVIA  "
   g_str_Parame = g_str_Parame & "           LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 202 AND S.PARDES_CODITE = G.SOLINM_TIPZON  "
   g_str_Parame = g_str_Parame & "          Where A.HIPCUO_TIPCRO = 1  "
   g_str_Parame = g_str_Parame & "            AND A.HIPCUO_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "            AND SUBSTR(A.HIPCUO_FECPAG,1,6) =  " & r_lng_FecAux
   'g_str_Parame = g_str_Parame & "            AND A.HIPCUO_FECPAG >= 20170101  "
   g_str_Parame = g_str_Parame & "            AND A.HIPCUO_FECPAG <= B.HIPMAE_FECCAN)  "
   g_str_Parame = g_str_Parame & "  ORDER BY COL_01  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Me.MousePointer = vbDefault
      grd_Listado.Redraw = True
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

'   'ubicar codigo
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT RPTTAB_VALNUM01, RPTTAB_VALNUM02 "
'   g_str_Parame = g_str_Parame & "   FROM RPT_TABLAS A "
'   g_str_Parame = g_str_Parame & "  Where A.RPTTAB_PERMES = " & l_int_PerMes
'   g_str_Parame = g_str_Parame & "    AND A.RPTTAB_PERANO = " & l_int_PerAno
'   g_str_Parame = g_str_Parame & "    AND A.RPTTAB_NOMBRE = 'FACTURA_ELECTRONICA' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, r_rst_GenAux, 3) Then
'      Me.MousePointer = vbDefault
'      grd_Listado.Redraw = True
'      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
'      Exit Sub
'   End If
'
'   If Not (r_rst_GenAux.BOF And r_rst_GenAux.EOF) Then
'      r_rst_GenAux.MoveFirst
'      r_int_FilIni = r_rst_GenAux!RPTTAB_VALNUM01
'      r_int_FilFin = r_rst_GenAux!RPTTAB_VALNUM02
'   Else
'      g_str_Parame = ""
'      g_str_Parame = g_str_Parame & " SELECT NVL((SELECT RPTTAB_VALNUM02 "
'      g_str_Parame = g_str_Parame & "               FROM (SELECT RPTTAB_VALNUM02 "
'      g_str_Parame = g_str_Parame & "                       FROM RPT_TABLAS A "
'      g_str_Parame = g_str_Parame & "                      WHERE A.RPTTAB_NOMBRE = 'FACTURA_ELECTRONICA' "
'      g_str_Parame = g_str_Parame & "                      ORDER BY RPTTAB_VALNUM02 DESC) A "
'      g_str_Parame = g_str_Parame & "              WHERE ROWNUM = 1),0) RPTTAB_VALNUM02 "
'      g_str_Parame = g_str_Parame & "   FROM DUAL "
'
'      If Not gf_EjecutaSQL(g_str_Parame, r_rst_GenAux, 3) Then
'         Me.MousePointer = vbDefault
'         grd_Listado.Redraw = True
'         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
'         Exit Sub
'      End If
'
'      r_int_FilIni = r_rst_GenAux!RPTTAB_VALNUM02 + 1
'      r_int_FilFin = r_int_FilIni
'
'      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
'         g_rst_Princi.MoveFirst
'         Do While Not g_rst_Princi.EOF
'            r_int_FilFin = r_int_FilFin + 1
'            g_rst_Princi.MoveNext
'         Loop
'         r_int_FilFin = r_int_FilFin - 1
'      End If
            
'      Call moddat_gs_FecSis
'
'      g_str_Parame = ""
'      g_str_Parame = g_str_Parame & "INSERT INTO RPT_TABLAS("
'      g_str_Parame = g_str_Parame & "      RPTTAB_PERMES,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_PERANO,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_CODIGO,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_NOMBRE,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_SITUAC,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_VALNUM01,"
'      g_str_Parame = g_str_Parame & "      RPTTAB_VALNUM02,"
'      g_str_Parame = g_str_Parame & "      SEGUSUCRE, "
'      g_str_Parame = g_str_Parame & "      SEGFECCRE, "
'      g_str_Parame = g_str_Parame & "      SEGHORCRE, "
'      g_str_Parame = g_str_Parame & "      SEGPLTCRE, "
'      g_str_Parame = g_str_Parame & "      SEGTERCRE, "
'      g_str_Parame = g_str_Parame & "      SEGSUCCRE) "
'      g_str_Parame = g_str_Parame & "  VALUES("
'      g_str_Parame = g_str_Parame & l_int_PerMes & ","
'      g_str_Parame = g_str_Parame & l_int_PerAno & ","
'      g_str_Parame = g_str_Parame & "(SELECT NVL(MAX(A.RPTTAB_CODIGO),0) + 1 FROM RPT_TABLAS A WHERE  A.RPTTAB_NOMBRE = 'FACTURA_ELECTRONICA'),"
'      g_str_Parame = g_str_Parame & "'FACTURA_ELECTRONICA',"
'      g_str_Parame = g_str_Parame & "1,"
'      g_str_Parame = g_str_Parame & r_int_FilIni & ","
'      g_str_Parame = g_str_Parame & r_int_FilFin & ","
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
'      g_str_Parame = g_str_Parame & "'" & Format(moddat_g_str_FecSis, "yyyymmdd") & "', "
'      g_str_Parame = g_str_Parame & "'" & Format(moddat_g_str_HorSis, "HHmmss") & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
'      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
'
'      If Not gf_EjecutaSQL(g_str_Parame, r_rst_GenAux, 2) Then
'         Exit Sub
'      End If
'   End If

   grd_Listado.Row = 1
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_NumFil = 1
'      r_int_filAux = r_int_FilIni
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listado.TextMatrix(r_int_NumFil, 0) = r_int_NumFil
         grd_Listado.TextMatrix(r_int_NumFil, 1) = Trim(g_rst_Princi!COL_01)
         grd_Listado.TextMatrix(r_int_NumFil, 2) = Trim(g_rst_Princi!COL_02)
         grd_Listado.TextMatrix(r_int_NumFil, 3) = "-" 'Trim(g_rst_Princi!COL_03)
         grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(r_int_NumFil, "00000000000000000000") 'Trim(g_rst_Princi!COL_04)
 '        If r_int_filAux <= r_int_FilFin Then
 '           grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(r_int_filAux, "00000000000000000000") 'Trim(g_rst_Princi!COL_04)
 '        Else
 '           grd_Listado.TextMatrix(r_int_NumFil, 4) = "                    " 'Trim(g_rst_Princi!COL_04)
 '        End If
         grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!COL_05, "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!COL_06, "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 7) = Trim(g_rst_Princi!COL_07)
         grd_Listado.TextMatrix(r_int_NumFil, 8) = Trim(g_rst_Princi!COL_08)
         grd_Listado.TextMatrix(r_int_NumFil, 9) = Trim(g_rst_Princi!COL_09)
         grd_Listado.TextMatrix(r_int_NumFil, 10) = Trim(g_rst_Princi!COL_10)
         grd_Listado.TextMatrix(r_int_NumFil, 11) = Trim(g_rst_Princi!COL_11)
         grd_Listado.TextMatrix(r_int_NumFil, 12) = Trim(g_rst_Princi!COL_12)
         grd_Listado.TextMatrix(r_int_NumFil, 13) = Trim(g_rst_Princi!COL_13)
         grd_Listado.TextMatrix(r_int_NumFil, 14) = Trim(g_rst_Princi!COL_14)
         grd_Listado.TextMatrix(r_int_NumFil, 15) = Trim(g_rst_Princi!COL_15)
         grd_Listado.TextMatrix(r_int_NumFil, 16) = Trim(g_rst_Princi!COL_16)
         grd_Listado.TextMatrix(r_int_NumFil, 17) = Trim(g_rst_Princi!COL_17)
         grd_Listado.TextMatrix(r_int_NumFil, 18) = Left(Trim(g_rst_Princi!COL_18), 100)
         grd_Listado.TextMatrix(r_int_NumFil, 19) = Trim(g_rst_Princi!COL_19)
         
         r_int_NumFil = r_int_NumFil + 1
         'r_int_filAux = r_int_filAux + 1
         grd_Listado.Rows = grd_Listado.Rows + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   grd_Listado.Rows = grd_Listado.Rows - 1
   
   Call fs_Activa(True)
   grd_Listado.Redraw = True
   Me.MousePointer = vbDefault
End Sub

Private Sub fs_ProcesarData_Financieros()
Dim r_int_NumFil  As Integer
   
   'valida ingresos
   If ipp_PerAno.Text <= 2015 And cmb_PerMes.ListIndex <= 4 Then
      MsgBox "Seleccione Año a partir del 2015 y Mes a partir de Junio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Screen.MousePointer = 0
      Exit Sub
   End If
   If ipp_PerAno.Text < 2015 Then
      MsgBox "Seleccione Año a partir del 2015.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_Listado.Redraw = False
   Me.MousePointer = vbHourglass
   
   'Procesa infomacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_RPT_RATIOFINANCIEROS("
   g_str_Parame = g_str_Parame & CInt(l_int_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(l_int_PerAno) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'REPORTE RATIOS FINANCIEROS', "
   g_str_Parame = g_str_Parame & "0)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Me.MousePointer = vbDefault
      grd_Listado.Redraw = True
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   fs_SeteaColumnas_Financieros
   
   'Consulta informacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_PERMES = '" & CInt(l_int_PerMes) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_PERANO = '" & CInt(l_int_PerAno) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'REPORTE RATIOS FINANCIEROS' "
   g_str_Parame = g_str_Parame & "    AND RPT_MONEDA = 0 "
   g_str_Parame = g_str_Parame & "  ORDER BY TO_NUMBER(RPT_CODIGO)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Me.MousePointer = vbDefault
      grd_Listado.Redraw = True
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_NumFil = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listado.TextMatrix(r_int_NumFil, 0) = Trim(g_rst_Princi!RPT_DESCRI)
         grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), "", g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), "", g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), "", g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), "", g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), "", g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), "", g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), "", g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), "", g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM09), "", g_rst_Princi!RPT_VALNUM09), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM10), "", g_rst_Princi!RPT_VALNUM10), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM11), "", g_rst_Princi!RPT_VALNUM11), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 12) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM12), "", g_rst_Princi!RPT_VALNUM12), "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_NumFil, 13) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM13), "", g_rst_Princi!RPT_VALNUM13), "###,###,##0.00")
         
         r_int_NumFil = r_int_NumFil + 1
         grd_Listado.Rows = grd_Listado.Rows + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listado.Rows = grd_Listado.Rows - 1
   
   With grd_Listado
      .Col = 0
      .Row = 1: .CellFontBold = True
      .Row = 5: .CellFontBold = True
      .Row = 11: .CellFontBold = True
      .Row = 15: .CellFontBold = True
      .Row = 21: .CellFontBold = True
   End With
         
   Call fs_Activa(True)
   grd_Listado.Redraw = True
   Me.MousePointer = vbDefault
End Sub

Private Sub fs_ProcesarData_Indicadores()
Dim l_str_PerMes     As String
Dim l_str_PerAno     As String
Dim r_dbl_TotSal     As Double
Dim r_dbl_TotCal     As Double
Dim r_dbl_TotCar     As Double
Dim r_dbl_PatEfec    As Double
Dim r_dbl_ValNum     As Double
Dim r_int_Column     As Integer

   grd_Listado.Rows = 1
   Call gs_LimpiaGrid(grd_Listado)
   Call fs_SeteaColumnas_Indicadores
   DoEvents
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = gf_Query(ipp_PerAno.Text)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listado.Row = 1
   Select Case l_int_PerMes
      Case 1
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
      Case 2
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
      Case 3
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
      Case 4
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
      Case 5
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
      Case 6
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
      Case 7
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
      Case 8
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 8) = Format(g_rst_Princi!TOT_08, "###,###,##0.00")
      Case 9
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 8) = Format(g_rst_Princi!TOT_08, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 9) = Format(g_rst_Princi!TOT_09, "###,###,##0.00")
      Case 10
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 8) = Format(g_rst_Princi!TOT_08, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 9) = Format(g_rst_Princi!TOT_09, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 10) = Format(g_rst_Princi!TOT_10, "###,###,##0.00")
      Case 11
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 8) = Format(g_rst_Princi!TOT_08, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 9) = Format(g_rst_Princi!TOT_09, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 10) = Format(g_rst_Princi!TOT_10, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 11) = Format(g_rst_Princi!TOT_11, "###,###,##0.00")
      Case 12
         grd_Listado.TextMatrix(1, 1) = Format(g_rst_Princi!TOT_01, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 2) = Format(g_rst_Princi!TOT_02, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 3) = Format(g_rst_Princi!TOT_03, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 4) = Format(g_rst_Princi!TOT_04, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 5) = Format(g_rst_Princi!TOT_05, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 6) = Format(g_rst_Princi!TOT_06, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 7) = Format(g_rst_Princi!TOT_07, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 8) = Format(g_rst_Princi!TOT_08, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 9) = Format(g_rst_Princi!TOT_09, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 10) = Format(g_rst_Princi!TOT_10, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 11) = Format(g_rst_Princi!TOT_11, "###,###,##0.00")
         grd_Listado.TextMatrix(1, 12) = Format(g_rst_Princi!TOT_12, "###,###,##0.00")
   End Select
   
   'Cartera Pesada
   l_str_PerMes = 1
   l_str_PerAno = ipp_PerAno.Text
   r_int_Column = 1
   
   Do While r_int_Column <= l_int_PerMes
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_CLAPRV IN (2,3,4)) AS CON_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_CLACLI IN (2,3,4)) AS SIN_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & ") AS TOT_CAR "
      g_str_Parame = g_str_Parame & "  FROM DUAL"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron operaciones.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_dbl_TotSal = IIf(IsNull(g_rst_Princi!SIN_ALI), 0, g_rst_Princi!SIN_ALI)
      r_dbl_TotCal = IIf(IsNull(g_rst_Princi!CON_ALI), 0, g_rst_Princi!CON_ALI)
      r_dbl_TotCar = IIf(IsNull(g_rst_Princi!TOT_CAR), 0, g_rst_Princi!TOT_CAR)
    
      If r_dbl_TotCar <> 0 Then grd_Listado.TextMatrix(2, r_int_Column) = Format((r_dbl_TotCal / r_dbl_TotCar) * 100, "##0.00")
      
      r_int_Column = r_int_Column + 1
      l_str_PerMes = l_str_PerMes + 1
   Loop
   
   g_str_Parame = "USP_CUR_GEN_EEBG ("
   g_str_Parame = g_str_Parame & l_int_PerMes & ", "
   g_str_Parame = g_str_Parame & CInt(l_int_PerAno) & ", 1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "
   
   'Ejecuta consulta
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEBG.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Adeudos/Pasivo Total
   Call fs_ProcesarData_Indicadores_Adeudos_Pasivo
   
   'Disponible/Activo Total
   Call fs_ProcesarData_Indicadores_Disponible_Activo
   
   'Posicion en ME
   r_dbl_PatEfec = 0
   l_str_PerMes = 1
   r_int_Column = 1
   
   Do While r_int_Column <= l_int_PerMes
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_CLAPRV IN (2,3,4)) AS CON_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_CLACLI IN (2,3,4)) AS SIN_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)) "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & ") AS TOT_CAR "
      g_str_Parame = g_str_Parame & "  FROM DUAL"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron operaciones.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_dbl_TotCar = IIf(IsNull(g_rst_Princi!TOT_CAR), 0, g_rst_Princi!TOT_CAR)
      
      'Utilizamos esta condicion para verificar si esta variable y parte del query anterior no sobrepasa el mes seleccionado
      If r_dbl_TotCar <> 0 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT CONLIM_PATEFE AS PE "
         g_str_Parame = g_str_Parame & "  FROM CTB_CONLIM "
         g_str_Parame = g_str_Parame & " WHERE CONLIM_CODMES = " & l_str_PerMes & " "
         g_str_Parame = g_str_Parame & "   AND CONLIM_CODANO = " & l_str_PerAno & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_dbl_PatEfec = g_rst_Princi!PE
         End If
         
         r_dbl_ValNum = 0
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "USP_HIS_REPAEF( "
         g_str_Parame = g_str_Parame & "'CTB_REPSBS_17', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"
         g_str_Parame = g_str_Parame & l_str_PerMes & ", "
         g_str_Parame = g_str_Parame & l_str_PerAno & ") "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         'Consulta de datos
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT REPAEF_POCABA  AS SUMA "
         g_str_Parame = g_str_Parame & "  FROM HIS_REPAEF "
         g_str_Parame = g_str_Parame & " WHERE REPAEF_PERMES = " & l_str_PerMes & " "
         g_str_Parame = g_str_Parame & "   AND REPAEF_PERANO = " & l_str_PerAno & " "
         g_str_Parame = g_str_Parame & "   AND REPAEF_NOMRPT = 'CTB_REPSBS_17'  "
         g_str_Parame = g_str_Parame & "   AND REPAEF_TERCRE = '" & modgen_g_str_NombPC & "' "
         g_str_Parame = g_str_Parame & "   AND REPAEF_NUMITE = 14 "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_dbl_ValNum = g_rst_Princi!SUMA
         End If
         
         grd_Listado.TextMatrix(6, r_int_Column) = Format((r_dbl_ValNum / r_dbl_PatEfec) * 100, "##0.00")
      End If
      
      r_int_Column = r_int_Column + 1
      l_str_PerMes = l_str_PerMes + 1
   Loop
   
   grd_Listado.Redraw = True
   Call gs_UbiIniGrid(grd_Listado)
   Call fs_Activa(True)
End Sub

Private Sub fs_ProcesarData_Indicadores_Adeudos_Pasivo()
Dim l_dbl_MtoEne As Double, l_dbl_MtoFeb As Double, l_dbl_MtoMar As Double
Dim l_dbl_MtoAbr As Double, l_dbl_MtoMay As Double, l_dbl_MtoJun As Double
Dim l_dbl_MtoJul As Double, l_dbl_MtoAgo As Double, l_dbl_MtoSet As Double
Dim l_dbl_MtoOct As Double, l_dbl_MtoNov As Double, l_dbl_MtoDic As Double
   
   On Error Resume Next
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!GRUPO = 11 Then
         If g_rst_Princi!SUBGRP = 1 And Trim(g_rst_Princi!INDTIPO) = "G" Then
            If l_str_CieAno <> l_int_PerAno Then
               Select Case l_int_PerMes
               Case 1
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
               Case 2
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
               Case 3
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
               Case 4
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
               Case 5
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
               Case 6
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
               Case 7
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
               Case 8
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
               Case 9
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
               Case 10
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
               Case 11
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
               Case 12
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
                  l_dbl_MtoDic = Format(g_rst_Princi!MES12, "###,###,##0.00")
                End Select
            Else
               Select Case l_str_CieMes
               Case 1
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
               Case 2
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
               Case 3
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
               Case 4
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
               Case 5
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
               Case 6
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
               Case 7
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
               Case 8
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
               Case 9
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
               Case 10
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
               Case 11
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
               Case 12
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
                  l_dbl_MtoDic = Format(g_rst_Princi!MES12, "###,###,##0.00")
               End Select
            End If
         End If
      End If

      If g_rst_Princi!GRUPO = 14 Then '13
         If g_rst_Princi!SUBGRP = 3 Then
            If l_str_CieAno <> l_int_PerAno Then
               Select Case l_int_PerMes
               Case 1
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
               Case 2
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
               Case 3
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
               Case 4
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
               Case 5
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
               Case 6
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
               Case 7
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
               Case 8
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
               Case 9
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
               Case 10
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
               Case 11
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
               Case 12
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 12) = Format((l_dbl_MtoDic / g_rst_Princi!MES12) * 100, "###,##0.00")
               End Select
            Else
               Select Case l_str_CieMes
               Case 1
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
               Case 2
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
               Case 3
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
               Case 4
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
               Case 5
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
               Case 6
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
               Case 7
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
               Case 8
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
               Case 9
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
               Case 10
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
               Case 11
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
               Case 12
                  grd_Listado.TextMatrix(4, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(4, 12) = Format((l_dbl_MtoDic / g_rst_Princi!MES12) * 100, "###,##0.00")
               End Select
            End If
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
End Sub

Private Sub fs_ProcesarData_Indicadores_Disponible_Activo()
   Dim l_dbl_MtoEne As Double, l_dbl_MtoFeb As Double, l_dbl_MtoMar As Double
   Dim l_dbl_MtoAbr As Double, l_dbl_MtoMay As Double, l_dbl_MtoJun As Double
   Dim l_dbl_MtoJul As Double, l_dbl_MtoAgo As Double, l_dbl_MtoSet As Double
   Dim l_dbl_MtoOct As Double, l_dbl_MtoNov As Double, l_dbl_MtoDic As Double
   
   On Error Resume Next
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!GRUPO = 1 Then
         If g_rst_Princi!SUBGRP = 1 And Trim(g_rst_Princi!INDTIPO) = "G" Then
            If l_str_CieAno <> l_int_PerAno Then
               Select Case l_int_PerMes
               Case 1
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
               Case 2
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
               Case 3
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
               Case 4
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
               Case 5
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
               Case 6
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
               Case 7
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
               Case 8
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
               Case 9
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
               Case 10
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
               Case 11
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
               Case 12
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
                  l_dbl_MtoDic = Format(g_rst_Princi!MES12, "###,###,##0.00")
               End Select
            Else
               Select Case l_str_CieMes
               Case 1
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
               Case 2
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
               Case 3
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
               Case 4
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
               Case 5
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
               Case 6
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
               Case 7
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
               Case 8
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
               Case 9
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
               Case 10
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
               Case 11
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
               Case 12
                  l_dbl_MtoEne = Format(g_rst_Princi!MES01, "###,###,##0.00")
                  l_dbl_MtoFeb = Format(g_rst_Princi!MES02, "###,###,##0.00")
                  l_dbl_MtoMar = Format(g_rst_Princi!MES03, "###,###,##0.00")
                  l_dbl_MtoAbr = Format(g_rst_Princi!MES04, "###,###,##0.00")
                  l_dbl_MtoMay = Format(g_rst_Princi!MES05, "###,###,##0.00")
                  l_dbl_MtoJun = Format(g_rst_Princi!MES06, "###,###,##0.00")
                  l_dbl_MtoJul = Format(g_rst_Princi!MES07, "###,###,##0.00")
                  l_dbl_MtoAgo = Format(g_rst_Princi!MES08, "###,###,##0.00")
                  l_dbl_MtoSet = Format(g_rst_Princi!MES09, "###,###,##0.00")
                  l_dbl_MtoOct = Format(g_rst_Princi!MES10, "###,###,##0.00")
                  l_dbl_MtoNov = Format(g_rst_Princi!MES11, "###,###,##0.00")
                  l_dbl_MtoDic = Format(g_rst_Princi!MES12, "###,###,##0.00")
               End Select
            End If
         End If
      End If

      If g_rst_Princi!GRUPO = 7 Then
         If g_rst_Princi!SUBGRP = 2 Then
            If l_str_CieAno <> l_int_PerAno Then
               Select Case l_int_PerMes
               Case 1
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
               Case 2
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
               Case 3
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
               Case 4
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
               Case 5
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
               Case 6
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
               Case 7
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
               Case 8
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
               Case 9
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
               Case 10
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
               Case 11
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
               Case 12
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 12) = Format((l_dbl_MtoDic / g_rst_Princi!MES12) * 100, "###,##0.00")
               End Select
            Else
               Select Case l_str_CieMes
               Case 1
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
               Case 2
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
               Case 3
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
               Case 4
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
               Case 5
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
               Case 6
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
               Case 7
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
               Case 8
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
               Case 9
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
               Case 10
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
               Case 11
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
               Case 12
                  grd_Listado.TextMatrix(5, 1) = Format((l_dbl_MtoEne / g_rst_Princi!MES01) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 2) = Format((l_dbl_MtoFeb / g_rst_Princi!MES02) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 3) = Format((l_dbl_MtoMar / g_rst_Princi!MES03) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 4) = Format((l_dbl_MtoAbr / g_rst_Princi!MES04) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 5) = Format((l_dbl_MtoMay / g_rst_Princi!MES05) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 6) = Format((l_dbl_MtoJun / g_rst_Princi!MES06) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 7) = Format((l_dbl_MtoJul / g_rst_Princi!MES07) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 8) = Format((l_dbl_MtoAgo / g_rst_Princi!MES08) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 9) = Format((l_dbl_MtoSet / g_rst_Princi!MES09) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 10) = Format((l_dbl_MtoOct / g_rst_Princi!MES10) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 11) = Format((l_dbl_MtoNov / g_rst_Princi!MES11) * 100, "###,##0.00")
                  grd_Listado.TextMatrix(5, 12) = Format((l_dbl_MtoDic / g_rst_Princi!MES12) * 100, "###,##0.00")
               End Select
            End If
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
End Sub

Private Sub fs_Procesar_Saldos()
Dim r_int_fila       As Integer

   grd_Listado.Rows = 1
   Call gs_LimpiaGrid(grd_Listado)
   Call fs_SeteaColumnas_Saldos
   DoEvents
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame + "SELECT ROWNUM AS ITEM, TRIM(E.PRODUC_DESCRI) AS PRODUCTO, TRIM(B.PARDES_DESCRI) AS MONEDA_PRESTAMO,"
   g_str_Parame = g_str_Parame + "       ROUND(A.HIPCIE_MTOPRE, 2) AS MONTO_PRESTAMO, ROUND(A.HIPCIE_SALCON + HIPCIE_SALCAP, 2) AS SALDO_PRESTAMO,"
   g_str_Parame = g_str_Parame + "       ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCON + HIPCIE_SALCAP, (HIPCIE_SALCON + HIPCIE_SALCAP) * HIPCIE_TIPCAM),2) AS SALDO_SOLES,"
   g_str_Parame = g_str_Parame + "       A.HIPCIE_CLAPRV AS CLASIFICACION"
   g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE A INNER JOIN MNT_PARDES  B ON B.PARDES_CODGRP = 204  AND B.PARDES_CODITE = A.HIPCIE_TIPMON"
   g_str_Parame = g_str_Parame + " INNER JOIN CRE_PRODUC E ON E.PRODUC_CODIGO = A.HIPCIE_CODPRD"
   g_str_Parame = g_str_Parame + " WHERE A.HIPCIE_PERANO = " & l_int_PerAno & " AND A.HIPCIE_PERMES = " & l_int_PerMes & " ORDER BY A.HIPCIE_NUMOPE"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listado.Row = 1
   r_int_fila = 1
   If Not g_rst_Princi.EOF And Not g_rst_Princi.BOF Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listado.TextMatrix(r_int_fila, 0) = Format(g_rst_Princi!Item, "###0")
         grd_Listado.TextMatrix(r_int_fila, 1) = g_rst_Princi!PRODUCTO
         grd_Listado.TextMatrix(r_int_fila, 2) = g_rst_Princi!MONEDA_PRESTAMO
         grd_Listado.TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!MONTO_PRESTAMO, "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_fila, 4) = Format(g_rst_Princi!SALDO_PRESTAMO, "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!SALDO_SOLES, "###,###,##0.00")
         grd_Listado.TextMatrix(r_int_fila, 6) = g_rst_Princi!CLASIFICACION
         
         grd_Listado.Rows = grd_Listado.Rows + 1
         r_int_fila = r_int_fila + 1
         
         g_rst_Princi.MoveNext
      Loop
      grd_Listado.Rows = grd_Listado.Rows - 1
   End If
   
   grd_Listado.Redraw = True
   Call gs_UbiIniGrid(grd_Listado)
   Call fs_Activa(True)
End Sub

Private Function gf_Query(r_str_Anio As String) As String
   gf_Query = ""
   gf_Query = gf_Query & " SELECT (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 1 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_01,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 2 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_02,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 3 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_03,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 4 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_04,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 5 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_05,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 6 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_06,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 7 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_07,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 8 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_08,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 9 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_09,                                                                                               "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 10 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_10,                                                                                              "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 11 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_11,                                                                                              "
   
   gf_Query = gf_Query & "        (SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVEN, HIPCIE_CAPVEN*HIPCIE_TIPCAM))/ SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG+HIPCIE_CAPVEN), (HIPCIE_CAPVIG+HIPCIE_CAPVEN)*HIPCIE_TIPCAM))*100, 2) "
   gf_Query = gf_Query & "           FROM CRE_HIPCIE "
   gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = 12 AND HIPCIE_PERANO = '" & r_str_Anio & "') AS TOT_12                                                                                               "
   
   gf_Query = gf_Query & "   FROM DUAL                                                                                                                                                                         "
End Function

Private Sub fs_Busca_UltCierre()
   'Obtiene ultimo cierre procesado
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & " FROM (SELECT DISTINCT HIPCIE_PERANO, HIPCIE_PERMES "
   g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "        ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_PERANO DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   l_str_CieMes = g_rst_Princi!HIPCIE_PERMES
   l_str_CieAno = g_rst_Princi!HIPCIE_PERANO
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmd_Proces_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
        
   Select Case cmb_TipRep.ListIndex
      Case 3:
         If MsgBox("Este proceso puede demorar en ejecutarse 15 minutos aproximadamente." & vbCrLf & "¿Está seguro que desea realizar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
      Case Else
         If MsgBox("Este proceso puede demorar algunos minutos." & vbCrLf & "¿Está seguro que desea realizar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
   End Select
   
   Screen.MousePointer = 11
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
   
   Select Case cmb_TipRep.ListIndex
      Case 0: Call fs_ProcesarData_CAR
      Case 1: Call fs_ProcesarData_Indicadores
      Case 2: Call fs_Procesar_Saldos
      Case 3: Call fs_ProcesarData_Financieros
      Case 4: Call fs_Procesar_FactElectronica
   End Select
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExcRes_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Select Case cmb_TipRep.ListIndex
      Case 0: Call fs_GenExcRes
      Case 1: Call fs_GenExc_Indicadores
      Case 2: Call fs_GenExc_Saldos
      Case 3: Call fs_GenExc_Financieros
      Case 4: Call fs_GenExc_FactElectronica
   End Select
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call fs_Busca_UltCierre
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   ipp_PerAno.Text = Year(date)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call gs_LimpiaGrid(grd_Listado)
   
   cmb_TipRep.AddItem "REPORTE CAR - RATIO DE CAPITAL AJUSTADO"
   cmb_TipRep.AddItem "REPORTE DE INDICADORES DE CUENTAS DE BALANCE"
   cmb_TipRep.AddItem "REPORTE DE SALDOS MENSUALES"
   cmb_TipRep.AddItem "REPORTE DE INDICADORES FINANCIEROS FMV"
   cmb_TipRep.AddItem "REPORTE DE FACTURAS ELECTRONICAS"
   cmb_TipRep.ListIndex = 0
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
   cmd_ExpExcRes.Enabled = estado
End Sub

Private Sub fs_SeteaColumnas_FactElectronica()
Dim r_int_NumCol     As Integer
   
   r_int_NumCol = 0
   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)
   grd_Listado.Clear
 
   grd_Listado.Cols = 20
   grd_Listado.ColWidth(0) = 600
   grd_Listado.ColWidth(1) = 1000
   grd_Listado.ColWidth(2) = 700
   grd_Listado.ColWidth(3) = 2020
   grd_Listado.ColWidth(4) = 2060
   grd_Listado.ColWidth(5) = 1300
   grd_Listado.ColWidth(6) = 1300
   grd_Listado.ColWidth(7) = 750
   grd_Listado.ColWidth(8) = 700
   grd_Listado.ColWidth(9) = 1000
   grd_Listado.ColWidth(10) = 2800
   grd_Listado.ColWidth(11) = 700
   grd_Listado.ColWidth(12) = 700
   grd_Listado.ColWidth(13) = 700
   grd_Listado.ColWidth(14) = 1000
   grd_Listado.ColWidth(15) = 1100
   grd_Listado.ColWidth(16) = 700
   grd_Listado.ColWidth(17) = 1050
   grd_Listado.ColWidth(18) = 2500
   grd_Listado.ColWidth(19) = 700
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = 0: grd_Listado.FixedRows = 0
               
   grd_Listado.ColAlignment(0) = flexAlignCenterCenter
   grd_Listado.ColAlignment(1) = flexAlignCenterCenter
   grd_Listado.ColAlignment(2) = flexAlignCenterCenter
   grd_Listado.ColAlignment(3) = flexAlignCenterCenter
   grd_Listado.ColAlignment(4) = flexAlignLeftCenter
   grd_Listado.ColAlignment(7) = flexAlignCenterCenter
   grd_Listado.ColAlignment(8) = flexAlignCenterCenter
   grd_Listado.ColAlignment(9) = flexAlignCenterCenter
   grd_Listado.ColAlignment(10) = flexAlignLeftCenter
   grd_Listado.ColAlignment(11) = flexAlignCenterCenter
   grd_Listado.ColAlignment(12) = flexAlignCenterCenter
   grd_Listado.ColAlignment(13) = flexAlignCenterCenter
   grd_Listado.ColAlignment(14) = flexAlignCenterCenter
   grd_Listado.ColAlignment(15) = flexAlignCenterCenter
   grd_Listado.ColAlignment(16) = flexAlignCenterCenter
   grd_Listado.ColAlignment(17) = flexAlignCenterCenter
   grd_Listado.ColAlignment(18) = flexAlignLeftCenter
   grd_Listado.ColAlignment(19) = flexAlignCenterCenter
   
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "NRO": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_01": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_02": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_03": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_04": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_05": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignRightCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_06": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignRightCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_07": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_08": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_09": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_10": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_11": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_12": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_13": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_14": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_15": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_16": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_17": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_18": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = r_int_NumCol:  grd_Listado.Text = "COL_19": r_int_NumCol = r_int_NumCol + 1: grd_Listado.CellAlignment = flexAlignCenterCenter
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
      
   grd_Listado.Row = 0
   With grd_Listado
      '.FixedCols = 1
      .FixedRows = 1
   End With
   
   grd_Listado.Redraw = True
End Sub

Private Sub fs_SeteaColumnas_CAR()
   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)
   
   'Ancho de columnas
   grd_Listado.Cols = 13
   grd_Listado.ColWidth(0) = 2460    ' DESCRIPCION
   grd_Listado.ColWidth(1) = 1350    ' MES 1
   grd_Listado.ColWidth(2) = 1350    ' MES 2
   grd_Listado.ColWidth(3) = 1350    ' MES 3
   grd_Listado.ColWidth(4) = 1350    ' MES 4
   grd_Listado.ColWidth(5) = 1350    ' MES 5
   grd_Listado.ColWidth(6) = 1350    ' MES 6
   grd_Listado.ColWidth(7) = 1350    ' MES 7
   grd_Listado.ColWidth(8) = 1350    ' MES 8
   grd_Listado.ColWidth(9) = 1350    ' MES 9
   grd_Listado.ColWidth(10) = 1350   ' MES 10
   grd_Listado.ColWidth(11) = 1350   ' MES 11
   grd_Listado.ColWidth(12) = 1350   ' MES 12
   grd_Listado.ColAlignment(0) = flexAlignLeftCenter
   grd_Listado.ColAlignment(1) = flexAlignRightCenter
   grd_Listado.ColAlignment(2) = flexAlignRightCenter
   grd_Listado.ColAlignment(3) = flexAlignRightCenter
   grd_Listado.ColAlignment(4) = flexAlignRightCenter
   grd_Listado.ColAlignment(5) = flexAlignRightCenter
   grd_Listado.ColAlignment(6) = flexAlignRightCenter
   grd_Listado.ColAlignment(7) = flexAlignRightCenter
   grd_Listado.ColAlignment(8) = flexAlignRightCenter
   grd_Listado.ColAlignment(9) = flexAlignRightCenter
   grd_Listado.ColAlignment(10) = flexAlignRightCenter
   grd_Listado.ColAlignment(11) = flexAlignRightCenter
   grd_Listado.ColAlignment(12) = flexAlignRightCenter
   
   'Cabecera
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Row = 0: grd_Listado.Text = ""
   grd_Listado.Col = 0: grd_Listado.Text = ""
   grd_Listado.Col = 1: grd_Listado.Text = "ENERO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 2: grd_Listado.Text = "FEBRERO":      grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 3: grd_Listado.Text = "MARZO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 4: grd_Listado.Text = "ABRIL":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 5: grd_Listado.Text = "MAYO":         grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 6: grd_Listado.Text = "JUNIO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 7: grd_Listado.Text = "JULIO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 8: grd_Listado.Text = "AGOSTO":       grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 9: grd_Listado.Text = "SETIEMBRE":    grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 10: grd_Listado.Text = "OCTUBRE":     grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 11: grd_Listado.Text = "NOVIEMBRE":   grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 12: grd_Listado.Text = "DICIEMBRE":   grd_Listado.CellAlignment = flexAlignCenterCenter
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "RIESGO DE CREDITO"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "RIESGO DE MERCADO"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "RIESGO OPERACIONAL"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "TOTAL RIESGOS"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "PATRIMONIO EFECTIVO"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "RATIO CAPITAL AJUSTADO (%)"
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "APALANCAMIENTO (veces)"
         
   With grd_Listado
'      .MergeCells = flexMergeFree
      .FixedCols = 1
      .FixedRows = 1
   End With
   grd_Listado.Redraw = True
End Sub

Private Sub fs_llenar(ByVal Columna As Integer, ByVal anno As Integer, ByVal anvigente, ByVal mesvigente)
Dim j As Integer
   
   grd_Listado.Row = 0
    
   If Columna = 12 Then
      j = -19
      grd_Listado.Col = 2 + Columna + j + 18:  grd_Listado.Text = "DIC-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S1:
      grd_Listado.Col = 2 + Columna + j + 17:  grd_Listado.Text = "NOV-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S2:
      grd_Listado.Col = 2 + Columna + j + 16:  grd_Listado.Text = "OCT-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S3:
      grd_Listado.Col = 2 + Columna + j + 15:  grd_Listado.Text = "SET-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S4:
      grd_Listado.Col = 2 + Columna + j + 14:  grd_Listado.Text = "AGO-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S5:
      grd_Listado.Col = 2 + Columna + j + 13:  grd_Listado.Text = "JUL-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S6:
      grd_Listado.Col = 2 + Columna + j + 12:  grd_Listado.Text = "JUN-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S7:
      grd_Listado.Col = 2 + Columna + j + 11:  grd_Listado.Text = "MAY-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S8:
      grd_Listado.Col = 2 + Columna + j + 10:  grd_Listado.Text = "ABR-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S9:
      grd_Listado.Col = 2 + Columna + j + 9:   grd_Listado.Text = "MAR-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S10:
      grd_Listado.Col = 2 + Columna + j + 8:   grd_Listado.Text = "FEB-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S11:
      grd_Listado.Col = 2 + Columna + j + 7:   grd_Listado.Text = "ENE-" & Right(anno, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
            
   ElseIf Columna = 11 Then
      j = -17
      GoTo S1
   ElseIf Columna = 10 Then
      j = -15
      GoTo S2
   ElseIf Columna = 9 Then
      j = -13
      GoTo S3
   ElseIf Columna = 8 Then
      j = -11
      GoTo S4
   ElseIf Columna = 7 Then
      j = -9
      GoTo S5
   ElseIf Columna = 6 Then
      j = -7
      GoTo S6
   ElseIf Columna = 5 Then
      j = -5
      GoTo S7
   ElseIf Columna = 4 Then
      j = -3
      GoTo S8
   ElseIf Columna = 3 Then
      j = -1
      GoTo S9
   ElseIf Columna = 2 Then
      j = 1
      GoTo S10
   ElseIf Columna = 1 Then
      j = 3
      GoTo S11
   End If
End Sub

Private Sub fs_SeteaColumnas_Financieros()
Dim p                As Integer
Dim AnioVigente      As Integer
Dim mesvigente       As Integer

   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)

   grd_Listado.Cols = 14
   grd_Listado.ColWidth(0) = 4300
   grd_Listado.ColWidth(1) = 1350
   grd_Listado.ColWidth(2) = 1350
   grd_Listado.ColWidth(3) = 1350
   grd_Listado.ColWidth(4) = 1350
   grd_Listado.ColWidth(5) = 1350
   grd_Listado.ColWidth(6) = 1350
   grd_Listado.ColWidth(7) = 1350
   grd_Listado.ColWidth(8) = 1350
   grd_Listado.ColWidth(9) = 1350
   grd_Listado.ColWidth(10) = 1350
   grd_Listado.ColWidth(11) = 1350
   grd_Listado.ColWidth(12) = 1350
   grd_Listado.ColWidth(13) = 1350
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = 0: grd_Listado.FixedRows = 0
         
   If CInt(l_int_PerMes) = 12 Then
      p = -10
      GoTo S12
    
S1:
      grd_Listado.Col = p:       grd_Listado.Text = "ENE-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S2:
      grd_Listado.Col = p + 1:   grd_Listado.Text = "FEB-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S3:
      grd_Listado.Col = p + 2:   grd_Listado.Text = "MAR-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S4:
      grd_Listado.Col = p + 3:   grd_Listado.Text = "ABR-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S5:
      grd_Listado.Col = p + 4:   grd_Listado.Text = "MAY-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S6:
      grd_Listado.Col = p + 5:   grd_Listado.Text = "JUN-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S7:
      grd_Listado.Col = p + 6:   grd_Listado.Text = "JUL-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S8:
      grd_Listado.Col = p + 7:   grd_Listado.Text = "AGO-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S9:
      grd_Listado.Col = p + 8:   grd_Listado.Text = "SET-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S10:
      grd_Listado.Col = p + 9:   grd_Listado.Text = "OCT-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S11:
      grd_Listado.Col = p + 10:  grd_Listado.Text = "NOV-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
S12:
      grd_Listado.Col = p + 11:  grd_Listado.Text = "DIC-" & Right(CInt(l_int_PerAno) - 1, 2): grd_Listado.CellAlignment = flexAlignCenterCenter
  
   ElseIf CInt(l_int_PerMes) = 11 Then
      p = -9
      GoTo S11
   ElseIf CInt(l_int_PerMes) = 10 Then
      p = -8
      GoTo S10
   ElseIf CInt(l_int_PerMes) = 9 Then
      p = -7
      GoTo S9
   ElseIf CInt(l_int_PerMes) = 8 Then
      p = -6
      GoTo S8
   ElseIf CInt(l_int_PerMes) = 7 Then
      p = -5
      GoTo S7
   ElseIf CInt(l_int_PerMes) = 6 Then
      p = -4
      GoTo S6
   ElseIf CInt(l_int_PerMes) = 5 Then
      p = -3
      GoTo S5
   ElseIf CInt(l_int_PerMes) = 4 Then
      p = -2
      GoTo S4
   ElseIf CInt(l_int_PerMes) = 3 Then
      p = -1
      GoTo S3
   ElseIf CInt(l_int_PerMes) = 2 Then
      p = 0
      GoTo S2
   ElseIf CInt(l_int_PerMes) = 1 Then
      p = 1
      GoTo S1
   End If
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   
   'Consulta para obtener el mes y año vigente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT MAX(PERMES_CODMES) AS MES, MAX(PERMES_CODANO) AS ANO "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "Error al ejecutar la consulta para obtener año vigente.", vbCritical, modgen_g_str_NomPlt
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      AnioVigente = g_rst_Genera!ANO
      mesvigente = g_rst_Genera!Mes
   End If

   Call fs_llenar(CInt(l_int_PerMes), CInt(l_int_PerAno), AnioVigente, mesvigente)
   
   grd_Listado.Col = 1
   With grd_Listado
      .FixedCols = 1
      .FixedRows = 1
   End With
   
   grd_Listado.Redraw = True
End Sub

Private Sub fs_SeteaColumnas_Indicadores()
   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)
   grd_Listado.SelectionMode = flexSelectionByRow
   grd_Listado.FocusRect = flexFocusNone
   grd_Listado.HighLight = flexHighlightAlways

   'Ancho de columnas
   grd_Listado.Cols = 13
   grd_Listado.ColWidth(0) = 2450    ' DESCRIPCION
   grd_Listado.ColWidth(1) = 1300    ' MES 1
   grd_Listado.ColWidth(2) = 1300    ' MES 2
   grd_Listado.ColWidth(3) = 1300    ' MES 3
   grd_Listado.ColWidth(4) = 1300    ' MES 4
   grd_Listado.ColWidth(5) = 1300    ' MES 5
   grd_Listado.ColWidth(6) = 1300    ' MES 6
   grd_Listado.ColWidth(7) = 1300    ' MES 7
   grd_Listado.ColWidth(8) = 1300    ' MES 8
   grd_Listado.ColWidth(9) = 1300    ' MES 9
   grd_Listado.ColWidth(10) = 1300   ' MES 10
   grd_Listado.ColWidth(11) = 1300   ' MES 11
   grd_Listado.ColWidth(12) = 1300   ' MES 12
   grd_Listado.ColAlignment(0) = flexAlignLeftCenter
   grd_Listado.ColAlignment(1) = flexAlignRightCenter
   grd_Listado.ColAlignment(2) = flexAlignRightCenter
   grd_Listado.ColAlignment(3) = flexAlignRightCenter
   grd_Listado.ColAlignment(4) = flexAlignRightCenter
   grd_Listado.ColAlignment(5) = flexAlignRightCenter
   grd_Listado.ColAlignment(6) = flexAlignRightCenter
   grd_Listado.ColAlignment(7) = flexAlignRightCenter
   grd_Listado.ColAlignment(8) = flexAlignRightCenter
   grd_Listado.ColAlignment(9) = flexAlignRightCenter
   grd_Listado.ColAlignment(10) = flexAlignRightCenter
   grd_Listado.ColAlignment(11) = flexAlignRightCenter
   grd_Listado.ColAlignment(12) = flexAlignRightCenter
      
   'Cabecera
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Row = 0: grd_Listado.Text = ""
   grd_Listado.Col = 0: grd_Listado.Text = "INDICADOR":    grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 1: grd_Listado.Text = "ENERO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 2: grd_Listado.Text = "FEBRERO":      grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 3: grd_Listado.Text = "MARZO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 4: grd_Listado.Text = "ABRIL":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 5: grd_Listado.Text = "MAYO":         grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 6: grd_Listado.Text = "JUNIO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 7: grd_Listado.Text = "JULIO":        grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 8: grd_Listado.Text = "AGOSTO":       grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 9: grd_Listado.Text = "SETIEMBRE":    grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 10: grd_Listado.Text = "OCTUBRE":     grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 11: grd_Listado.Text = "NOVIEMBRE":   grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 12: grd_Listado.Text = "DICIEMBRE":   grd_Listado.CellAlignment = flexAlignCenterCenter
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "MOROSIDAD"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "CARTERA PESADA"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "LIQUIDEZ"
   
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "     ADEUDOS / PASIVO TOTAL"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "     DISPONIBLE / ACTIVO TOTAL"

   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Col = 0:   grd_Listado.Text = "POSICION EN ME"

   With grd_Listado
'      .MergeCells = flexMergeFree
      .FixedCols = 1
      .FixedRows = 1
   End With
   grd_Listado.Redraw = True
End Sub

Private Sub fs_SeteaColumnas_Saldos()
   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)
   grd_Listado.Clear
   'Ancho de columnas
   grd_Listado.Cols = 7
   grd_Listado.ColWidth(0) = 700
   grd_Listado.ColWidth(1) = 4700
   grd_Listado.ColWidth(2) = 2000
   grd_Listado.ColWidth(3) = 1500
   grd_Listado.ColWidth(4) = 1500
   grd_Listado.ColWidth(5) = 1500
   grd_Listado.ColWidth(6) = 1500
   grd_Listado.ColAlignment(0) = flexAlignRightCenter
   grd_Listado.ColAlignment(1) = flexAlignLeftCenter
   grd_Listado.ColAlignment(2) = flexAlignLeftCenter
   grd_Listado.ColAlignment(3) = flexAlignRightCenter
   grd_Listado.ColAlignment(4) = flexAlignRightCenter
   grd_Listado.ColAlignment(5) = flexAlignRightCenter
   grd_Listado.ColAlignment(6) = flexAlignRightCenter
      
   'Cabecera
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.Row = grd_Listado.Rows - 1
   grd_Listado.Row = 0: grd_Listado.Text = ""
   grd_Listado.Col = 0: grd_Listado.Text = "ITEM":             grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 1: grd_Listado.Text = "PRODUCTO":         grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 2: grd_Listado.Text = "MONEDA":           grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 3: grd_Listado.Text = "MONTO PRESTAMO":   grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 4: grd_Listado.Text = "SALDO PRESTAMO":   grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 5: grd_Listado.Text = "SALDO SOLES":      grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Col = 6: grd_Listado.Text = "CLASIFICACION":    grd_Listado.CellAlignment = flexAlignCenterCenter
   grd_Listado.Rows = grd_Listado.Rows + 1
   grd_Listado.FixedRows = 1
   grd_Listado.FixedCols = 0
   grd_Listado.Redraw = True
End Sub

Private Sub fs_ProcesarData_CAR()
Dim r_int_NumFil  As Integer

   grd_Listado.Redraw = False
   Call gs_LimpiaGrid(grd_Listado)
   Call fs_SeteaColumnas_CAR
   DoEvents
   
   'Procesa infomacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_RPT_RATIOCAPAJUS("
   g_str_Parame = g_str_Parame & CInt(l_int_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(l_int_PerAno) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'RATIOCAPAJUST', "
   g_str_Parame = g_str_Parame & "0)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Consulta informacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_PERMES = '" & CInt(l_int_PerMes) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_PERANO = '" & CInt(l_int_PerAno) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'RATIOCAPAJUST' "
   g_str_Parame = g_str_Parame & "    AND RPT_MONEDA = 0 "
   g_str_Parame = g_str_Parame & "  ORDER BY RPT_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_NumFil = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         Select Case l_int_PerMes
            Case 1
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
            Case 2
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
            Case 3
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
            Case 4
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
            Case 5
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
            Case 6
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
            Case 7
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
            Case 8
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
            Case 9
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09), "###,###,##0.00")
            Case 10
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10), "###,###,##0.00")
            Case 11
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11), "###,###,##0.00")
            Case 12
               grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11), "###,###,##0.00")
               grd_Listado.TextMatrix(r_int_NumFil, 12) = Format(IIf(IsNull(g_rst_Princi!RPT_VALNUM12), 0, g_rst_Princi!RPT_VALNUM12), "###,###,##0.00")
         End Select
         
         r_int_NumFil = r_int_NumFil + 1
         g_rst_Princi.MoveNext
         
         If g_rst_Princi.EOF Then
            Select Case l_int_PerMes
               Case 1
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
               Case 2
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
               Case 3
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
               Case 4
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
               Case 5
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
               Case 6
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
               Case 7
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
               Case 8
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 8) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(grd_Listado.TextMatrix(4, 8) / grd_Listado.TextMatrix(5, 8), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(0, "###,###,##0.00")
               Case 9
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 8) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(grd_Listado.TextMatrix(4, 8) / grd_Listado.TextMatrix(5, 8), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 9) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(grd_Listado.TextMatrix(4, 9) / grd_Listado.TextMatrix(5, 9), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(0, "###,###,##0.00")
               Case 10
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 8) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(grd_Listado.TextMatrix(4, 8) / grd_Listado.TextMatrix(5, 8), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 9) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(grd_Listado.TextMatrix(4, 9) / grd_Listado.TextMatrix(5, 9), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 10) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(grd_Listado.TextMatrix(4, 10) / grd_Listado.TextMatrix(5, 10), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(0, "###,###,##0.00")
               Case 11
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 8) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(grd_Listado.TextMatrix(4, 8) / grd_Listado.TextMatrix(5, 8), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 9) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(grd_Listado.TextMatrix(4, 9) / grd_Listado.TextMatrix(5, 9), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 10) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(grd_Listado.TextMatrix(4, 10) / grd_Listado.TextMatrix(5, 10), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 11) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(grd_Listado.TextMatrix(4, 11) / grd_Listado.TextMatrix(5, 11), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(0, "###,###,##0.00")
               Case 12
                  If grd_Listado.TextMatrix(5, 1) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(grd_Listado.TextMatrix(4, 1) / grd_Listado.TextMatrix(5, 1), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 1) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 2) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(grd_Listado.TextMatrix(4, 2) / grd_Listado.TextMatrix(5, 2), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 2) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 3) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(grd_Listado.TextMatrix(4, 3) / grd_Listado.TextMatrix(5, 3), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 3) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 4) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(grd_Listado.TextMatrix(4, 4) / grd_Listado.TextMatrix(5, 4), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 4) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 5) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(grd_Listado.TextMatrix(4, 5) / grd_Listado.TextMatrix(5, 5), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 5) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 6) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(grd_Listado.TextMatrix(4, 6) / grd_Listado.TextMatrix(5, 6), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 6) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 7) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(grd_Listado.TextMatrix(4, 7) / grd_Listado.TextMatrix(5, 7), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 7) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 8) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(grd_Listado.TextMatrix(4, 8) / grd_Listado.TextMatrix(5, 8), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 8) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 9) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(grd_Listado.TextMatrix(4, 9) / grd_Listado.TextMatrix(5, 9), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 9) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 10) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(grd_Listado.TextMatrix(4, 10) / grd_Listado.TextMatrix(5, 10), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 10) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 11) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(grd_Listado.TextMatrix(4, 11) / grd_Listado.TextMatrix(5, 11), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 11) = Format(0, "###,###,##0.00")
                  If grd_Listado.TextMatrix(5, 12) <> "0.00" Then grd_Listado.TextMatrix(r_int_NumFil, 12) = Format(grd_Listado.TextMatrix(4, 12) / grd_Listado.TextMatrix(5, 12), "###,###,##0.00") Else grd_Listado.TextMatrix(r_int_NumFil, 12) = Format(0, "###,###,##0.00")
            End Select
         End If
      Loop
   End If
   
   Call fs_Activa(True)
   grd_Listado.Redraw = True
End Sub

Private Sub fs_GenExc_Financieros()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
   
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "RATIOS FINANCIEROS:"
      .Range(.Cells(1, 2), .Cells(1, 2)).Font.Bold = True

'  r_obj_Excel.Visible = True

      For r_int_Contad = 1 To 13
         .Cells(2, r_int_Contad + 2) = "'" & grd_Listado.TextMatrix(0, r_int_Contad)
      Next
      
      For r_int_Contad = 1 To 39
         .Cells(r_int_Contad + 2, 2) = "'" & grd_Listado.TextMatrix(r_int_Contad, 0)
         .Cells(r_int_Contad + 2, 3) = grd_Listado.TextMatrix(r_int_Contad, 1)
         .Cells(r_int_Contad + 2, 4) = grd_Listado.TextMatrix(r_int_Contad, 2)
         .Cells(r_int_Contad + 2, 5) = grd_Listado.TextMatrix(r_int_Contad, 3)
         .Cells(r_int_Contad + 2, 6) = grd_Listado.TextMatrix(r_int_Contad, 4)
         .Cells(r_int_Contad + 2, 7) = grd_Listado.TextMatrix(r_int_Contad, 5)
         .Cells(r_int_Contad + 2, 8) = grd_Listado.TextMatrix(r_int_Contad, 6)
         .Cells(r_int_Contad + 2, 9) = grd_Listado.TextMatrix(r_int_Contad, 7)
         .Cells(r_int_Contad + 2, 10) = grd_Listado.TextMatrix(r_int_Contad, 8)
         .Cells(r_int_Contad + 2, 11) = grd_Listado.TextMatrix(r_int_Contad, 9)
         .Cells(r_int_Contad + 2, 12) = grd_Listado.TextMatrix(r_int_Contad, 10)
         .Cells(r_int_Contad + 2, 13) = grd_Listado.TextMatrix(r_int_Contad, 11)
         .Cells(r_int_Contad + 2, 14) = grd_Listado.TextMatrix(r_int_Contad, 12)
         .Cells(r_int_Contad + 2, 15) = grd_Listado.TextMatrix(r_int_Contad, 13)
      Next
        
      .Range(.Cells(2, 1), .Cells(2, 20)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 20)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 3
      .Columns("B").ColumnWidth = 50
'      .Columns("C").ColumnWidth = 10
'      .Columns("D").ColumnWidth = 10
'      .Columns("E").ColumnWidth = 10
'      .Columns("F").ColumnWidth = 10
'      .Columns("G").ColumnWidth = 10
'      .Columns("H").ColumnWidth = 10
'      .Columns("I").ColumnWidth = 10
'      .Columns("J").ColumnWidth = 10
'      .Columns("K").ColumnWidth = 10
'      .Columns("L").ColumnWidth = 10
'      .Columns("M").ColumnWidth = 10
'      .Columns("N").ColumnWidth = 10
'      .Columns("O").ColumnWidth = 10
      
'      .Range(.Cells(4, 2), .Cells(4, 14)).Interior.Color = RGB(146, 208, 80)
'      .Range(.Cells(4, 2), .Cells(4, 14)).Font.Bold = True
'      .Range(.Cells(4, 3), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(5, 2), .Cells(12, 2)).Font.Bold = True


      .Range(.Cells(3, 2), .Cells(3, 2)).Font.Bold = True
      .Range(.Cells(7, 2), .Cells(7, 2)).Font.Bold = True
      .Range(.Cells(13, 2), .Cells(13, 2)).Font.Bold = True
      .Range(.Cells(17, 2), .Cells(17, 2)).Font.Bold = True
      .Range(.Cells(23, 2), .Cells(23, 2)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 2)).Font.Underline = True
      .Range(.Cells(7, 2), .Cells(7, 2)).Font.Underline = True
      .Range(.Cells(13, 2), .Cells(13, 2)).Font.Underline = True
      .Range(.Cells(17, 2), .Cells(17, 2)).Font.Underline = True
      .Range(.Cells(23, 2), .Cells(23, 2)).Font.Underline = True

      .Columns("C").ColumnWidth = 11
      .Columns("C").HorizontalAlignment = xlHAlignRight
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 11
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 11
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 11
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 11
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 11
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 11
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 11
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 11
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 11
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 11
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 11
      .Columns("O").NumberFormat = "###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight

      .Range(.Cells(1, 1), .Cells(25, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(25, 99)).Font.Size = 11
      
      .Range(.Cells(27, 3), .Cells(41, 99)).Font.Name = "Calibri"
      .Range(.Cells(27, 3), .Cells(41, 99)).Font.Size = 7
      
'
'      r_int_NumFil = 2
'      For r_int_Contad = 1 To grd_Listado.Rows - 1
'         .Cells(r_int_NumFil + 3, 2) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 0), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 3) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 1), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 4) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 2), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 5) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 3), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 6) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 4), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 7) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 5), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 8) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 6), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 9) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 7), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 10) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 8), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 11) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 9), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 12) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 10), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 13) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 11), "###,###,##0.00")
'         .Cells(r_int_NumFil + 3, 14) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 12), "###,###,##0.00")
'
'         .Range(.Cells(r_int_NumFil + 2, 2), .Cells(r_int_NumFil + 3, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
'         .Range(.Cells(r_int_NumFil + 2, 2), .Cells(r_int_NumFil + 3, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'
'         r_int_NumFil = r_int_NumFil + 1
'      Next r_int_Contad
'
'      For r_int_Contad = 2 To 15
'         .Range(.Cells(4, r_int_Contad), .Cells(10, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'      Next
'
'      .Range(.Cells(7, 3), .Cells(7, 14)).Merge
'      .Range(.Cells(7, 3), .Cells(7, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 3), .Cells(2, 15)).HorizontalAlignment = xlHAlignCenter
      .Cells(3, 1).Select
      r_obj_Excel.ActiveWindow.FreezePanes = True
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_FactElectronica()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_str_FecRpt        As String
      
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(l_int_PerMes, "00") & "/" & l_int_PerAno
   
   'Verifica que exista ruta
   'If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
   '   MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Columns("A").ColumnWidth = 8:   .Columns("A").HorizontalAlignment = xlHAlignCenter 'NRO
      .Columns("B").ColumnWidth = 15:  .Columns("B").HorizontalAlignment = xlHAlignCenter 'COL_01
      .Columns("C").ColumnWidth = 8:   .Columns("C").HorizontalAlignment = xlHAlignCenter 'COL_02
      .Columns("D").ColumnWidth = 25:  .Columns("D").HorizontalAlignment = xlHAlignCenter 'COL_03
      .Columns("E").ColumnWidth = 25:  .Columns("E").HorizontalAlignment = xlHAlignLeft   'COL_04
      .Columns("F").ColumnWidth = 16:  .Columns("F").HorizontalAlignment = xlHAlignRight: .Columns("F").NumberFormat = "###,###,##0.00" 'COL_05
      .Columns("G").ColumnWidth = 16:  .Columns("G").HorizontalAlignment = xlHAlignRight: .Columns("G").NumberFormat = "###,###,##0.00" 'COL_06
      .Columns("H").ColumnWidth = 7:   .Columns("H").HorizontalAlignment = xlHAlignCenter 'COL_07
      .Columns("I").ColumnWidth = 7:   .Columns("I").HorizontalAlignment = xlHAlignCenter 'COL_08
      .Columns("J").ColumnWidth = 14:  .Columns("J").HorizontalAlignment = xlHAlignCenter 'COL_09
      .Columns("K").ColumnWidth = 52:  .Columns("K").HorizontalAlignment = xlHAlignLeft   'COL_10
      .Columns("L").ColumnWidth = 7:   .Columns("L").HorizontalAlignment = xlHAlignCenter 'COL_11
      .Columns("M").ColumnWidth = 7:   .Columns("M").HorizontalAlignment = xlHAlignCenter 'COL_12
      .Columns("N").ColumnWidth = 7:   .Columns("N").HorizontalAlignment = xlHAlignCenter 'COL_13
      .Columns("O").ColumnWidth = 13:  .Columns("O").HorizontalAlignment = xlHAlignCenter 'COL_14
      .Columns("P").ColumnWidth = 14:  .Columns("P").HorizontalAlignment = xlHAlignCenter 'COL_15
      .Columns("Q").ColumnWidth = 7:   .Columns("Q").HorizontalAlignment = xlHAlignCenter 'COL_16
      .Columns("R").ColumnWidth = 13:  .Columns("R").HorizontalAlignment = xlHAlignCenter 'COL_17
      .Columns("S").ColumnWidth = 88:  .Columns("S").HorizontalAlignment = xlHAlignLeft   'COL_18
      .Columns("T").ColumnWidth = 7:   .Columns("T").HorizontalAlignment = xlHAlignCenter 'COL_19
      
      
      .Cells(2, 2) = "REPORTE DE FACTURAS ELECTRONICAS " & _
      UCase("Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(l_int_PerAno, "0000"))
      .Range(.Cells(2, 2), .Cells(2, 19)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 19)).Merge
      
      .Range(.Cells(4, 1), .Cells(4, 20)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 20)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 20)).HorizontalAlignment = xlHAlignCenter
      
      For r_int_Contad = 1 To grd_Listado.Rows
         .Cells(r_int_Contad + 3, 1) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 0)
         .Cells(r_int_Contad + 3, 2) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 1)
         .Cells(r_int_Contad + 3, 3) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 2)
         .Cells(r_int_Contad + 3, 4) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 3)
         .Cells(r_int_Contad + 3, 5) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 4)
         .Cells(r_int_Contad + 3, 6) = grd_Listado.TextMatrix(r_int_Contad - 1, 5)
         .Cells(r_int_Contad + 3, 7) = grd_Listado.TextMatrix(r_int_Contad - 1, 6)
         .Cells(r_int_Contad + 3, 8) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 7)
         .Cells(r_int_Contad + 3, 9) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 8)
         .Cells(r_int_Contad + 3, 10) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 9)
         .Cells(r_int_Contad + 3, 11) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 10)
         .Cells(r_int_Contad + 3, 12) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 11)
         .Cells(r_int_Contad + 3, 13) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 12)
         .Cells(r_int_Contad + 3, 14) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 13)
         .Cells(r_int_Contad + 3, 15) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 14)
         .Cells(r_int_Contad + 3, 16) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 15)
         .Cells(r_int_Contad + 3, 17) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 16)
         .Cells(r_int_Contad + 3, 18) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 17)
         .Cells(r_int_Contad + 3, 19) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 18)
         .Cells(r_int_Contad + 3, 20) = "'" & grd_Listado.TextMatrix(r_int_Contad - 1, 19)
      Next
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
   Call fs_GenAch_FactElectronica
   
End Sub

Private Function fs_GenAch_FactElectronica() As String
Dim r_int_NumRes   As Integer
Dim r_str_NomRes1  As String
Dim r_str_CadAux   As String
Dim r_int_Contad   As Long
Dim R_STR_CONSTT   As String

Dim r_int_UltMes   As Integer
Dim r_int_PerMes   As Integer
Dim r_int_PerAno   As Integer

   fs_GenAch_FactElectronica = ""
   
   If grd_Listado.Rows = 1 Then
      Exit Function
   End If
   '----------Creando Archivo----------
   r_int_PerMes = Format(grd_Listado.TextMatrix(1, 1), "mm")
   r_int_PerAno = Format(grd_Listado.TextMatrix(1, 1), "yyyy")
   r_int_UltMes = ff_Ultimo_Dia_Mes(r_int_PerMes, r_int_PerAno)
   r_str_CadAux = Format(grd_Listado.TextMatrix(1, 1), "yyyymm") & Format(r_int_UltMes, "00")
   
   r_str_NomRes1 = moddat_g_str_RutLoc & "\20511904162-IH-" & r_str_CadAux & "-01.TXT"
   R_STR_CONSTT = "|"
   
   r_int_NumRes = FreeFile
   Open r_str_NomRes1 For Output As r_int_NumRes
   
        For r_int_Contad = 1 To grd_Listado.Rows - 1
            Print #1, grd_Listado.TextMatrix(r_int_Contad, 1); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 2); R_STR_CONSTT; _
                      "-"; R_STR_CONSTT; Format(grd_Listado.TextMatrix(r_int_Contad, 0), "00000000000000000000"); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 5); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 6); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 7); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 8); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 9); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 10); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 11); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 12); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 13); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 14); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 15); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 16); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 17); R_STR_CONSTT; _
                      grd_Listado.TextMatrix(r_int_Contad, 18); R_STR_CONSTT; grd_Listado.TextMatrix(r_int_Contad, 19); R_STR_CONSTT
        Next
   
   Close #1
      
   fs_GenAch_FactElectronica = Trim(r_str_NomRes1)
   MsgBox "El archivo ha sido creado. " & Trim(r_str_NomRes1), vbInformation, modgen_g_str_NomPlt
End Function

Private Sub fs_GenExc_Indicadores()
Dim r_obj_Excel         As Excel.Application
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer
   
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(l_int_PerMes, "00") & "/" & l_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "INDICADORES RELACIONADOS A CUENTAS DE BALANCE"
      .Range(.Cells(1, 2), .Cells(1, 2)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(l_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 2)).Font.Bold = True
      
      .Cells(4, 2) = "INDICADOR"
      .Range(.Cells(4, 2), .Cells(4, 2)).HorizontalAlignment = xlHAlignCenter
      .Cells(4, 3) = "'" & "ENE " & Right(l_int_PerAno, 2)
      .Cells(4, 4) = "'" & "FEB " & Right(l_int_PerAno, 2)
      .Cells(4, 5) = "'" & "MAR " & Right(l_int_PerAno, 2)
      .Cells(4, 6) = "'" & "ABR " & Right(l_int_PerAno, 2)
      .Cells(4, 7) = "'" & "MAY " & Right(l_int_PerAno, 2)
      .Cells(4, 8) = "'" & "JUN " & Right(l_int_PerAno, 2)
      .Cells(4, 9) = "'" & "JUL " & Right(l_int_PerAno, 2)
      .Cells(4, 10) = "'" & "AGO " & Right(l_int_PerAno, 2)
      .Cells(4, 11) = "'" & "SET " & Right(l_int_PerAno, 2)
      .Cells(4, 12) = "'" & "OCT " & Right(l_int_PerAno, 2)
      .Cells(4, 13) = "'" & "NOV " & Right(l_int_PerAno, 2)
      .Cells(4, 14) = "'" & "DIC " & Right(l_int_PerAno, 2)
      
      .Range(.Cells(4, 2), .Cells(4, 14)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 14)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 2), .Cells(12, 2)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 33
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignRight
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 12
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 12
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 12
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 12
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 12
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 12
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 12
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 12
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 12
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 12
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 12
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 1 To grd_Listado.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 0), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 3) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 1), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 4) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 2), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 5) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 3), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 6) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 4), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 7) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 5), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 8) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 6), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 9) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 7), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 10) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 8), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 11) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 9), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 12) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 10), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 13) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 11), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 14) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 12), "###,###,##0.00")
         
         .Range(.Cells(r_int_NumFil + 2, 2), .Cells(r_int_NumFil + 3, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 2, 2), .Cells(r_int_NumFil + 3, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
      
      For r_int_Contad = 2 To 15
         .Range(.Cells(4, r_int_Contad), .Cells(10, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      Next
      
      .Range(.Cells(7, 3), .Cells(7, 14)).Merge
      .Range(.Cells(7, 3), .Cells(7, 14)).HorizontalAlignment = xlHAlignCenter
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Saldos()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer
   
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 3) = "REPORTE DE SALDOS DEL MES DE " & cmb_PerMes.Text & " DEL AÑO " & ipp_PerAno.Text
      .Range(.Cells(2, 3), .Cells(2, 3)).Font.Bold = True
      
      .Cells(4, 2) = "PRODUCTO"
      .Cells(4, 3) = "MONEDA PRESTAMO"
      .Cells(4, 4) = "MONTO PRESTAMO"
      .Cells(4, 5) = "SALDO PRESTAMO"
      .Cells(4, 6) = "SALDO SOLES"
      .Cells(4, 7) = "CLASIFICACION"

      .Range(.Cells(4, 1), .Cells(4, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 7)).Font.Bold = True
      .Range("A4:G4").HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 7
      .Columns("B").ColumnWidth = 60
      .Columns("C").ColumnWidth = 25
      .Columns("D").ColumnWidth = 22
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 22
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 22
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 15

      r_int_NumFil = 2
      For r_int_Contad = 1 To grd_Listado.Rows - 1
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
         
         .Cells(r_int_NumFil + 3, 1) = r_int_NumFil - 1
         .Cells(r_int_NumFil + 3, 2) = grd_Listado.TextMatrix(r_int_NumFil - 1, 1)
         .Cells(r_int_NumFil + 3, 3) = grd_Listado.TextMatrix(r_int_NumFil - 1, 2)
         .Cells(r_int_NumFil + 3, 4) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 3), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 5) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 4), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 6) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 5), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 7) = grd_Listado.TextMatrix(r_int_NumFil - 1, 6)

         .Range(.Cells(r_int_NumFil + 2, 1), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 2, 1), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 2, 1), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 2, 1), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 2, 1), .Cells(r_int_NumFil + 3, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous

         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer
   
   l_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   l_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(l_int_PerMes, "00") & "/" & l_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE RATIO DE CAPITAL AJUSTADO"
      .Range(.Cells(1, 2), .Cells(1, 2)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(l_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 2)).Font.Bold = True
      
      .Cells(4, 2) = "EJERCICIOS"
      .Cells(4, 3) = "'" & "ENE " & Right(l_int_PerAno, 2)
      .Cells(4, 4) = "'" & "FEB " & Right(l_int_PerAno, 2)
      .Cells(4, 5) = "'" & "MAR " & Right(l_int_PerAno, 2)
      .Cells(4, 6) = "'" & "ABR " & Right(l_int_PerAno, 2)
      .Cells(4, 7) = "'" & "MAY " & Right(l_int_PerAno, 2)
      .Cells(4, 8) = "'" & "JUN " & Right(l_int_PerAno, 2)
      .Cells(4, 9) = "'" & "JUL " & Right(l_int_PerAno, 2)
      .Cells(4, 10) = "'" & "AGO " & Right(l_int_PerAno, 2)
      .Cells(4, 11) = "'" & "SET " & Right(l_int_PerAno, 2)
      .Cells(4, 12) = "'" & "OCT " & Right(l_int_PerAno, 2)
      .Cells(4, 13) = "'" & "NOV " & Right(l_int_PerAno, 2)
      .Cells(4, 14) = "'" & "DIC " & Right(l_int_PerAno, 2)
      
      .Range(.Cells(4, 2), .Cells(4, 14)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 14)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 2), .Cells(12, 2)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 37
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignRight
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 15
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 1 To grd_Listado.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 0), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 3) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 1), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 4) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 2), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 5) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 3), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 6) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 4), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 7) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 5), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 8) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 6), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 9) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 7), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 10) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 8), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 11) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 9), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 12) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 10), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 13) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 11), "###,###,##0.00")
         .Cells(r_int_NumFil + 3, 14) = Format(grd_Listado.TextMatrix(r_int_NumFil - 1, 12), "###,###,##0.00")
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub grd_Listado_SelChange()
   If grd_Listado.Rows > 2 Then
      grd_Listado.RowSel = grd_Listado.Row
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub
