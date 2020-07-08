VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   4185
   ClientTop       =   5565
   ClientWidth     =   7170
   Icon            =   "GesCtb_frm_004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   7011
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
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Height          =   570
            Left            =   630
            TabIndex        =   12
            Top             =   45
            Width           =   4275
            _Version        =   65536
            _ExtentX        =   7541
            _ExtentY        =   1005
            _StockProps     =   15
            Caption         =   "Reporte de Saldos de Créditos Hipotecarios"
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
            Left            =   60
            Picture         =   "GesCtb_frm_004.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Left            =   630
            Picture         =   "GesCtb_frm_004.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_004.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "GesCtb_frm_004.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1230
            Top             =   30
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2445
         Left            =   30
         TabIndex        =   14
         Top             =   1440
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   4313
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
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   1140
            TabIndex        =   6
            Top             =   2100
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Permes 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1350
            Width           =   5895
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todas las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   5895
         End
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   1050
            Width           =   1995
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   1140
            TabIndex        =   5
            Top             =   1740
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9999"
            MinValue        =   "1900"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   90
            TabIndex        =   18
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   1410
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub chk_Empres_Click()
   If chk_Empres.Value = 1 Then
      cmb_Empres.ListIndex = -1
      cmb_Empres.Enabled = False
      If cmb_TipPro.Enabled Then
         Call gs_SetFocus(cmb_TipPro)
      Else
         Call gs_SetFocus(cmb_PerMes)
      End If
   ElseIf chk_Empres.Value = 0 Then
      cmb_Empres.Enabled = True
      Call gs_SetFocus(cmb_Empres)
   End If
End Sub

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      cmb_PerMes.ListIndex = -1
      cmb_PerMes.Enabled = False
      ipp_PerAno.Value = 0
      ipp_PerAno.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf Chk_FecAct.Value = 0 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub chk_TipPro_Click()
   If chk_TipPro.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(cmb_PerMes)
   ElseIf chk_TipPro.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_Empres_Click()
   If cmb_TipPro.Enabled Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If Chk_FecAct.Value = 0 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Me.Enabled = False
   If cmb_PerMes.ListIndex = -1 Then
      Call fs_GenExc_FecAct
   Else
      Call fs_GenExc_Period
   End If
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If Chk_FecAct.Value = 0 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   If cmb_PerMes.ListIndex = -1 Then
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
      crp_Imprim.DataFiles(1) = "CLI_DATGEN"
      crp_Imprim.DataFiles(2) = "CRE_PRODUC"
      crp_Imprim.SelectionFormula = "{CRE_HIPMAE.HIPMAE_SITUAC} = 2 "
      
      If chk_TipPro.Value = 0 Then
         crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "AND {CRE_HIPMAE.HIPMAE_CODPRD} = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
      End If
      
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_02.RPT"
      crp_Imprim.Action = 1
   Else
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "CLI_DATGEN"
      crp_Imprim.DataFiles(1) = "CRE_PRODUC"
      crp_Imprim.DataFiles(2) = "CRE_HIPCIE"
      crp_Imprim.SelectionFormula = "{CRE_HIPCIE.HIPCIE_PERMES} = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPCIE.HIPCIE_PERANO} = " & Format(ipp_PerAno.Text, "0000") & ""
      
      If chk_TipPro.Value = 0 Then
         crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "AND {CRE_HIPCIE.HIPCIE_CODPRD} = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
      End If
      
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_03.RPT"
      crp_Imprim.Action = 1
   End If
         
         
         
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = -1
   chk_Empres.Value = 0
   cmb_TipPro.ListIndex = -1
   chk_TipPro.Value = 0
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub fs_GenExc_FecAct()
Dim r_obj_Excel      As Excel.Application
Dim r_str_PerMes     As String
Dim r_str_PerAno     As String
Dim r_int_ConVer     As Integer
Dim r_str_NumSol     As String
Dim r_dbl_PBPPer     As Double
Dim r_dbl_TipCam     As Double

  
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_NUMOPE, HIPMAE_TDOCLI, HIPMAE_NDOCLI, HIPMAE_FECDES, HIPMAE_MONEDA, HIPMAE_MTOPRE, "
   g_str_Parame = g_str_Parame & "       HIPMAE_INTCAP, HIPMAE_TOTPRE, HIPMAE_TASINT, HIPMAE_PLAANO, HIPMAE_SALCAP, HIPMAE_SALCON, HIPMAE_PRXVCT, "
   g_str_Parame = g_str_Parame & "       HIPMAE_ULTVCT, HIPMAE_ULTPAG, HIPMAE_VCTANT, HIPMAE_DIAMOR, HIPMAE_TIPGAR, HIPMAE_MONGAR, HIPMAE_MTOGAR, "
   g_str_Parame = g_str_Parame & "       HIPMAE_ACUDIF, HIPMAE_CONHIP, HIPMAE_TDOCYG, HIPMAE_NDOCYG, DATGEN_DIRELE, TRIM(N.PARDES_DESCRI) AS GENERO, "
   g_str_Parame = g_str_Parame & "       TRIM(O.PARDES_DESCRI) AS ESTADOCIVIL, TRIM(I.SUBPRD_DESCRI) AS SUBPRODUCTO, TRIM(J.PARDES_DESCRI) AS TIPO_EVAL, "
   g_str_Parame = g_str_Parame & "       TRIM(C.PARDES_DESCRI) AS MONEDA, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE, HIPMAE_GARLIN, PRODUC_DESCRI, "
   g_str_Parame = g_str_Parame & "       HIPGAR_FECCON, EVALEG_FECBLQ_INM, HIPMAE_NUMSOL, HIPMAE_CUOFIJ, NVL(EVACRE_INGTOT, 0) AS EVACRE_INGTOT, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_PRO = 7 THEN TRIM(K.DATGEN_RAZSOC) ELSE TRIM(SOLINM_RAZSOC_PRO) END,'-') AS NOM_PROMOTOR, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_CON = 7 THEN TRIM(L.DATGEN_RAZSOC) ELSE TRIM(SOLINM_RAZSOC_CON) END,'-') AS NOM_CONSTRUCTOR, "
   g_str_Parame = g_str_Parame & "       (SELECT PARDES_DESCRI FROM MNT_PARDES WHERE PARDES_CODGRP = '008' AND PARDES_CODITE= DATGEN_OCUPAC ) AS REGLABORAL, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HIPMAE_TIPGAR NOT IN (1,2) THEN"
   g_str_Parame = g_str_Parame & "           (SELECT PARDES_DESCRI FROM MNT_PARDES WHERE PARDES_CODGRP = '505' AND PARDES_CODITE= HIPMAE_BCOGAR ) "
   g_str_Parame = g_str_Parame & "       END AS BANCOGAR, "
   g_str_Parame = g_str_Parame & "       NVL(Q.EVATAS_VALREA_INM,0)+NVL(Q.EVATAS_VALREA_ES1,0)+NVL(Q.EVATAS_VALREA_ES2,0)+NVL(Q.EVATAS_VALREA_DEP,0) AS VAL_REALIZACION,'D' AS TIPO   "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN   ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC   ON PRODUC_CODIGO = HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = '204' AND C.PARDES_CODITE = HIPMAE_MONEDA "
   g_str_Parame = g_str_Parame & "  LEFT OUTER JOIN CRE_HIPGAR ON HIPGAR_NUMOPE = HIPMAE_NUMOPE AND HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "  LEFT OUTER JOIN TRA_EVALEG ON EVALEG_NUMSOL = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM   ON SOLINM_NUMSOL = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN K ON K.DATGEN_EMPTDO = SOLINM_TIPDOC_PRO AND K.DATGEN_EMPNDO = SOLINM_NUMDOC_PRO "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES N ON N.PARDES_CODGRP = '207' AND N.PARDES_CODITE = DATGEN_CODSEX "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES O ON O.PARDES_CODGRP = '205' AND O.PARDES_CODITE = DATGEN_ESTCIV "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN L ON L.DATGEN_EMPTDO = SOLINM_TIPDOC_CON AND L.DATGEN_EMPNDO = SOLINM_NUMDOC_CON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN TRA_EVACRE M ON M.EVACRE_NUMSOL = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE P ON P.SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SUBPRD I ON I.SUBPRD_CODPRD = HIPMAE_CODPRD AND I.SUBPRD_CODSUB = HIPMAE_CODSUB "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES J ON J.PARDES_CODGRP = '038' AND J.PARDES_CODITE = P.SOLMAE_TIPEVA "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS Q ON Q.EVATAS_NUMSOL = HIPMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2 "
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "AND HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_CODPRD ASC, HIPMAE_MONEDA ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "TIPO"
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "SUB-PRODUCTO"
      .Cells(1, 5) = "TIPO EVALUACION"
      .Cells(1, 6) = "OPERACION"
      .Cells(1, 7) = "DOC. TITULAR"
      .Cells(1, 8) = "NOMBRE CLIENTE"
      .Cells(1, 9) = "GENERO"
      .Cells(1, 10) = "ESTADO CIVIL"
      .Cells(1, 11) = "DOC. CONYUGE"
      .Cells(1, 12) = "F. DESEMBOLSO"
      .Cells(1, 13) = "CONSEJERO"
      .Cells(1, 14) = "MONEDA"
      .Cells(1, 15) = "MTO. PRESTAMO"
      .Cells(1, 16) = "INT. CAPIT."
      .Cells(1, 17) = "TOTAL PRESTAMO"
      .Cells(1, 18) = "T. INTERES"
      .Cells(1, 19) = "PLAZO"
      .Cells(1, 20) = "SALDO CAPITAL"
      .Cells(1, 21) = "SALDO TC"
      .Cells(1, 22) = "SALDO PBP"
      .Cells(1, 23) = "TOTAL SALDO"
      .Cells(1, 24) = "F. PROX. VCTO."
      .Cells(1, 25) = "F. ULT. VCTO."
      .Cells(1, 26) = "F. ULT. PAGO"
      .Cells(1, 27) = "F. VCTO ANT."
      .Cells(1, 28) = "DIA ATR."
      .Cells(1, 29) = "TIPO GARANTIA"
      .Cells(1, 30) = "BANCO GARANTIA"
      .Cells(1, 31) = "GARANTIA S/."
      .Cells(1, 32) = "GARANTIA US$."
      
      .Cells(1, 33) = "VALOR NETO REALIZACION S/."
      .Cells(1, 34) = "VALOR NETO REALIZACION US$."
      .Cells(1, 35) = "INT. DIFERIDO"
      .Cells(1, 36) = "PROYECTO MI CASITA"
      .Cells(1, 37) = "NOMBRE DEL PROYECTO"
      .Cells(1, 38) = "DIRECCIÓN"
      .Cells(1, 39) = "COD.EXCEP.(CREDITOS)"
      .Cells(1, 40) = "FINANCIAMIENTO"
      .Cells(1, 41) = "F. CONSTITUCION"
      .Cells(1, 42) = "NOMBRE DEL PROMOTOR"
      .Cells(1, 43) = "NOMBRE DEL CONSTRUCTOR"
      .Cells(1, 44) = "EXPOSICION RCC"
      .Cells(1, 45) = "SOBRE ENDEUDAMIENTO"
      .Cells(1, 46) = "RÉGIMEN LABORAL"
      .Cells(1, 47) = "CORREO ELECTRONICO"
      
      .Range(.Cells(1, 1), .Cells(1, 47)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 47)).HorizontalAlignment = xlHAlignCenter
      
      'ITEM
      .Columns("A").ColumnWidth = 6
      'TIPO
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      'PRODUCTO
      .Columns("C").ColumnWidth = 42
      'SUB-PRODUCTO
      .Columns("D").ColumnWidth = 60
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      'TIPO EVALUACION
      .Columns("E").ColumnWidth = 30
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      'OPERACION
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      'DOC. TITULAR
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      'NOMBRE CLIENTE
      .Columns("H").ColumnWidth = 45
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      'GENERO
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      'ESTADO CIVIL
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      'DOC. CONYUGE
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      'F.DESEMBOLSO
      .Columns("L").ColumnWidth = 16
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      'CONSEJERO
      .Columns("M").ColumnWidth = 16
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      'MONEDA
      .Columns("N").ColumnWidth = 24
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      'MTO. PRESTAMO
      .Columns("O").ColumnWidth = 18
      'INT. CAPIT.
      .Columns("P").ColumnWidth = 12
      'TOTAL PRESTAMO
      .Columns("Q").ColumnWidth = 18
      'T. INTERES
      .Columns("R").ColumnWidth = 12
      'PLAZO
      .Columns("S").ColumnWidth = 12
      'SALDO CAPITAL
      .Columns("T").ColumnWidth = 16
      'SALDO TC
      .Columns("U").ColumnWidth = 16
      'SALDO PBP
      .Columns("V").ColumnWidth = 12
      'TOTAL SALDO
      .Columns("W").ColumnWidth = 16
      'F.PROX.VCTO.
      .Columns("X").ColumnWidth = 16
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      'F.ULT.VCTO.
      .Columns("Y").ColumnWidth = 16
      .Columns("Y").HorizontalAlignment = xlHAlignCenter
      'F.ULT.PAGO
      .Columns("Z").ColumnWidth = 16
      .Columns("Z").HorizontalAlignment = xlHAlignCenter
      'F.VCTO.ANTERIOR
      .Columns("AA").ColumnWidth = 16
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      'DIA ATR.
      .Columns("AB").ColumnWidth = 10
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      'TIPO GARANTIA
      .Columns("AC").ColumnWidth = 30
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      'BANCO GARANTIA
      .Columns("AD").ColumnWidth = 30
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      'GARANTIA S/.
      .Columns("AE").ColumnWidth = 16
      'GARANTIA US$
      .Columns("AF").ColumnWidth = 16
      
      'VALOR DE REALIZACION S/.
      .Columns("AG").ColumnWidth = 27
      'VALOR DE REALIZACION US$
      .Columns("AH").ColumnWidth = 29
      'INT. DIFERIDO
      .Columns("AI").ColumnWidth = 15
      'PROYECTO MICASITA
      .Columns("AJ").ColumnWidth = 20
      .Columns("AJ").HorizontalAlignment = xlHAlignCenter
      'NOMBRE DEL PROYECTO
      .Columns("AK").ColumnWidth = 50
      'DIRECCION
      .Columns("AL").ColumnWidth = 0  '125
      'COD.EXCEP.(CREDITO)
      .Columns("AM").ColumnWidth = 21
      .Columns("AM").HorizontalAlignment = xlHAlignCenter
      'FINANCIAMIENTO
      .Columns("AN").ColumnWidth = 17
      .Columns("AN").HorizontalAlignment = xlHAlignCenter
      'F.CONSTITUCION
      .Columns("AO").ColumnWidth = 17
      .Columns("AO").HorizontalAlignment = xlHAlignCenter
      'NOMBRE DEL PROMOTOR
      .Columns("AP").ColumnWidth = 45
      'NOMBRE DEL CONSTRUCTOR
      .Columns("AQ").ColumnWidth = 45
      'EXPOSICION RCC
      .Columns("AR").ColumnWidth = 16
      .Columns("AR").HorizontalAlignment = xlHAlignCenter
      'SOBRE ENDEUDAMIENTO
      .Columns("AS").ColumnWidth = 24
      .Columns("AS").HorizontalAlignment = xlHAlignCenter
      'REGIMEN LABORAL
      .Columns("AT").ColumnWidth = 39
      'CORREO ELECTRONICA
      .Columns("AU").ColumnWidth = 0  '35
   End With
   
   'Obtiene periodo de la ultima carga del RCC
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES "
   g_str_Parame = g_str_Parame & "          FROM CLI_RCCCAB "
   g_str_Parame = g_str_Parame & "        ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) "
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY RCCCAB_PERANO DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   r_str_PerMes = g_rst_GenAux!RCCCAB_PERMES
   r_str_PerAno = g_rst_GenAux!RCCCAB_PERANO
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   'Obtiene el tipo de cambio del dia
   r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   r_dbl_TipCam = "3.26"
   If r_dbl_TipCam = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio para el dia de hoy. No se puede procesar la información.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Carga excel con la informacion del cursor principal
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!TIPO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SUBPRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!TIPO_EVAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!GENERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!ESTADOCIVIL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(g_rst_Princi!HIPMAE_TDOCYG) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCYG)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!HIPMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!Moneda)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!HIPMAE_INTCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!HIPMAE_TOTPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CStr(g_rst_Princi!HIPMAE_PLAANO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!HIPMAE_SALCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_dbl_PBPPer = ff_Calcula_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(r_dbl_PBPPer, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_UlTVCT)))
      
      If g_rst_Princi!HIPMAE_ULTPAG > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_ULTPAG)))
      End If
      If g_rst_Princi!HIPMAE_VCTANT > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_VCTANT)))
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = g_rst_Princi!HIPMAE_DIAMOR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Trim(g_rst_Princi!BANCOGAR)
      
      If g_rst_Princi!HIPMAE_MONGAR = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      End If
      If g_rst_Princi!HIPMAE_TIPGAR = 3 Or g_rst_Princi!HIPMAE_TIPGAR = 6 Then
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
         End If
      End If
      'cambio 2018-01-17
      If g_rst_Princi!HIPMAE_MONEDA = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!VAL_REALIZACION, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!VAL_REALIZACION, "###,###,##0.00")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(g_rst_Princi!HIPMAE_ACUDIF, "###,###,##0.00")
      'INICIALIZA VARIABLE SOLICITUD
      r_str_NumSol = "0"
      
      'OBTENER LOS DATOS DEL PROYECTO MI CASITA
      g_str_Parame = "SELECT * FROM CRE_SOLINM "
      g_str_Parame = g_str_Parame & "JOIN CRE_HIPMAE ON (HIPMAE_NUMSOL = SOLINM_NUMSOL) "
      g_str_Parame = g_str_Parame & "WHERE HIPMAE_NUMOPE = '" & g_rst_Princi!HIPMAE_NUMOPE & "' "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_str_NumSol = g_rst_GenAux!SOLINM_NUMSOL
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = moddat_gf_Consulta_ParDes("214", g_rst_GenAux!SOLINM_PRYMCS)
         
         If g_rst_GenAux!SOLINM_TABPRY = 2 Then
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD)) > 0 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
            Else
               If Len(Trim(g_rst_GenAux!SOLINM_PRYNOM)) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Trim(g_rst_GenAux!SOLINM_PRYNOM & "")
               End If
            End If
         Else
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD & "")) > 0 Then
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
            End If
         End If
      
         'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = moddat_gf_Consulta_ParDes("201", CStr(g_rst_GenAux!SOLINM_TIPVIA)) & " " & Trim(g_rst_GenAux!SOLINM_NOMVIA) & " " & Trim(g_rst_GenAux!SOLINM_NUMVIA) & _
         '               IIf(Len(Trim(g_rst_GenAux!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_GenAux!SOLINM_INTDPT) & ")", "") & IIf(Len(Trim(g_rst_GenAux!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_GenAux!SOLINM_TIPZON)) & " " & Trim(g_rst_GenAux!SOLINM_NOMZON), "")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = ""
      End If
      
      'OBTIENE NUMERO DE EXCEPCION SI LA HUBIERA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TRA_SEGEXC "
      g_str_Parame = g_str_Parame & " WHERE SEGEXC_NUMSOL = '" & g_rst_Princi!HIPMAE_NUMSOL & "' "
      g_str_Parame = g_str_Parame & "   AND SEGEXC_CODINS = 21"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF Or g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = g_rst_GenAux!SEGEXC_MOTEXC
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = ""
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = ""
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
            
      '----------------
      If g_rst_Princi!HIPMAE_GARLIN = "000002" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = "BID"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = ""
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = ""
      If g_rst_Princi!HIPMAE_TIPGAR = 1 Then
         If Len(Trim(g_rst_Princi!HIPGAR_FECCON)) > 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECCON)))
         End If
      ElseIf g_rst_Princi!HIPMAE_TIPGAR = 2 Then
         If Len(Trim(g_rst_Princi!EVALEG_FECBLQ_INM)) > 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM)))
         End If
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Trim(g_rst_Princi!NOM_PROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Trim(g_rst_Princi!NOM_CONSTRUCTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = moddat_gf_Consulta_ExposicionRCC(g_rst_Princi!HIPMAE_TDOCLI, g_rst_Princi!HIPMAE_NDOCLI, g_rst_Princi!HIPMAE_MONEDA, r_dbl_TipCam, g_rst_Princi!HIPMAE_CUOFIJ, g_rst_Princi!EVACRE_INGTOT)
      
      'Valida sobre endeudamiento del titular y del conyuge
      If moddat_gf_Consulta_SobreEndeudamiento(g_rst_Princi!HIPMAE_TDOCLI, g_rst_Princi!HIPMAE_NDOCLI, r_str_PerMes, r_str_PerAno) = "1" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = "SI - TIT"
         If Len(Trim(CStr(g_rst_Princi!HIPMAE_TDOCYG))) > 0 And Len(Trim(g_rst_Princi!HIPMAE_NDOCYG)) > 0 Then
            If moddat_gf_Consulta_SobreEndeudamiento(g_rst_Princi!HIPMAE_TDOCYG, g_rst_Princi!HIPMAE_NDOCYG, r_str_PerMes, r_str_PerAno) = "1" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = "SI - AMB"
            End If
         End If
      Else
         If Len(Trim(CStr(g_rst_Princi!HIPMAE_TDOCYG))) > 0 And Len(Trim(g_rst_Princi!HIPMAE_NDOCYG)) > 0 Then
            If moddat_gf_Consulta_SobreEndeudamiento(g_rst_Princi!HIPMAE_TDOCYG, g_rst_Princi!HIPMAE_NDOCYG, r_str_PerMes, r_str_PerAno) = "1" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = "SI - CYG"
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = "NO"
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = "NO"
         End If
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Trim(g_rst_Princi!REGLABORAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 47) = ""  'Trim(g_rst_Princi!DATGEN_DIRELE)
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Period()
Dim r_obj_Excel      As Excel.Application
Dim r_str_PerMes     As String
Dim r_str_PerAno     As String
Dim r_int_ConVer     As Integer
Dim r_dbl_TipCam     As Double
Dim r_str_FecCie     As String
Dim r_bol_FlgRCC     As Boolean
   
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " USP_RPT_SALDOS_CREDHIP ("
    g_str_Parame = g_str_Parame & "'" & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & "', "
    g_str_Parame = g_str_Parame & "'" & Format(ipp_PerAno.Text, "0000") & "', "
    g_str_Parame = g_str_Parame & "'REPORTE SALDOS CREHIP', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
    g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "') "
         
    DoEvents: DoEvents: DoEvents
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then 'g_rst_GenAux
       Exit Sub
    End If
    DoEvents: DoEvents: DoEvents
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "TIPO"
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "SUB-PRODUCTO"
      .Cells(1, 5) = "TIPO EVALUACION"
      .Cells(1, 6) = "MODALIDAD DE PRESTAMO"
      .Cells(1, 7) = "PRIMERA VIVIENDA"
      .Cells(1, 8) = "FINANCIAMIENTO"
      .Cells(1, 9) = "OPERACION"
      .Cells(1, 10) = "DNI CONSEJERO"
      .Cells(1, 11) = "CONSEJERO"
      .Cells(1, 12) = "CODIGO SBS"
      .Cells(1, 13) = "DOC. TITULAR"
      .Cells(1, 14) = "NOMBRE CLIENTE"
      .Cells(1, 15) = "FEC. NACIMIENTO"
      .Cells(1, 16) = "PAIS DE RESIDENCIA"
      .Cells(1, 17) = "CORREO ELECTRONICO"
      .Cells(1, 18) = "GENERO"
      .Cells(1, 19) = "NIVEL EDUCACION"
      .Cells(1, 20) = "PROFESION"
      .Cells(1, 21) = "RÉGIMEN LABORAL"
      .Cells(1, 22) = "TIPO DE RENTA"
      .Cells(1, 23) = "SECTOR ECONOMICO"
      
      .Cells(1, 24) = "CODIGO CIIU"
      .Cells(1, 25) = "CIIU"
      .Cells(1, 26) = "CENTRO LABORAL"
      .Cells(1, 27) = "DIRECCION LABORAL"
      .Cells(1, 28) = "DEPARTAMENTO LABORAL"
      .Cells(1, 29) = "DISTRITO LABORAL"
      .Cells(1, 30) = "UBIGEO LABORAL"
      
      .Cells(1, 31) = "INGRESO_LIQUIDO"
      .Cells(1, 32) = "INGRESO_NETO"
      .Cells(1, 33) = "ESTADO CIVIL"
      .Cells(1, 34) = "DOC. CONYUGE"
      .Cells(1, 35) = "F. SOLICITUD"
      .Cells(1, 36) = "F. DESEMBOLSO"
      .Cells(1, 37) = "MONEDA"
      .Cells(1, 38) = "VALOR VIVIENDA"
      
      .Cells(1, 39) = "APORTE PROPIO"
      .Cells(1, 40) = "MTO. PBP"
      .Cells(1, 41) = "MTO. BBP"
      .Cells(1, 42) = "MTO. BMS"
      .Cells(1, 43) = "MTO. AFP"
      
      .Cells(1, 44) = "GASTOS DE CIERRE"
      .Cells(1, 45) = "MTO. PRESTAMO"
      .Cells(1, 46) = "INT. CAPIT."
      .Cells(1, 47) = "TOTAL PRESTAMO"
      .Cells(1, 48) = "MONTO CUOTA"
      .Cells(1, 49) = "T. INTERES"
      .Cells(1, 50) = "INTERES COFIDE"
      .Cells(1, 51) = "COMISION COFIDE"
      .Cells(1, 52) = "PLAZO (AÑOS)"
      .Cells(1, 53) = "P. GRACIA"
      .Cells(1, 54) = "DIA PAGO"
      .Cells(1, 55) = "TIPO SEGURO"
      .Cells(1, 56) = "CUOTAS DOBLES"
      .Cells(1, 57) = "CUO. PAGADAS"
      .Cells(1, 58) = "CUO.PENDIENTES"
      .Cells(1, 59) = "SALDO CAPITAL"
      .Cells(1, 60) = "SALDO TC"
      .Cells(1, 61) = "SALDO PBP"
      .Cells(1, 62) = "TOTAL SALDO"
      .Cells(1, 63) = "TOTAL SALDO S/."
      .Cells(1, 64) = "CAPITAL VENCIDO"
      .Cells(1, 65) = "CAPITAL VIGENTE"
      .Cells(1, 66) = "F. PROX. VCTO."
      .Cells(1, 67) = "F. ULT. VCTO."
      .Cells(1, 68) = "F. ULT. PAGO"
      .Cells(1, 69) = "F. VCTO ANT."
      .Cells(1, 70) = "INT. DEVENGADO"
      .Cells(1, 71) = "INT. DIFERIDO"
      .Cells(1, 72) = "CUOTAS DIFERIDAS"
      .Cells(1, 73) = "DIA ATR."
      .Cells(1, 74) = "CUO. VENC."
      .Cells(1, 75) = "FECHA TASACION"
      .Cells(1, 76) = "MONEDA TASACION"
      .Cells(1, 77) = "SUMA SEGURADA"
      .Cells(1, 78) = "EXPOSICION RCC"
      .Cells(1, 79) = "SOBRE ENDEUDAMIENTO"
      .Cells(1, 80) = "TIPO GARANTIA"
      .Cells(1, 81) = "F. CONSTITUCION"
      .Cells(1, 82) = "FEC. EMISION C/F"
      .Cells(1, 83) = "ULT. VCTO. C/F"
      .Cells(1, 84) = "BANCO GARANTIA"
      .Cells(1, 85) = "GARANTIA S/."
      .Cells(1, 86) = "GARANTIA US$."
      
      .Cells(1, 87) = "VALOR NETO REALIZACION S/."
      .Cells(1, 88) = "VALOR NETO REALIZACION US$."
      
      .Cells(1, 89) = "MENOR S/. (G.PREFERIDA)"
      .Cells(1, 90) = "MENOR US$.(G.PREFERIDA)"
      
      .Cells(1, 91) = "GARANTIA NO PREFERIDA S/. "
      .Cells(1, 92) = "GARANTIA NO PREFERIDA US$."
      
      .Cells(1, 93) = "BLOQUEO >90 DIAS"
      .Cells(1, 94) = "HIPOTECA MATRIZ"
      .Cells(1, 95) = "F. PRESENTACION"
      .Cells(1, 96) = "F. INSCRIPCION"
      .Cells(1, 97) = "DOCUMENTO REGISTRAL"
      .Cells(1, 98) = "COD.EXCEP.(CREDITOS)"
      .Cells(1, 99) = "MOTIVO EXCEPCION"
      .Cells(1, 100) = "DESCRIPCION"
      .Cells(1, 101) = "ENT. REPORTADAS (CIERRE)"
      .Cells(1, 102) = "MONTO REPORTADO (CIERRE)"
      .Cells(1, 103) = "PROYECTO MI CASITA"
      .Cells(1, 104) = "NOMBRE DEL PROYECTO"
      .Cells(1, 105) = "NOMBRE DEL PROMOTOR"
      .Cells(1, 106) = "NOMBRE DEL CONSTRUCTOR"
      .Cells(1, 107) = "DIRECCIÓN DEL INMUEBLE"
      .Cells(1, 108) = "DEPARTAMENTO DEL INMUEBLE"
      .Cells(1, 109) = "DISTRITO DEL INMUEBLE"
      .Cells(1, 110) = "UBIGEO DEL INMUEBLE"
      .Cells(1, 111) = "C/F PARAGUA"

      .Range(.Cells(1, 1), .Cells(1, 111)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 111)).HorizontalAlignment = xlHAlignCenter
      
      'ITEM
      .Columns("A").ColumnWidth = 6
      'TIPO
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      'PRODUCTO
      .Columns("C").ColumnWidth = 42
      'SUB-PRODUCTO
      .Columns("D").ColumnWidth = 60
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      'TIPO EVALUACION
      .Columns("E").ColumnWidth = 30
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      'MODALIDAD DE PRESTAMO
      .Columns("F").ColumnWidth = 37
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      'PRIMERA VIVIENDA
      .Columns("G").ColumnWidth = 18
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      'FINANCIAMIENTO
      .Columns("H").ColumnWidth = 17
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      'OPERACION
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      'DNI CONSEJERO
      .Columns("J").ColumnWidth = 16
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      'CONSEJERO
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      'CODIGO SBS
      .Columns("L").ColumnWidth = 15
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      'DOC. TITULAR
      .Columns("M").ColumnWidth = 15
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      'NOMBRE CLIENTE
      .Columns("N").ColumnWidth = 45
      .Columns("N").HorizontalAlignment = xlHAlignLeft
      'FEC. NACIMIENTO
      .Columns("O").ColumnWidth = 16
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      'PAIS DE RESIDENCIA
      .Columns("P").ColumnWidth = 26
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      'CORREO ELECTRONICO
      .Columns("Q").ColumnWidth = 0  '35
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      'GENERO
      .Columns("R").ColumnWidth = 15
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      'NIVEL EDUCACION
      .Columns("S").ColumnWidth = 17
      'PROFESION
      .Columns("T").ColumnWidth = 50
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      'REGIMEN LABORAL
      .Columns("U").ColumnWidth = 39
      'TIPO DE RENTA
      .Columns("V").ColumnWidth = 19
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      'SECTOR ECONOMICO
      .Columns("W").ColumnWidth = 33
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      
      'CODIGO CIIU
      .Columns("X").ColumnWidth = 15
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      'CIIU
      .Columns("Y").ColumnWidth = 70
      'CENTRO LABORAL
      .Columns("Z").ColumnWidth = 0
      'DIRECCIÓN LABORAL
      .Columns("AA").ColumnWidth = 0  '70
      'DEPARTAMENTO LABORAL
      .Columns("AB").ColumnWidth = 27
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      'DISTRITO LABORAL
      .Columns("AC").ColumnWidth = 30
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      'UBIGEO LABORAL
      .Columns("AD").ColumnWidth = 30
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      
      'INGRESO_LIQUIDO
      .Columns("AE").ColumnWidth = 0  '17
      'INGRESO_NETO
      .Columns("AF").ColumnWidth = 0  '17
      'ESTADO CIVIL
      .Columns("AG").ColumnWidth = 15
      .Columns("AG").HorizontalAlignment = xlHAlignCenter
      'DOC. CONYUGE
      .Columns("AH").ColumnWidth = 16
      .Columns("AH").HorizontalAlignment = xlHAlignCenter
      'F.SOLICITUD
      .Columns("AI").ColumnWidth = 16
      .Columns("AI").HorizontalAlignment = xlHAlignCenter
      'F.DESEMBOLSO
      .Columns("AJ").ColumnWidth = 16
      .Columns("AJ").HorizontalAlignment = xlHAlignCenter
      'MONEDA
      .Columns("AK").ColumnWidth = 24
      .Columns("AK").HorizontalAlignment = xlHAlignCenter
      'VALOR VIVIENDA
      .Columns("AL").ColumnWidth = 18
      
      'APORTE PROPIO
      .Columns("AM").ColumnWidth = 18
      'MTO. PBP
      .Columns("AN").ColumnWidth = 18
      'MTO. BBP
      .Columns("AO").ColumnWidth = 18
      'MTO. BMS
      .Columns("AP").ColumnWidth = 18
      'MTO. AFP
      .Columns("AQ").ColumnWidth = 18
      
      'GASTOS DE CIERRE
      .Columns("AR").ColumnWidth = 18
      'MTO. PRESTAMO
      .Columns("AS").ColumnWidth = 18
      'INT. CAPIT.
      .Columns("AT").ColumnWidth = 12
      'TOTAL PRESTAMO
      .Columns("AU").ColumnWidth = 18
      'MONTO CUOTA
      .Columns("AV").ColumnWidth = 18
      'T. INTERES
      .Columns("AW").ColumnWidth = 12
      'INTERES COFIDE
      .Columns("AX").ColumnWidth = 15
      'COMISION COFIDE
      .Columns("AY").ColumnWidth = 17
      'PLAZO
      .Columns("AZ").ColumnWidth = 14
      'P. GRACIA
      .Columns("BA").ColumnWidth = 12
      .Columns("BA").HorizontalAlignment = xlHAlignCenter
      'DIA PAGO
      .Columns("BB").ColumnWidth = 12
      .Columns("BB").HorizontalAlignment = xlHAlignCenter
      'TIPO SEGURO
      .Columns("BC").ColumnWidth = 22
      'CUOTAS DOBLES
      .Columns("BD").ColumnWidth = 18
      .Columns("BD").HorizontalAlignment = xlHAlignCenter
      'CUO. PAGADAS
      .Columns("BE").ColumnWidth = 16
      'CUO.PENDIENTES
      .Columns("BF").ColumnWidth = 16
      'SALDO CAPITAL
      .Columns("BG").ColumnWidth = 16
      'SALDO TC
      .Columns("BH").ColumnWidth = 16
      'SALDO PBP
      .Columns("BI").ColumnWidth = 12
      'TOTAL SALDO
      .Columns("BJ").ColumnWidth = 16
      'TOTAL SALDO S/.
      .Columns("BK").ColumnWidth = 16
      'CAPITAL VENCIDO
      .Columns("BL").ColumnWidth = 17
      'CAPITAL VIGENTE
      .Columns("BM").ColumnWidth = 17
      'F.PROX.VCTO.
      .Columns("BN").ColumnWidth = 16
      .Columns("BN").HorizontalAlignment = xlHAlignCenter
      'F.ULT.VCTO.
      .Columns("BO").ColumnWidth = 16
      .Columns("BO").HorizontalAlignment = xlHAlignCenter
      'F.ULT.PAGO
      .Columns("BP").ColumnWidth = 16
      .Columns("BP").HorizontalAlignment = xlHAlignCenter
      'F.VCTO.ANTERIOR
      .Columns("BQ").ColumnWidth = 16
      .Columns("BQ").HorizontalAlignment = xlHAlignCenter
      'INT. DEVENGADO
      .Columns("BR").ColumnWidth = 16
      'INT. DIFERIDO
      .Columns("BS").ColumnWidth = 15
      'CUO. DIFERIDAS
      .Columns("BT").ColumnWidth = 18
      .Columns("BT").HorizontalAlignment = xlHAlignCenter
      'DIAS ATRASADOS
      .Columns("BU").ColumnWidth = 10
      .Columns("BU").HorizontalAlignment = xlHAlignCenter
      'CUO. VENCIDAS
      .Columns("BV").ColumnWidth = 12
      .Columns("BV").HorizontalAlignment = xlHAlignCenter
      'FECHA DE TASACION
      .Columns("BW").ColumnWidth = 17
      .Columns("BW").HorizontalAlignment = xlHAlignCenter
      'MONEDA DE TASACION
      .Columns("BX").ColumnWidth = 21
      .Columns("BX").HorizontalAlignment = xlHAlignCenter
      'SUMA ASEGURADA
      .Columns("BY").ColumnWidth = 17
      'EXPOSICION RCC
      .Columns("BZ").ColumnWidth = 16
      .Columns("BZ").HorizontalAlignment = xlHAlignCenter
      'SOBRE ENDEUDAMIENTO
      .Columns("CA").ColumnWidth = 24
      .Columns("CA").HorizontalAlignment = xlHAlignCenter
      'TIPO GARANTIA
      .Columns("CB").ColumnWidth = 30
      .Columns("CB").HorizontalAlignment = xlHAlignCenter
      'F.CONSTITUCION
      .Columns("CC").ColumnWidth = 17
      .Columns("CC").HorizontalAlignment = xlHAlignCenter
      'FEC. EMISION C/F
      .Columns("CD").ColumnWidth = 17
      .Columns("CD").HorizontalAlignment = xlHAlignCenter
      'FECHA ULTIMO VENCIMIENTO C/F
      .Columns("CE").ColumnWidth = 15
      .Columns("CE").HorizontalAlignment = xlHAlignCenter
      'BANCO GARANTIA
      .Columns("CF").ColumnWidth = 30
      .Columns("CF").HorizontalAlignment = xlHAlignCenter
      'GARANTIA S/.
      .Columns("CG").ColumnWidth = 16
      'GARANTIA US$
      .Columns("CH").ColumnWidth = 16
      
      'VALOR REALIZACION S/.
      .Columns("CI").ColumnWidth = 27
      'VALOR REALIZACION US$
      .Columns("CJ").ColumnWidth = 29
      
      'MENOR S/. (G.PREFERIDA)
      .Columns("CK").ColumnWidth = 27
      'MENOR US$ (G.PREFERIDA)
      .Columns("CL").ColumnWidth = 27
      
      'GARANTIA NO PREFERIDA S/.
      .Columns("CM").ColumnWidth = 27
      'GARANTIA NO PREFERIDA US$
      .Columns("CN").ColumnWidth = 27
      
      'BLOQUEO REGISTRAL > 90 DIAS
      .Columns("CO").ColumnWidth = 18
      'HIPOTECA MATRIZ
      .Columns("CP").ColumnWidth = 18
      .Columns("CP").HorizontalAlignment = xlHAlignCenter

      'F. PRESENTACION
      .Columns("CQ").ColumnWidth = 16
      .Columns("CQ").HorizontalAlignment = xlHAlignCenter
      'F. INSCRIPCION
      .Columns("CR").ColumnWidth = 16
      .Columns("CR").HorizontalAlignment = xlHAlignCenter
      'DOCUMENTO REGISTRAL
      .Columns("CS").ColumnWidth = 75
      'COD.EXCEP.(CREDITO)
      .Columns("CT").ColumnWidth = 21
      .Columns("CT").HorizontalAlignment = xlHAlignCenter
      'MOTIVO DE EXCEPCION
      .Columns("CU").ColumnWidth = 60
      'DESCRIPCION
      .Columns("CV").ColumnWidth = 80
      .Columns("CV").HorizontalAlignment = xlHAlignCenter
      'ENT. REPORTADAS (CIERRE)
      .Columns("CW").ColumnWidth = 27
      .Columns("CW").HorizontalAlignment = xlHAlignCenter
      'MONTO REPORTADO (CIERRE)
      .Columns("CX").ColumnWidth = 27
      'PROYECTO MICASITA
      .Columns("CY").ColumnWidth = 20
      .Columns("CY").HorizontalAlignment = xlHAlignCenter
      'NOMBRE DEL PROYECTO
      .Columns("CZ").ColumnWidth = 50
      'NOMBRE DEL PROMOTOR
      .Columns("DA").ColumnWidth = 45
      'NOMBRE DEL CONSTRUCTOR
      .Columns("DB").ColumnWidth = 45
      .Columns("DB").HorizontalAlignment = xlHAlignCenter
      'DIRECCION
      .Columns("DC").ColumnWidth = 0  '120
      .Columns("DC").HorizontalAlignment = xlHAlignCenter
      'DEPARTAMENTO
      .Columns("DD").ColumnWidth = 30
      .Columns("DD").HorizontalAlignment = xlHAlignCenter
      'DISTRITO
      .Columns("DE").ColumnWidth = 35
      .Columns("DE").HorizontalAlignment = xlHAlignCenter
      'UBIGEO INMUEBLE
      .Columns("DF").ColumnWidth = 35
      .Columns("DF").HorizontalAlignment = xlHAlignCenter
      'C/F PARAGUA
      .Columns("DG").ColumnWidth = 20
      .Columns("DG").HorizontalAlignment = xlHAlignCenter
   End With
   
   'Obtiene periodo anterior para la consulta del RCC
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1 Then
      r_str_PerMes = 12
      r_str_PerAno = CInt(ipp_PerAno.Text) - 1
   Else
      r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) - 1
      r_str_PerAno = CInt(ipp_PerAno.Text)
   End If
   
   'Obtiene Fecha del Cierre
   r_str_FecCie = Format(CInt(ipp_PerAno.Text), "####") & Format(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), "0#") & Format(ff_Ultimo_Dia_Mes(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(ipp_PerAno.Text)), "00")
   
   'Carga excel con la informacion del cursor principal
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!TIPO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SUBPRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!TIPO_EVAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!MOD_PRESTAMO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!PRI_VIVIENDA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!FINANCIAMIENTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "'" & Trim(g_rst_Princi!DNI_CONSEJERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!CONSEJERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!CODIGO_SBS)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = g_rst_Princi!DOC_TITULAR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = g_rst_Princi!NOMBRE_CLIENTE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!FEC_NACIMIENTO))))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!PAIS_RESIDENCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = ""   'Trim(g_rst_Princi!CORREO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Trim(g_rst_Princi!GENERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Trim(g_rst_Princi!NIVEL_ESTUDIO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(g_rst_Princi!PROFESION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Trim(g_rst_Princi!REGLABORAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Trim(g_rst_Princi!TIPO_RENTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!SECTOR_ECONOMICO)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = "'" & Format(Trim(g_rst_Princi!CODCIUU), "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Trim(g_rst_Princi!CIUU)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = ""  'Trim(g_rst_Princi!CENTRO_LABORAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = ""  'Trim(g_rst_Princi!DIRECCION_LABORAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Trim(g_rst_Princi!DEPARTAMENTO_LABORAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Trim(g_rst_Princi!DISTRITO_LABORAL)
      If CLng(Trim(g_rst_Princi!UBIGEO_LABORAL & "") & "0") > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = "'" & Format(Trim(g_rst_Princi!UBIGEO_LABORAL), "000000")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = ""
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = "0.00"   'Format(g_rst_Princi!INGRESO_LIQUIDO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = "0.00"   'Format(g_rst_Princi!INGRESO_NETO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(g_rst_Princi!ESTADO_CIVIL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = g_rst_Princi!DOC_CONYUGE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_SOLICITUD)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_DESEMBOLSO)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Trim(g_rst_Princi!Moneda)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(g_rst_Princi!VALOR_VIVIENDA, "###,###,##0.00")
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(g_rst_Princi!APORTE_PROPIO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(g_rst_Princi!MTO_PBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(g_rst_Princi!MTO_BBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(g_rst_Princi!MTO_BMS, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Format(g_rst_Princi!MTO_AFP, "###,###,##0.00")
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = Format(g_rst_Princi!GASTOS_CIERRE, "###,###,##0.00")
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format(g_rst_Princi!MTO_PRESTAMO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Format(g_rst_Princi!INT_CAPIT, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 47) = Format(g_rst_Princi!TOTAL_PRESTAMO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 48) = Format(g_rst_Princi!MONTO_CUOTA, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(g_rst_Princi!TASA_INTERES, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 50) = Format(g_rst_Princi!INTERES_COFIDE, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 51) = Format(g_rst_Princi!COMISION_COFIDE, "##0.0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 52) = CStr(g_rst_Princi!PLAZO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 53) = g_rst_Princi!PER_GRACIA
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 54) = g_rst_Princi!DIA_PAGO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = Trim(g_rst_Princi!TIPO_SEGURO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Trim(g_rst_Princi!CUOTAS_DOBLES)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = g_rst_Princi!CUOTAS_PAGADAS
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = g_rst_Princi!CUOTAS_PEND
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 59) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = Format(g_rst_Princi!SALDO_TC, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = Format(g_rst_Princi!SALDO_PBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(g_rst_Princi!TOTAL_SALDO, "###,###,##0.00")
      If g_rst_Princi!TIPO_MONEDA = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = Format(g_rst_Princi!TOTAL_SALDO * g_rst_Princi!TIPO_CAMBIO, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = Format(g_rst_Princi!TOTAL_SALDO, "###,###,##0.00")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 64) = Format(g_rst_Princi!CAPITAL_VENCIDO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = Format(g_rst_Princi!CAPITAL_VIGENTE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 66) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_PROX_VCTO)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 67) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_ULT_VCTO)))
      If g_rst_Princi!FEC_ULT_PAGO > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 68) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_ULT_PAGO)))
      End If
      If g_rst_Princi!FEC_VCTO_ANT > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 69) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_VCTO_ANT)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 70) = Format(g_rst_Princi!INT_DEVENGADO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 71) = Format(g_rst_Princi!INT_DIFERIDO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 72) = Format(g_rst_Princi!CUOTAS_DIFERIDAS, "##0")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 73) = g_rst_Princi!DIA_ATRASO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 74) = g_rst_Princi!CUOTAS_VENCIDAS
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 75) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_TASACION)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 76) = Trim(g_rst_Princi!MONEDA_TASACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 77) = Format(g_rst_Princi!SUMA_ASEGURADA, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 78) = g_rst_Princi!EXPOSICION_RCC
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 79) = g_rst_Princi!SOBRE_ENDEUD
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 80) = Trim(g_rst_Princi!TIPO_GARANTIA)
      If Not IsNull(g_rst_Princi!FEC_CONSTITUCION) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 81) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_CONSTITUCION)))
      End If
      If Not IsNull(g_rst_Princi!FEC_EMISION) Then
         If Not g_rst_Princi!FEC_EMISION = 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 82) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_EMISION)))
         End If
      End If
      If Trim(g_rst_Princi!ULT_VCTO) <> 0 Then
        r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 83) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!ULT_VCTO)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 84) = Trim(g_rst_Princi!BANCO_GARANTIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 85) = Format(g_rst_Princi!GARANTIA_SOL, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 86) = Format(g_rst_Princi!GARANTIA_DOL, "###,###,##0.00")
      
      '000000000000000000000
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 87) = Format(g_rst_Princi!VAL_REALIZA_SOL, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 88) = Format(g_rst_Princi!VAL_REALIZA_DOL, "###,###,##0.00")
                                          
      If Trim(g_rst_Princi!TIPO_GARANTIA) = "HIPOTECA" Or Trim(g_rst_Princi!TIPO_GARANTIA) = "GARANTIA HIPOTECARIA" Or Trim(g_rst_Princi!TIPO_GARANTIA) = "CARTA FIANZA" Or Trim(g_rst_Princi!TIPO_GARANTIA) = "BLOQUEO REGISTRAL" Then
         
         If g_rst_Princi!CLASIFICACION_CLIENTE = 3 Or g_rst_Princi!CLASIFICACION_CLIENTE = 4 Then
            If ff_Calcula_Clasifica_Dudosa(g_rst_Princi!OPERACION) = True Or ff_Calcula_Clasifica_Perdida(g_rst_Princi!OPERACION) = True Then
               GoTo Saltar
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 91) = Format(g_rst_Princi!GARANTIA_SOL, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 92) = Format(g_rst_Princi!GARANTIA_DOL, "###,###,##0.00")
            End If
         Else
Saltar:
            If g_rst_Princi!GARANTIA_SOL > 0 Then
               If g_rst_Princi!GARANTIA_SOL <= g_rst_Princi!VAL_REALIZA_SOL Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 89) = Format(g_rst_Princi!GARANTIA_SOL, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 89) = Format(g_rst_Princi!VAL_REALIZA_SOL, "###,###,##0.00")
               End If
            End If
            
            If g_rst_Princi!GARANTIA_DOL > 0 Then
               If g_rst_Princi!GARANTIA_DOL <= g_rst_Princi!VAL_REALIZA_DOL Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 90) = Format(g_rst_Princi!GARANTIA_DOL, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 90) = Format(g_rst_Princi!VAL_REALIZA_DOL, "###,###,##0.00")
               End If
            End If
         End If
         
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 91) = Format(g_rst_Princi!GARANTIA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 92) = Format(g_rst_Princi!GARANTIA_DOL, "###,###,##0.00")
      End If
      
      If Trim(g_rst_Princi!TIPO_GARANTIA) = "BLOQUEO REGISTRAL" And g_rst_Princi!NUMDIA_BLOQ_REG > 90 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 91) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 85), "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 92) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 86), "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 89) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 90) = ""
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 93) = Format(g_rst_Princi!BLOQMAY_90DIAS, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 94) = g_rst_Princi!MATRIZ
      If (IsNull(g_rst_Princi!F_PRESENTACION) = False) Then
          r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 95) = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!F_PRESENTACION))))
      End If
      If (IsNull(g_rst_Princi!F_INSCRIPCION) = False) Then
          r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 96) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!F_INSCRIPCION)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 97) = Trim(g_rst_Princi!DOC_REGISTRAL)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 98) = g_rst_Princi!COD_EXCEP
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 99) = Trim(g_rst_Princi!MOT_EXCEPCION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 100) = Trim(g_rst_Princi!DESCRIPCION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 101) = g_rst_Princi!ENT_REPORTADAS
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 102) = Format(g_rst_Princi!MONTO_REPORTADO, "###,###,##0.00")
      
      'Proyecto miCasita
      If Not IsNull(g_rst_Princi!PROYECTO_MICASITA) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 103) = Trim(g_rst_Princi!PROYECTO_MICASITA)
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 103) = ""
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 104) = Trim(g_rst_Princi!NOMBRE_PROYECTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 105) = CStr(g_rst_Princi!NOM_PROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 106) = CStr(g_rst_Princi!NOM_CONSTRUCTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 107) = ""   'Trim(g_rst_Princi!DIRECCION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 108) = Trim(g_rst_Princi!DEPARTAMENTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 109) = Trim(g_rst_Princi!DISTRITO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 110) = "'" & Format(Trim(g_rst_Princi!UBIGEO_INMUEBLE), "000000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 111) = "NO"
            
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function ff_Calcula_Clasifica_Dudosa(ByVal p_NumOpe As String) As Boolean
   
   ff_Calcula_Clasifica_Dudosa = False
   
   '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
      g_str_CadCnx = ""
      g_str_CadCnx = g_str_CadCnx & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
      g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
      g_str_CadCnx = g_str_CadCnx & "          FROM CRE_HIPCIE "
      g_str_CadCnx = g_str_CadCnx & "         WHERE HIPCIE_PERMES > 0 "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_PERANO > 2010 "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_CLAPRV = 3 "
      g_str_CadCnx = g_str_CadCnx & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      g_str_CadCnx = g_str_CadCnx & " WHERE ROWNUM < 37 "
      g_str_CadCnx = g_str_CadCnx & " GROUP BY HIPCIE_CLAPRV "
      
      If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
         Exit Function
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         If g_rst_GenAux!CONTADOR <= 36 Then
            ff_Calcula_Clasifica_Dudosa = True
         End If
      End If
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing

 End Function
 Private Function ff_Calcula_Clasifica_Perdida(ByVal p_NumOpe As String) As Boolean
 
   ff_Calcula_Clasifica_Perdida = False
   
   '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
      g_str_CadCnx = ""
      g_str_CadCnx = g_str_CadCnx & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
      g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
      g_str_CadCnx = g_str_CadCnx & "          FROM CRE_HIPCIE "
      g_str_CadCnx = g_str_CadCnx & "         WHERE HIPCIE_PERMES > 0 "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_PERANO > 2009 "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_CLAPRV = 4 "
      g_str_CadCnx = g_str_CadCnx & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      g_str_CadCnx = g_str_CadCnx & " WHERE ROWNUM < 25 "
      g_str_CadCnx = g_str_CadCnx & " GROUP BY HIPCIE_CLAPRV "
      
      If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
         Exit Function
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         If g_rst_GenAux!CONTADOR <= 24 Then
            ff_Calcula_Clasifica_Perdida = True
         End If
      End If
End Function
Private Function ff_Calcula_PBPPerdido(ByVal p_NumOpe As String) As Double
Dim r_rst_Genera        As ADODB.Recordset

   ff_Calcula_PBPPerdido = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(HIPCUO_CAPBBP) - SUM(HIPCUO_CBPPAG) AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_CAPBBP > 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   r_rst_Genera.MoveFirst
   If Not IsNull(r_rst_Genera!total) Then
      ff_Calcula_PBPPerdido = r_rst_Genera!total
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

