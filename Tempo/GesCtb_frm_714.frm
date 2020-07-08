VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   3945
   ClientTop       =   2715
   ClientWidth     =   5445
   Icon            =   "GesCtb_frm_714.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4260
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
         TabIndex        =   1
         Top             =   60
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Height          =   270
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Anexo N° 9"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   3
            Top             =   270
            Width           =   4545
            _Version        =   65536
            _ExtentX        =   8017
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0109-03 Posiciones Afectas al Riesgo Cambiario"
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
            Picture         =   "GesCtb_frm_714.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_714.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_714.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_714.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4770
            Picture         =   "GesCtb_frm_714.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_714.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2610
            Top             =   90
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
         Height          =   855
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   12
            Top             =   420
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
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim r_dbl_Evalua(200)   As Double
   Dim r_str_Denomi(20)    As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String
   Dim r_dbl_MulUso        As Double
   Dim r_str_Cadena        As String
   Dim r_str_NomRes        As String
   Dim r_str_ParAux        As String
   Dim r_dbl_Volati        As Double
   Dim r_dbl_TipCam        As Double
   Dim r_dbl_PatEfe        As Double
   Dim r_int_Cantid        As Integer
   Dim r_int_FlgRpr        As Integer
   
Private Sub cmd_Imprim_Click()
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
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
  Call fs_GenRpt
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   r_dbl_Volati = 0.01
      
   Screen.MousePointer = 0
End Sub
   
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Limpia()

   Dim r_int_PerMes  As Integer
   Dim r_int_PerAno  As Integer

   r_int_PerMes = Month(date)
   r_int_PerAno = Year(date)
   
   If Month(date) = 12 Then
      r_int_PerMes = 1
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If
 
   Call gs_BuscarCombo_Item(cmb_PerMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   
End Sub

Private Sub fs_Inicia()
         
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
End Sub

Private Sub cmd_ExpArc_Click()

   Dim r_int_MsgBox As Integer
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_PARICA", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("¿Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_GenArc
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_GenArc
  
End Sub

Private Sub cmd_ExpExc_Click()

   Dim r_int_MsgBox As Integer
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_PARICA", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("¿Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_GenExc
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_GenExc
  
End Sub

Private Sub fs_GenDat()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi

   r_str_Denomi(0) = "Dólar Americano"
   r_str_Denomi(1) = "Libra Esterlina"
   r_str_Denomi(2) = "Yen Japonés"
   r_str_Denomi(3) = "Dólar Canadienes"
   r_str_Denomi(4) = "Euro"
   r_str_Denomi(5) = "Otras Divisas"
   r_str_Denomi(6) = "Oro"
   r_str_Denomi(7) = "TOTAL VaR"
   r_str_Denomi(8) = "TOTAL 3*VaR"
   r_str_Denomi(9) = "TOTAL 3*VaR/P.E"

   Erase r_dbl_Evalua()
   
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & r_int_PerAno & " "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_dbl_TipCam = g_rst_Princi!HIPCIE_TIPCAM
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_PatEfe = 0
   
   g_str_Parame = "SELECT * FROM HIS_PATEFE WHERE "
   g_str_Parame = g_str_Parame & "PATEFE_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "PATEFE_PERANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "PATEFE_NUMITE = 0 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_dbl_PatEfe = 0
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_dbl_PatEfe = g_rst_Princi!PATEFE_MTOSOL
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
         
   g_str_Parame = "SELECT * FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "MES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "ANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "(CNTA_CTBL LIKE '1%' OR "
   g_str_Parame = g_str_Parame & "CNTA_CTBL LIKE '2%') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
         
         If Mid(Trim(g_rst_Princi!CNTA_CTBL), 3, 1) = "2" Or Mid(Trim(g_rst_Princi!CNTA_CTBL), 3, 1) = "2" Then
            r_int_Contad = 0
            r_dbl_Evalua(r_int_Contad) = r_dbl_Evalua(r_int_Contad) + Format((g_rst_Princi!IMP_MOVSOL * r_dbl_TipCam), "###########0.00")
            
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_Evalua(1) = r_dbl_Volati
         
   r_dbl_Evalua(2) = r_dbl_Evalua(0) + r_dbl_Evalua(1)
   
   r_dbl_Evalua(68) = Format((r_dbl_Evalua(2) / 1000) * 2.33 * Sqr(10), "############0.00")
   r_dbl_Evalua(77) = Format(r_dbl_Evalua(68) * 3, "############0.00")
   r_dbl_Evalua(86) = Format(r_dbl_Evalua(77) / r_dbl_PatEfe, "###0.0000000")
         
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer
   Dim r_str_TipMon     As String
   
   Screen.MousePointer = 11
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
                  
      
      .Range(.Cells(1, 1), .Cells(4, 3)).Font.Bold = True
            
      .Cells(1, 1) = "SUPERINTENDENCIA DE BANCA, SEGUROS Y AFP"
      .Cells(3, 1) = "EMPRESA: Edpyme MiCasita S.A."
      .Cells(4, 1) = "CODIGO: 240"

      .Cells(2, 3) = "ANEXO Nº9"
      .Cells(3, 3) = "POSICIONES AFECTAS A RIESGO CAMBIARIO"
      .Cells(4, 3) = "(Expresado En Nuevos Soles)"
      .Cells(5, 3) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      
      .Range(.Cells(2, 3), .Cells(5, 3)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 3), .Cells(2, 8)).Merge
      .Range(.Cells(3, 3), .Cells(3, 8)).Merge
      .Range(.Cells(4, 3), .Cells(4, 8)).Merge
      .Range(.Cells(5, 3), .Cells(5, 8)).Merge
            
      .Cells(8, 1) = "III. MODELOS DE VALOR EN RIESGO"
      .Range(.Cells(8, 1), .Cells(8, 1)).Font.Bold = True
                     
      .Cells(10, 1) = "Divisas"
      .Cells(10, 2) = "Modelo Regulatorio"
      .Cells(11, 2) = "Posición Global ME (H)"
      .Cells(11, 3) = "Volatilidad 21/"
      .Cells(11, 4) = "Valor en Riesgo (VaR)"
      .Cells(12, 4) = "Pos. Global 22/"
      .Cells(12, 5) = "Gamma"
      .Cells(12, 6) = "Vega"
      .Cells(12, 7) = "Total VaR"
      .Cells(10, 8) = "Modelo Interno"
      .Cells(11, 8) = "Exposición en M.E. 24/"
      .Cells(11, 9) = "Volatilidad 25/"
      .Cells(11, 10) = "VaR 26/"

      .Cells(14, 1) = "Dólar Americano"
      .Cells(15, 1) = "Libra Esterlina"
      .Cells(16, 1) = "Yen Japonés"
      .Cells(17, 1) = "Dólar Canadienes"
      .Cells(18, 1) = "Euro"
      .Cells(19, 1) = "Otras Divisas"
      .Cells(20, 1) = "Oro"
      .Cells(22, 1) = "TOTAL VaR"
      .Cells(23, 1) = "TOTAL 3*VaR"
      .Cells(24, 1) = "TOTAL 3*VaR/P.E"
     
      .Range(.Cells(10, 1), .Cells(12, 1)).Merge
      .Range(.Cells(10, 2), .Cells(10, 7)).Merge
      .Range(.Cells(11, 2), .Cells(12, 2)).Merge
      .Range(.Cells(11, 3), .Cells(12, 3)).Merge
      .Range(.Cells(11, 4), .Cells(11, 7)).Merge
      .Range(.Cells(10, 8), .Cells(10, 10)).Merge
      .Range(.Cells(11, 8), .Cells(12, 8)).Merge
      .Range(.Cells(11, 9), .Cells(12, 9)).Merge
      .Range(.Cells(11, 10), .Cells(12, 10)).Merge
            
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 10)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 10)).WrapText = True
      .Range(.Cells(10, 1), .Cells(12, 10)).Font.Bold = True
      
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(21, 1), .Cells(24, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(22, 1), .Cells(24, 10)).Font.Bold = True

      
      .Range(.Cells(22, 2), .Cells(24, 6)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(13, 7), .Cells(21, 7)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(22, 8), .Cells(24, 9)).Interior.Color = RGB(0, 0, 0)
            
      .Range(.Cells(28, 2), .Cells(29, 9)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(28, 2), .Cells(28, 3)).Merge
      .Range(.Cells(28, 5), .Cells(28, 6)).Merge
      .Range(.Cells(28, 9), .Cells(28, 9)).Merge
      
      .Range(.Cells(29, 2), .Cells(29, 3)).Merge
      .Range(.Cells(29, 5), .Cells(29, 6)).Merge
      .Range(.Cells(29, 9), .Cells(29, 9)).Merge
      
      .Range(.Cells(28, 2), .Cells(28, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 5), .Cells(28, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 9), .Cells(28, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(28, 2) = "Sr. Roberto Baba Yamamoto"
      .Cells(28, 5) = "Rossana Meza Bustamante"
      .Cells(28, 9) = "Javier Delgado Blanco"
      
      .Cells(29, 2) = "Gerente General"
      .Cells(29, 5) = "CPC Nº33526"
      .Cells(29, 9) = "Unidad de Riesgos"
      
      .Range(.Cells(32, 1), .Cells(32, 3)).Merge
      
      .Cells(32, 1) = "(1) Eliminadas las secciones I y II mediante la Resol. SBS Nº6328-2009 del 18.06.2009."
                                        
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 45
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 20
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 25
      .Columns("J").ColumnWidth = 15
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 15
            
      .Columns("B:M").NumberFormat = "###,###,##0.00"
      
      .Cells(24, 7).NumberFormat = "###,###,##0.0000000"
                   
   End With
         
   r_int_ConTem = 0
      
   For r_int_Contad = 14 To 24 Step 1
      For r_int_ConAux = 2 To 10 Step 1
         If r_int_Contad <> 21 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad, r_int_ConAux) = r_dbl_Evalua(r_int_ConTem)
            r_int_ConTem = r_int_ConTem + 1

         End If
      Next
            
   Next
   
   Call fs_GeneDB
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenArc()
  
   Dim r_int_NumRes     As Integer
   Dim r_int_CodEmp     As Integer
     
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   Dim r_dbl_MulUso     As Double
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\03" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".109"
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
   
      
   g_str_Parame = "SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "WHERE EMPGRP_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      r_int_CodEmp = g_rst_Princi!EMPGRP_CODSBS
   
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Print #r_int_NumRes, Format(109, "0000") & Format(3, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_str_Cadena = ""
   
   For r_int_Contad = 0 To 8 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_Contad), "########0.00"), 1, "0", 18)
   Next
   
   Print #r_int_NumRes, "0001" & "02" & r_str_Cadena
   
   r_str_Cadena = ""
   
   For r_int_Contad = 54 To 62 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_Contad), "########0.00"), 1, "0", 18)
   Next

   Print #r_int_NumRes, "0100" & "  " & r_str_Cadena
      
   r_int_ConAux = 63
   
   For r_int_Contad = 1000 To 1200 Step 100
      r_str_Cadena = ""
   
      For r_int_ConTem = 0 To 8 Step 1
         If r_int_Contad = 1200 And r_int_ConTem = 5 Then
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.0000000"), 1, "0", 18)
         Else
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
         End If
         r_int_ConAux = r_int_ConAux + 1
      Next
      
      Print #r_int_NumRes, Format(r_int_Contad, "0000") & "00" & r_str_Cadena
      
   Next
          
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
      
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt

End Sub

Private Sub fs_GeneDB()

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM HIS_PARICA WHERE "
   g_str_Parame = g_str_Parame & "PARICA_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "PARICA_PERANO = " & r_int_PerAno & "  "
   'g_str_Parame = g_str_Parame & "PARICA_USUCRE = '" & modgen_g_str_CodUsu & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = 0
   
   For r_int_Contad = 0 To 9 Step 1
   
      r_str_Cadena = "USP_HIS_PARICA ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_Contad) & "', "
      
      For r_int_ConTem = 0 To 8 Step 1
         If r_int_ConTem = 8 Then
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux) & ") "
         Else
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux) & ", "
         End If
         r_int_ConAux = r_int_ConAux + 1
      Next
          
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_PARICA.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
   Next

End Sub

Private Sub fs_GenRpt()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer
   Dim r_str_TipMon     As String
   
   Screen.MousePointer = 11
   
   Call fs_GenDat
   'Call fs_GeneDB
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
                  
      
      .Range(.Cells(1, 1), .Cells(4, 3)).Font.Bold = True
            
      .Cells(1, 1) = "SUPERINTENDENCIA DE BANCA, SEGUROS Y AFP"
      .Cells(3, 1) = "EMPRESA: Edpyme MiCasita S.A."
      .Cells(4, 1) = "CODIGO: 240"

      .Cells(2, 3) = "ANEXO Nº9"
      .Cells(3, 3) = "POSICIONES AFECTAS A RIESGO CAMBIARIO"
      .Cells(4, 3) = "(Expresado En Nuevos Soles)"
      .Cells(5, 3) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      
      .Range(.Cells(2, 3), .Cells(5, 3)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 3), .Cells(2, 8)).Merge
      .Range(.Cells(3, 3), .Cells(3, 8)).Merge
      .Range(.Cells(4, 3), .Cells(4, 8)).Merge
      .Range(.Cells(5, 3), .Cells(5, 8)).Merge
            
      .Cells(8, 1) = "III. MODELOS DE VALOR EN RIESGO"
      .Range(.Cells(8, 1), .Cells(8, 1)).Font.Bold = True
                     
      .Cells(10, 1) = "Divisas"
      .Cells(10, 2) = "Modelo Regulatorio"
      .Cells(11, 2) = "Posición Global ME (H)"
      .Cells(11, 3) = "Volatilidad 21/"
      .Cells(11, 4) = "Valor en Riesgo (VaR)"
      .Cells(12, 4) = "Pos. Global 22/"
      .Cells(12, 5) = "Gamma"
      .Cells(12, 6) = "Vega"
      .Cells(12, 7) = "Total VaR"
      .Cells(10, 8) = "Modelo Interno"
      .Cells(11, 8) = "Exposición en M.E. 24/"
      .Cells(11, 9) = "Volatilidad 25/"
      .Cells(11, 10) = "VaR 26/"

      .Cells(14, 1) = "Dólar Americano"
      .Cells(15, 1) = "Libra Esterlina"
      .Cells(16, 1) = "Yen Japonés"
      .Cells(17, 1) = "Dólar Canadienes"
      .Cells(18, 1) = "Euro"
      .Cells(19, 1) = "Otras Divisas"
      .Cells(20, 1) = "Oro"
      .Cells(22, 1) = "TOTAL VaR"
      .Cells(23, 1) = "TOTAL 3*VaR"
      .Cells(24, 1) = "TOTAL 3*VaR/P.E"
     
      .Range(.Cells(10, 1), .Cells(12, 1)).Merge
      .Range(.Cells(10, 2), .Cells(10, 7)).Merge
      .Range(.Cells(11, 2), .Cells(12, 2)).Merge
      .Range(.Cells(11, 3), .Cells(12, 3)).Merge
      .Range(.Cells(11, 4), .Cells(11, 7)).Merge
      .Range(.Cells(10, 8), .Cells(10, 10)).Merge
      .Range(.Cells(11, 8), .Cells(12, 8)).Merge
      .Range(.Cells(11, 9), .Cells(12, 9)).Merge
      .Range(.Cells(11, 10), .Cells(12, 10)).Merge
            
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 10)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 10)).WrapText = True
      .Range(.Cells(10, 1), .Cells(12, 10)).Font.Bold = True
      
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(24, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(21, 1), .Cells(24, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(22, 1), .Cells(24, 10)).Font.Bold = True

      
      .Range(.Cells(22, 2), .Cells(24, 6)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(13, 7), .Cells(21, 7)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(22, 8), .Cells(24, 9)).Interior.Color = RGB(0, 0, 0)
            
      .Range(.Cells(28, 2), .Cells(29, 9)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(28, 2), .Cells(28, 3)).Merge
      .Range(.Cells(28, 5), .Cells(28, 6)).Merge
      .Range(.Cells(28, 9), .Cells(28, 9)).Merge
      
      .Range(.Cells(29, 2), .Cells(29, 3)).Merge
      .Range(.Cells(29, 5), .Cells(29, 6)).Merge
      .Range(.Cells(29, 9), .Cells(29, 9)).Merge
      
      .Range(.Cells(28, 2), .Cells(28, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 5), .Cells(28, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 9), .Cells(28, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(28, 2) = "Sr. Roberto Baba Yamamoto"
      .Cells(28, 5) = "Rossana Meza Bustamante"
      .Cells(28, 9) = "Javier Delgado Blanco"
      
      .Cells(29, 2) = "Gerente General"
      .Cells(29, 5) = "CPC Nº33526"
      .Cells(29, 9) = "Unidad de Riesgos"
      
      .Range(.Cells(32, 1), .Cells(32, 3)).Merge
      
      .Cells(32, 1) = "(1) Eliminadas las secciones I y II mediante la Resol. SBS Nº6328-2009 del 18.06.2009."
                                        
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 45
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 20
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 25
      .Columns("J").ColumnWidth = 15
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 15
            
      .Columns("B:M").NumberFormat = "###,###,##0.00"
      
      .Cells(24, 7).NumberFormat = "###,###,##0.0000000"
                   
   End With
         
   r_int_ConTem = 0
      
   For r_int_Contad = 14 To 24 Step 1
      For r_int_ConAux = 2 To 10 Step 1
         If r_int_Contad <> 21 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad, r_int_ConAux) = r_dbl_Evalua(r_int_ConTem)
            r_int_ConTem = r_int_ConTem + 1

         End If
      Next
            
   Next
   
   Call fs_GeneDB
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

  
End Sub


Private Sub fs_GenDat_DB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi

   r_str_Denomi(0) = "Dólar Americano"
   r_str_Denomi(1) = "Libra Esterlina"
   r_str_Denomi(2) = "Yen Japonés"
   r_str_Denomi(3) = "Dólar Canadienes"
   r_str_Denomi(4) = "Euro"
   r_str_Denomi(5) = "Otras Divisas"
   r_str_Denomi(6) = "Oro"
   r_str_Denomi(7) = "TOTAL VaR"
   r_str_Denomi(8) = "TOTAL 3*VaR"
   r_str_Denomi(9) = "TOTAL 3*VaR/P.E"

   Erase r_dbl_Evalua()
   
   g_str_Parame = "SELECT * FROM HIS_PARICA WHERE "
   g_str_Parame = g_str_Parame & "PARICA_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "PARICA_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY PARICA_NUMITE ASC "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_ConAux = -1
   
      Do While Not g_rst_Princi.EOF
         
         r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!PARICA_MRPGME
         r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!PARICA_MRVOLA
         r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!PARICA_MRVRPG
         r_dbl_Evalua(r_int_ConAux + 4) = g_rst_Princi!PARICA_MGVRGA
         r_dbl_Evalua(r_int_ConAux + 5) = g_rst_Princi!PARICA_MRVRVE
         r_dbl_Evalua(r_int_ConAux + 6) = g_rst_Princi!PARICA_MRVRTO
         r_dbl_Evalua(r_int_ConAux + 7) = g_rst_Princi!PARICA_MIEXPO
         r_dbl_Evalua(r_int_ConAux + 8) = g_rst_Princi!PARICA_MIVOLA
         r_dbl_Evalua(r_int_ConAux + 9) = g_rst_Princi!PARICA_MOINVA
         
         r_int_ConAux = r_int_ConAux + 9
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'r_dbl_Evalua(1) = r_dbl_Volati
         
   'r_dbl_Evalua(2) = r_dbl_Evalua(0) + r_dbl_Evalua(1)
   
   'r_dbl_Evalua(68) = Format((r_dbl_Evalua(2) / 1000) * 2.33 * Sqr(10), "############0.00")
   'r_dbl_Evalua(77) = Format(r_dbl_Evalua(68) * 3, "############0.00")
   'r_dbl_Evalua(86) = Format(r_dbl_Evalua(77) / r_dbl_PatEfe, "###0.0000000")
         
End Sub





