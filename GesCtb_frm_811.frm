VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   2670
   ClientTop       =   5145
   ClientWidth     =   7200
   Icon            =   "GesCtb_frm_811.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   6429
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
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
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
            Height          =   300
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Riesgo Hipotecario"
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
            Picture         =   "GesCtb_frm_811.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   780
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
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
            Left            =   6510
            Picture         =   "GesCtb_frm_811.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_811.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_811.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
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
         Height          =   2115
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
         _ExtentY        =   3731
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
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1380
            Width           =   2265
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   90
            Width           =   5895
         End
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todos las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   10
            Top             =   450
            Width           =   1995
         End
         Begin VB.CheckBox chk_TipRie 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   9
            Top             =   1080
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipRie 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   750
            Width           =   5895
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1140
            TabIndex        =   16
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
            Left            =   90
            TabIndex        =   17
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   90
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   12
            Top             =   750
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera

Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_Empres_Click()
   If chk_Empres.Value = 1 Then
      cmb_Empres.ListIndex = -1
      cmb_Empres.Enabled = False
   
      If cmb_TipRie.Enabled Then
         Call gs_SetFocus(cmb_TipRie)
      Else
         Call gs_SetFocus(cmb_PerMes)
      End If
   
   ElseIf chk_Empres.Value = 0 Then
      cmb_Empres.Enabled = True
      Call gs_SetFocus(cmb_Empres)
   End If
End Sub

Private Sub chk_TipRie_Click()
   If chk_TipRie.Value = 1 Then
      cmb_TipRie.ListIndex = -1
      cmb_TipRie.Enabled = False
      
      Call gs_SetFocus(cmb_PerMes)
      
   ElseIf chk_TipRie.Value = 0 Then
      cmb_TipRie.Enabled = True
      
      Call gs_SetFocus(cmb_TipRie)
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_TipRie.Enabled Then
      Call gs_SetFocus(cmb_TipRie)
   Else
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
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
   
   If chk_TipRie.Value = 0 Then
      If cmb_TipRie.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipRie)
         Exit Sub
      End If
   End If
   
   'If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
   '   MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(ipp_FecFin)
   '   Exit Sub
   'End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_TipMon As String
      
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
      
   If chk_TipRie.Value = 0 Then
      If cmb_TipRie.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipRie)
         Exit Sub
      End If
   End If
   
   Call fs_GenExc_34
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call fs_Inicia
      
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR CIIU")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 1
      
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR UBICACIONES GEOGRAFICAS")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 2
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR TIPO DE PROYECTOS")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 3
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR PROYECTOS VINCULADOS")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 4
      
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR TIPO DE ACTIVIDAD ECONOMICA")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 5
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR TIPO DE EVALUACION CREDITICIA")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 6
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR PRODUCTO")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 7
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR DIA DE ATRASO")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 8
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR DIA DE ATRASO")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 9
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR TIPO DE GARANTIA")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 10
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR PLAZO")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 11
   
   cmb_TipRie.AddItem ("DISTRIBUCION DE SALDOS POR TASA DE INTERES")
   cmb_TipRie.ItemData(cmb_TipRie.NewIndex) = 12
      
   
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = -1
   chk_Empres.Value = 0

   cmb_TipRie.ListIndex = -1
   chk_TipRie.Value = 0
   
   
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   
'--    Para Distribución de Saldos por Ubicaciones Geograficas
'      select hipmae_ubigeo, pardes_descri, hipmae_moneda, count(hipmae_salcap), sum(hipmae_salcap + hipmae_salcon)
'      from cre_hipmae a, cli_datgen b, mnt_pardes c
'      Where
'      hipmae_situac = 2             and
'      hipmae_tdocli = datgen_tipdoc and
'      hipmae_ndocli = datgen_numdoc and
'      pardes_codgrp = '101' and
'      hipmae_ubigeo = pardes_codite
'      Group By
'      hipmae_ubigeo , pardes_descri, hipmae_moneda
   
   

   g_str_Parame = "SELECT HIPCIE_UBIGEO, PARDES_DESCRI, HIPCIE_TIPMON, COUNT(HIPCIE_SALCAP), SUM(HIPCIE_SALCAP + HIPCIE_SALCON) FROM CRE_HIPCIE A, CRE_HIPCUO B, MNT_PARDES C WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = DATGEN_NUMDOC AND "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '101' AND "
   g_str_Parame = g_str_Parame & "HIPMAE_UBIGEO = PARDES_CODITE AND "
   
'   If chk_Empres.Value = 0 Then
'      g_str_Parame = g_str_Parame & "HIPMAE_PROCRE = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
'   End If
'
'   If chk_TipRie.Value = 0 Then
'      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipRie.ListIndex + 1).Genera_Codigo & "' AND "
'   End If
'
'   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
'   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
      
   g_str_Parame = g_str_Parame & "GROUP BY HIPCIE_UBIGEO, PARDES_DESCRIDATGEN_NOMBRE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "OPERACION"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "OPERACION MIVIVIENDA"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "NRO CUOTA"
      .Cells(1, 8) = "TIPO DE MONEDA"
      .Cells(1, 9) = "CAPITAL"
      .Cells(1, 10) = "INTERES"
      .Cells(1, 11) = "COM. COFIDE"
         
      .Range(.Cells(1, 1), .Cells(1, 11)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 11)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 28
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 17
      .Columns("F").ColumnWidth = 43
      .Columns("G").ColumnWidth = 11
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 12
      .Columns("J").ColumnWidth = 12
      .Columns("K").ColumnWidth = 12
                 
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
         
      If g_rst_Princi!HIPMAE_CODPRD = "003" And g_rst_Princi!HIPCUO_TIPCRO = 5 Then
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!HIPMAE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!HIPMAE_OPEMVI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
               
         If g_rst_Princi!hipmae_moneda = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "SOLES"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "DOLARES AMERICANOS"
         End If
                    
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
            
         r_int_ConVer = r_int_ConVer + 1
            
      ElseIf g_rst_Princi!HIPMAE_CODPRD = "004" And g_rst_Princi!HIPCUO_TIPCRO = 3 Then
               
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!HIPMAE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!HIPMAE_OPEMVI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
               
         If g_rst_Princi!hipmae_moneda = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "SOLES"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "DOLARES AMERICANOS"
         End If
                    
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         
         r_int_ConVer = r_int_ConVer + 1
               
      ElseIf g_rst_Princi!HIPMAE_CODPRD = "007" And g_rst_Princi!HIPCUO_TIPCRO = 3 Then
                     
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!HIPMAE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!HIPMAE_OPEMVI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
               
         If g_rst_Princi!hipmae_moneda = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "SOLES"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "DOLARES AMERICANOS"
         End If
                    
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         
         r_int_ConVer = r_int_ConVer + 1
                     
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_34()

   Dim r_obj_Excel      As Excel.Application
   
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_ConVer     As Integer
   Dim r_str_NumCta     As String
         
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   Screen.MousePointer = 11
   
   'Creando Archivo
   g_str_Parame = "SELECT CNTA_CTBL, SUM(IMP_MOVSOL) AS MONTO FROM CNTBL_ASIENTO_DET "
   g_str_Parame = g_str_Parame & "WHERE ANO = " & r_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & r_str_PerMes & " "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      .DrawingObjects(1).Left = 20
      .DrawingObjects(1).Top = 20
      
      .Range(.Cells(1, 5), .Cells(2, 5)).HorizontalAlignment = xlHAlignRight
      .Cells(1, 5) = "Dpto. de Tecnología e Informática"
      .Cells(2, 5) = "Desarrollo de Sistemas"
       
      .Range(.Cells(5, 1), .Cells(5, 5)).Merge
      .Range(.Cells(5, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(5, 1) = "Hipotecas Constituidas por Fecha de Registro"
   
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "INDICADOR BASICO"
      .Cells(7, 3) = "CUENTA CONTABLE"
      .Cells(7, 4) = "DESCRIPCION"
      .Cells(7, 5) = "IMPORTE"
       
      .Range(.Cells(7, 1), .Cells(7, 5)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
     
      .Columns("B").ColumnWidth = 25
      '.Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").NumberFormat = "@"
      
      .Columns("D").ColumnWidth = 50
      '.Columns("D").HorizontalAlignment = xlHAlignCenter
            
      .Columns("E").ColumnWidth = 15
      '.Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("E").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(7, 1), .Cells(7, 5)).VerticalAlignment = xlHAlignFill
            
   End With
      
   g_rst_Princi.MoveFirst
      
   r_int_ConVer = 8
   
   Do While Not g_rst_Princi.EOF
      
      If Left(g_rst_Princi!CNTA_CTBL, 2) = "41" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "42" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "49" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "51" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "52" Then
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 7
         
         r_str_NumCta = ff_DesCue(Trim(g_rst_Princi!CNTA_CTBL))
         
         If Left(g_rst_Princi!CNTA_CTBL, 2) = "51" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "INGRESOS FINANCIEROS"
         ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "52" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "INGRESOS POR SERVICIOS"
         ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "41" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "GASTOS FINANCIEROS"
         ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "42" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "GASTOS POR SERVICIOS"
         ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "49" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "GASTOS POR SERVICIOS"
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!CNTA_CTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(r_str_NumCta)
         
         If Left(g_rst_Princi!CNTA_CTBL, 2) = "41" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "42" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "49" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!MONTO * -1
         ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "51" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "52" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!MONTO
         End If
         
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.Range(r_obj_Excel.Cells(1, 1), r_obj_Excel.Cells(r_int_ConVer, 99)).Font.Name = "Arial"
   r_obj_Excel.Range(r_obj_Excel.Cells(1, 1), r_obj_Excel.Cells(r_int_ConVer, 99)).Font.Size = 10
   
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(8, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

End Sub

Private Function ff_DesCue(ByVal p_NumCta As String) As String
                     
   g_str_Parame = "SELECT * FROM CNTBL_CNTA WHERE "
   g_str_Parame = g_str_Parame & "CNTA_CTBL = '" & p_NumCta & "' "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
                                   
         ff_DesCue = Trim(g_rst_Listas!DESC_CNTA)
         
         g_rst_Listas.MoveNext
                  
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function
 
