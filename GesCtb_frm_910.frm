VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_SdoCie_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "GesCtb_frm_910.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2745
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   4842
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
         TabIndex        =   6
         Top             =   60
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
            Height          =   315
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Diferencias de Saldos en Cierre"
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
            Left            =   60
            Picture         =   "GesCtb_frm_910.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
            Left            =   5160
            Picture         =   "GesCtb_frm_910.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_910.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1215
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   2143
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3975
         End
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   810
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
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   840
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_SdoCie_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Period)
   End If
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

Private Sub cmd_Proces_Click()
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   If ipp_PerAno.Text < 2009 Then
      MsgBox "Debe registrar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el reporte ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   
   Call fs_Reporte_Diferencias(Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00"), CInt(Format(ipp_PerAno.Text, "0000")))
   
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_SetFocus(cmb_Empres)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   ipp_PerAno.Text = Year(date)
   cmb_Empres.ListIndex = 0
End Sub

Private Sub fs_Reporte_Diferencias(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_MesAnt     As Integer
Dim r_int_AnoAnt     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_dbl_SdoRes     As Double
Dim r_dbl_Difere     As Double

   r_str_FecIni = CStr(p_PerAno) + Format(p_PerMes, "00") + "01"
   r_str_FecFin = CStr(p_PerAno) + Format(p_PerMes, "00") + CStr(ff_Ultimo_Dia_Mes(p_PerMes, p_PerAno))
   
   If p_PerMes = 1 Then
      r_int_MesAnt = 12
      r_int_AnoAnt = p_PerAno - 1
   Else
      r_int_MesAnt = p_PerMes - 1
      r_int_AnoAnt = p_PerAno
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CREDITO AS OPERACION, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT B.HIPMAE_SITUAC "
   g_str_Parame = g_str_Parame & "              FROM CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & "             WHERE B.HIPMAE_NUMOPE = A.CREDITO), 0) AS ESTADO, "
   g_str_Parame = g_str_Parame & "       A.CAPITAL_DESEMBOLSADO + A.CAPITAL_INTERES - A.CAPITAL_AMORTIZADO AS SALDO_NUEVO, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT SUM(C.HIPPAG_CAPITA + C.HIPPAG_CAPBBP) "
   g_str_Parame = g_str_Parame & "              FROM CRE_HIPPAG C "
   g_str_Parame = g_str_Parame & "             WHERE C.HIPPAG_NUMOPE = A.CREDITO "
   g_str_Parame = g_str_Parame & "               AND C.HIPPAG_FECPAG >= " & r_str_FecIni & " AND C.HIPPAG_FECPAG <= " & r_str_FecFin & "),0) AS PAGOS_MES, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT SUM(E.PPGCAB_MTOAPL + E.PPGCAB_PBPPER) "
   g_str_Parame = g_str_Parame & "              FROM CRE_PPGCAB E "
   g_str_Parame = g_str_Parame & "             WHERE E.PPGCAB_NUMOPE = A.CREDITO "
   g_str_Parame = g_str_Parame & "               AND E.PPGCAB_FECPPG >= " & r_str_FecIni & " AND E.PPGCAB_FECPPG <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "               AND E.PPGCAB_FECPRO > 0 AND E.PPGCAB_TIPPPG = 1), 0) AS PREPAGOS_MES, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT D2.DETPBP_CAPCLI "
   g_str_Parame = g_str_Parame & "              FROM CRE_DETPBP D2 "
   g_str_Parame = g_str_Parame & "             WHERE D2.DETPBP_NUMOPE = A.CREDITO "
   g_str_Parame = g_str_Parame & "               AND D2.DETPBP_PERMES = A.MES    AND D2.DETPBP_PERANO = A.ANO), 0) AS PAGO_PBP, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT D2.DETPBP_FLGPBP "
   g_str_Parame = g_str_Parame & "              FROM CRE_DETPBP D2 "
   g_str_Parame = g_str_Parame & "             WHERE D2.DETPBP_NUMOPE = A.CREDITO "
   g_str_Parame = g_str_Parame & "               AND D2.DETPBP_PERMES = A.MES    AND D2.DETPBP_PERANO = A.ANO), 0) AS FLAG_PBP_MES, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT F.CAPITAL_DESEMBOLSADO + F.CAPITAL_INTERES - F.CAPITAL_AMORTIZADO "
   g_str_Parame = g_str_Parame & "              FROM CREDITO_CIERRE_FINMES F "
   g_str_Parame = g_str_Parame & "             WHERE F.ANO = " & r_int_AnoAnt & " AND F.MES = " & r_int_MesAnt & " AND F.CREDITO = A.CREDITO),0) AS SALDO_ANTERIOR "
   g_str_Parame = g_str_Parame & "  FROM CREDITO_CIERRE_FINMES A "
   g_str_Parame = g_str_Parame & " WHERE A.ANO = " & p_PerAno
   g_str_Parame = g_str_Parame & "   AND A.MES = " & p_PerMes
   g_str_Parame = g_str_Parame & "   AND A.PRODUCTO <> '008' "
   g_str_Parame = g_str_Parame & "ORDER BY A.CREDITO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 5) = "REPORTE DE DIFERENCIA DE SALDOS AL MES DE " & Trim(cmb_Period.Text) & " DEL " & CStr(ipp_PerAno.Text)
      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "OPERACION"
      .Cells(3, 3) = "PADRON MES ANTERIOR"
      .Cells(3, 4) = "PAGOS DEL MES"
      .Cells(3, 5) = "PREPAGOS DEL MES"
      .Cells(3, 6) = "PBP DEL MES"
      .Cells(3, 7) = "SALDO RESULTANTE"
      .Cells(3, 8) = "PADRON MES ACTUAL"
      .Cells(3, 9) = "DIFERENCIAS"
      
      .Range(.Cells(1, 1), .Cells(3, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 18
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 16
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 18
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 14
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 18
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 20
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      .Columns("I").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 4
   
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!FLAG_PBP_MES = 2 Then
         r_dbl_SdoRes = g_rst_Princi!SALDO_ANTERIOR - g_rst_Princi!PAGOS_MES - g_rst_Princi!PREPAGOS_MES
      Else
         r_dbl_SdoRes = g_rst_Princi!SALDO_ANTERIOR - g_rst_Princi!PAGOS_MES - g_rst_Princi!PREPAGOS_MES - g_rst_Princi!PAGO_PBP
      End If
      r_dbl_Difere = g_rst_Princi!SALDO_NUEVO - r_dbl_SdoRes
      
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 3
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Format(g_rst_Princi!SALDO_ANTERIOR, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Format(g_rst_Princi!PAGOS_MES, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Format(g_rst_Princi!PREPAGOS_MES, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Format(g_rst_Princi!PAGO_PBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Format(r_dbl_SdoRes, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!SALDO_NUEVO, "#0.00")
      If g_rst_Princi!SALDO_ANTERIOR = 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(0, "#0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(r_dbl_Difere, "#0.00")
      End If
                              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
