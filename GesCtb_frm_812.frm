VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   6585
   ClientTop       =   3915
   ClientWidth     =   7170
   Icon            =   "GesCtb_frm_812.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3705
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   6535
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
         TabIndex        =   10
         Top             =   60
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
            Height          =   270
            Left            =   600
            TabIndex        =   11
            Top             =   120
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Padron de Deudores"
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
            Picture         =   "GesCtb_frm_812.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   780
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
            Picture         =   "GesCtb_frm_812.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_812.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "GesCtb_frm_812.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2145
         Left            =   30
         TabIndex        =   13
         Top             =   1470
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   3784
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
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   1050
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
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todas las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin VB.ComboBox cmb_Permes 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1350
            Width           =   5895
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
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   720
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
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   1410
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   90
            TabIndex        =   14
            Top             =   1800
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_12"
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
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Perido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc_PadDeu2
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
   cmd_Imprim.Enabled = False
   
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

Private Sub fs_GenExc_PadDeu2()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_NumSol     As String
Dim r_int_TipMon     As Integer
Dim r_str_FecCam     As String
Dim r_str_TipCam     As String
Dim r_str_CheCgo     As String
Dim r_str_CtaCgo     As String
Dim r_str_BanCgo     As String
Dim r_int_MesCie     As Integer
Dim r_int_AnoCie     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_FecPpg     As String
Dim r_str_NomCli     As String
Dim r_str_FecCie     As String
Dim r_str_PerMes     As String
Dim r_str_PerAno     As String
Dim r_int_TipGar     As Integer

   'Inicializa Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "NRO. OPERACION"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "F. DESEMBOLSO"
      .Cells(1, 7) = "TIPO DE MONEDA"
      .Cells(1, 8) = "DIAS DE ATRASO"
      .Cells(1, 9) = "CUOTAS PAGADAS (MESES)"
      .Cells(1, 10) = "PLAZO (AÑOS)"
      .Cells(1, 11) = "SITUACION DEL CREDITO"
      .Cells(1, 12) = "TIPO DE GARANTIA"
      .Cells(1, 13) = "GARANTIA  S/."
      .Cells(1, 14) = "GARANTIA US$"
      
      .Cells(1, 15) = "VALOR ACTUALIZADO"
      
      .Cells(1, 16) = "MENOR S/."
      .Cells(1, 17) = "MENOR US$"
      
      .Cells(1, 18) = "CLASIF. INTERNA"
      .Cells(1, 19) = "CLASIF. ALINEADA"
      .Cells(1, 20) = "CLASIF. PROVIS."
      .Cells(1, 21) = "TASA PROVIS."
      .Cells(1, 22) = "PROV. GENERICA"
      .Cells(1, 23) = "PROV. GENERICA RC"
      .Cells(1, 24) = "PROV. ESPECIFICA"
      .Cells(1, 25) = "PROV. RIESGO. CAMB."
      .Cells(1, 26) = "PROV. PROCICLICA"
      .Cells(1, 27) = "PROV. PROCICLICA RC"
      .Cells(1, 28) = "PROV. VOLUNTARIA"
      .Cells(1, 29) = "PROV. RIESGO PAIS"
      .Cells(1, 30) = "APLICACION PROCIC."
      .Cells(1, 31) = "DVG. VIGENTE"
      .Cells(1, 32) = "DVG. VIGENTE SOLES"
      .Cells(1, 33) = "DVG. VENCIDO"
      .Cells(1, 34) = "DVG. VENCIDO SOLES"
      .Cells(1, 35) = "DVG. REFINANCIADO"
      .Cells(1, 36) = "DVG. REFINANCIADO SOLES"
      .Cells(1, 37) = "INTERES DIFERIDO"
      .Cells(1, 38) = "CAPITAL JUDICIAL"
      .Cells(1, 39) = "CAPITAL REFINANCIADO"
      .Cells(1, 40) = "CAP. VENC. REFINAN."
      .Cells(1, 41) = "CAPITAL VENCIDO"
      .Cells(1, 42) = "CAPITAL VIGENTE"
      .Cells(1, 43) = "DVG. PBP"
      .Cells(1, 44) = "PROV. INT. COFIDE"
      .Cells(1, 45) = "PROV. COM. COFIDE"
      .Cells(1, 46) = "TASA DE INTERES ACTIVA"
      .Cells(1, 47) = "TASA DE INTERES PASIVA"
      .Cells(1, 48) = "SALDO CAPITAL"
      .Cells(1, 49) = "SALDO CAPITAL S/."
      .Cells(1, 50) = "INT. x DIAS"
      .Cells(1, 51) = "COBERTURA FMV"
      .Cells(1, 52) = "COBERTURA FMV RC"
      .Cells(1, 53) = "COBERTURA FMV AUTOLIQUIDABLE"
      .Cells(1, 54) = "INT. CUO. CIERRE"
      .Cells(1, 55) = "FEC. VCT. CIERRE"
      .Cells(1, 56) = "SAL. CAP. CIERRE"
      .Cells(1, 57) = "PRIMERA VIVIENDA"
      .Cells(1, 58) = "ADQUISICION DE PRIMERA VIVIENDA CON HIPOTECA"
      .Cells(1, 59) = "OTROS CON HIPOTECA INSCRITA"
      .Cells(1, 60) = "ADQUISICION DE PRIMERA VIVIENDA SIN HIPOTECA"
      .Cells(1, 61) = "OTROS SIN HIPOTECA INSCRITA"
      .Cells(1, 62) = "OTRAS GARANTIAS NO PREFERIDAS - BLOQUEO > 90 DIAS"
      .Cells(1, 63) = "PROVIS. REQUERIDA"
      .Cells(1, 64) = "PROVIS. CONSTITUIDA"
      .Cells(1, 65) = "FINANCIAMIENTO"
      .Cells(1, 66) = "PASIVO TNC"
      .Cells(1, 67) = "PASIVO TC"
      .Cells(1, 68) = "HIPOTECA MATRIZ"
      .Range(.Cells(1, 1), .Cells(1, 68)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 68)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 50
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 14
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 45
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 18
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 20
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 25
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12
      .Columns("N").ColumnWidth = 12
      .Columns("O").ColumnWidth = 18
      .Columns("P").ColumnWidth = 15
      .Columns("Q").ColumnWidth = 15
      .Columns("R").ColumnWidth = 14
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 15
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").ColumnWidth = 13
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 12
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 14
      .Columns("W").ColumnWidth = 17
      .Columns("X").ColumnWidth = 15
      .Columns("Y").ColumnWidth = 18
      .Columns("Z").ColumnWidth = 16
      .Columns("AA").ColumnWidth = 18
      .Columns("AB").ColumnWidth = 16
      .Columns("AC").ColumnWidth = 17
      .Columns("AD").ColumnWidth = 17
      .Columns("AE").ColumnWidth = 12
      .Columns("AF").ColumnWidth = 17
      .Columns("AG").ColumnWidth = 12
      .Columns("AH").ColumnWidth = 18
      .Columns("AI").ColumnWidth = 18
      .Columns("AJ").ColumnWidth = 21
      .Columns("AK").ColumnWidth = 16
      .Columns("AL").ColumnWidth = 16
      .Columns("AM").ColumnWidth = 19
      .Columns("AN").ColumnWidth = 17
      .Columns("AO").ColumnWidth = 14
      .Columns("AP").ColumnWidth = 15
      .Columns("AQ").ColumnWidth = 16
      .Columns("AR").ColumnWidth = 14
      .Columns("AS").ColumnWidth = 16
      .Columns("AT").ColumnWidth = 20
      .Columns("AU").ColumnWidth = 20
      .Columns("AV").ColumnWidth = 16
      .Columns("AW").ColumnWidth = 16
      .Columns("AX").ColumnWidth = 11
      .Columns("AY").ColumnWidth = 15
      .Columns("AZ").ColumnWidth = 17
      .Columns("BA").ColumnWidth = 28
      .Columns("BB").ColumnWidth = 15
      .Columns("BC").ColumnWidth = 14
      .Columns("BC").HorizontalAlignment = xlHAlignCenter
      .Columns("BD").ColumnWidth = 15
      .Columns("BE").ColumnWidth = 16
      .Columns("BE").HorizontalAlignment = xlHAlignCenter
      .Columns("BF").ColumnWidth = 40
      .Columns("BG").ColumnWidth = 27
      .Columns("BH").ColumnWidth = 40
      .Columns("BI").ColumnWidth = 26
      .Columns("BJ").ColumnWidth = 43
      .Columns("BK").ColumnWidth = 18
      .Columns("BL").ColumnWidth = 18
      .Columns("BM").ColumnWidth = 18
      .Columns("BM").HorizontalAlignment = xlHAlignCenter
      .Columns("BN").ColumnWidth = 15
      .Columns("BO").ColumnWidth = 15
      .Columns("BP").ColumnWidth = 15
      .Columns("BP").HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 68)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(1, 68)).Font.Size = 8
   End With
   
   '*** INICIALIZA VARIABLES
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = ipp_PerAno.Text
   r_int_MesCie = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_AnoCie = CInt(ipp_PerAno.Text)
   r_str_FecIni = Format(r_int_AnoCie, "0000") & Format(r_int_MesCie, "00") & "01"
   r_str_FecFin = Format(r_int_AnoCie, "0000") & Format(r_int_MesCie, "00") & "31"
   r_str_NomCli = ""
   
   '*** INFORMACION DE CREDITOS HIPOTECARIOS
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_NUMOPE, TRIM(B.PRODUC_DESCRI) AS PRODUC_DESCRI, A.HIPCIE_TDOCLI, A.HIPCIE_NDOCLI, A.HIPCIE_CODPRD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMCLIENTE, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_TIPMON, A.HIPCIE_FECDES, TRIM(D.PARDES_DESCRI) AS MONEDA_DESEM, A.HIPCIE_DIAMOR, G.HIPMAE_GARLIN, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CUOPAG, A.HIPCIE_SITCRE, A.HIPCIE_TIPGAR, A.HIPCIE_CLACLI, A.HIPCIE_CLAALI, A.HIPCIE_PRVVOL, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CLAPRV, A.HIPCIE_PRVGEN, A.HIPCIE_PRVGEN_RC, A.HIPCIE_PRVESP, A.HIPCIE_PRVCAM, A.HIPCIE_FLGJUD, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_PRVCIC, A.HIPCIE_PRVCIC_RC, A.HIPCIE_DEVVIG, A.HIPCIE_TIPCAM, A.HIPCIE_DEVVEN, A.HIPCIE_TIPCAM, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_INTDIF, A.HIPCIE_CAPVEN, A.HIPCIE_CAPVIG, A.HIPCIE_ACUDVG, A.HIPCIE_ACUDVC, A.HIPCIE_PRENCO, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_DEVPBP, A.HIPCIE_PRVICO, A.HIPCIE_PRVCCO, A.HIPCIE_TASINT, A.HIPCIE_SALCAP, A.HIPCIE_SALCON, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CBRFMV, A.HIPCIE_CBRFMV_RC, A.HIPCIE_PRVRIP, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN A.HIPCIE_TIPGAR = 6 THEN J.HIPDES_TIPMON ELSE A.HIPCIE_MONGAR END) AS HIPCIE_MONGAR, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN A.HIPCIE_TIPGAR = 6 THEN J.HIPDES_DESMPR ELSE A.HIPCIE_MTOGAR END) AS HIPCIE_MTOGAR, TRIM(E.PARDES_DESCRI) AS GARANTIA, "
   g_str_Parame = g_str_Parame & "       F.HIPCUO_INTERE, F.HIPCUO_FECVCT, F.HIPCUO_SALCAP, A.HIPCIE_FLGJUD, A.HIPCIE_APLCIC, TRIM(H.PARDES_DESCRI) AS PRI_VIVIENDA, "
   g_str_Parame = g_str_Parame & "       EVALEG_FECBLQ_INM, HIPMAE_PRIVIV, HIPCIE_FLGREF, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN A.HIPCIE_CODPRD IN ('001','003') THEN HIPCIE_TASMVI ELSE HIPCIE_TASCOF END) HIPCIE_TASMVICOF,"
   g_str_Parame = g_str_Parame & "       (SELECT ROUND(DECODE(MITAB.MONEDA, 1, MITAB.PROV_REQUERIDA, MITAB.PROV_REQUERIDA * MITAB.TIPO_CAMBIO), 2) AS PROVISION_REQUERIDA "
   g_str_Parame = g_str_Parame & "          FROM (SELECT CASE WHEN HP.HIPCIE_CLAPRV = 0 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.007 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.007 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 1 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.025 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.050 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 2 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.125 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.25 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 3 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.3 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.6 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 4 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.6 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 1 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                       END AS PROV_REQUERIDA,"
   g_str_Parame = g_str_Parame & "                       HP.HIPCIE_TIPMON AS MONEDA, HP.HIPCIE_TIPCAM AS TIPO_CAMBIO,HP.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "                  FROM CRE_HIPCIE HP "
   g_str_Parame = g_str_Parame & "                 WHERE HP.HIPCIE_PERMES = " & CInt(r_str_PerMes)
   g_str_Parame = g_str_Parame & "                   AND HP.HIPCIE_PERANO = " & CInt(r_str_PerAno) & ") MITAB WHERE MITAB.HIPCIE_NUMOPE = A.HIPCIE_NUMOPE) AS PROVISION_REQUERIDA, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPCIE_TIPMON, 1, A.HIPCIE_PRVGEN+A.HIPCIE_PRVESP+A.HIPCIE_PRVCIC+A.HIPCIE_PRVGEN_RC+A.HIPCIE_PRVCIC_RC+A.HIPCIE_PRVVOL, (A.HIPCIE_PRVGEN+A.HIPCIE_PRVESP+A.HIPCIE_PRVCIC+A.HIPCIE_PRVGEN_RC+A.HIPCIE_PRVCIC_RC+A.HIPCIE_PRVVOL)*A.HIPCIE_TIPCAM) AS PROVISION_CONSTITUIDA, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT TIPPRV_PORCEN "
   g_str_Parame = g_str_Parame & "              FROM CTB_TIPPRV "
   g_str_Parame = g_str_Parame & "             WHERE TIPPRV_TIPPRV = '2' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLACRE = '13' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLAGAR = 2),0) AS TASA_CG , "
   g_str_Parame = g_str_Parame & "       NVL((SELECT TIPPRV_PORCEN "
   g_str_Parame = g_str_Parame & "              FROM CTB_TIPPRV "
   g_str_Parame = g_str_Parame & "             WHERE TIPPRV_TIPPRV = '2' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLACRE = '13' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLAGAR = 1),0) AS TASA_SG, "
   g_str_Parame = g_str_Parame & "       K.SOLMAE_PLAANO AS PLAZO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 3 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPNCO "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 3 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_TNC, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 4 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= TO_CHAR(ADD_MONTHS(TO_DATE(" & r_str_FecIni & ", 'YYYY/MM/DD'), -5), 'YYYYMMDD') AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPCON "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 4 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= TO_CHAR(ADD_MONTHS(TO_DATE(" & r_str_FecIni & ", 'YYYY/MM/DD'), -5), 'YYYYMMDD') AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_TC, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 5 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPNCO "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 5 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_CME, "
   g_str_Parame = g_str_Parame & "       CASE WHEN G.HIPMAE_HIPMTZ = 3 THEN 'NO LEVANTADA' "
   g_str_Parame = g_str_Parame & "            WHEN G.HIPMAE_HIPMTZ = 2 THEN 'LEVANTADA' "
   g_str_Parame = g_str_Parame & "            ELSE 'NO APLICA' "
   g_str_Parame = g_str_Parame & "       END AS HIPOTECA_MATRIZ, "
   g_str_Parame = g_str_Parame & "       M.ACTGAR_VALACZ AS VALOR_ACTUALIZADO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A  "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.HIPCIE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = A.HIPCIE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(A.HIPCIE_NDOCLI) "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPCIE_TIPMON "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.HIPCIE_TIPGAR "
   g_str_Parame = g_str_Parame & "         LEFT JOIN CRE_HIPCUO F ON F.HIPCUO_NUMOPE = A.HIPCIE_NUMOPE AND F.HIPCUO_TIPCRO = 1 AND F.HIPCUO_FECVCT >= " & r_str_FecIni & " AND F.HIPCUO_FECVCT <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE G ON G.HIPMAE_NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = '214' AND H.PARDES_CODITE = G.HIPMAE_PRIVIV"
   g_str_Parame = g_str_Parame & "         LEFT OUTER JOIN TRA_EVALEG I ON I.EVALEG_NUMSOL = G.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "         LEFT JOIN CRE_HIPDES J ON J.HIPDES_NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLMAE K ON K.SOLMAE_NUMERO = G.HIPMAE_NUMSOL "
   
   g_str_Parame = g_str_Parame & "         LEFT JOIN (  SELECT N.ACTGAR_NUMOPE AS NUMOPE, N.ACTGAR_ANOPRO AS ANIO, MAX(N.ACTGAR_MESPRO) AS MES "
   g_str_Parame = g_str_Parame & "                        FROM CRE_ACTGAR N "
   g_str_Parame = g_str_Parame & "                             INNER JOIN CRE_HIPCIE ON HIPCIE_NUMOPE = ACTGAR_NUMOPE AND HIPCIE_PERMES = " & CStr(r_int_MesCie) & " AND HIPCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                             INNER JOIN (SELECT NUMOPE, MAX(ACTGAR_ANOPRO) AS ANIO "
   g_str_Parame = g_str_Parame & "                                           FROM (SELECT DISTINCT ACTGAR_NUMOPE AS NUMOPE, ACTGAR_ANOPRO, ACTGAR_MESPRO "
   g_str_Parame = g_str_Parame & "                                                   FROM CRE_ACTGAR "
   g_str_Parame = g_str_Parame & "                                                        INNER JOIN CRE_HIPCIE ON HIPCIE_NUMOPE = ACTGAR_NUMOPE AND HIPCIE_PERMES = " & CStr(r_int_MesCie) & "  AND HIPCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                                                  WHERE ACTGAR_VALACZ > 0 "
   g_str_Parame = g_str_Parame & "                                                  ORDER BY ACTGAR_NUMOPE ASC, ACTGAR_ANOPRO DESC, ACTGAR_MESPRO DESC) "
   g_str_Parame = g_str_Parame & "                                          GROUP BY NUMOPE ) O ON O.NUMOPE = N.ACTGAR_NUMOPE AND O.ANIO = N.ACTGAR_ANOPRO "
   g_str_Parame = g_str_Parame & "                       Where N.ACTGAR_VALACZ > 0 "
   g_str_Parame = g_str_Parame & "                       GROUP BY ACTGAR_NUMOPE, ACTGAR_ANOPRO "
   g_str_Parame = g_str_Parame & "                     ) L ON L.NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "         LEFT JOIN CRE_ACTGAR M ON M.ACTGAR_NUMOPE = L.NUMOPE AND M.ACTGAR_ANOPRO = L.ANIO AND M.ACTGAR_MESPRO = L.MES "
      
   g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERMES = " & CStr(r_int_MesCie) & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   If chk_TipPro.Value = 0 Then
     g_str_Parame = g_str_Parame & " AND A.HIPCIE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
   End If
   g_str_Parame = g_str_Parame & "ORDER BY A.HIPCIE_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_ConVer = 2
      r_str_FecCie = Format(CInt(ipp_PerAno.Text), "####") & Format(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), "0#") & Format(ff_Ultimo_Dia_Mes(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(ipp_PerAno.Text)), "00")
      
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!HIPCIE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPCIE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPCIE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMCLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_FECDES))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!MONEDA_DESEM)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!HIPCIE_DIAMOR, "00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCIE_CUOPAG, "00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = g_rst_Princi!PLAZO_PRESTAMO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!HIPCIE_SITCRE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!GARANTIA)
         
         If g_rst_Princi!HIPCIE_MONGAR = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(0, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!VALOR_ACTUALIZADO, "###,###,##0.00")
         
         If g_rst_Princi!HIPCIE_MONGAR = 1 Then
            If r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) < r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13), "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15), "###,###,##0.00")
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
         Else
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
            If r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) < r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14), "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15), "###,###,##0.00")
            End If
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Trim(g_rst_Princi!HIPCIE_CLACLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Trim(g_rst_Princi!HIPCIE_CLAALI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(g_rst_Princi!HIPCIE_CLAPRV)
         
         '-----------------------------------EXCEPCIONES SOBRE LA TASA---------------------------------------------
         r_int_TipGar = 0
         '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            g_str_CadCnx = ""
            g_str_CadCnx = g_str_CadCnx & "SELECT COUNT(*) AS CONTADOR "
            g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_CLACLI "
            g_str_CadCnx = g_str_CadCnx & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
            g_str_CadCnx = g_str_CadCnx & "                  FROM CRE_HIPCIE "
            g_str_CadCnx = g_str_CadCnx & "                 WHERE HIPCIE_PERMES > 0 "
            g_str_CadCnx = g_str_CadCnx & "                   AND HIPCIE_PERANO > 2014 "
            g_str_CadCnx = g_str_CadCnx & "                   AND HIPCIE_NUMOPE = '" & g_rst_Princi!HIPCIE_NUMOPE & "' "
            g_str_CadCnx = g_str_CadCnx & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            g_str_CadCnx = g_str_CadCnx & "         WHERE ROWNUM < 37) "
            g_str_CadCnx = g_str_CadCnx & " WHERE HIPCIE_CLACLI = 3 "
            
            If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
               Exit Sub
            End If
            
            If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               g_rst_GenAux.MoveFirst
               If g_rst_GenAux!CONTADOR = 36 Then
                  r_int_TipGar = 5
               End If
            End If
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
         End If
         
         '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            g_str_CadCnx = ""
            g_str_CadCnx = g_str_CadCnx & "SELECT COUNT(*) AS CONTADOR "
            g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_CLACLI "
            g_str_CadCnx = g_str_CadCnx & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
            g_str_CadCnx = g_str_CadCnx & "                  FROM CRE_HIPCIE "
            g_str_CadCnx = g_str_CadCnx & "                 WHERE HIPCIE_PERMES > 0 "
            g_str_CadCnx = g_str_CadCnx & "                   AND HIPCIE_PERANO > 2015 "
            g_str_CadCnx = g_str_CadCnx & "                   AND HIPCIE_NUMOPE = '" & g_rst_Princi!HIPCIE_NUMOPE & "' "
            g_str_CadCnx = g_str_CadCnx & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            g_str_CadCnx = g_str_CadCnx & "         WHERE ROWNUM < 25) "
            g_str_CadCnx = g_str_CadCnx & " WHERE HIPCIE_CLACLI = 4 "
            
            If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
               Exit Sub
            End If
            
            If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               g_rst_GenAux.MoveFirst
               If g_rst_GenAux!CONTADOR = 24 Then
                  r_int_TipGar = 5
               End If
            End If
            
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
         End If
         
         If (r_int_TipGar = 5) Then
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDbl(Format(g_rst_Princi!TASA_SG, "###,###,##0.00"))
         Else
             If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 4) Then
                 r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDbl(Format(g_rst_Princi!TASA_CG, "###,###,##0.00"))
             Else
                 r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDbl(Format(g_rst_Princi!TASA_SG, "###,###,##0.00"))
             End If
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!HIPCIE_PRVGEN, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!HIPCIE_PRVGEN_RC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPCIE_PRVESP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPCIE_PRVCAM, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!HIPCIE_PRVCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Format(g_rst_Princi!HIPCIE_PRVCIC_RC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Format(g_rst_Princi!HIPCIE_PRVVOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!HIPCIE_PRVRIP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Format(g_rst_Princi!HIPCIE_APLCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
         
         If g_rst_Princi!HIPCIE_FLGREF = 0 Then
            If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPCIE_ACUDVG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
            If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!HIPCIE_ACUDVC * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = 0
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = 0
            
            If g_rst_Princi!HIPCIE_ACUDVC > 0 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!HIPCIE_ACUDVC * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
               End If
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!HIPCIE_ACUDVG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
               End If
            End If
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(g_rst_Princi!HIPCIE_INTDIF, "###,###,##0.00")
         If g_rst_Princi!HIPCIE_FLGJUD = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(0, "###,###,##0.00")
         Else
            If g_rst_Princi!HIPCIE_FLGREF = 1 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(g_rst_Princi!HIPCIE_CAPVIG, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(g_rst_Princi!HIPCIE_CAPVEN, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(g_rst_Princi!HIPCIE_CAPVEN, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(g_rst_Princi!HIPCIE_CAPVIG, "###,###,##0.00")
            End If
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Format(g_rst_Princi!HIPCIE_DEVPBP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = Format(g_rst_Princi!HIPCIE_PRVICO, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format(g_rst_Princi!HIPCIE_PRVCCO, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Format(g_rst_Princi!HIPCIE_TASINT, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 47) = Format(g_rst_Princi!HIPCIE_TASMVICOF, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 48) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
         
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON), "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 50) = Format(((((g_rst_Princi!HIPCIE_TASINT / 100 + 1) ^ (1 / 360)) - 1) * (g_rst_Princi!HIPCIE_SALCAP)), "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 51) = Format(g_rst_Princi!HIPCIE_CBRFMV, "###,###,##0.00")
         
         'If g_rst_Princi!HIPCIE_FLGJUD = 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 52) = Format(g_rst_Princi!HIPCIE_CBRFMV_RC, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 53) = Format(0, "###,###,##0.00")
         'Else
         '   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(0, "###,###,##0.00")
         '   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 50) = Format(g_rst_Princi!HIPCIE_CBRFMV_RC, "###,###,##0.00")
         'End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 54) = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         
         If Not IsNull(g_rst_Princi!HIPCUO_FECVCT) Then
            r_str_FecPpg = fs_BuscaPrepago(g_rst_Princi!HIPCIE_NUMOPE, r_str_FecIni, r_str_FecFin)
            If r_str_FecPpg = "0" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = "'" & gf_FormatoFecha(r_str_FecPpg)
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_FECDES))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Format(g_rst_Princi!HIPCIE_PRENCO, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = g_rst_Princi!PRI_VIVIENDA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 59) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = Format(0, "###,###,##0.00")
         
         If g_rst_Princi!HIPMAE_PRIVIV = 1 Then
            If g_rst_Princi!HIPCIE_FECDES > 20130101 Then
               If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               Else
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               End If
            End If
         Else
            If g_rst_Princi!HIPCIE_FECDES > 20130101 Then
               If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 59) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 59) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               Else
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               End If
            End If
         End If
         
         If g_rst_Princi!HIPCIE_TIPGAR = 2 Then
            If DateDiff("d", Format(g_rst_Princi!EVALEG_FECBLQ_INM, "####/##/##"), Format(r_str_FecCie, "####/##/##")) > 90 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(0, "###,###,##0.00")
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(0, "###,###,##0.00")
         End If

         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = IIf(IsNull(g_rst_Princi!PROVISION_REQUERIDA), "0.00", Format(g_rst_Princi!PROVISION_REQUERIDA, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 64) = IIf(IsNull(g_rst_Princi!PROVISION_CONSTITUIDA), "0.00", Format(g_rst_Princi!PROVISION_CONSTITUIDA, "###,###,##0.00"))
         
         '****************
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 64) = ""
         Select Case g_rst_Princi!HIPMAE_GARLIN
            Case "000001": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "MIVIVIENDA"
            Case "000002":
                   If CLng(r_str_PerAno & Format(r_str_PerMes, "00")) < CLng("201610") Then
                      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "BID"
                   Else
                      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "NINGUNO"
                   End If
            Case "000003": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "F RAMIREZ PRADO"
            Case "000004": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "COFICASA"
            Case "999999": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 65) = "NINGUNO"
         End Select
         '****************
         
         If g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "002" Or g_rst_Princi!HIPCIE_CODPRD = "006" Or g_rst_Princi!HIPCIE_CODPRD = "011" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 66) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 67) = Format(0, "###,###,##0.00")
         Else
            If g_rst_Princi!HIPCIE_CODPRD = "003" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 66) = Format(g_rst_Princi!PASIVO_CME, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 67) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 66) = Format(g_rst_Princi!PASIVO_TNC, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 67) = Format(g_rst_Princi!PASIVO_TC, "###,###,##0.00")
            End If
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 68) = Trim(g_rst_Princi!HIPOTECA_MATRIZ)
         
         r_int_ConVer = r_int_ConVer + 1
         r_str_NomCli = Trim(g_rst_Princi!NOMCLIENTE) '(*)guarda en variable nombre cliente
         g_rst_Princi.MoveNext
                  
         If g_rst_Princi.EOF Then
            '(*)si es fin de archivo sale del bucle
            Exit Do
         Else
            '(*)compara variable nombre cliente con el proximo cliente,
            'y adelanta un registro si comparacion entre clientes es igual.
            If r_str_NomCli = Trim(g_rst_Princi!NOMCLIENTE) Then
               g_rst_Princi.MoveNext
            End If
         End If
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '*** INFORMACION DE CREDITOS INMMOBILIARIOS
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.COMCIE_NUMOPE, TRIM(B.PRODUC_DESCRI) AS PRODUC_DESCRI, A.COMCIE_TDOCLI, A.COMCIE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_RAZSOC) AS NOMEMPRESA, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_TIPMON, A.COMCIE_FECDES, TRIM(D.PARDES_DESCRI) AS MONEDA_DESEM, A.COMCIE_DIAMOR, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CUOPAG, A.COMCIE_SITCRE, A.COMCIE_TIPGAR, A.COMCIE_CLACLI, A.COMCIE_CLAALI, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CLAPRV, A.COMCIE_PRVGEN, A.COMCIE_PRVESP, A.COMCIE_PRVCAM, A.COMCIE_FLGJUD, COMCIE_FLGREF, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_PRVCIC, A.COMCIE_DEVVIG, A.COMCIE_TIPCAM, A.COMCIE_DEVVEN, A.COMCIE_INTDIF, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CAPVEN, A.COMCIE_CAPVIG, A.COMCIE_ACUDVG, A.COMCIE_ACUDVC, A.COMCIE_TASINT, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_SALCAP, A.COMCIE_MONGAR, A.COMCIE_MTOGAR, TRIM(E.PARDES_DESCRI) AS GARANTIA "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.COMCIE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = A.COMCIE_TDOCLI AND TRIM(C.DATGEN_EMPNDO) = TRIM(A.COMCIE_NDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.COMCIE_TIPMON "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.COMCIE_TIPGAR "
   g_str_Parame = g_str_Parame & " WHERE A.COMCIE_PERMES = " & CStr(r_int_MesCie) & " "
   g_str_Parame = g_str_Parame & "   AND A.COMCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "   AND A.COMCIE_SITUAC = 1 "
   'g_str_Parame = g_str_Parame & "   AND A.COMCIE_TIPGAR = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY A.COMCIE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!COMCIE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!COMCIE_TDOCLI) & "-" & Trim(g_rst_Princi!COMCIE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMEMPRESA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!COMCIE_FECDES))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!MONEDA_DESEM)
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!COMCIE_DIAMOR, "00") '"0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!COMCIE_CUOPAG, "00") '"0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "1"
                  
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!COMCIE_SITCRE) '"1"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!GARANTIA)
         
         If g_rst_Princi!COMCIE_MONGAR = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(0, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Trim(g_rst_Princi!COMCIE_CLACLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Trim(g_rst_Princi!COMCIE_CLAALI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(g_rst_Princi!COMCIE_CLAPRV)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = 0                                    'TASA PROVIS.
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!COMCIE_PRVGEN, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = "0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!COMCIE_PRVESP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!COMCIE_PRVCAM, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!COMCIE_PRVCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = "0"                                  'PROV. PROCICLICA RC
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = "0"                                  'PROV. VOLUNTARIA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = "0"                                  'RIESGO PAIS
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = "0"                                  'APLICACION PROCIC.
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!COMCIE_ACUDVG, "###,###,##0.00")
         If g_rst_Princi!COMCIE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!COMCIE_ACUDVG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!COMCIE_ACUDVG, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!COMCIE_ACUDVC, "###,###,##0.00")
         If g_rst_Princi!COMCIE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!COMCIE_ACUDVC * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!COMCIE_ACUDVC, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = "0"                                  'DVG. REFINANCIADO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = "0"                                  'DVG. REFINANCIADO SOLES
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(g_rst_Princi!COMCIE_INTDIF, "###,###,##0.00")

         If g_rst_Princi!comcie_flgjud = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(g_rst_Princi!COMCIE_SALCAP, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(0, "###,###,##0.00")
         Else
             If g_rst_Princi!COMCIE_FLGREF = 1 Then
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(g_rst_Princi!COMCIE_CAPVIG, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(g_rst_Princi!COMCIE_CAPVEN, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(0, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(0, "###,###,##0.00")
             Else
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(0, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(g_rst_Princi!COMCIE_CAPVEN, "###,###,##0.00")
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(g_rst_Princi!COMCIE_CAPVIG, "###,###,##0.00")
             End If
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Format(g_rst_Princi!COMCIE_TASINT, "###,###,##0.00")  'TASA DE INTERES ACTIVA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 47) = Format(0, "###,###,##0.00") 'TASA DE INTERES PASIVA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 48) = Format(g_rst_Princi!COMCIE_SALCAP, "###,###,##0.00") 'SALDO CAPITAL
         If g_rst_Princi!COMCIE_TIPMON = 1 Then 'SALDO CAPITAL S/.
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(g_rst_Princi!COMCIE_SALCAP, "###,###,##0.00")
            Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(g_rst_Princi!COMCIE_SALCAP * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 50) = Format(((((g_rst_Princi!COMCIE_TASINT / 100 + 1) ^ (1 / 360)) - 1) * (g_rst_Princi!COMCIE_SALCAP)), "###,###,##0.00")
         
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_PadDeu2_OLD()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_NumSol     As String
Dim r_int_TipMon     As Integer
Dim r_str_FecCam     As String
Dim r_str_TipCam     As String
Dim r_str_CheCgo     As String
Dim r_str_CtaCgo     As String
Dim r_str_BanCgo     As String
Dim r_int_MesCie     As Integer
Dim r_int_AnoCie     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_FecPpg     As String
Dim r_str_NomCli     As String
Dim r_str_FecCie     As String
Dim r_str_PerMes     As String
Dim r_str_PerAno     As String
Dim r_int_TipGar     As Integer

   'Inicializa Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "NRO. OPERACION"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "F. DESEMBOLSO"
      .Cells(1, 7) = "TIPO DE MONEDA"
      .Cells(1, 8) = "DIAS DE ATRASO"
      .Cells(1, 9) = "CUOTAS PAGADAS (MESES)"
      .Cells(1, 10) = "PLAZO (AÑOS)"
      .Cells(1, 11) = "SITUACION DEL CREDITO"
      .Cells(1, 12) = "TIPO DE GARANTIA"
      .Cells(1, 13) = "GARANTIA  S/."
      .Cells(1, 14) = "GARANTIA US$"
      .Cells(1, 15) = "CLASIF. INTERNA"
      .Cells(1, 16) = "CLASIF. ALINEADA"
      .Cells(1, 17) = "CLASIF. PROVIS."
      .Cells(1, 18) = "TASA PROVIS."
      .Cells(1, 19) = "PROV. GENERICA"
      .Cells(1, 20) = "PROV. GENERICA RC"
      .Cells(1, 21) = "PROV. ESPECIFICA"
      .Cells(1, 22) = "PROV. RIESGO. CAMB."
      .Cells(1, 23) = "PROV. PROCICLICA"
      .Cells(1, 24) = "PROV. PROCICLICA RC"
      .Cells(1, 25) = "PROV. VOLUNTARIA"
      .Cells(1, 26) = "PROV. RIESGO PAIS"
      .Cells(1, 27) = "APLICACION PROCIC."
      .Cells(1, 28) = "DVG. VIGENTE"
      .Cells(1, 29) = "DVG. VIGENTE SOLES"
      .Cells(1, 30) = "DVG. VENCIDO"
      .Cells(1, 31) = "DVG. VENCIDO SOLES"
      .Cells(1, 32) = "DVG. REFINANCIADO"
      .Cells(1, 33) = "DVG. REFINANCIADO SOLES"
      .Cells(1, 34) = "INTERES DIFERIDO"
      .Cells(1, 35) = "CAPITAL JUDICIAL"
      .Cells(1, 36) = "CAPITAL REFINANCIADO"
      .Cells(1, 37) = "CAP. VENC. REFINAN."
      .Cells(1, 38) = "CAPITAL VENCIDO"
      .Cells(1, 39) = "CAPITAL VIGENTE"
      .Cells(1, 40) = "DVG. PBP"
      .Cells(1, 41) = "PROV. INT. COFIDE"
      .Cells(1, 42) = "PROV. COM. COFIDE"
      .Cells(1, 43) = "TASA DE INTERES"
      .Cells(1, 44) = "SALDO CAPITAL"
      .Cells(1, 45) = "SALDO CAPITAL S/."
      .Cells(1, 46) = "INT. x DIAS"
      .Cells(1, 47) = "COBERTURA FMV"
      .Cells(1, 48) = "COBERTURA FMV RC"
      .Cells(1, 49) = "COBERTURA FMV AUTOLIQUIDABLE"
      .Cells(1, 50) = "INT. CUO. CIERRE"
      .Cells(1, 51) = "FEC. VCT. CIERRE"
      .Cells(1, 52) = "SAL. CAP. CIERRE"
      .Cells(1, 53) = "PRIMERA VIVIENDA"
      .Cells(1, 54) = "ADQUISICION DE PRIMERA VIVIENDA CON HIPOTECA"
      .Cells(1, 55) = "OTROS CON HIPOTECA INSCRITA"
      .Cells(1, 56) = "ADQUISICION DE PRIMERA VIVIENDA SIN HIPOTECA"
      .Cells(1, 57) = "OTROS SIN HIPOTECA INSCRITA"
      .Cells(1, 58) = "OTRAS GARANTIAS NO PREFERIDAS - BLOQUEO > 90 DIAS"
      .Cells(1, 59) = "PROVIS. REQUERIDA"
      .Cells(1, 60) = "PROVIS. CONSTITUIDA"
      .Cells(1, 61) = "FINANCIAMIENTO"
      .Cells(1, 62) = "PASIVO TNC"
      .Cells(1, 63) = "PASIVO TC"
      .Cells(1, 64) = "HIPOTECA MATRIZ"
      
      .Range(.Cells(1, 1), .Cells(1, 64)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 64)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 50
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 14
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 45
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 18
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 20
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 25
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12
      .Columns("N").ColumnWidth = 12
      .Columns("O").ColumnWidth = 14
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 15
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 12
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 14
      .Columns("T").ColumnWidth = 17
      .Columns("U").ColumnWidth = 15
      .Columns("V").ColumnWidth = 18
      .Columns("W").ColumnWidth = 16
      .Columns("X").ColumnWidth = 18
      .Columns("Y").ColumnWidth = 16
      .Columns("Z").ColumnWidth = 17
      .Columns("AA").ColumnWidth = 17
      .Columns("AB").ColumnWidth = 12
      .Columns("AC").ColumnWidth = 17
      .Columns("AD").ColumnWidth = 12
      .Columns("AE").ColumnWidth = 18
      .Columns("AF").ColumnWidth = 18
      .Columns("AG").ColumnWidth = 21
      .Columns("AH").ColumnWidth = 16
      .Columns("AI").ColumnWidth = 16
      .Columns("AJ").ColumnWidth = 19
      .Columns("AK").ColumnWidth = 17
      .Columns("AL").ColumnWidth = 14
      .Columns("AM").ColumnWidth = 15
      .Columns("AN").ColumnWidth = 16
      .Columns("AO").ColumnWidth = 14
      .Columns("AP").ColumnWidth = 16
      .Columns("AQ").ColumnWidth = 13
      .Columns("AR").ColumnWidth = 16
      .Columns("AS").ColumnWidth = 16
      .Columns("AT").ColumnWidth = 11
      .Columns("AU").ColumnWidth = 15
      .Columns("AV").ColumnWidth = 17
      .Columns("AW").ColumnWidth = 28
      .Columns("AX").ColumnWidth = 15
      .Columns("AY").ColumnWidth = 14
      .Columns("AY").HorizontalAlignment = xlHAlignCenter
      .Columns("AZ").ColumnWidth = 15
      .Columns("BA").ColumnWidth = 16
      .Columns("BA").HorizontalAlignment = xlHAlignCenter
      .Columns("BB").ColumnWidth = 40
      .Columns("BC").ColumnWidth = 27
      .Columns("BD").ColumnWidth = 40
      .Columns("BE").ColumnWidth = 26
      .Columns("BF").ColumnWidth = 43
      .Columns("BG").ColumnWidth = 18
      .Columns("BH").ColumnWidth = 18
      .Columns("BI").ColumnWidth = 18
      .Columns("BI").HorizontalAlignment = xlHAlignCenter
      .Columns("BJ").ColumnWidth = 15
      .Columns("BK").ColumnWidth = 15
      .Columns("BL").ColumnWidth = 15
      .Columns("BL").HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 64)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(1, 64)).Font.Size = 8
   End With
   
   '*** INICIALIZA VARIABLES
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = ipp_PerAno.Text
   r_int_MesCie = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_AnoCie = CInt(ipp_PerAno.Text)
   r_str_FecIni = Format(r_int_AnoCie, "0000") & Format(r_int_MesCie, "00") & "01"
   r_str_FecFin = Format(r_int_AnoCie, "0000") & Format(r_int_MesCie, "00") & "31"
   r_str_NomCli = ""
   
   '*** INFORMACION DE CREDITOS HIPOTECARIOS
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_NUMOPE, TRIM(B.PRODUC_DESCRI) AS PRODUC_DESCRI, A.HIPCIE_TDOCLI, A.HIPCIE_NDOCLI, A.HIPCIE_CODPRD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMCLIENTE, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_TIPMON, A.HIPCIE_FECDES, TRIM(D.PARDES_DESCRI) AS MONEDA_DESEM, A.HIPCIE_DIAMOR, G.HIPMAE_GARLIN, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CUOPAG, A.HIPCIE_SITCRE, A.HIPCIE_TIPGAR, A.HIPCIE_CLACLI, A.HIPCIE_CLAALI, A.HIPCIE_PRVVOL, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CLAPRV, A.HIPCIE_PRVGEN, A.HIPCIE_PRVGEN_RC, A.HIPCIE_PRVESP, A.HIPCIE_PRVCAM, A.HIPCIE_FLGJUD, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_PRVCIC, A.HIPCIE_PRVCIC_RC, A.HIPCIE_DEVVIG, A.HIPCIE_TIPCAM, A.HIPCIE_DEVVEN, A.HIPCIE_TIPCAM, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_INTDIF, A.HIPCIE_CAPVEN, A.HIPCIE_CAPVIG, A.HIPCIE_ACUDVG, A.HIPCIE_ACUDVC, A.HIPCIE_PRENCO, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_DEVPBP, A.HIPCIE_PRVICO, A.HIPCIE_PRVCCO, A.HIPCIE_TASINT, A.HIPCIE_SALCAP, A.HIPCIE_SALCON, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CBRFMV, A.HIPCIE_CBRFMV_RC, A.HIPCIE_PRVRIP, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN A.HIPCIE_TIPGAR = 6 THEN J.HIPDES_TIPMON ELSE A.HIPCIE_MONGAR END) AS HIPCIE_MONGAR, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN A.HIPCIE_TIPGAR = 6 THEN J.HIPDES_DESMPR ELSE A.HIPCIE_MTOGAR END) AS HIPCIE_MTOGAR, TRIM(E.PARDES_DESCRI) AS GARANTIA, "
   g_str_Parame = g_str_Parame & "       F.HIPCUO_INTERE, F.HIPCUO_FECVCT, F.HIPCUO_SALCAP, A.HIPCIE_FLGJUD, A.HIPCIE_APLCIC, TRIM(H.PARDES_DESCRI) AS PRI_VIVIENDA, "
   g_str_Parame = g_str_Parame & "       EVALEG_FECBLQ_INM, HIPMAE_PRIVIV, HIPCIE_FLGREF, "
   g_str_Parame = g_str_Parame & "       (SELECT ROUND(DECODE(MITAB.MONEDA, 1, MITAB.PROV_REQUERIDA, MITAB.PROV_REQUERIDA * MITAB.TIPO_CAMBIO), 2) AS PROVISION_REQUERIDA "
   g_str_Parame = g_str_Parame & "          FROM (SELECT CASE WHEN HP.HIPCIE_CLAPRV = 0 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.007 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.007 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 1 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.025 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.050 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 2 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.125 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.25 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 3 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.3 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.6 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                            WHEN HP.HIPCIE_CLAPRV = 4 THEN CASE WHEN HP.HIPCIE_TIPGAR IN (1,2,4) "
   g_str_Parame = g_str_Parame & "                                                      THEN (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 0.6 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                      ELSE (HP.HIPCIE_SALCAP+HP.HIPCIE_SALCON-HP.HIPCIE_INTDIF-HP.HIPCIE_CBRFMV-HP.HIPCIE_CBRFMV_RC) * 1 + (HP.HIPCIE_CBRFMV_RC) * 0.7/100 "
   g_str_Parame = g_str_Parame & "                                                 END "
   g_str_Parame = g_str_Parame & "                       END AS PROV_REQUERIDA,"
   g_str_Parame = g_str_Parame & "                       HP.HIPCIE_TIPMON AS MONEDA, HP.HIPCIE_TIPCAM AS TIPO_CAMBIO,HP.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "                  FROM CRE_HIPCIE HP "
   g_str_Parame = g_str_Parame & "                 WHERE HP.HIPCIE_PERMES = " & CInt(r_str_PerMes)
   g_str_Parame = g_str_Parame & "                   AND HP.HIPCIE_PERANO = " & CInt(r_str_PerAno) & ") MITAB WHERE MITAB.HIPCIE_NUMOPE = A.HIPCIE_NUMOPE) AS PROVISION_REQUERIDA, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPCIE_TIPMON, 1, A.HIPCIE_PRVGEN+A.HIPCIE_PRVESP+A.HIPCIE_PRVCIC+A.HIPCIE_PRVGEN_RC+A.HIPCIE_PRVCIC_RC+A.HIPCIE_PRVVOL, (A.HIPCIE_PRVGEN+A.HIPCIE_PRVESP+A.HIPCIE_PRVCIC+A.HIPCIE_PRVGEN_RC+A.HIPCIE_PRVCIC_RC+A.HIPCIE_PRVVOL)*A.HIPCIE_TIPCAM) AS PROVISION_CONSTITUIDA, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT TIPPRV_PORCEN "
   g_str_Parame = g_str_Parame & "              FROM CTB_TIPPRV "
   g_str_Parame = g_str_Parame & "             WHERE TIPPRV_TIPPRV = '2' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLACRE = '13' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLAGAR = 2),0) AS TASA_CG , "
   g_str_Parame = g_str_Parame & "       NVL((SELECT TIPPRV_PORCEN "
   g_str_Parame = g_str_Parame & "              FROM CTB_TIPPRV "
   g_str_Parame = g_str_Parame & "             WHERE TIPPRV_TIPPRV = '2' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLACRE = '13' "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "               AND TIPPRV_CLAGAR = 1),0) AS TASA_SG, "
   g_str_Parame = g_str_Parame & "       K.SOLMAE_PLAANO AS PLAZO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 3 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPNCO "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 3 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_TNC, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 4 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= TO_CHAR(ADD_MONTHS(TO_DATE(" & r_str_FecIni & ", 'YYYY/MM/DD'), -5), 'YYYYMMDD') AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPCON "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 4 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= TO_CHAR(ADD_MONTHS(TO_DATE(" & r_str_FecIni & ", 'YYYY/MM/DD'), -5), 'YYYYMMDD') AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_TC, "
   g_str_Parame = g_str_Parame & "       CASE WHEN NVL((SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 5 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                         AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & "), 0) = 0 "
   g_str_Parame = g_str_Parame & "            THEN HIPMAE_IMPNCO "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT NVL(CUOCIE_SALCAP, 0) FROM CRE_CUOCIE WHERE CUOCIE_PERMES = " & CStr(r_int_MesCie) & " AND CUOCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_NUMOPE = HIPCIE_NUMOPE AND CUOCIE_TIPCRO = 5 AND CUOCIE_NUMCUO > 0 "
   g_str_Parame = g_str_Parame & "                     AND CUOCIE_FECVCT >= " & r_str_FecIni & " AND CUOCIE_FECVCT <= " & r_str_FecFin & ") "
   g_str_Parame = g_str_Parame & "       END AS PASIVO_CME, "
   g_str_Parame = g_str_Parame & "       CASE WHEN G.HIPMAE_HIPMTZ = 3 THEN 'NO LEVANTADA' "
   g_str_Parame = g_str_Parame & "            WHEN G.HIPMAE_HIPMTZ = 2 THEN 'LEVANTADA' "
   g_str_Parame = g_str_Parame & "            ELSE 'NO APLICA' "
   g_str_Parame = g_str_Parame & "       END AS HIPOTECA_MATRIZ "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.HIPCIE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = A.HIPCIE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(A.HIPCIE_NDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPCIE_TIPMON "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.HIPCIE_TIPGAR "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPCUO F ON F.HIPCUO_NUMOPE = A.HIPCIE_NUMOPE AND F.HIPCUO_TIPCRO = 1 AND F.HIPCUO_FECVCT >= " & r_str_FecIni & " AND F.HIPCUO_FECVCT <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE G ON G.HIPMAE_NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = '214' AND H.PARDES_CODITE = G.HIPMAE_PRIVIV"
   g_str_Parame = g_str_Parame & "  LEFT OUTER JOIN TRA_EVALEG I ON I.EVALEG_NUMSOL = G.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPDES J ON J.HIPDES_NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE K ON K.SOLMAE_NUMERO = G.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERMES = " & CStr(r_int_MesCie) & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   If chk_TipPro.Value = 0 Then
     g_str_Parame = g_str_Parame & " AND A.HIPCIE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
   End If
   g_str_Parame = g_str_Parame & "ORDER BY A.HIPCIE_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_ConVer = 2
      r_str_FecCie = Format(CInt(ipp_PerAno.Text), "####") & Format(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), "0#") & Format(ff_Ultimo_Dia_Mes(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(ipp_PerAno.Text)), "00")
      
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!HIPCIE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPCIE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPCIE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMCLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_FECDES))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!MONEDA_DESEM)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!HIPCIE_DIAMOR, "00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCIE_CUOPAG, "00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = g_rst_Princi!PLAZO_PRESTAMO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!HIPCIE_SITCRE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!GARANTIA)
         
         If g_rst_Princi!HIPCIE_MONGAR = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(0, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!HIPCIE_CLACLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!HIPCIE_CLAALI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Trim(g_rst_Princi!HIPCIE_CLAPRV)
         
         '-----------------------------------EXCEPCIONES SOBRE LA TASA---------------------------------------------
         r_int_TipGar = 0
         '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            g_str_CadCnx = ""
            g_str_CadCnx = g_str_CadCnx & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
            g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
            g_str_CadCnx = g_str_CadCnx & "          FROM CRE_HIPCIE "
            g_str_CadCnx = g_str_CadCnx & "         WHERE HIPCIE_PERMES > 0 "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_PERANO > 2010 "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_NUMOPE = '" & g_rst_Princi!HIPCIE_NUMOPE & "' "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_CLAPRV = 3 "
            g_str_CadCnx = g_str_CadCnx & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            g_str_CadCnx = g_str_CadCnx & " WHERE ROWNUM < 37 "
            g_str_CadCnx = g_str_CadCnx & " GROUP BY HIPCIE_CLAPRV "
            
            If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
               Exit Sub
            End If
            
            If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               g_rst_GenAux.MoveFirst
               If g_rst_GenAux!CONTADOR = 36 Then
                  r_int_TipGar = 5
               End If
            End If
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
         End If
         
         '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            g_str_CadCnx = ""
            g_str_CadCnx = g_str_CadCnx & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
            g_str_CadCnx = g_str_CadCnx & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
            g_str_CadCnx = g_str_CadCnx & "          FROM CRE_HIPCIE "
            g_str_CadCnx = g_str_CadCnx & "         WHERE HIPCIE_PERMES > 0 "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_PERANO > 2009 "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_NUMOPE = '" & g_rst_Princi!HIPCIE_NUMOPE & "' "
            g_str_CadCnx = g_str_CadCnx & "           AND HIPCIE_CLAPRV = 4 "
            g_str_CadCnx = g_str_CadCnx & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            g_str_CadCnx = g_str_CadCnx & " WHERE ROWNUM < 25 "
            g_str_CadCnx = g_str_CadCnx & " GROUP BY HIPCIE_CLAPRV "
            
            If Not gf_EjecutaSQL(g_str_CadCnx, g_rst_GenAux, 3) Then
               Exit Sub
            End If
            
            If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               g_rst_GenAux.MoveFirst
               If g_rst_GenAux!CONTADOR = 24 Then
                  r_int_TipGar = 5
               End If
            End If
            
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
         End If
         
         
         If (r_int_TipGar = 5) Then
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDbl(Format(g_rst_Princi!TASA_SG, "###,###,##0.00"))
         Else
             If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 4) Then
                 r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDbl(Format(g_rst_Princi!TASA_CG, "###,###,##0.00"))
             Else
                 r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDbl(Format(g_rst_Princi!TASA_SG, "###,###,##0.00"))
             End If
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!HIPCIE_PRVGEN, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!HIPCIE_PRVGEN_RC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!HIPCIE_PRVESP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!HIPCIE_PRVCAM, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!HIPCIE_PRVCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPCIE_PRVCIC_RC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPCIE_PRVVOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!HIPCIE_PRVRIP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Format(g_rst_Princi!HIPCIE_APLCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
         
         If g_rst_Princi!HIPCIE_FLGREF = 0 Then
            If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!HIPCIE_ACUDVG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
            If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!HIPCIE_ACUDVC * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = 0
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = 0
            
            If g_rst_Princi!HIPCIE_ACUDVC > 0 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!HIPCIE_ACUDVC * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!HIPCIE_ACUDVC, "###,###,##0.00")
               End If
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!HIPCIE_ACUDVG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!HIPCIE_ACUDVG, "###,###,##0.00")
               End If
            End If
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(g_rst_Princi!HIPCIE_INTDIF, "###,###,##0.00")
         If g_rst_Princi!HIPCIE_FLGJUD = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
         Else
            If g_rst_Princi!HIPCIE_FLGREF = 1 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!HIPCIE_CAPVIG, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(g_rst_Princi!HIPCIE_CAPVEN, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(0, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(g_rst_Princi!HIPCIE_CAPVEN, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(g_rst_Princi!HIPCIE_CAPVIG, "###,###,##0.00")
            End If
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(g_rst_Princi!HIPCIE_DEVPBP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(g_rst_Princi!HIPCIE_PRVICO, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(g_rst_Princi!HIPCIE_PRVCCO, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Format(g_rst_Princi!HIPCIE_TASINT, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
         
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON), "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Format(((((g_rst_Princi!HIPCIE_TASINT / 100 + 1) ^ (1 / 360)) - 1) * (g_rst_Princi!HIPCIE_SALCAP)), "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 47) = Format(g_rst_Princi!HIPCIE_CBRFMV, "###,###,##0.00")
         If g_rst_Princi!HIPCIE_FLGJUD = 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 48) = Format(g_rst_Princi!HIPCIE_CBRFMV_RC, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(0, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 48) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 49) = Format(g_rst_Princi!HIPCIE_CBRFMV_RC, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 50) = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         
         If Not IsNull(g_rst_Princi!HIPCUO_FECVCT) Then
            r_str_FecPpg = fs_BuscaPrepago(g_rst_Princi!HIPCIE_NUMOPE, r_str_FecIni, r_str_FecFin)
            If r_str_FecPpg = "0" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 51) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 51) = "'" & gf_FormatoFecha(r_str_FecPpg)
            End If
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 52) = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 51) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_FECDES))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 52) = Format(g_rst_Princi!HIPCIE_PRENCO, "###,###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 53) = g_rst_Princi!PRI_VIVIENDA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 54) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = Format(0, "###,###,##0.00")
         
         If g_rst_Princi!HIPMAE_PRIVIV = 1 Then
            If g_rst_Princi!HIPCIE_FECDES > 20130101 Then
               If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 54) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 54) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               Else
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 56) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               End If
            End If
         Else
            If g_rst_Princi!HIPCIE_FECDES > 20130101 Then
               If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 55) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               Else
                  If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  End If
               End If
            End If
         End If
         
         If g_rst_Princi!HIPCIE_TIPGAR = 2 Then
            If DateDiff("d", Format(g_rst_Princi!EVALEG_FECBLQ_INM, "####/##/##"), Format(r_str_FecCie, "####/##/##")) > 90 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 57) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = Format(0, "###,###,##0.00")
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 58) = Format(0, "###,###,##0.00")
         End If

         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 59) = IIf(IsNull(g_rst_Princi!PROVISION_REQUERIDA), "0.00", Format(g_rst_Princi!PROVISION_REQUERIDA, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = IIf(IsNull(g_rst_Princi!PROVISION_CONSTITUIDA), "0.00", Format(g_rst_Princi!PROVISION_CONSTITUIDA, "###,###,##0.00"))
         
         '****************
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 60) = ""
         Select Case g_rst_Princi!HIPMAE_GARLIN
            Case "000001": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = "MIVIVIENDA"
            Case "000002": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = "BID"
            Case "000003": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = "F RAMIREZ PRADO"
            Case "000004": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = "COFICASA"
            Case "999999": r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 61) = "NINGUNO"
         End Select
         '****************
         
         If g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "002" Or g_rst_Princi!HIPCIE_CODPRD = "006" Or g_rst_Princi!HIPCIE_CODPRD = "011" Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = Format(0, "###,###,##0.00")
         Else
            If g_rst_Princi!HIPCIE_CODPRD = "003" Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(g_rst_Princi!PASIVO_CME, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = Format(0, "###,###,##0.00")
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 62) = Format(g_rst_Princi!PASIVO_TNC, "###,###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 63) = Format(g_rst_Princi!PASIVO_TC, "###,###,##0.00")
            End If
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 64) = Trim(g_rst_Princi!HIPOTECA_MATRIZ)
         
         r_int_ConVer = r_int_ConVer + 1
         r_str_NomCli = Trim(g_rst_Princi!NOMCLIENTE) '(*)guarda en variable nombre cliente
         g_rst_Princi.MoveNext
                  
         If g_rst_Princi.EOF Then
            '(*)si es fin de archivo sale del bucle
            Exit Do
         Else
            '(*)compara variable nombre cliente con el proximo cliente,
            'y adelanta un registro si comparacion entre clientes es igual.
            If r_str_NomCli = Trim(g_rst_Princi!NOMCLIENTE) Then
               g_rst_Princi.MoveNext
            End If
         End If
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '*** INFORMACION DE CREDITOS INMMOBILIARIOS
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.COMCIE_NUMOPE, TRIM(B.PRODUC_DESCRI) AS PRODUC_DESCRI, A.COMCIE_TDOCLI, A.COMCIE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_RAZSOC) AS NOMEMPRESA, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_TIPMON, A.COMCIE_FECDES, TRIM(D.PARDES_DESCRI) AS MONEDA_DESEM, A.COMCIE_DIAMOR, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CUOPAG, A.COMCIE_SITCRE, A.COMCIE_TIPGAR, A.COMCIE_CLACLI, A.COMCIE_CLAALI, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CLAPRV, A.COMCIE_PRVGEN, A.COMCIE_PRVESP, A.COMCIE_PRVCAM, A.COMCIE_FLGJUD, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_PRVCIC, A.COMCIE_DEVVIG, A.COMCIE_TIPCAM, A.COMCIE_DEVVEN, A.COMCIE_INTDIF, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CAPVEN, A.COMCIE_CAPVIG, A.COMCIE_ACUDVG, A.COMCIE_ACUDVC, A.COMCIE_TASINT, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_SALCAP, A.COMCIE_MONGAR, A.COMCIE_MTOGAR, TRIM(E.PARDES_DESCRI) AS GARANTIA "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.COMCIE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = A.COMCIE_TDOCLI AND TRIM(C.DATGEN_EMPNDO) = TRIM(A.COMCIE_NDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.COMCIE_TIPMON "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.COMCIE_TIPGAR "
   g_str_Parame = g_str_Parame & " WHERE A.COMCIE_PERMES = " & CStr(r_int_MesCie) & " "
   g_str_Parame = g_str_Parame & "   AND A.COMCIE_PERANO = " & Format(r_int_AnoCie, "0000") & " "
   g_str_Parame = g_str_Parame & "   AND A.COMCIE_TIPGAR = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY A.COMCIE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!COMCIE_NUMOPE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!COMCIE_TDOCLI) & "-" & Trim(g_rst_Princi!COMCIE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMEMPRESA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!COMCIE_FECDES))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!MONEDA_DESEM)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "1"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!GARANTIA)
         If g_rst_Princi!COMCIE_MONGAR = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(0, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(0, "###,###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!COMCIE_CLACLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!COMCIE_CLAALI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!COMCIE_CLAPRV)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0                                    'TASA PROVIS.
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!COMCIE_PRVGEN, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = "0"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!COMCIE_PRVESP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!COMCIE_PRVCAM, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!COMCIE_PRVCIC, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = "0"                                  'PROV. PROCICLICA RC
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = "0"                                  'PROV. VOLUNTARIA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = "0"                                  'APLICACION PROCIC.
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!COMCIE_ACUDVG, "###,###,##0.00")
         If g_rst_Princi!COMCIE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Format(g_rst_Princi!COMCIE_ACUDVG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Format(g_rst_Princi!COMCIE_ACUDVG, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Format(g_rst_Princi!COMCIE_ACUDVC, "###,###,##0.00")
         If g_rst_Princi!COMCIE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!COMCIE_ACUDVC * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!COMCIE_ACUDVC, "###,###,##0.00")
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = "0"                                  'DVG. REFINANCIADO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = "0"                                  'DVG. REFINANCIADO SOLES
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Format(g_rst_Princi!COMCIE_INTDIF, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(0, "###,###,##0.00")          'CAPITAL JUDICIAL
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format(0, "###,###,##0.00")          'CAPITAL REFINANCIADO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format(g_rst_Princi!COMCIE_CAPVEN, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Format(g_rst_Princi!COMCIE_CAPVIG, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Format(g_rst_Princi!COMCIE_TASINT, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Format(g_rst_Princi!COMCIE_SALCAP, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Format(((((g_rst_Princi!COMCIE_TASINT / 100 + 1) ^ (1 / 360)) - 1) * (g_rst_Princi!COMCIE_SALCAP)), "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 44) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 45) = Format(0, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 46) = Format(0, "###,###,##0.00")
         
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function fs_BuscaPrepago(ByVal p_NumOpe As String, ByVal p_FecIni As String, ByVal p_FecFin As String) As String
Dim r_str_Parama     As String
Dim r_rst_PrePag     As ADODB.Recordset

   fs_BuscaPrepago = "0"
   
   r_str_Parama = ""
   r_str_Parama = r_str_Parama & "SELECT NVL(PPGCAB_FECPPG, 0) AS FECHA_PREPAGO "
   r_str_Parama = r_str_Parama & "  FROM CRE_PPGCAB "
   r_str_Parama = r_str_Parama & " WHERE PPGCAB_NUMOPE  = '" & p_NumOpe & "' "
   r_str_Parama = r_str_Parama & "   AND PPGCAB_FECPPG >= " & p_FecIni & " "
   r_str_Parama = r_str_Parama & "   AND PPGCAB_FECPPG <= " & p_FecFin & " "
   r_str_Parama = r_str_Parama & " ORDER BY PPGCAB_FECPPG DESC "
   
   If Not gf_EjecutaSQL(r_str_Parama, r_rst_PrePag, 3) Then
      Exit Function
   End If
   
   If r_rst_PrePag.BOF And r_rst_PrePag.EOF Then
      r_rst_PrePag.Close
      Set r_rst_PrePag = Nothing
      Exit Function
   End If
   
   r_rst_PrePag.MoveFirst
   fs_BuscaPrepago = r_rst_PrePag!FECHA_PREPAGO
   
   r_rst_PrePag.Close
   Set r_rst_PrePag = Nothing
End Function
