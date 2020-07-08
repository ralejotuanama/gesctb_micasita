VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   9690
   ClientTop       =   6480
   ClientWidth     =   4770
   Icon            =   "GesCtb_frm_009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   4207
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
         TabIndex        =   5
         Top             =   60
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
            TabIndex        =   6
            Top             =   30
            Width           =   3795
            _Version        =   65536
            _ExtentX        =   6694
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Control de Limites Globales e Individuales"
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
            TabIndex        =   7
            Top             =   270
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte 13"
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
            Picture         =   "GesCtb_frm_009.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
            Left            =   4080
            Picture         =   "GesCtb_frm_009.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_009.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1440
            TabIndex        =   1
            Top             =   450
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
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   90
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_salMay        As Double
Dim l_dbl_SalMen        As Double
Dim l_dbl_SalSol        As Double
Dim l_dbl_SalDol        As Double
Dim l_str_PerMes        As String
Dim l_str_PerAno        As String
Dim l_str_FecIni        As String
Dim l_str_FecFin        As String
Dim l_dbl_CapRes        As Double
Dim l_dbl_PatEfe        As Double

Private Function fp_UltimoDia(ByVal p_Fec_xMes As String, p_Fec_xAno As String) As String
Dim r_Int_Dia As Integer
   
   Select Case CInt(p_Fec_xMes)
       Case 1, 3, 5, 7, 8, 10, 12
            r_Int_Dia = 31
       Case 2
            If (CInt(p_Fec_xAno) - 2008) Mod 4 = 0 Then
                'Año bisiesto 2008, 2012,2016, 2020...etc
                    r_Int_Dia = 29
                Else
                    r_Int_Dia = 28
            End If
       Case 4, 6, 9, 11
            r_Int_Dia = 30
   End Select
   
   fp_UltimoDia = Str(r_Int_Dia) + "/" + p_Fec_xMes + "/" + p_Fec_xAno
End Function

Private Sub cmd_ExpExc_Click()
   l_dbl_CapRes = 0
   l_dbl_PatEfe = 0
   l_dbl_salMay = 0
   l_dbl_SalMen = 0
   l_dbl_SalSol = 0
   l_dbl_SalDol = 0
   l_str_PerMes = ""
   l_str_PerAno = ""
   l_str_FecIni = ""
   l_str_FecFin = ""

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
   
   l_str_PerAno = Format(ipp_PerAno.Text, "0000")
   l_str_PerMes = Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00")
   l_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   l_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
      
   If ff_Buscar_1 = 0 Then
      MsgBox "Debe procesar el mes y año seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If ff_Buscar_2 = 0 Then
      MsgBox "Debe de llenar el mantenedor de Limtes Globales del mes y Año Seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   'Call fs_GenExc(l_str_FecIni, l_str_FecFin)
   Call fs_GenExc2
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
End Sub

Private Function ff_Buscar_1() As Integer
   ff_Buscar_1 = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG "
   g_str_Parame = g_str_Parame & "  FROM TMP_LIMGLO  "
   g_str_Parame = g_str_Parame & " WHERE LIMGLO_PERMES = " & l_str_PerMes & " "
   g_str_Parame = g_str_Parame & "   AND LIMGLO_PERANO = " & l_str_PerAno & " "
   g_str_Parame = g_str_Parame & "   AND LIMGLO_CODSUC = 001 "
   g_str_Parame = g_str_Parame & "   AND LIMGLO_CODEMP = 000001 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_Buscar_1 = g_rst_Princi!TOTREG
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Function ff_Buscar_2() As Integer
   ff_Buscar_2 = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG "
   g_str_Parame = g_str_Parame & "  FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & " WHERE CONLIM_CODMES = " & l_str_PerMes & " "
   g_str_Parame = g_str_Parame & "   AND CONLIM_CODANO = " & l_str_PerAno & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_Buscar_2 = g_rst_Princi!TOTREG
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Function ff_TipCam() As Double
   ff_TipCam = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM OPE_TIPCAM "
   g_str_Parame = g_str_Parame & " WHERE TIPCAM_CODIGO = 2 "
   g_str_Parame = g_str_Parame & "   AND TIPCAM_FECDIA = '" & l_str_FecFin & "' "
   g_str_Parame = g_str_Parame & " ORDER BY TIPCAM_FECDIA DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_TipCam = g_rst_Princi!TIPCAM_COMPRA
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Function ff_ConLim() As Double
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & " WHERE CONLIM_CODANO = " & l_str_PerAno & " "
   g_str_Parame = g_str_Parame & "   AND CONLIM_CODMES = " & l_str_PerMes & " "
   g_str_Parame = g_str_Parame & " ORDER BY CONLIM_CODANO, CONLIM_CODMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_dbl_CapRes = g_rst_Princi!CONLIM_CAPRES
      l_dbl_PatEfe = g_rst_Princi!CONLIM_PATEFE
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Sub fs_GenExc(ByVal p_FecIni As String, ByVal p_FecFin As String)
Dim r_obj_Excel      As Excel.Application
Dim r_rst_Client     As ADODB.Recordset
Dim r_int_ConVer     As Integer
Dim r_int_FecVct     As Double
Dim r_dbl_TncMen_01  As Double
Dim r_dbl_TncMay_01  As Double
Dim r_dbl_TcMen_01   As Double
Dim r_dbl_TcMay_01   As Double
Dim r_dbl_SalMen     As Double
Dim r_dbl_SalMay     As Double
Dim r_dbl_MtoMen     As Double
Dim r_dbl_MtoMay     As Double
Dim r_dbl_SinGar     As Double
Dim r_dbl_ConGar     As Double
Dim r_dbl_LimGlo     As Double
Dim r_dbl_LimInd     As Double
Dim p_SalMen         As Double
Dim p_salMay         As Double
Dim r_dbl_SalCap     As Double
Dim r_dbl_TipCam     As Double
Dim r_dbl_ConLim     As Double
Dim r_int_SinGar     As Integer
Dim r_int_ConGar     As Integer
   
   Call ff_ConLim
   r_dbl_TipCam = ff_TipCam
   r_dbl_LimGlo = 7 / 100 * l_dbl_CapRes
   r_dbl_LimInd = r_dbl_LimGlo * 5 / 100
   r_dbl_SinGar = 0.1 * l_dbl_PatEfe
   r_dbl_ConGar = 0.15 * l_dbl_PatEfe
   
   '-- Consulta listado de Clientes cierre
   g_str_Parame = ff_Query(l_str_PerMes, l_str_PerAno)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Client, 3) Then
      Exit Sub
   End If

   If r_rst_Client.BOF And r_rst_Client.EOF Then
      r_rst_Client.Close
      Set r_rst_Client = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(6, 1) = "CONTROL DE LIMITES GLOBALES E INDIVIDUALES APLICABLES A LAS EMPRESAS DEL SISTEMA FINANCIERO"
      .Cells(7, 1) = "REPORTE 13"
      .Cells(10, 1) = "Al: " & Right(p_FecFin, 2) & "/" & Mid(p_FecFin, 5, 2) & "/" & Left(p_FecFin, 4)
      .Cells(10, 5) = "Tipo de Cambio"
      .Cells(10, 8) = r_dbl_TipCam
      .Cells(12, 1) = "Capital y Reserva Legal : "
      .Cells(12, 2) = l_dbl_CapRes
      .Cells(12, 5) = "Patrimonio Efectivo: (un mes anterior)"
      .Cells(12, 8) = l_dbl_PatEfe
      .Cells(14, 1) = "Limite Global: 7% del Patrimonio Efectivo:"
      .Cells(14, 2) = Format(r_dbl_LimGlo, "###,###,##0.00")
      .Cells(15, 1) = "Limite Individual: 7% del Patrimonio Efectivo del 5%:"
      .Cells(15, 2) = Format(r_dbl_LimInd, "###,###,##0.00")
      
      .Cells(18, 1) = "CREDITOS A DIRECTORES Y TRABAJADORES DE LA EMPRESA:"
      .Cells(19, 1) = "Apellidos y Nombres"
      .Cells(19, 2) = "Credito"
      .Cells(19, 3) = "Credito"
      .Cells(19, 4) = "TEA"
      .Cells(19, 5) = "Limite"
      .Cells(19, 6) = "Control"
      .Cells(19, 7) = "Limite"
      .Cells(19, 8) = "Control"
      .Cells(20, 2) = "Dolares"
      .Cells(20, 3) = "Soles"
      .Cells(20, 4) = "%"
      .Cells(20, 5) = "Individual %"
      .Cells(20, 7) = "Global %"
      .Cells(21, 1) = "Total"
      .Cells(21, 2) = "0.00"
      .Cells(21, 3) = "0.00"
      .Cells(21, 7) = "0.00"
      .Cells(21, 8) = "Si Procede"
      
      .Cells(23, 1) = "MAYOR FINANCIAMIENTO A TRABAJADOR:"
      .Cells(24, 1) = "Apellidos y Nombres"
      .Cells(24, 2) = "Credito"
      .Cells(24, 3) = "Credito"
      .Cells(24, 4) = "TEA"
      .Cells(24, 5) = "Limite"
      .Cells(24, 6) = "Control"
      .Cells(25, 2) = "Dolares"
      .Cells(25, 3) = "Soles"
      .Cells(25, 5) = "Individual %"
      
      .Cells(28, 1) = "MAYOR FINANCIAMIENTO A TRABAJADOR:"
      .Cells(29, 1) = "Apellidos y Nombres"
      .Cells(29, 2) = "Credito"
      .Cells(29, 3) = "Credito"
      .Cells(29, 4) = "TEA"
      .Cells(29, 5) = "Limite 30 %"
      .Cells(29, 6) = "Limite"
      .Cells(29, 7) = "Control"
      .Cells(29, 8) = "Disponible"
      .Cells(30, 2) = "Dolares"
      .Cells(30, 3) = "Soles"
      .Cells(30, 6) = "Global %"
      .Cells(31, 1) = "Total"
      .Cells(31, 2) = "0.00"
      .Cells(31, 3) = "0.00"
      .Cells(31, 5) = Format(l_dbl_PatEfe * 0.3, "###,###,##0.00")
      .Cells(31, 6) = "0.00"
      .Cells(31, 7) = "Si Procede"
      .Cells(31, 8) = Format(l_dbl_PatEfe * 0.3, "###,###,##0.00")
      
      .Cells(34, 1) = "FINANCIAMIENTO CON O SIN GARANTIA:"
      .Cells(35, 1) = "Para créditos Sin Garantia: 10% del Patrimonio Efectivo->"
      .Cells(35, 2) = Format(r_dbl_SinGar, "###,###,##0.00")
      .Cells(36, 1) = "Para créditos Con Hipoteca: 15% del Patrimonio Efectivo->"
      .Cells(36, 2) = Format(r_dbl_ConGar, "###,###,##0.00")
      
      .Cells(39, 1) = "Apellidos y Nombres"
      .Cells(39, 2) = "Saldo"
      .Cells(40, 2) = "Capital US$"
      .Cells(39, 3) = "Saldo"
      .Cells(40, 3) = "Capital S/."
      .Cells(39, 4) = "Garantia"
      .Cells(39, 5) = "Disponible"
      .Cells(39, 6) = "Limite"
      .Cells(39, 7) = "control"
      
      .Cells(52, 1) = "*En este reporte se ordena la columna 'Saldo Capital' de mayor a menor."
      .Cells(54, 1) = "PRESTAMOS, CONTINGENTES Y OPERACIONES DE ARRENDAMIENTO FINANCIERO:"
      .Cells(55, 1) = "Cuando el vencimiento ocurra en un plazo mayor a 1 año:"
      .Cells(56, 4) = "Importe Maximo:"
      .Cells(57, 1) = "Hasta 4 veces el Patrimonio Efectivo (Art.200, Numeral 7)"
      .Cells(58, 2) = "Saldo"
      .Cells(59, 2) = "Capital US$"
      .Cells(58, 3) = "Saldo"
      .Cells(59, 3) = "Capital S/."
      .Cells(58, 4) = "Saldo"
      
      .Cells(60, 1) = "Créditos con plazo mayor a un 1 año"
      .Cells(61, 1) = "Créditos con plazo menor a un 1 año"
      .Cells(63, 1) = "Creditos Comerciales"
      .Cells(64, 1) = "Totales"
      .Cells(60, 5) = "Veces"
      .Cells(61, 5) = "Veces"
      .Cells(63, 5) = "Veces"
      
      .Range("A6:H6").Merge
      .Range("A7:H7").Merge
      .Range("A6:H6").Font.Bold = True
      .Range("A7:H7").Font.Bold = True
      .Range("A6:H6").HorizontalAlignment = xlHAlignCenter
      .Range("A7:H7").HorizontalAlignment = xlHAlignCenter
      
      .Range("E12:G12").Merge
      .Range("A18:H18").Merge
      .Cells(18, 1).Font.Bold = True
      .Range("A64:H64").Font.Bold = True
      
      'CREDITOS A DIRECTORES Y TRABAJADORES DE LA EMPRESA
      .Range("B20:H20").HorizontalAlignment = xlHAlignCenter
      .Range("A21:H21").Font.Bold = True
      
      'MAYOR FINANCIAMIENTO A TRABAJADOR
      .Range("A23:H23").Merge
      .Range("A23:H23").Font.Bold = True
      
      'FINANCIAMIENTO A PERSONAS VINCULADAS
      .Range("A28:H28").Merge
      .Range("A28:H28").Font.Bold = True
      .Range("A31:H31").Font.Bold = True
                         
      'FINANCIAMIENTO CON O SIN GARANTIA
      .Range("A34:H34").Merge
      .Range("A34:H34").Font.Bold = True
      
      .Range("A54:H54").Merge
      .Range("A54:H54").Font.Bold = True
      
      '.Range("D39:D40", "G39:G40").WrapText = True
      .Range("B19:H19", "B20:H20").HorizontalAlignment = xlHAlignCenter
      .Range("B24:H24", "B25:H25").HorizontalAlignment = xlHAlignCenter
      .Range("B29:H29", "B30:H30").HorizontalAlignment = xlHAlignCenter
      .Range("B39:G39", "B40:G40").HorizontalAlignment = xlHAlignCenter
      .Range("B58:D58", "B59:D59").HorizontalAlignment = xlHAlignCenter
      .Range("D41:D50").HorizontalAlignment = xlHAlignCenter
      .Range("G41:G50").HorizontalAlignment = xlHAlignCenter
      .Range("E60:E64").HorizontalAlignment = xlHAlignCenter
      
      .Cells(21, 8).HorizontalAlignment = xlHAlignCenter
      .Cells(31, 7).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 57.71
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 14
      .Columns("D").ColumnWidth = 16
      .Columns("E").ColumnWidth = 14
      .Columns("F").ColumnWidth = 12
      .Columns("G").ColumnWidth = 12
      .Columns("H").ColumnWidth = 12
      
      .Range(.Cells(18, 1), .Cells(18, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(18, 1), .Cells(18, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(18, 1), .Cells(18, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(18, 1), .Cells(18, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(23, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(23, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(23, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(23, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(54, 1), .Cells(54, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(54, 1), .Cells(54, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(54, 1), .Cells(54, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(54, 1), .Cells(54, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(2000, 10)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(2000, 10)).Font.Size = 8
   End With
   
   r_int_ConVer = 41
   r_dbl_MtoMen = 0
   r_dbl_MtoMay = 0
   
   r_rst_Client.MoveFirst
   Do While Not r_rst_Client.EOF
      If r_rst_Client!MONTO > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_rst_Client!NOMBRE
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Format(r_rst_Client!MONTO / r_dbl_TipCam, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Format(r_rst_Client!MONTO, "###,###,##0.00")
         
         If r_rst_Client!TIPGAR = 1 Or r_rst_Client!TIPGAR = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "Si"
            r_int_ConGar = r_int_ConGar + 1
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "No"
            r_int_SinGar = r_int_SinGar + 1
         End If
         
         If r_rst_Client!TIPGAR = 1 Or r_rst_Client!TIPGAR = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Format(r_dbl_ConGar - r_rst_Client!MONTO, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Format(r_dbl_SinGar - r_rst_Client!MONTO, "###,###,##0.00")
         End If
         
         If r_rst_Client!TIPGAR = 1 Or r_rst_Client!TIPGAR = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Format((r_rst_Client!MONTO / l_dbl_PatEfe) * 100, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Format((r_rst_Client!MONTO / l_dbl_PatEfe) * 100, "###,###,##0.00")
         End If
         
         If r_rst_Client!TIPGAR = 1 Or r_rst_Client!TIPGAR = 2 Then
            If r_rst_Client!MONTO / l_dbl_PatEfe * 100 > 15 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "No Procede"
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "Si Procede"
            End If
         Else
            If r_rst_Client!MONTO / l_dbl_PatEfe * 100 > 10 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "No Procede"
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "Si Procede"
            End If
         End If
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      r_rst_Client.MoveNext
      DoEvents
   Loop
   
   Call ff_SalCap_01
   Call ff_SalCap_02
   
   r_obj_Excel.ActiveSheet.Cells(56, 5) = Format(l_dbl_PatEfe * 4, "###,###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 2) = Format(l_dbl_salMay / r_dbl_TipCam, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(61, 2) = Format(l_dbl_SalMen / r_dbl_TipCam, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 3) = Format(l_dbl_salMay, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(61, 3) = Format(l_dbl_SalMen, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(63, 2) = Format(l_dbl_SalDol, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(63, 3) = Format(l_dbl_SalSol, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(64, 3) = Format(l_dbl_salMay + l_dbl_SalMen + l_dbl_SalSol, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 4) = Format(l_dbl_salMay / l_dbl_PatEfe, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(61, 4) = Format(l_dbl_SalMen / l_dbl_PatEfe, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(63, 4) = Format(l_dbl_SalSol / l_dbl_PatEfe, "###,###,##0.00")
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc2()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Conta      As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      l_str_PerAno = Format(ipp_PerAno.Text, "0000")
      l_str_PerMes = Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00")
      
      .Cells(2, 3) = "CONTROL DE LIMITES GLOBALES E INDIVIDUALES APLICABLES A LAS EMPRESAS DEL SISTEMA FINANCIERO"
      .Cells(3, 6) = "REPORTE 13"
      .Range(.Cells(2, 3), .Cells(2, 3)).Font.Bold = True
      .Range(.Cells(3, 6), .Cells(3, 6)).Font.Bold = True
      .Cells(6, 1) = "Al: "
      .Cells(6, 2) = fp_UltimoDia(l_str_PerMes, l_str_PerAno)
      .Cells(6, 8) = "Tipo de Cambio"
      
      'Prepara SP
      g_str_Parame = "USP_CUR_GEN_EEBG ("
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", 1)  "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      'Obtiene tipo de cambio
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT DISTINCT HIPCIE_TIPCAM FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_PERMES = " & l_str_PerMes & ""
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(6, 11) = Format(g_rst_Listas!HIPCIE_TIPCAM, "##0.##0")
      End If
      
      .Cells(8, 1) = "Capital y Reserva Legal : "
      .Cells(8, 8) = "Patrimonio Efectivo: (un mes anterior)"
      .Cells(10, 1) = "Limite Global: 7% del Patrimonio Efectivo:"
      .Cells(11, 1) = "Limite Individual: 7% del Patrimonio Efectivo del 5%:"
      
      'Obtiene Patrimonio efectivo
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CONLIM_CAPRES, CONLIM_PATEFE FROM CTB_CONLIM "
      g_str_Parame = g_str_Parame & " WHERE CONLIM_CODANO = " & l_str_PerAno & " AND CONLIM_CODMES = " & l_str_PerMes & ""
      g_str_Parame = g_str_Parame & " ORDER BY CONLIM_CODANO, CONLIM_CODMES ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(8, 5) = g_rst_Listas!CONLIM_CAPRES
         .Cells(8, 11) = Format(g_rst_Listas!CONLIM_PATEFE, "###,###,##0.00")
         .Cells(10, 5) = Format(g_rst_Listas!CONLIM_CAPRES * 0.07, "###,###,##0.00")
         .Cells(11, 5) = Format(g_rst_Listas!CONLIM_CAPRES * 0.07 * 0.05, "###,###,##0.00")
      End If
               
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      .Columns("A").ColumnWidth = 10
      .Columns("B").ColumnWidth = 10
      .Columns("C").ColumnWidth = 12
      .Columns("D").ColumnWidth = 12
      .Columns("E").ColumnWidth = 12
      .Columns("F").ColumnWidth = 12
      .Columns("G").ColumnWidth = 12
      .Columns("H").ColumnWidth = 12
      .Columns("I").ColumnWidth = 12
      .Columns("J").ColumnWidth = 12
      .Columns("K").ColumnWidth = 12
      
      '*******************
      'CREDITOS DIRECTORES
      .Cells(14, 1) = "CREDITOS A DIRECTORES Y TRABAJADORES DE LA EMPRESA:"
      .Range(.Cells(14, 1), .Cells(14, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(14, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(14, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(14, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(14, 11)).Font.Bold = True
            
      .Cells(15, 1) = "Apellidos y Nombres"
      .Cells(15, 5) = "Credito"
      .Cells(15, 6) = "Credito"
      .Cells(15, 7) = "TEA"
      .Cells(15, 8) = "Limite"
      .Cells(15, 9) = "Control"
      .Cells(15, 10) = "Limite"
      .Cells(15, 11) = "Control"
      .Range(.Cells(15, 1), .Cells(15, 11)).Font.Bold = True

      .Range("E15:E15").HorizontalAlignment = xlHAlignCenter
      .Range("E16:E16").HorizontalAlignment = xlHAlignCenter
      .Range("F15:F15").HorizontalAlignment = xlHAlignCenter
      .Range("F16:F16").HorizontalAlignment = xlHAlignCenter
      .Range("G15:G15").HorizontalAlignment = xlHAlignCenter
      .Range("G16:G16").HorizontalAlignment = xlHAlignCenter
      .Range("H15:H15").HorizontalAlignment = xlHAlignCenter
      .Range("H16:H16").HorizontalAlignment = xlHAlignCenter
      .Range("I15:I15").HorizontalAlignment = xlHAlignCenter
      .Range("I16:I16").HorizontalAlignment = xlHAlignCenter
      .Range("J15:J15").HorizontalAlignment = xlHAlignCenter
      .Range("J16:J16").HorizontalAlignment = xlHAlignCenter
      .Range("K15:K15").HorizontalAlignment = xlHAlignCenter
      
      .Cells(16, 5) = "Dolares"
      .Cells(16, 6) = "Soles"
      .Cells(16, 7) = "%"
      .Cells(16, 8) = "Individual %"
      .Cells(16, 10) = "Global %"
      .Cells(17, 1) = "Cta.Cble. 151709010103 Prestamos Administrativos"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT MES" & l_str_PerMes & " FROM TT_EEBG WHERE CNTACTBLE = '151709010103'"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(17, 6) = g_rst_Listas.Fields(0)
      End If
      
      .Cells(17, 10) = "7" & "%"
      .Cells(17, 11) = Format((.Cells(17, 6) / .Cells(8, 11)) * 100, "##0.00")
      .Cells(18, 1) = "Total"
      .Range(.Cells(18, 1), .Cells(18, 11)).Font.Bold = True
      
      '*********************
      'CREDITOS TRABAJADORES
      .Cells(21, 1) = "MAYOR FINANCIAMIENTO A TRABAJADOR, 3 MAYORES:"
      .Range(.Cells(21, 1), .Cells(21, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(21, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(21, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(21, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(21, 11)).Font.Bold = True
      
      .Cells(22, 1) = "Apellidos y Nombres"
      .Cells(22, 5) = "Credito"
      .Cells(22, 6) = "Credito"
      .Cells(22, 7) = "TEA"
      .Cells(22, 8) = "Limite"
      .Cells(22, 9) = "Control"
      .Cells(23, 5) = "Dolares"
      .Cells(23, 6) = "Soles"
      .Cells(23, 8) = "Individual %"
      .Range(.Cells(22, 1), .Cells(22, 11)).Font.Bold = True
      .Range(.Cells(23, 1), .Cells(23, 11)).Font.Bold = True
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT SUBSTR(TRIM(PARVAL_DESCRI),1,6) AS PERIODO,"
      g_str_Parame = g_str_Parame + "       TRIM(SUBSTR(PARVAL_DESCRI,10,100)) AS NOMBRE, PARVAL_CANTID AS MONTO"
      g_str_Parame = g_str_Parame + "  FROM MNT_PARVAL"
      g_str_Parame = g_str_Parame + " WHERE PARVAL_CODGRP = 600 AND PARVAL_CODITE <> '000'"
      g_str_Parame = g_str_Parame + "       AND SUBSTR(PARVAL_DESCRI,1,4)='" & l_str_PerAno & "' AND SUBSTR(PARVAL_DESCRI,5,2)='" & l_str_PerMes & "'"
      g_str_Parame = g_str_Parame + " ORDER BY PARVAL_DESCRI"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            .Cells(24 + r_int_Conta, 1) = g_rst_Listas!NOMBRE
            .Cells(24 + r_int_Conta, 6) = g_rst_Listas!MONTO
                       
            r_int_Conta = r_int_Conta + 1
            g_rst_Listas.MoveNext
         
            .Cells(24, 8) = "0.35"
            .Cells(25, 8) = "0.35"
            .Cells(26, 8) = "0.35"
            .Cells(24, 9) = Format((.Cells(24, 6) / .Cells(8, 11)) * 100, "##0.00")
            .Cells(25, 9) = Format((.Cells(25, 6) / .Cells(8, 11)) * 100, "##0.00")
            .Cells(26, 9) = Format((.Cells(26, 6) / .Cells(8, 11)) * 100, "##0.00")
         Loop
      End If
      
      .Range("E22:E22").HorizontalAlignment = xlHAlignCenter
      .Range("E23:E23").HorizontalAlignment = xlHAlignCenter
      .Range("F22:F22").HorizontalAlignment = xlHAlignCenter
      .Range("F23:F23").HorizontalAlignment = xlHAlignCenter
      .Range("G22:G22").HorizontalAlignment = xlHAlignCenter
      .Range("H22:H22").HorizontalAlignment = xlHAlignCenter
      .Range("H23:H23").HorizontalAlignment = xlHAlignCenter
      .Range("I22:I22").HorizontalAlignment = xlHAlignCenter
      
      '***********************
      'FINANCIAMIENTO CLIENTES
      .Cells(30, 1) = "FINANCIAMIENTO CON O SIN GARANTIA, 3 MAYORES PERSONAS JURIDICAS Y PERSONAS NATURALES:"
      .Range(.Cells(30, 1), .Cells(30, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(30, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(30, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(30, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(30, 11)).Font.Bold = True
      .Cells(31, 1) = "Para créditos Sin Garantia: 10% del Patrimonio Efectivo->"
      .Cells(32, 1) = "Para créditos Con Hipoteca: 15% del Patrimonio Efectivo->"
      .Cells(31, 5) = .Cells(8, 5) * 0.1
      .Cells(32, 5) = .Cells(8, 5) * 0.15

      .Cells(34, 1) = "Apellidos y Nombres"
      .Cells(34, 5) = "Saldo"
      .Cells(35, 5) = "Capital US$"
      .Cells(34, 6) = "Saldo"
      .Cells(35, 6) = "Capital S/."
      .Cells(34, 7) = "Garantia"
      .Cells(34, 8) = "Disponible"
      .Cells(34, 9) = "Limite"
      .Cells(34, 10) = "Control"
      .Range(.Cells(34, 1), .Cells(34, 11)).Font.Bold = True
      .Range(.Cells(35, 1), .Cells(35, 11)).Font.Bold = True
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT A.COMCIE_NUMOPE AS OPERACION, TRIM(B.DATGEN_RAZSOC) AS EMPRESA,"
      g_str_Parame = g_str_Parame + "       ROUND((A.COMCIE_SALCAP + COMCIE_LINCRE) / A.COMCIE_TIPCAM, 2) AS MONTO_DOLARES, "
      g_str_Parame = g_str_Parame + "       A.COMCIE_SALCAP + A.COMCIE_LINCRE AS MONTO_SOLES"
      g_str_Parame = g_str_Parame + "  FROM CRE_COMCIE A INNER JOIN EMP_DATGEN B ON B.DATGEN_EMPTDO=A.COMCIE_TDOCLI"
      g_str_Parame = g_str_Parame + "       AND B.DATGEN_EMPNDO=A.COMCIE_NDOCLI"
      g_str_Parame = g_str_Parame + " WHERE A.COMCIE_PERANO=" & l_str_PerAno & " AND A.COMCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame + "       AND A.COMCIE_TIPGAR = 1"
      g_str_Parame = g_str_Parame + " ORDER BY 3 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      r_int_Conta = 0
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            .Cells(36 + r_int_Conta, 1) = r_int_Conta + 1 & "." & g_rst_Listas!EMPRESA
            .Cells(36 + r_int_Conta, 5) = g_rst_Listas!MONTO_DOLARES
            .Cells(36 + r_int_Conta, 6) = g_rst_Listas!MONTO_SOLES
            .Cells(36 + r_int_Conta, 8) = Format(.Cells(32, 5) - .Cells(36 + r_int_Conta, 6), "###,###,##0.00")
            .Cells(36 + r_int_Conta, 9) = Format((.Cells(36 + r_int_Conta, 6) / .Cells(8, 11)) * 100, "##0.00")
            
            r_int_Conta = r_int_Conta + 1
            g_rst_Listas.MoveNext
         Loop
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_NUMOPE AS OPERACION, HIPCIE_TIPGAR AS TIPO_GARANTIA, "
      g_str_Parame = g_str_Parame + "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
      g_str_Parame = g_str_Parame + "       ROUND(DECODE(HIPCIE_TIPMON , 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)/HIPCIE_TIPCAM, 2) AS MONTO_DOLARES,"
      g_str_Parame = g_str_Parame + "       DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP + HIPCIE_SALCON, (HIPCIE_SALCAP + HIPCIE_SALCON) * HIPCIE_TIPCAM) AS MONTO_SOLES"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & ""
      g_str_Parame = g_str_Parame + " ORDER BY DECODE(HIPCIE_TIPMON , 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM) DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      r_int_Conta = 0
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            .Cells(40 + r_int_Conta, 1) = r_int_Conta + 1 & "." & g_rst_Listas!NOMBRE_CLIENTE
            .Cells(40 + r_int_Conta, 5) = g_rst_Listas!MONTO_DOLARES
            .Cells(40 + r_int_Conta, 6) = g_rst_Listas!MONTO_SOLES
            .Cells(40 + r_int_Conta, 8) = Format(.Cells(32, 5) - .Cells(40 + r_int_Conta, 6), "###,###,##0.00")
            .Cells(40 + r_int_Conta, 9) = Format((.Cells(40 + r_int_Conta, 6) / .Cells(8, 11)) * 100, "##0.00")
                       
            r_int_Conta = r_int_Conta + 1
            If r_int_Conta = 3 Then Exit Do
            g_rst_Listas.MoveNext
         Loop
      End If
      
      .Cells(36, 7) = "Si"
      .Cells(37, 7) = "Si"
      .Cells(38, 7) = "Si"
      .Cells(40, 7) = "Si"
      .Cells(41, 7) = "Si"
      .Cells(42, 7) = "Si"
      .Range("G36:G36").HorizontalAlignment = xlHAlignCenter
      .Range("G37:G37").HorizontalAlignment = xlHAlignCenter
      .Range("G38:G38").HorizontalAlignment = xlHAlignCenter
      .Range("G40:G40").HorizontalAlignment = xlHAlignCenter
      .Range("G41:G41").HorizontalAlignment = xlHAlignCenter
      .Range("G42:G42").HorizontalAlignment = xlHAlignCenter
      
      .Cells(36, 10) = "Si Procede"
      .Cells(37, 10) = "Si Procede"
      .Cells(38, 10) = "Si Procede"
      .Cells(40, 10) = "Si Procede"
      .Cells(41, 10) = "Si Procede"
      .Cells(42, 10) = "Si Procede"
      .Range("E34:E34").HorizontalAlignment = xlHAlignCenter
      .Range("E35:E35").HorizontalAlignment = xlHAlignCenter
      .Range("F34:F34").HorizontalAlignment = xlHAlignCenter
      .Range("F35:F35").HorizontalAlignment = xlHAlignCenter
      .Range("G34:G34").HorizontalAlignment = xlHAlignCenter
      .Range("H34:H34").HorizontalAlignment = xlHAlignCenter
      .Range("I34:I34").HorizontalAlignment = xlHAlignCenter
      .Range("J34:J34").HorizontalAlignment = xlHAlignCenter
      .Cells(44, 1) = "*En este reporte se ordena la columna 'Saldo Capital' de mayor a menor."
      
      '***************************
      'INVERSION MUEBLES INMUEBLES
      .Cells(47, 1) = "INVERSION EN MUEBLES E INMUEBLES:"
      .Range(.Cells(47, 1), .Cells(47, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(47, 1), .Cells(47, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(47, 1), .Cells(47, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(47, 1), .Cells(47, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(47, 1), .Cells(47, 11)).Font.Bold = True

      .Cells(48, 1) = "Descripcion"
      .Cells(48, 5) = "Cta.Cble."
      .Cells(48, 6) = "S/."
      .Cells(48, 7) = "Ratio"
      .Range(.Cells(48, 1), .Cells(48, 11)).Font.Bold = True
      
      .Cells(49, 1) = "Mobiliario"
      .Cells(50, 1) = "Equipos de computación"
      .Cells(51, 1) = "Vehículos"
      .Cells(52, 1) = "Otros bienes y equipos de oficina"
      .Cells(53, 1) = "Instalaciones de bienes alquilados"
      .Cells(54, 1) = "(Deprec.acum.de mobiliario)"
      .Cells(55, 1) = "(Deprec.acum.de equipos de computo)"
      .Cells(56, 1) = "(Deprec.acum.otros bienes y equipo)"
      .Cells(57, 1) = "(Deprec.acum.otros vehículos)"
      .Cells(58, 1) = "(Amortiz.acum.inst. y mejoras)"
      .Cells(59, 5) = "Total"
      .Range(.Cells(59, 5), .Cells(59, 6)).Font.Bold = True
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CNTACTBLE, MES" & l_str_PerMes & " FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('181301010101','181302010101','181309010101',"
      g_str_Parame = g_str_Parame & "                     '181701010101','181903010101','181903010102',"
      g_str_Parame = g_str_Parame & "                     '181903010103','181907010101','181401010101','181904010101')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            Select Case g_rst_Listas!CNTACTBLE
               Case "181301010101": .Cells(49, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(49, 6) = g_rst_Listas.Fields(1)
               Case "181302010101": .Cells(50, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(50, 6) = g_rst_Listas.Fields(1)
               Case "181401010101": .Cells(51, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(51, 6) = g_rst_Listas.Fields(1)
               Case "181309010101": .Cells(52, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(52, 6) = g_rst_Listas.Fields(1)
               Case "181701010101": .Cells(53, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(53, 6) = g_rst_Listas.Fields(1)
               Case "181903010101": .Cells(54, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(54, 6) = g_rst_Listas.Fields(1)
               Case "181903010102": .Cells(55, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(55, 6) = g_rst_Listas.Fields(1)
               Case "181903010103": .Cells(56, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(56, 6) = g_rst_Listas.Fields(1)
               Case "181904010101": .Cells(57, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(57, 6) = g_rst_Listas.Fields(1)
               Case "181907010101": .Cells(58, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(58, 6) = g_rst_Listas.Fields(1)
            End Select
         g_rst_Listas.MoveNext
         Loop
      End If
      
      .Cells(59, 6).Formula = "=SUM(F49:F58)"
      .Cells(49, 7) = Format((.Cells(59, 6) / .Cells(8, 11)) * 100, "##0.00")
      .Range("E48:E48").HorizontalAlignment = xlHAlignCenter
      .Range("F48:F48").HorizontalAlignment = xlHAlignCenter
      .Range("G48:G48").HorizontalAlignment = xlHAlignCenter
      
      '*****************
      'INVERSION MUEBLES
      .Cells(62, 1) = "INVERSION EN INMUEBLES:"
      .Range(.Cells(62, 1), .Cells(62, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(62, 1), .Cells(62, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(62, 1), .Cells(62, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(62, 1), .Cells(62, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(62, 1), .Cells(62, 11)).Font.Bold = True
      
      .Cells(63, 1) = "Descripcion"
      .Cells(63, 5) = "Cta.Cble."
      .Cells(63, 6) = "S/."
      .Cells(63, 7) = "Ratio"
      .Range(.Cells(63, 1), .Cells(63, 11)).Font.Bold = True
      
      .Cells(64, 1) = "Instalaciones en bienes alquilados"
      .Cells(65, 1) = "(Amortiz.acum.inst. y mejoras)"
      .Cells(66, 5) = "Total"
      .Range(.Cells(66, 5), .Cells(66, 6)).Font.Bold = True
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CNTACTBLE, MES" & l_str_PerMes & " FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('181701010101','181907010101')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            Select Case g_rst_Listas!CNTACTBLE
               Case "181701010101": .Cells(64, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(64, 6) = g_rst_Listas.Fields(1)
               Case "181907010101": .Cells(65, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(65, 6) = g_rst_Listas.Fields(1)
            End Select
         g_rst_Listas.MoveNext
         Loop
      End If
      
      .Cells(66, 6).Formula = "=SUM(F64:F65)"
      .Cells(64, 7) = Format((.Cells(66, 6) / .Cells(8, 11)) * 100, "##0.00")
      .Range("E63:E63").HorizontalAlignment = xlHAlignCenter
      .Range("F63:F63").HorizontalAlignment = xlHAlignCenter
      .Range("G63:G63").HorizontalAlignment = xlHAlignCenter
      
      '*****************
      'INVERSION MUEBLES
      .Cells(69, 1) = "INVERSION EN MUEBLES:"
      .Range(.Cells(69, 1), .Cells(69, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(69, 1), .Cells(69, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(69, 1), .Cells(69, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(69, 1), .Cells(69, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(69, 1), .Cells(69, 11)).Font.Bold = True

      .Cells(70, 1) = "Descripcion"
      .Cells(70, 5) = "Cta.Cble."
      .Cells(70, 6) = "S/."
      .Cells(70, 7) = "Ratio"
      .Range(.Cells(70, 1), .Cells(70, 11)).Font.Bold = True
      
      .Cells(71, 1) = "Mobiliario"
      .Cells(72, 1) = "Equipos de computacion"
      .Cells(73, 1) = "Otros bienes y equipos de oficina"
      .Cells(74, 1) = "Vehículos"
      .Cells(75, 1) = "(Deprec.acum.de mobiliario)"
      .Cells(76, 1) = "(Deprec.acum.de equipos de computo)"
      .Cells(77, 1) = "(Deprec.acum.otros bienes y equipos)"
      .Cells(78, 1) = "(Deprec.acum.otros vehículos)"
      .Cells(79, 5) = "Total"
      .Range(.Cells(79, 5), .Cells(79, 6)).Font.Bold = True
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CNTACTBLE, MES" & l_str_PerMes & " FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('181301010101','181302010101','181309010101',"
      g_str_Parame = g_str_Parame & "                     '181903010101','181903010102','181903010103','181401010101','181904010101')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         Do While Not g_rst_Listas.EOF
            Select Case g_rst_Listas!CNTACTBLE
               Case "181301010101": .Cells(71, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(71, 6) = g_rst_Listas.Fields(1)
               Case "181302010101": .Cells(72, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(72, 6) = g_rst_Listas.Fields(1)
               Case "181309010101": .Cells(73, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(73, 6) = g_rst_Listas.Fields(1)
               Case "181903010101": .Cells(74, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(74, 6) = g_rst_Listas.Fields(1)
               Case "181401010101": .Cells(75, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(75, 6) = g_rst_Listas.Fields(1)
               Case "181903010102": .Cells(76, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(76, 6) = g_rst_Listas.Fields(1)
               Case "181903010103": .Cells(77, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(77, 6) = g_rst_Listas.Fields(1)
               Case "181904010101": .Cells(78, 5) = "'" & g_rst_Listas!CNTACTBLE: .Cells(78, 6) = g_rst_Listas.Fields(1)
            End Select
         g_rst_Listas.MoveNext
         Loop
      End If
      
      .Cells(79, 6).Formula = "=SUM(F71:F78)"
      .Cells(71, 7) = Format((.Cells(79, 6) / .Cells(8, 11)) * 100, "##0.00")
      .Range("E70:E70").HorizontalAlignment = xlHAlignCenter
      .Range("F70:F70").HorizontalAlignment = xlHAlignCenter
      .Range("G70:G70").HorizontalAlignment = xlHAlignCenter
      
      '*********
      'DEPOSITOS
      .Cells(82, 1) = "DEPOSITOS EN BANCOS, 3 MAYORES"
      .Range(.Cells(82, 1), .Cells(82, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(82, 1), .Cells(82, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(82, 1), .Cells(82, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(82, 1), .Cells(82, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(82, 1), .Cells(82, 11)).Font.Bold = True

      .Cells(83, 1) = "Descripcion"
      .Cells(83, 5) = "Cta.Cble."
      .Cells(83, 6) = "S/."
      .Cells(83, 7) = "Ratio"
      .Range(.Cells(83, 1), .Cells(83, 11)).Font.Bold = True
      
      .Cells(84, 1) = "1.BBVA"
      .Cells(85, 1) = "2.HSBC"
      .Cells(86, 1) = "3.CREDITO"
      .Cells(87, 1) = "4.INTERBANK"
      .Cells(84, 5) = "'" & "11030106"
      .Cells(85, 5) = "'" & "11070932"
      .Cells(86, 5) = "'" & "11030103"
      .Cells(87, 5) = "'" & "11030104"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT SUM(MES" & l_str_PerMes & ") AS BBVA FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('111301060102','111301060201','112301060102','112301060202')"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(84, 6) = Format(g_rst_Listas!BBVA, "###,###,##0.00")
      End If
      .Cells(84, 7) = (.Cells(84, 6) / .Cells(8, 11)) * 100
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT SUM(MES" & l_str_PerMes & ") AS HSBC FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('111709320101','112709320101')"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(85, 6) = Format(g_rst_Listas!HSBC, "###,###,##0.00")
      End If
      .Cells(85, 7) = (.Cells(85, 6) / .Cells(8, 11)) * 100

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT SUM(MES" & l_str_PerMes & ") AS BCP FROM TT_EEBG "
      g_str_Parame = g_str_Parame & " WHERE CNTACTBLE IN ('111301030101','112301030101')"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(86, 6) = Format(g_rst_Listas!BCP, "###,###,##0.00")
      End If
      .Cells(86, 7) = (.Cells(86, 6) / .Cells(8, 11)) * 100
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT SUM(MES" & l_str_PerMes & ") AS IBK FROM TT_EEBG "
      g_str_Parame = g_str_Parame + " WHERE CNTACTBLE IN ('111301040201')"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         .Cells(87, 6) = Format(g_rst_Listas!IBK, "###,###,##0.00")
      End If
      .Cells(87, 7) = (.Cells(87, 6) / .Cells(8, 11)) * 100
      
      .Range("A1:K87").Font.Name = "Calibri"
      .Range("A1:K87").Font.Size = 10
      .Range("E81:E83").HorizontalAlignment = xlHAlignCenter
      .Range("F81:F83").HorizontalAlignment = xlHAlignCenter
      .Range("G81:G83").HorizontalAlignment = xlHAlignCenter
   End With
  
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function ff_Query(ByVal p_PerMes As String, ByVal p_PerAno As String) As String
   ff_Query = ""
   ff_Query = ff_Query & "SELECT * FROM "
   ff_Query = ff_Query & "(SELECT * FROM (SELECT HIPCIE_TDOCLI AS TIPDOC, HIPCIE_NDOCLI AS NUMDOC, HIPCIE_TIPGAR AS TIPGAR, TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE, ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS MONTO "
   ff_Query = ff_Query & "                  FROM CRE_HIPCIE "
   ff_Query = ff_Query & "                 INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI "
   ff_Query = ff_Query & "                 WHERE HIPCIE_PERMES = " & p_PerMes & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_TIPGAR IN (1,2) "
   ff_Query = ff_Query & "                 ORDER BY MONTO DESC) "
   ff_Query = ff_Query & "  WHERE ROWNUM < 4 "
   ff_Query = ff_Query & " UNION "
   ff_Query = ff_Query & " SELECT * FROM (SELECT HIPCIE_TDOCLI AS TIPDOC, HIPCIE_NDOCLI AS NUMDOC, HIPCIE_TIPGAR AS TIPGAR, TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE, ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS MONTO "
   ff_Query = ff_Query & "                  FROM CRE_HIPCIE "
   ff_Query = ff_Query & "                 INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI "
   ff_Query = ff_Query & "                 WHERE HIPCIE_PERMES = " & p_PerMes & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_TIPGAR NOT IN (1,2) "
   ff_Query = ff_Query & "                 ORDER BY MONTO DESC) "
   ff_Query = ff_Query & "  WHERE ROWNUM < 4 "
   ff_Query = ff_Query & " UNION "
   ff_Query = ff_Query & " SELECT * FROM (SELECT COMCIE_TDOCLI AS TIPDOC, COMCIE_NDOCLI AS NUMDOC, COMCIE_TIPGAR AS TIPGAR, TRIM(DATGEN_NOMCOM) AS NOMBRE, SUM(COMCIE_SALCAP+COMCIE_LINCRE) AS MONTO FROM CRE_COMCIE "
   ff_Query = ff_Query & "                 INNER JOIN EMP_DATGEN ON DATGEN_EMPTDO = COMCIE_TDOCLI AND DATGEN_EMPNDO = COMCIE_NDOCLI "
   ff_Query = ff_Query & "                 WHERE COMCIE_PERMES = " & p_PerMes & " AND COMCIE_PERANO = " & p_PerAno & " "
   ff_Query = ff_Query & "                 GROUP BY COMCIE_TDOCLI, COMCIE_NDOCLI, COMCIE_TIPGAR, DATGEN_NOMCOM ORDER BY MONTO DESC) "
   ff_Query = ff_Query & "  WHERE ROWNUM < 4) "
   ff_Query = ff_Query & " ORDER BY MONTO DESC "
End Function

Private Sub ff_SalCap_01()
   g_str_Parame = ""
   g_str_Parame = "SELECT LIMGLO_SALMAY, LIMGLO_SALMEN FROM TMP_LIMGLO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         If g_rst_Listas!LIMGLO_SALMAY <> " " Then
            l_dbl_salMay = l_dbl_salMay + g_rst_Listas!LIMGLO_SALMAY
            l_dbl_SalMen = l_dbl_SalMen + g_rst_Listas!LIMGLO_SALMEN
         End If
         g_rst_Listas.MoveNext
      Loop
   End If

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_SalCap_02()
   g_str_Parame = ""
   g_str_Parame = "SELECT LIMGLO_SALDOL, LIMGLO_SALSOL FROM TMP_LIMGLO WHERE LIMGLO_TDOCLI = 7 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         l_dbl_SalSol = l_dbl_SalSol + g_rst_Listas!LIMGLO_SALSOL
         l_dbl_SalDol = l_dbl_SalDol + g_rst_Listas!LIMGLO_SALDOL
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

'Private Function ff_TncMen_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
'   ff_TncMen_01 = 0
'
'   g_str_Parame = "select * from cre_hipcuo where "
'   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
'   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
'   g_str_Parame = g_str_Parame & "hipcuo_fecvct <= " & l_int_FecVct & " and "
'   g_str_Parame = g_str_Parame & "hipcuo_situac = 2 "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'       Exit Function
'   End If
'
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'
'      Do While Not g_rst_Listas.EOF
'         ff_TncMen_01 = ff_TncMen_01 + g_rst_Listas!HIPCUO_CAPITA
'         g_rst_Listas.MoveNext
'      Loop
'   End If
'
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function

'Private Function ff_FecVct(ByVal p_NumSol As String, ByVal l_int_NumCuo As Integer) As Double
'   ff_FecVct = 0
'
'   g_str_Parame = "select * from cre_hipcuo where "
'   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
'   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
'   g_str_Parame = g_str_Parame & "hipcuo_numcuo = " & l_int_NumCuo + 12
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'       Exit Function
'   End If
'
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'
'      Do While Not g_rst_Listas.EOF
'         ff_FecVct = g_rst_Listas!HIPCUO_FECVCT
'         g_rst_Listas.MoveNext
'      Loop
'   End If
'
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function

'Private Function ff_TncMay_01(ByVal p_NumSol As String, ByVal l_int_NumCuo As Integer) As Double
'   ff_TncMay_01 = 0
'
'   g_str_Parame = "select * from cre_hipcuo where "
'   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
'   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
'   g_str_Parame = g_str_Parame & "hipcuo_numcuo > " & l_int_NumCuo + 12
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'       Exit Function
'   End If
'
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'
'      Do While Not g_rst_Listas.EOF
'         ff_TncMay_01 = ff_TncMay_01 + g_rst_Listas!HIPCUO_CAPITA
'         g_rst_Listas.MoveNext
'      Loop
'   End If
'
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function

'Private Function ff_TcMen_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
'   ff_TcMen_01 = 0
'
'   g_str_Parame = "select * from cre_hipcuo where "
'   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
'
'   'If p_NumSol = "0040700001" Or p_NumSol = "0040700002" Or p_NumSol = "0040700003" Or p_NumSol = "0040700004" Then
'   '   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 4 and "
'   'Else
'      g_str_Parame = g_str_Parame & "hipcuo_tipcro = 2 and "
'   'End If
'
'   g_str_Parame = g_str_Parame & "hipcuo_fecvct <= " & l_int_FecVct & " and "
'   g_str_Parame = g_str_Parame & "hipcuo_situac = 2"
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'       Exit Function
'   End If
'
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'
'      Do While Not g_rst_Listas.EOF
'         ff_TcMen_01 = ff_TcMen_01 + g_rst_Listas!HIPCUO_CAPITA
'         g_rst_Listas.MoveNext
'      Loop
'   End If
'
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function

'Private Function ff_TcMay_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
'   ff_TcMay_01 = 0
'
'   g_str_Parame = "select * from cre_hipcuo where "
'   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
'
'   'If p_NumSol = "0040700001" Or p_NumSol = "0040700002" Or p_NumSol = "0040700003" Or p_NumSol = "0040700004" Then
'   '   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 4 and "
'   'Else
'      g_str_Parame = g_str_Parame & "hipcuo_tipcro = 2 and "
'   'End If
'
'   g_str_Parame = g_str_Parame & "hipcuo_fecvct > " & l_int_FecVct
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
'       Exit Function
'   End If
'
'   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
'      g_rst_Listas.MoveFirst
'
'      Do While Not g_rst_Listas.EOF
'         ff_TcMay_01 = ff_TcMay_01 + g_rst_Listas!HIPCUO_CAPITA
'         g_rst_Listas.MoveNext
'      Loop
'   End If
'
'   g_rst_Listas.Close
'   Set g_rst_Listas = Nothing
'End Function
