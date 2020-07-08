VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   8070
   ClientTop       =   2775
   ClientWidth     =   4440
   Icon            =   "GesCtb_frm_814.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4485
      _Version        =   65536
      _ExtentX        =   7911
      _ExtentY        =   4154
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
         Left            =   30
         TabIndex        =   6
         Top             =   60
         Width           =   4365
         _Version        =   65536
         _ExtentX        =   7699
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
            Height          =   300
            Left            =   570
            TabIndex        =   7
            Top             =   300
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Créditos Según Dias de Incumplimiento"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   570
            TabIndex        =   12
            Top             =   30
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte N° 14"
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
            Picture         =   "GesCtb_frm_814.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   4365
         _Version        =   65536
         _ExtentX        =   7699
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
            Left            =   30
            Picture         =   "GesCtb_frm_814.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_814.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   3750
            Picture         =   "GesCtb_frm_814.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   4365
         _Version        =   65536
         _ExtentX        =   7699
         _ExtentY        =   1455
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
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
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_FecIni        As String
Dim r_str_FecFin        As String

Private Sub cmd_ExpArc_Click()
   
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
   
   'Confirmación
   If MsgBox("¿Está seguro de Generar Archivo de Texto?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   Call fs_GenArc(r_str_FecIni, r_str_FecFin)
   
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
   
   Screen.MousePointer = 0
   
End Sub

Private Sub cmd_ExpExc_Click()
      
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
   
   'Confirmación
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   Call fs_GenExc(r_str_FecIni, r_str_FecFin)
   
End Sub

Private Sub fs_Inicia()
         
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
End Sub

Private Sub ipp_Ano_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
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

Private Sub fs_GenExc(ByVal p_FecIni As String, ByVal p_FecFin As String)

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer
   
   Dim r_str_PerMes     As String
   Dim r_str_TipMon     As String
   
   'Creditos Hipotecarios
   Dim r_dbl_CaVeHp     As Double
   Dim r_dbl_CaVgHp     As Double
   
   Dim r_dbl_VeHp01     As Double
   Dim r_dbl_VeHp02     As Double
   Dim r_dbl_VeHp03     As Double
   Dim r_dbl_VeHp04     As Double
   Dim r_dbl_VeHp05     As Double
   Dim r_dbl_VeHp06     As Double
   Dim r_dbl_VeHp07     As Double
   Dim r_dbl_VeHp08     As Double
   Dim r_dbl_VeHp09     As Double
   Dim r_dbl_VeHp10     As Double
   Dim r_dbl_VeHp11     As Double
   Dim r_dbl_VeHp12     As Double
   Dim r_dbl_VeHp13     As Double
   Dim r_dbl_VeHp14     As Double
   Dim r_dbl_VeHp15     As Double
   Dim r_dbl_VeHp16     As Double
   Dim r_dbl_SnVgHp     As Double
   Dim r_dbl_SeVgHp     As Double
   
   Dim r_dbl_SlHp01     As Double
   Dim r_dbl_SlHp02     As Double
   Dim r_dbl_SlHp03     As Double
   Dim r_dbl_SlHp04     As Double
   Dim r_dbl_SlHp05     As Double
   Dim r_dbl_SlHp06     As Double
   Dim r_dbl_SlHp07     As Double
   Dim r_dbl_SlHp08     As Double
   Dim r_dbl_SlHp09     As Double
   Dim r_dbl_SlHp10     As Double
   Dim r_dbl_SlHp11     As Double
   Dim r_dbl_SlHp12     As Double
   Dim r_dbl_SlHp13     As Double
   Dim r_dbl_SlHp14     As Double
   Dim r_dbl_SlHp15     As Double
   Dim r_dbl_SlHp16     As Double
         
   'Creditos Comerciales
   Dim r_dbl_CaVeCo     As Double
   Dim r_dbl_CaVgCo     As Double
   
   Dim r_dbl_VeCo01     As Double
   Dim r_dbl_VeCo02     As Double
   Dim r_dbl_VeCo03     As Double
   Dim r_dbl_VeCo04     As Double
   Dim r_dbl_VeCo05     As Double
   Dim r_dbl_VeCo06     As Double
   Dim r_dbl_VeCo07     As Double
   Dim r_dbl_VeCo08     As Double
   Dim r_dbl_VeCo09     As Double
   Dim r_dbl_VeCo10     As Double
   Dim r_dbl_VeCo11     As Double
   Dim r_dbl_VeCo12     As Double
   Dim r_dbl_VeCo13     As Double
   Dim r_dbl_VeCo14     As Double
   Dim r_dbl_VeCo15     As Double
   Dim r_dbl_VeCo16     As Double
   Dim r_dbl_SnVgCo     As Double
   Dim r_dbl_SeVgCo     As Double
   
   Dim r_dbl_SlCo01     As Double
   Dim r_dbl_SlCo02     As Double
   Dim r_dbl_SlCo03     As Double
   Dim r_dbl_SlCo04     As Double
   Dim r_dbl_SlCo05     As Double
   Dim r_dbl_SlCo06     As Double
   Dim r_dbl_SlCo07     As Double
   Dim r_dbl_SlCo08     As Double
   Dim r_dbl_SlCo09     As Double
   Dim r_dbl_SlCo10     As Double
   Dim r_dbl_SlCo11     As Double
   Dim r_dbl_SlCo12     As Double
   Dim r_dbl_SlCo13     As Double
   Dim r_dbl_SlCo14     As Double
   Dim r_dbl_SlCo15     As Double
   Dim r_dbl_SlCo16     As Double
         
         
   'Creditos Comerciales
   g_str_Parame = "SELECT COMCIE_NUMOPE, COMCIE_CAPVIG, COMCIE_CAPVEN, COMCIE_TIPCAM, COMCIE_TIPMON, COMCIE_DIAMOR FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 15
   
   Do While Not g_rst_Princi.EOF
                
      If g_rst_Princi!COMCIE_TIPMON = 1 Then
         If g_rst_Princi!COMCIE_CAPVEN = 0 Then
            r_dbl_SnVgCo = r_dbl_SnVgCo + g_rst_Princi!COMCIE_CAPVIG
         Else
            If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
               r_dbl_VeCo01 = r_dbl_VeCo01 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo01 = r_dbl_SlCo01 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
               r_dbl_VeCo02 = r_dbl_VeCo02 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo02 = r_dbl_SlCo02 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
               r_dbl_VeCo03 = r_dbl_VeCo03 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo03 = r_dbl_SlCo03 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
               r_dbl_VeCo04 = r_dbl_VeCo04 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo04 = r_dbl_SlCo04 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
               r_dbl_VeCo05 = r_dbl_VeCo05 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo05 = r_dbl_SlCo05 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
               r_dbl_VeCo06 = r_dbl_VeCo06 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo06 = r_dbl_SlCo06 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
               r_dbl_VeCo07 = r_dbl_VeCo07 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo07 = r_dbl_SlCo07 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
               r_dbl_VeCo08 = r_dbl_VeCo08 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo08 = r_dbl_SlCo08 + g_rst_Princi!COMCIE_CAPVIG
            End If
         End If
      ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
         If g_rst_Princi!COMCIE_CAPVEN = 0 Then
            r_dbl_SeVgCo = r_dbl_SeVgCo + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
               r_dbl_VeCo09 = r_dbl_VeCo09 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo09 = r_dbl_SlCo09 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
               r_dbl_VeCo10 = r_dbl_VeCo10 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo10 = r_dbl_SlCo10 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
               r_dbl_VeCo11 = r_dbl_VeCo11 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo11 = r_dbl_SlCo11 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
               r_dbl_VeCo12 = r_dbl_VeCo12 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo12 = r_dbl_SlCo12 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
               r_dbl_VeCo13 = r_dbl_VeCo13 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo13 = r_dbl_SlCo13 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
               r_dbl_VeCo14 = r_dbl_VeCo14 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo14 = r_dbl_SlCo14 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
               r_dbl_VeCo15 = r_dbl_VeCo15 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo15 = r_dbl_SlCo15 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
               r_dbl_VeCo16 = r_dbl_VeCo16 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo16 = r_dbl_SlCo16 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            End If
            
         End If
      End If
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
         
         
   'Credito Hipotecario
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_CAPVIG, HIPCIE_CAPVEN, HIPCIE_TIPCAM, HIPCIE_TIPMON, HIPCIE_DIAMOR FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 15
   
   Do While Not g_rst_Princi.EOF
        
        
      If g_rst_Princi!HIPCIE_TIPMON = 1 Then
         If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
            r_dbl_SnVgHp = r_dbl_SnVgHp + g_rst_Princi!HIPCIE_CAPVIG
         Else
            If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
               r_dbl_VeHp01 = r_dbl_VeHp01 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp01 = r_dbl_SlHp01 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
               r_dbl_VeHp02 = r_dbl_VeHp02 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp02 = r_dbl_SlHp02 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
               r_dbl_VeHp03 = r_dbl_VeHp03 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp03 = r_dbl_SlHp03 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
               r_dbl_VeHp04 = r_dbl_VeHp04 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp04 = r_dbl_SlHp04 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
               r_dbl_VeHp05 = r_dbl_VeHp05 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp05 = r_dbl_SlHp05 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
               r_dbl_VeHp06 = r_dbl_VeHp06 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp06 = r_dbl_SlHp06 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
               r_dbl_VeHp07 = r_dbl_VeHp07 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp07 = r_dbl_SlHp07 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
               r_dbl_VeHp08 = r_dbl_VeHp08 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp08 = r_dbl_SlHp08 + g_rst_Princi!HIPCIE_CAPVIG
            End If
         End If
      ElseIf g_rst_Princi!HIPCIE_TIPMON = 2 Then
         If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
            r_dbl_SeVgHp = r_dbl_SeVgHp + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         Else
            If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
               r_dbl_VeHp09 = r_dbl_VeHp09 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp09 = r_dbl_SlHp09 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
               r_dbl_VeHp10 = r_dbl_VeHp10 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp10 = r_dbl_SlHp10 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
               r_dbl_VeHp11 = r_dbl_VeHp11 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp11 = r_dbl_SlHp11 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
               r_dbl_VeHp12 = r_dbl_VeHp12 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp12 = r_dbl_SlHp12 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
               r_dbl_VeHp13 = r_dbl_VeHp13 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp13 = r_dbl_SlHp13 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
               r_dbl_VeHp14 = r_dbl_VeHp14 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp14 = r_dbl_SlHp14 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
               r_dbl_VeHp15 = r_dbl_VeHp15 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp15 = r_dbl_SlHp15 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
               r_dbl_VeHp16 = r_dbl_VeHp16 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp16 = r_dbl_SlHp16 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            End If
            
         End If
      End If
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      
      .Cells(1, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
      .Range(.Cells(1, 2), .Cells(1, 5)).Merge
      .Range(.Cells(1, 2), .Cells(1, 5)).HorizontalAlignment = xlHAlignLeft
            
      .Cells(2, 19).HorizontalAlignment = xlHAlignRight
      .Cells(2, 19) = "REPORTE 14"
             
      .Range(.Cells(5, 2), .Cells(8, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(14, 21)).Font.Bold = True
      .Cells(6, 2).HorizontalAlignment = xlHAlignLeft
            
      .Cells(5, 2) = "CREDITOS SEGUN DIAS DE INCUMPLIMIENTO"
      .Cells(6, 2) = "CODIGO: 240"
      .Cells(7, 2) = "Al " & Right(p_FecFin, 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Left(p_FecFin, 4)
      .Cells(8, 2) = "(En Soles)"
      
      .Range("C10") = "Saldo de Créditos sin atraso"
      .Range("D10") = "Incumplimiento 1/"
      .Range("T10") = "Saldo Total de Créditos 2/"
      
      .Range("B14") = "I MONEDA NACIONAL 3/"
      .Range("B15") = "Comerciales 4/"
      .Range("B16") = "Microempresas 5/"
      .Range("B17") = "Consumo 6/"
      .Range("B18") = "Hipotecarios para Vivienda 7/"
      .Range("B19") = "Sobregiros 8/"
      .Range("B20") = "Arrendamiento Financiero 9/"
      
      .Range("B22") = "II MONEDA EXTRANJERA"
      .Range("B23") = "Comerciales 4/"
      .Range("B24") = "Microempresas 5/"
      .Range("B25") = "Consumo 6/"
      .Range("B26") = "Hipotecarios para Vivienda 7/"
      .Range("B27") = "Sobregiros 8/"
      .Range("B28") = "Arrendamiento Financiero 9/"
      .Range("E39") = "Sr. Roberto Baba Yamamoto"
      .Range("L39") = "Srta. Rossana Meza Bustamente"
      .Range("E40") = "GERENTE GENERAL"
      .Range("L40") = "CONTADOR GENERAL"
      
      .Range("B22:T22").Font.Bold = True
      
      .Range("B30") = "TOTAL (I + II)"
      .Range("B30:T30").Font.Bold = True
      
      .Range(.Cells(5, 2), .Cells(5, 20)).Merge
      .Range(.Cells(7, 2), .Cells(7, 20)).Merge
      .Range(.Cells(8, 2), .Cells(8, 20)).Merge
           
      .Range("B10:B12").Merge
      .Range("C10:C12").Merge
      .Range("D10:S10").Merge
      .Range("T10:T12").Merge
      
      'Firmas
      .Range("E39:G39").Merge
      .Range("L39:N39").Merge
      .Range("L40:N40").Merge
      .Range("E40:G40").Merge
      
      .Range("E39:G39").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("L39:N39").Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(39, 3), .Cells(40, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(39, 3), .Cells(40, 12)).VerticalAlignment = xlVAlignCenter
      
      .Range("D11") = "De 1 a 15 dias"
      .Range("D11:E11").Merge
            
      .Range("F11") = "Entre 16 y 30 dias"
      .Range("F11:G11").Merge
            
      .Range("H11") = "Entre 31 y 60 dias"
      .Range("H11:I11").Merge
            
      .Range("J11") = "Entre 61 y 90 dias"
      .Range("J11:K11").Merge
            
      .Range("L11") = "Entre 91 y 120 dias"
      .Range("L11:M11").Merge
            
      .Range("N11") = "Entre 121 y 180 dias"
      .Range("N11:O11").Merge
            
      .Range("P11") = "Entre 181 y 365 dias"
      .Range("P11:Q11").Merge
            
      .Range("R11") = "Mayor a 365 dias"
      .Range("R11:S11").Merge
      
      
      .Range(.Cells(10, 2), .Cells(14, 20)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(10, 2), .Cells(14, 20)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 2), .Cells(14, 20)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 2), .Cells(14, 20)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 2), .Cells(14, 20)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range("D11:S11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D11:S11").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range("B13:T13").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B13:T13").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range("B22:T22").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B22:T22").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range("B30:T30").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B30:T30").Borders(xlEdgeBottom).LineStyle = xlContinuous

      
            
      For r_int_ConVar = 2 To 20 Step 1
         .Range(.Cells(13, 2), .Cells(30, r_int_ConVar)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, 2), .Cells(30, r_int_ConVar)).Borders(xlEdgeRight).LineStyle = xlContinuous
      Next
      
      For r_int_ConVar = 4 To 18 Step 2
         .Cells(12, r_int_ConVar) = "Porcion no Amortizada"
         .Cells(12, r_int_ConVar + 1) = "Saldo"
      Next
      
      
      .Range(.Cells(10, 3), .Cells(12, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 21)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 21)).VerticalAlignment = xlHAlignFill
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 2
      .Columns("B").ColumnWidth = 39
      .Columns("C").ColumnWidth = 14
      .Columns("T").ColumnWidth = 14
      .Columns("U").ColumnWidth = 2
      
      .Columns("C:T").NumberFormat = "###,###,##0.00"
      
             
   End With
   
   'Moneda Nacional
   'Hipotecario
   r_obj_Excel.ActiveSheet.Cells(18, 3) = r_dbl_SnVgHp
   
   r_obj_Excel.ActiveSheet.Cells(18, 4) = r_dbl_VeHp01
   r_obj_Excel.ActiveSheet.Cells(18, 5) = r_dbl_SlHp01
   r_obj_Excel.ActiveSheet.Cells(18, 6) = r_dbl_VeHp02
   r_obj_Excel.ActiveSheet.Cells(18, 7) = r_dbl_SlHp02
   r_obj_Excel.ActiveSheet.Cells(18, 8) = r_dbl_VeHp03
   r_obj_Excel.ActiveSheet.Cells(18, 9) = r_dbl_SlHp03
   r_obj_Excel.ActiveSheet.Cells(18, 10) = r_dbl_VeHp04
   r_obj_Excel.ActiveSheet.Cells(18, 11) = r_dbl_SlHp04
   r_obj_Excel.ActiveSheet.Cells(18, 12) = r_dbl_VeHp05
   r_obj_Excel.ActiveSheet.Cells(18, 13) = r_dbl_SlHp05
   r_obj_Excel.ActiveSheet.Cells(18, 14) = r_dbl_VeHp06
   r_obj_Excel.ActiveSheet.Cells(18, 15) = r_dbl_SlHp06
   r_obj_Excel.ActiveSheet.Cells(18, 16) = r_dbl_VeHp07
   r_obj_Excel.ActiveSheet.Cells(18, 17) = r_dbl_SlHp07
   r_obj_Excel.ActiveSheet.Cells(18, 18) = r_dbl_VeHp08
   r_obj_Excel.ActiveSheet.Cells(18, 19) = r_dbl_SlHp08
   
   r_obj_Excel.ActiveSheet.Cells(18, 20) = r_dbl_SnVgHp + r_dbl_VeHp01 + r_dbl_VeHp02 + r_dbl_VeHp03 + _
                                           r_dbl_VeHp04 + r_dbl_VeHp05 + r_dbl_VeHp06 + r_dbl_VeHp07 + r_dbl_VeHp08 + _
                                           r_dbl_SlHp01 + r_dbl_SlHp02 + r_dbl_SlHp03 + r_dbl_SlHp04 + r_dbl_SlHp05 + _
                                           r_dbl_SlHp06 + r_dbl_SlHp07 + r_dbl_SlHp08
   
   'Comercial
   r_obj_Excel.ActiveSheet.Cells(15, 3) = r_dbl_SnVgCo
   
   r_obj_Excel.ActiveSheet.Cells(15, 4) = r_dbl_VeCo01
   r_obj_Excel.ActiveSheet.Cells(15, 5) = r_dbl_SlCo01
   r_obj_Excel.ActiveSheet.Cells(15, 6) = r_dbl_VeCo02
   r_obj_Excel.ActiveSheet.Cells(15, 7) = r_dbl_SlCo02
   r_obj_Excel.ActiveSheet.Cells(15, 8) = r_dbl_VeCo03
   r_obj_Excel.ActiveSheet.Cells(15, 9) = r_dbl_SlCo03
   r_obj_Excel.ActiveSheet.Cells(15, 10) = r_dbl_VeCo04
   r_obj_Excel.ActiveSheet.Cells(15, 11) = r_dbl_SlCo04
   r_obj_Excel.ActiveSheet.Cells(15, 12) = r_dbl_VeCo05
   r_obj_Excel.ActiveSheet.Cells(15, 13) = r_dbl_SlCo05
   r_obj_Excel.ActiveSheet.Cells(15, 14) = r_dbl_VeCo06
   r_obj_Excel.ActiveSheet.Cells(15, 15) = r_dbl_SlCo06
   r_obj_Excel.ActiveSheet.Cells(15, 16) = r_dbl_VeCo07
   r_obj_Excel.ActiveSheet.Cells(15, 17) = r_dbl_SlCo07
   r_obj_Excel.ActiveSheet.Cells(15, 18) = r_dbl_VeCo08
   r_obj_Excel.ActiveSheet.Cells(15, 19) = r_dbl_SlCo08
   
   r_obj_Excel.ActiveSheet.Cells(15, 20) = r_dbl_SnVgCo + r_dbl_VeCo01 + r_dbl_VeCo02 + r_dbl_VeCo03 + _
                                           r_dbl_VeCo04 + r_dbl_VeCo05 + r_dbl_VeCo06 + r_dbl_VeCo07 + r_dbl_VeCo08 + _
                                           r_dbl_SlCo01 + r_dbl_SlCo02 + r_dbl_SlCo03 + r_dbl_SlCo04 + r_dbl_SlCo05 + _
                                           r_dbl_SlCo06 + r_dbl_SlCo07 + r_dbl_SlCo08
   
   
   'Moneda Extranjera
   'Hipotecario
   r_obj_Excel.ActiveSheet.Cells(26, 3) = r_dbl_SeVgHp
   
   r_obj_Excel.ActiveSheet.Cells(26, 4) = r_dbl_VeHp09
   r_obj_Excel.ActiveSheet.Cells(26, 5) = r_dbl_SlHp09
   r_obj_Excel.ActiveSheet.Cells(26, 6) = r_dbl_VeHp10
   r_obj_Excel.ActiveSheet.Cells(26, 7) = r_dbl_SlHp10
   r_obj_Excel.ActiveSheet.Cells(26, 8) = r_dbl_VeHp11
   r_obj_Excel.ActiveSheet.Cells(26, 9) = r_dbl_SlHp11
   r_obj_Excel.ActiveSheet.Cells(26, 10) = r_dbl_VeHp12
   r_obj_Excel.ActiveSheet.Cells(26, 11) = r_dbl_SlHp12
   r_obj_Excel.ActiveSheet.Cells(26, 12) = r_dbl_VeHp13
   r_obj_Excel.ActiveSheet.Cells(26, 13) = r_dbl_SlHp13
   r_obj_Excel.ActiveSheet.Cells(26, 14) = r_dbl_VeHp14
   r_obj_Excel.ActiveSheet.Cells(26, 15) = r_dbl_SlHp14
   r_obj_Excel.ActiveSheet.Cells(26, 16) = r_dbl_VeHp15
   r_obj_Excel.ActiveSheet.Cells(26, 17) = r_dbl_SlHp15
   r_obj_Excel.ActiveSheet.Cells(26, 18) = r_dbl_VeHp16
   r_obj_Excel.ActiveSheet.Cells(26, 19) = r_dbl_SlHp16
   
   r_obj_Excel.ActiveSheet.Cells(26, 20) = r_dbl_SeVgHp + r_dbl_VeHp09 + r_dbl_VeHp10 + r_dbl_VeHp11 + _
                                           r_dbl_VeHp12 + r_dbl_VeHp13 + r_dbl_VeHp14 + r_dbl_VeHp15 + r_dbl_VeHp16 + _
                                           r_dbl_SlHp09 + r_dbl_SlHp10 + r_dbl_SlHp11 + r_dbl_SlHp12 + r_dbl_SlHp13 + _
                                           r_dbl_SlHp14 + r_dbl_SlHp15 + r_dbl_SlHp16
   
   
   'Comercial
   r_obj_Excel.ActiveSheet.Cells(23, 3) = r_dbl_SeVgCo
   
   r_obj_Excel.ActiveSheet.Cells(23, 4) = r_dbl_VeCo09
   r_obj_Excel.ActiveSheet.Cells(23, 5) = r_dbl_SlCo09
   r_obj_Excel.ActiveSheet.Cells(23, 6) = r_dbl_VeCo10
   r_obj_Excel.ActiveSheet.Cells(23, 7) = r_dbl_SlCo10
   r_obj_Excel.ActiveSheet.Cells(23, 8) = r_dbl_VeCo11
   r_obj_Excel.ActiveSheet.Cells(23, 9) = r_dbl_SlCo11
   r_obj_Excel.ActiveSheet.Cells(23, 10) = r_dbl_VeCo12
   r_obj_Excel.ActiveSheet.Cells(23, 11) = r_dbl_SlCo12
   r_obj_Excel.ActiveSheet.Cells(23, 12) = r_dbl_VeCo13
   r_obj_Excel.ActiveSheet.Cells(23, 13) = r_dbl_SlCo13
   r_obj_Excel.ActiveSheet.Cells(23, 14) = r_dbl_VeCo14
   r_obj_Excel.ActiveSheet.Cells(23, 15) = r_dbl_SlCo14
   r_obj_Excel.ActiveSheet.Cells(23, 16) = r_dbl_VeCo15
   r_obj_Excel.ActiveSheet.Cells(23, 17) = r_dbl_SlCo15
   r_obj_Excel.ActiveSheet.Cells(23, 18) = r_dbl_VeCo16
   r_obj_Excel.ActiveSheet.Cells(23, 19) = r_dbl_SlCo16
   
   r_obj_Excel.ActiveSheet.Cells(23, 20) = r_dbl_SeVgCo + r_dbl_VeCo09 + r_dbl_VeCo10 + r_dbl_VeCo11 + _
                                           r_dbl_VeCo12 + r_dbl_VeCo13 + r_dbl_VeCo14 + r_dbl_VeCo15 + r_dbl_VeCo16 + _
                                           r_dbl_SlCo09 + r_dbl_SlCo10 + r_dbl_SlCo11 + r_dbl_SlCo12 + r_dbl_SlCo13 + _
                                           r_dbl_SlCo14 + r_dbl_SlCo15 + r_dbl_SlCo16
   
   
   'TOTALES MONEDA NACIONAL
   r_obj_Excel.ActiveSheet.Cells(14, 3) = r_dbl_SnVgHp + r_dbl_SnVgCo
      
   r_obj_Excel.ActiveSheet.Cells(14, 4) = r_dbl_VeHp01 + r_dbl_VeCo01
   r_obj_Excel.ActiveSheet.Cells(14, 5) = r_dbl_SlHp01 + r_dbl_SlCo01
   
   r_obj_Excel.ActiveSheet.Cells(14, 6) = r_dbl_VeHp02 + r_dbl_VeCo02
   r_obj_Excel.ActiveSheet.Cells(14, 7) = r_dbl_SlHp02 + r_dbl_SlCo02
   
   r_obj_Excel.ActiveSheet.Cells(14, 8) = r_dbl_VeHp03 + r_dbl_VeCo03
   r_obj_Excel.ActiveSheet.Cells(14, 9) = r_dbl_SlHp03 + r_dbl_SlCo03
   
   r_obj_Excel.ActiveSheet.Cells(14, 10) = r_dbl_VeHp04 + r_dbl_VeCo04
   r_obj_Excel.ActiveSheet.Cells(14, 11) = r_dbl_SlHp04 + r_dbl_SlCo04
   
   r_obj_Excel.ActiveSheet.Cells(14, 12) = r_dbl_VeHp05 + r_dbl_VeCo05
   r_obj_Excel.ActiveSheet.Cells(14, 13) = r_dbl_SlHp05 + r_dbl_SlCo05
   
   r_obj_Excel.ActiveSheet.Cells(14, 14) = r_dbl_VeHp06 + r_dbl_VeCo06
   r_obj_Excel.ActiveSheet.Cells(14, 15) = r_dbl_SlHp06 + r_dbl_SlCo06
   
   r_obj_Excel.ActiveSheet.Cells(14, 16) = r_dbl_VeHp07 + r_dbl_VeCo07
   r_obj_Excel.ActiveSheet.Cells(14, 17) = r_dbl_SlHp07 + r_dbl_SlCo07
   
   r_obj_Excel.ActiveSheet.Cells(14, 18) = r_dbl_VeHp08 + r_dbl_VeCo08
   r_obj_Excel.ActiveSheet.Cells(14, 19) = r_dbl_SlHp08 + r_dbl_SlCo08
   
   r_obj_Excel.ActiveSheet.Cells(14, 20) = r_dbl_SnVgHp + r_dbl_VeHp01 + r_dbl_VeHp02 + r_dbl_VeHp03 + r_dbl_VeHp04 + r_dbl_VeHp05 + r_dbl_VeHp06 + _
                                           r_dbl_VeHp07 + r_dbl_VeHp08 + r_dbl_SlHp01 + r_dbl_SlHp02 + r_dbl_SlHp03 + r_dbl_SlHp04 + r_dbl_SlHp05 + _
                                           r_dbl_SlHp06 + r_dbl_SlHp07 + r_dbl_SlHp08 + r_dbl_SnVgCo + r_dbl_VeCo01 + r_dbl_VeCo02 + r_dbl_VeCo03 + _
                                           r_dbl_VeCo04 + r_dbl_VeCo05 + r_dbl_VeCo06 + r_dbl_VeCo07 + r_dbl_VeCo08 + r_dbl_SlCo01 + r_dbl_SlCo02 + _
                                           r_dbl_SlCo03 + r_dbl_SlCo04 + r_dbl_SlCo05 + r_dbl_SlCo06 + r_dbl_SlCo07 + r_dbl_SlCo08
   
   
   'TOTALES MONEDA EXTRANJERA
   r_obj_Excel.ActiveSheet.Cells(22, 3) = r_dbl_SeVgHp + r_dbl_SeVgCo
      
   r_obj_Excel.ActiveSheet.Cells(22, 4) = r_dbl_VeHp09 + r_dbl_VeCo09
   r_obj_Excel.ActiveSheet.Cells(22, 5) = r_dbl_SlHp09 + r_dbl_SlCo09
      
   r_obj_Excel.ActiveSheet.Cells(22, 6) = r_dbl_VeHp10 + r_dbl_VeCo10
   r_obj_Excel.ActiveSheet.Cells(22, 7) = r_dbl_SlHp10 + r_dbl_SlCo10
   
   r_obj_Excel.ActiveSheet.Cells(22, 8) = r_dbl_VeHp11 + r_dbl_VeCo11
   r_obj_Excel.ActiveSheet.Cells(22, 9) = r_dbl_SlHp11 + r_dbl_SlCo11
   
   r_obj_Excel.ActiveSheet.Cells(22, 10) = r_dbl_VeHp12 + r_dbl_VeCo12
   r_obj_Excel.ActiveSheet.Cells(22, 11) = r_dbl_SlHp12 + r_dbl_SlCo12
   
   r_obj_Excel.ActiveSheet.Cells(22, 12) = r_dbl_VeHp13 + r_dbl_VeCo13
   r_obj_Excel.ActiveSheet.Cells(22, 13) = r_dbl_SlHp13 + r_dbl_SlCo13
   
   r_obj_Excel.ActiveSheet.Cells(22, 14) = r_dbl_VeHp14 + r_dbl_VeCo14
   r_obj_Excel.ActiveSheet.Cells(22, 15) = r_dbl_SlHp14 + r_dbl_SlCo14
   
   r_obj_Excel.ActiveSheet.Cells(22, 16) = r_dbl_VeHp15 + r_dbl_VeCo15
   r_obj_Excel.ActiveSheet.Cells(22, 17) = r_dbl_SlHp15 + r_dbl_SlCo15
   
   r_obj_Excel.ActiveSheet.Cells(22, 18) = r_dbl_VeHp16 + r_dbl_VeCo16
   r_obj_Excel.ActiveSheet.Cells(22, 19) = r_dbl_SlHp16 + r_dbl_SlCo16
      
   r_obj_Excel.ActiveSheet.Cells(22, 20) = r_dbl_SeVgHp + r_dbl_VeHp09 + r_dbl_VeHp10 + r_dbl_VeHp11 + r_dbl_VeHp12 + r_dbl_VeHp13 + r_dbl_VeHp14 + _
                                           r_dbl_VeHp15 + r_dbl_VeHp16 + r_dbl_SlHp09 + r_dbl_SlHp10 + r_dbl_SlHp11 + r_dbl_SlHp12 + r_dbl_SlHp13 + _
                                           r_dbl_SlHp14 + r_dbl_SlHp15 + r_dbl_SlHp16 + r_dbl_SeVgCo + r_dbl_VeCo09 + r_dbl_VeCo10 + r_dbl_VeCo11 + _
                                           r_dbl_VeCo12 + r_dbl_VeCo13 + r_dbl_VeCo14 + r_dbl_VeCo15 + r_dbl_VeCo16 + r_dbl_SlCo09 + r_dbl_SlCo10 + _
                                           r_dbl_SlCo11 + r_dbl_SlCo12 + r_dbl_SlCo13 + r_dbl_SlCo14 + r_dbl_SlCo15 + r_dbl_SlCo16
   
   'TOTALES
   r_obj_Excel.ActiveSheet.Cells(30, 3) = r_dbl_SeVgHp + r_dbl_SeVgCo + r_dbl_SnVgHp + r_dbl_SnVgCo
   
        
   r_obj_Excel.ActiveSheet.Cells(30, 4) = r_dbl_VeHp01 + r_dbl_VeCo01 + r_dbl_VeHp09 + r_dbl_VeCo09
   r_obj_Excel.ActiveSheet.Cells(30, 5) = r_dbl_SlHp01 + r_dbl_SlCo01 + r_dbl_SlHp09 + r_dbl_SlCo09
   
   r_obj_Excel.ActiveSheet.Cells(30, 6) = r_dbl_VeHp02 + r_dbl_VeCo02 + r_dbl_VeHp10 + r_dbl_VeCo10
   r_obj_Excel.ActiveSheet.Cells(30, 7) = r_dbl_SlHp02 + r_dbl_SlCo02 + r_dbl_SlHp10 + r_dbl_SlCo10
   
   r_obj_Excel.ActiveSheet.Cells(30, 8) = r_dbl_VeHp03 + r_dbl_VeCo03 + r_dbl_VeHp11 + r_dbl_VeCo11
   r_obj_Excel.ActiveSheet.Cells(30, 9) = r_dbl_SlHp03 + r_dbl_SlCo03 + r_dbl_SlHp11 + r_dbl_SlCo11
   
   r_obj_Excel.ActiveSheet.Cells(30, 10) = r_dbl_VeHp04 + r_dbl_VeCo04 + r_dbl_VeHp12 + r_dbl_VeCo12
   r_obj_Excel.ActiveSheet.Cells(30, 11) = r_dbl_SlHp04 + r_dbl_SlCo04 + r_dbl_SlHp12 + r_dbl_SlCo12
   
   r_obj_Excel.ActiveSheet.Cells(30, 12) = r_dbl_VeHp05 + r_dbl_VeCo05 + r_dbl_VeHp13 + r_dbl_VeCo13
   r_obj_Excel.ActiveSheet.Cells(30, 13) = r_dbl_SlHp05 + r_dbl_SlCo05 + r_dbl_SlHp13 + r_dbl_SlCo13
   
   r_obj_Excel.ActiveSheet.Cells(30, 14) = r_dbl_VeHp06 + r_dbl_VeCo06 + r_dbl_VeHp14 + r_dbl_VeCo14
   r_obj_Excel.ActiveSheet.Cells(30, 15) = r_dbl_SlHp06 + r_dbl_SlCo06 + r_dbl_SlHp14 + r_dbl_SlCo14
   
   r_obj_Excel.ActiveSheet.Cells(30, 16) = r_dbl_VeHp07 + r_dbl_VeCo07 + r_dbl_VeHp15 + r_dbl_VeCo15
   r_obj_Excel.ActiveSheet.Cells(30, 17) = r_dbl_SlHp07 + r_dbl_SlCo07 + r_dbl_SlHp15 + r_dbl_SlCo15
   
   r_obj_Excel.ActiveSheet.Cells(30, 18) = r_dbl_VeHp08 + r_dbl_VeCo08 + r_dbl_VeHp16 + r_dbl_VeCo16
   r_obj_Excel.ActiveSheet.Cells(30, 19) = r_dbl_SlHp08 + r_dbl_SlCo08 + r_dbl_SlHp16 + r_dbl_SlCo16
   
   
   r_obj_Excel.ActiveSheet.Cells(30, 20) = r_dbl_SeVgHp + r_dbl_VeHp09 + r_dbl_VeHp10 + r_dbl_VeHp11 + r_dbl_VeHp12 + r_dbl_VeHp13 + r_dbl_VeHp14 + _
                                           r_dbl_VeHp15 + r_dbl_VeHp16 + r_dbl_SlHp09 + r_dbl_SlHp10 + r_dbl_SlHp11 + r_dbl_SlHp12 + r_dbl_SlHp13 + _
                                           r_dbl_SlHp14 + r_dbl_SlHp15 + r_dbl_SlHp16 + r_dbl_SeVgCo + r_dbl_VeCo09 + r_dbl_VeCo10 + r_dbl_VeCo11 + _
                                           r_dbl_VeCo12 + r_dbl_VeCo13 + r_dbl_VeCo14 + r_dbl_VeCo15 + r_dbl_VeCo16 + r_dbl_SlCo09 + r_dbl_SlCo10 + _
                                           r_dbl_SlCo11 + r_dbl_SlCo12 + r_dbl_SlCo13 + r_dbl_SlCo14 + r_dbl_SlCo15 + r_dbl_SlCo16 + r_dbl_SnVgHp + _
                                           r_dbl_VeHp01 + r_dbl_VeHp02 + r_dbl_VeHp03 + r_dbl_VeHp04 + r_dbl_VeHp05 + r_dbl_VeHp06 + r_dbl_VeHp07 + _
                                           r_dbl_VeHp08 + r_dbl_SlHp01 + r_dbl_SlHp02 + r_dbl_SlHp03 + r_dbl_SlHp04 + r_dbl_SlHp05 + _
                                           r_dbl_SlHp06 + r_dbl_SlHp07 + r_dbl_SlHp08 + r_dbl_SnVgCo + r_dbl_VeCo01 + r_dbl_VeCo02 + r_dbl_VeCo03 + _
                                           r_dbl_VeCo04 + r_dbl_VeCo05 + r_dbl_VeCo06 + r_dbl_VeCo07 + r_dbl_VeCo08 + r_dbl_SlCo01 + r_dbl_SlCo02 + _
                                           r_dbl_SlCo03 + r_dbl_SlCo04 + r_dbl_SlCo05 + r_dbl_SlCo06 + r_dbl_SlCo07 + r_dbl_SlCo08
   
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenArc(ByVal p_FecIni As String, ByVal p_FecFin As String)
   
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_ConGen     As Integer
   
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   'Creditos Hipotecarios
   Dim r_dbl_CaVeHp     As Double
   Dim r_dbl_CaVgHp     As Double
   
   Dim r_dbl_VeHp01     As Double
   Dim r_dbl_VeHp02     As Double
   Dim r_dbl_VeHp03     As Double
   Dim r_dbl_VeHp04     As Double
   Dim r_dbl_VeHp05     As Double
   Dim r_dbl_VeHp06     As Double
   Dim r_dbl_VeHp07     As Double
   Dim r_dbl_VeHp08     As Double
   Dim r_dbl_VeHp09     As Double
   Dim r_dbl_VeHp10     As Double
   Dim r_dbl_VeHp11     As Double
   Dim r_dbl_VeHp12     As Double
   Dim r_dbl_VeHp13     As Double
   Dim r_dbl_VeHp14     As Double
   Dim r_dbl_VeHp15     As Double
   Dim r_dbl_VeHp16     As Double
   Dim r_dbl_SnVgHp     As Double
   Dim r_dbl_SeVgHp     As Double
   
   Dim r_dbl_SlHp01     As Double
   Dim r_dbl_SlHp02     As Double
   Dim r_dbl_SlHp03     As Double
   Dim r_dbl_SlHp04     As Double
   Dim r_dbl_SlHp05     As Double
   Dim r_dbl_SlHp06     As Double
   Dim r_dbl_SlHp07     As Double
   Dim r_dbl_SlHp08     As Double
   Dim r_dbl_SlHp09     As Double
   Dim r_dbl_SlHp10     As Double
   Dim r_dbl_SlHp11     As Double
   Dim r_dbl_SlHp12     As Double
   Dim r_dbl_SlHp13     As Double
   Dim r_dbl_SlHp14     As Double
   Dim r_dbl_SlHp15     As Double
   Dim r_dbl_SlHp16     As Double
         
   'Creditos Comerciales
   Dim r_dbl_CaVeCo     As Double
   Dim r_dbl_CaVgCo     As Double
   
   Dim r_dbl_VeCo01     As Double
   Dim r_dbl_VeCo02     As Double
   Dim r_dbl_VeCo03     As Double
   Dim r_dbl_VeCo04     As Double
   Dim r_dbl_VeCo05     As Double
   Dim r_dbl_VeCo06     As Double
   Dim r_dbl_VeCo07     As Double
   Dim r_dbl_VeCo08     As Double
   Dim r_dbl_VeCo09     As Double
   Dim r_dbl_VeCo10     As Double
   Dim r_dbl_VeCo11     As Double
   Dim r_dbl_VeCo12     As Double
   Dim r_dbl_VeCo13     As Double
   Dim r_dbl_VeCo14     As Double
   Dim r_dbl_VeCo15     As Double
   Dim r_dbl_VeCo16     As Double
   Dim r_dbl_SnVgCo     As Double
   Dim r_dbl_SeVgCo     As Double
   
   Dim r_dbl_SlCo01     As Double
   Dim r_dbl_SlCo02     As Double
   Dim r_dbl_SlCo03     As Double
   Dim r_dbl_SlCo04     As Double
   Dim r_dbl_SlCo05     As Double
   Dim r_dbl_SlCo06     As Double
   Dim r_dbl_SlCo07     As Double
   Dim r_dbl_SlCo08     As Double
   Dim r_dbl_SlCo09     As Double
   Dim r_dbl_SlCo10     As Double
   Dim r_dbl_SlCo11     As Double
   Dim r_dbl_SlCo12     As Double
   Dim r_dbl_SlCo13     As Double
   Dim r_dbl_SlCo14     As Double
   Dim r_dbl_SlCo15     As Double
   Dim r_dbl_SlCo16     As Double
   
   Dim r_dbl_MoSoCo     As Double
   Dim r_dbl_MtoSol     As Double
   Dim r_dbl_MoSoHp     As Double
   Dim r_dbl_MtoDol     As Double
   Dim r_dbl_MoDoHp     As Double
   Dim r_dbl_MoDoCo     As Double
       
   r_str_NomRes = "C:\01" & Right(r_str_FecFin, 6) & ".214"
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   'Creditos Comerciales
   g_str_Parame = "SELECT COMCIE_NUMOPE, COMCIE_CAPVIG, COMCIE_CAPVEN, COMCIE_TIPCAM, COMCIE_TIPMON, COMCIE_DIAMOR FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   'r_int_ConVer = 15
   
   Do While Not g_rst_Princi.EOF
                
      If g_rst_Princi!COMCIE_TIPMON = 1 Then
         If g_rst_Princi!COMCIE_CAPVEN = 0 Then
            r_dbl_SnVgCo = r_dbl_SnVgCo + g_rst_Princi!COMCIE_CAPVIG
         Else
            If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
               r_dbl_VeCo01 = r_dbl_VeCo01 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo01 = r_dbl_SlCo01 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
               r_dbl_VeCo02 = r_dbl_VeCo02 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo02 = r_dbl_SlCo02 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
               r_dbl_VeCo03 = r_dbl_VeCo03 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo03 = r_dbl_SlCo03 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
               r_dbl_VeCo04 = r_dbl_VeCo04 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo04 = r_dbl_SlCo04 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
               r_dbl_VeCo05 = r_dbl_VeCo05 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo05 = r_dbl_SlCo05 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
               r_dbl_VeCo06 = r_dbl_VeCo06 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo06 = r_dbl_SlCo06 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
               r_dbl_VeCo07 = r_dbl_VeCo07 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo07 = r_dbl_SlCo07 + g_rst_Princi!COMCIE_CAPVIG
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
               r_dbl_VeCo08 = r_dbl_VeCo08 + g_rst_Princi!COMCIE_CAPVEN
               r_dbl_SlCo08 = r_dbl_SlCo08 + g_rst_Princi!COMCIE_CAPVIG
            End If
         End If
      ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
         If g_rst_Princi!COMCIE_CAPVEN = 0 Then
            r_dbl_SeVgCo = r_dbl_SeVgCo + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
         Else
            If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
               r_dbl_VeCo09 = r_dbl_VeCo09 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo09 = r_dbl_SlCo09 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
               r_dbl_VeCo10 = r_dbl_VeCo10 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo10 = r_dbl_SlCo10 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
               r_dbl_VeCo11 = r_dbl_VeCo11 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo11 = r_dbl_SlCo11 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
               r_dbl_VeCo12 = r_dbl_VeCo12 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo12 = r_dbl_SlCo12 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
               r_dbl_VeCo13 = r_dbl_VeCo13 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo13 = r_dbl_SlCo13 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
               r_dbl_VeCo14 = r_dbl_VeCo14 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo14 = r_dbl_SlCo14 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
               r_dbl_VeCo15 = r_dbl_VeCo15 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo15 = r_dbl_SlCo15 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
               r_dbl_VeCo16 = r_dbl_VeCo16 + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlCo16 = r_dbl_SlCo16 + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
            End If
            
         End If
      End If
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
         
         
   'Credito Hipotecario
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_CAPVIG, HIPCIE_CAPVEN, HIPCIE_TIPCAM, HIPCIE_TIPMON, HIPCIE_DIAMOR FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   'r_int_ConVer = 15
   
   Do While Not g_rst_Princi.EOF
        
        
      If g_rst_Princi!HIPCIE_TIPMON = 1 Then
         If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
            r_dbl_SnVgHp = r_dbl_SnVgHp + g_rst_Princi!HIPCIE_CAPVIG
         Else
            If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
               r_dbl_VeHp01 = r_dbl_VeHp01 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp01 = r_dbl_SlHp01 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
               r_dbl_VeHp02 = r_dbl_VeHp02 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp02 = r_dbl_SlHp02 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
               r_dbl_VeHp03 = r_dbl_VeHp03 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp03 = r_dbl_SlHp03 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
               r_dbl_VeHp04 = r_dbl_VeHp04 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp04 = r_dbl_SlHp04 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
               r_dbl_VeHp05 = r_dbl_VeHp05 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp05 = r_dbl_SlHp05 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
               r_dbl_VeHp06 = r_dbl_VeHp06 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp06 = r_dbl_SlHp06 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
               r_dbl_VeHp07 = r_dbl_VeHp07 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp07 = r_dbl_SlHp07 + g_rst_Princi!HIPCIE_CAPVIG
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
               r_dbl_VeHp08 = r_dbl_VeHp08 + g_rst_Princi!HIPCIE_CAPVEN
               r_dbl_SlHp08 = r_dbl_SlHp08 + g_rst_Princi!HIPCIE_CAPVIG
            End If
         End If
      ElseIf g_rst_Princi!HIPCIE_TIPMON = 2 Then
         If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
            r_dbl_SeVgHp = r_dbl_SeVgHp + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         Else
            If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
               r_dbl_VeHp09 = r_dbl_VeHp09 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp09 = r_dbl_SlHp09 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
               r_dbl_VeHp10 = r_dbl_VeHp10 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp10 = r_dbl_SlHp10 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
               r_dbl_VeHp11 = r_dbl_VeHp11 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp11 = r_dbl_SlHp11 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
               r_dbl_VeHp12 = r_dbl_VeHp12 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp12 = r_dbl_SlHp12 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
               r_dbl_VeHp13 = r_dbl_VeHp13 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp13 = r_dbl_SlHp13 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
               r_dbl_VeHp14 = r_dbl_VeHp14 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp14 = r_dbl_SlHp14 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
               r_dbl_VeHp15 = r_dbl_VeHp15 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp15 = r_dbl_SlHp15 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
               r_dbl_VeHp16 = r_dbl_VeHp16 + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               r_dbl_SlHp16 = r_dbl_SlHp16 + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            End If
            
         End If
      End If
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_MoSoCo = r_dbl_SnVgCo + r_dbl_VeCo01 + r_dbl_VeCo02 + r_dbl_VeCo03 + r_dbl_VeCo04 + r_dbl_VeCo05 + r_dbl_VeCo06 + _
                  r_dbl_VeCo07 + r_dbl_VeCo08 + r_dbl_SlCo01 + r_dbl_SlCo02 + r_dbl_SlCo03 + r_dbl_SlCo04 + r_dbl_SlCo05 + r_dbl_SlCo06 + r_dbl_SlCo07 + r_dbl_SlCo08
      
   r_dbl_MoSoHp = r_dbl_SnVgHp + r_dbl_VeHp01 + r_dbl_VeHp02 + r_dbl_VeHp03 + r_dbl_VeHp04 + r_dbl_VeHp05 + r_dbl_VeHp06 + r_dbl_VeHp07 + r_dbl_VeHp08 + _
                  r_dbl_SlHp01 + r_dbl_SlHp02 + r_dbl_SlHp03 + r_dbl_SlHp04 + r_dbl_SlHp05 + r_dbl_SlHp06 + r_dbl_SlHp07 + r_dbl_SlHp08
   
   
   r_dbl_MoDoCo = r_dbl_SeVgCo + r_dbl_VeCo09 + r_dbl_VeCo10 + r_dbl_VeCo11 + r_dbl_VeCo12 + r_dbl_VeCo13 + r_dbl_VeCo14 + r_dbl_VeCo15 + r_dbl_VeCo16 + _
                  r_dbl_SlCo09 + r_dbl_SlCo10 + r_dbl_SlCo11 + r_dbl_SlCo12 + r_dbl_SlCo13 + r_dbl_SlCo14 + r_dbl_SlCo15 + r_dbl_SlCo16
   
   r_dbl_MoDoHp = r_dbl_SeVgHp + r_dbl_VeHp09 + r_dbl_VeHp10 + r_dbl_VeHp11 + r_dbl_VeHp12 + r_dbl_VeHp13 + r_dbl_VeHp14 + r_dbl_VeHp15 + r_dbl_VeHp16 + _
                  r_dbl_SlHp09 + r_dbl_SlHp10 + r_dbl_SlHp11 + r_dbl_SlHp12 + r_dbl_SlHp13 + r_dbl_SlHp14 + r_dbl_SlHp15 + r_dbl_SlHp16
   
   r_dbl_MtoSol = r_dbl_MoSoHp + r_dbl_MoSoCo
   r_dbl_MtoDol = r_dbl_MoDoHp + r_dbl_MoDoCo
   
      
   Print #r_int_NumRes, Format(214, "0000") & Format(1, "00") & Format(240, "00000") & r_str_FecFin & Format(12, "000")
    
   For r_int_ConGen = 10 To 150 Step 10
      r_int_Contad = 0
      r_str_Cadena = ""
  
      For r_int_Contad = 1 To 270 Step 1
         r_str_Cadena = r_str_Cadena & "0"
      Next
  
      If r_int_ConGen = 10 Then     'MONEDA NACIONAL
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SnVgHp + r_dbl_SnVgCo), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo01 + r_dbl_VeHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo01 + r_dbl_SlHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo02 + r_dbl_VeHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo02 + r_dbl_SlHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo03 + r_dbl_VeHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo03 + r_dbl_SlHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo04 + r_dbl_VeHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo04 + r_dbl_SlHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo05 + r_dbl_VeHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo05 + r_dbl_SlHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo06 + r_dbl_VeHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo06 + r_dbl_SlHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo07 + r_dbl_VeHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo07 + r_dbl_SlHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo08 + r_dbl_VeHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo08 + r_dbl_SlHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MtoSol), "###,###,##0.00")), 15)
  
      ElseIf r_int_ConGen = 20 Then     'COMERCIALES
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SnVgCo), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MoSoCo), "###,###,##0.00")), 15)
   
      ElseIf r_int_ConGen = 50 Then     'HIPOTECARIO
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SnVgHp), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MoSoHp), "###,###,##0.00")), 15)
      
      ElseIf r_int_ConGen = 80 Then     'DOLARES AMERICANOS
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SeVgCo + r_dbl_SeVgHp), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo09 + r_dbl_VeHp09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo09 + r_dbl_SlHp09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo10 + r_dbl_VeHp10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo10 + r_dbl_SlHp10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo11 + r_dbl_VeHp11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo11 + r_dbl_SlHp11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo12 + r_dbl_VeHp12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo12 + r_dbl_SlHp12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo13 + r_dbl_VeHp13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo13 + r_dbl_SlHp13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo14 + r_dbl_VeHp14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo14 + r_dbl_SlHp14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo15 + r_dbl_VeHp15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo15 + r_dbl_SlHp15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo16 + r_dbl_VeHp16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo16 + r_dbl_SlHp16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MtoDol), "###,###,##0.00")), 15)
      
      ElseIf r_int_ConGen = 90 Then     'COMERCIALES
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SeVgCo), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MoDoCo), "###,###,##0.00")), 15)
      
      ElseIf r_int_ConGen = 120 Then     'HIPOTECARIO
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SeVgHp), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp09), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp10), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp11), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp12), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp13), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp14), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp15), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeHp16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlHp16), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MoDoHp), "###,###,##0.00")), 15)
      
      ElseIf r_int_ConGen = 150 Then     'TOTAL GENERAL
         
         r_str_Cadena = modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SeVgCo + r_dbl_SeVgHp + r_dbl_SnVgHp + r_dbl_SnVgCo), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo09 + r_dbl_VeHp09 + r_dbl_VeCo01 + r_dbl_VeHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo09 + r_dbl_SlHp09 + r_dbl_SlCo01 + r_dbl_SlHp01), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo10 + r_dbl_VeHp10 + r_dbl_VeCo02 + r_dbl_VeHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo10 + r_dbl_SlHp10 + r_dbl_SlCo02 + r_dbl_SlHp02), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo11 + r_dbl_VeHp11 + r_dbl_VeCo03 + r_dbl_VeHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo11 + r_dbl_SlHp11 + r_dbl_SlCo03 + r_dbl_SlHp03), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo12 + r_dbl_VeHp12 + r_dbl_VeCo04 + r_dbl_VeHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo12 + r_dbl_SlHp12 + r_dbl_SlCo04 + r_dbl_SlHp04), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo13 + r_dbl_VeHp13 + r_dbl_VeCo05 + r_dbl_VeHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo13 + r_dbl_SlHp13 + r_dbl_SlCo05 + r_dbl_SlHp05), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo14 + r_dbl_VeHp14 + r_dbl_VeCo06 + r_dbl_VeHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo14 + r_dbl_SlHp14 + r_dbl_SlCo06 + r_dbl_SlHp06), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo15 + r_dbl_VeHp15 + r_dbl_VeCo07 + r_dbl_VeHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo15 + r_dbl_SlHp15 + r_dbl_SlCo07 + r_dbl_SlHp07), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_VeCo16 + r_dbl_VeHp16 + r_dbl_VeCo08 + r_dbl_VeHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_SlCo16 + r_dbl_SlHp16 + r_dbl_SlCo08 + r_dbl_SlHp08), "###,###,##0.00")), 15)
         r_str_Cadena = r_str_Cadena & modtac_gs_Genera_NumCar(modtac_gs_Cadena_ExtSal(Format(CDbl(r_dbl_MtoDol + r_dbl_MtoSol), "###,###,##0.00")), 15)
  
      End If
      
      Print #r_int_NumRes, Format(r_int_ConGen, "0000") & r_str_Cadena
      
   Next
   
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt
   
   
End Sub

