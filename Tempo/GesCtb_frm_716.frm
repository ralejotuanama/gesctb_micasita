VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_18 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   5460
   ClientTop       =   6240
   ClientWidth     =   9045
   Icon            =   "GesCtb_frm_716.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9075
      _Version        =   65536
      _ExtentX        =   16007
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
         TabIndex        =   8
         Top             =   60
         Width           =   8985
         _Version        =   65536
         _ExtentX        =   15849
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
            TabIndex        =   9
            Top             =   30
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte N° 2-D"
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
            TabIndex        =   10
            Top             =   270
            Width           =   8265
            _Version        =   65536
            _ExtentX        =   14579
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0202-50 Requerimientos de Patrimonio Efectivo por Riesgo de Credito, Mercado y Operacional"
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
            Picture         =   "GesCtb_frm_716.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   8985
         _Version        =   65536
         _ExtentX        =   15849
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
            Picture         =   "GesCtb_frm_716.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_716.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_716.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8370
            Picture         =   "GesCtb_frm_716.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_716.frx":11AE
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
         TabIndex        =   12
         Top             =   1470
         Width           =   8985
         _Version        =   65536
         _ExtentX        =   15849
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
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
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
Attribute VB_Name = "frm_RepSbs_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim r_dbl_Evalua(100)   As Double
   Dim r_str_Denomi(50)    As String
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_LIMGLO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_LIMGLO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

   r_str_Denomi(0) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Crédito"
   r_str_Denomi(1) = "Método Estándar"
   r_str_Denomi(2) = "Riesgo de Tasa de Interés"
   r_str_Denomi(3) = "Riesgo de Precio"
   r_str_Denomi(4) = "Riesgo Cambiario"
   r_str_Denomi(5) = "Riesgo de Commodities"
   r_str_Denomi(6) = "Método de Modelos Internos"
   r_str_Denomi(7) = "VAR Total"
   r_str_Denomi(8) = "Promedio VAR últimos 60 días útiles"
   r_str_Denomi(9) = "Factor"
   r_str_Denomi(10) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Mercado"
   r_str_Denomi(11) = "Método del Indicador Básico"
   r_str_Denomi(12) = "Método Estándar Alternativo"
   r_str_Denomi(13) = "Métodos Avanzados"
   r_str_Denomi(14) = "Total Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
   r_str_Denomi(15) = "Requerimiento de Patrimonio Efectivo Total"
   r_str_Denomi(16) = "Total Patrimonio Efectivo de Nivel 1"
   r_str_Denomi(17) = "Asignado a cubrir riesgo de crédito"
   r_str_Denomi(18) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(19) = "Asignado a cubrir riesgo operacional"
   r_str_Denomi(20) = "Total Patrimonio Efectivo de Nivel 2"
   r_str_Denomi(21) = "Asignado a cubrir riesgo de crédito"
   r_str_Denomi(22) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(23) = "Asignado a cubrir riesgo operacional"
   r_str_Denomi(24) = "Total Patrimonio Efectivo de Nivel 3"
   r_str_Denomi(25) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(26) = "Total Patrimonio Efectivo"
   r_str_Denomi(27) = "Ratio de Capital Global(%)"

   Erase r_dbl_Evalua()
   
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
                  
      .Cells(2, 1) = "CÁLCULO DEL LÍMITE GLOBAL A QUE SE REFIERE EL PRIMER PÁRRAFO DEL ARTÍCULO 199º Y LA VIGÉSIMA"
      .Cells(3, 1) = "(Artículos 186º AL 194º y Vigésima Cuarta Dispocisión Transitoria de la Ley General Nº26702)"
      .Cells(4, 1) = "(En Miles de Nuevos Soles)"
      .Cells(6, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      .Cells(8, 1) = "EMPRESA: Edpyme MiCasita S.A."
      .Range(.Cells(2, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(6, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 1), .Cells(2, 3)).Merge
      .Range(.Cells(3, 1), .Cells(3, 3)).Merge
      .Range(.Cells(4, 1), .Cells(4, 3)).Merge
      .Range(.Cells(6, 1), .Cells(6, 3)).Merge
      .Range(.Cells(2, 1), .Cells(6, 1)).WrapText = True

      .Cells(10, 1) = "(I)Requerimiento de Patrimonio Efectivo por Riesgo de Crédito:"
      .Cells(11, 2) = "APR"
      .Cells(11, 3) = "Total4"
      .Cells(12, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Crédito"
      
      .Cells(14, 1) = "(II)Requerimiento de Patrimonio Efectivo por Riesgo de Mercado:"
      .Cells(15, 2) = "APR"
      .Cells(15, 3) = "Total4"
      .Cells(16, 1) = "Método Estándar"
      .Cells(17, 1) = "Riesgo de Tasa de Interés"
      .Cells(18, 1) = "Riesgo de Precio"
      .Cells(19, 1) = "Riesgo Cambiario"
      .Cells(20, 1) = "Riesgo de Commodities"
      .Cells(21, 1) = "Método de Modelos Internos"
      .Cells(22, 1) = "VAR Total"
      .Cells(23, 1) = "Promedio VAR últimos 60 días útiles"
      .Cells(24, 1) = "Factor"
      .Cells(25, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Mercado"
      
      .Range(.Cells(14, 1), .Cells(14, 3)).Font.Bold = True
      .Range(.Cells(15, 1), .Cells(15, 3)).Font.Bold = True
      .Range(.Cells(16, 1), .Cells(16, 3)).Font.Bold = True
      .Range(.Cells(21, 1), .Cells(21, 3)).Font.Bold = True
      .Range(.Cells(25, 1), .Cells(25, 3)).Font.Bold = True
      
      .Cells(27, 1) = "(III)Requerimiento de Patrimonio Efectivo por Riesgo Operacional:"
      .Cells(28, 2) = "APR"
      .Cells(28, 3) = "Total4"
      .Cells(29, 1) = "Método del Indicador Básico"
      .Cells(30, 1) = "Método Estándar Alternativo"
      .Cells(31, 1) = "Métodos Avanzados"
      .Cells(32, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
      
      .Range(.Cells(27, 1), .Cells(32, 3)).Font.Bold = True
      
      .Cells(34, 1) = "(IV)Requerimiento de Patrimonio Efectivo Total:"
      
      .Cells(36, 1) = "(V)Patrimonio Efectivo5"
      .Cells(37, 2) = "Total"
      .Cells(38, 1) = "Total Patrimonio Efectivo de Nivel 1"
      .Cells(39, 1) = "Asignado a cubrir riesgo de crédito"
      .Cells(40, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(41, 1) = "Asignado a cubrir riesgo operacional"
      .Cells(42, 1) = "Total Patrimonio Efectivo de Nivel 2"
      .Cells(43, 1) = "Asignado a cubrir riesgo de crédito"
      .Cells(44, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(45, 1) = "Asignado a cubrir riesgo operacional"
      .Cells(46, 1) = "Total Patrimonio Efectivo de Nivel 3"
      .Cells(47, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(48, 1) = "Total Patrimonio Efectivo6"
      .Cells(50, 1) = "(VI)Ratio de Capital Global(%)"
      
      .Cells(50, 3) = "(*)"
      .Cells(50, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(51, 1) = "(*)"
      
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(11, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 1), .Cells(12, 3)).VerticalAlignment = xlVAlignCenter
            
      .Range(.Cells(10, 1), .Cells(12, 3)).Font.Bold = True
      
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(15, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(15, 1), .Cells(25, 3)).VerticalAlignment = xlVAlignCenter
      '.Range(.Cells(11, 1), .Cells(12, 3)).WrapText = True
      '.Range(.Cells(14, 1), .Cells(25, 3)).Font.Bold = True
      
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(28, 1), .Cells(32, 3)).VerticalAlignment = xlVAlignCenter
      '.Range(.Cells(11, 1), .Cells(12, 3)).WrapText = True
      '.Range(.Cells(14, 1), .Cells(25, 3)).Font.Bold = True
      
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Font.Bold = True
      
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      '.Range(.Cells(37, 1), .Cells(48, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(37, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(37, 1), .Cells(48, 2)).VerticalAlignment = xlVAlignCenter
      
      .Range(.Cells(36, 1), .Cells(38, 2)).Font.Bold = True
      .Range(.Cells(42, 1), .Cells(42, 2)).Font.Bold = True
      .Range(.Cells(46, 1), .Cells(46, 2)).Font.Bold = True
      .Range(.Cells(48, 1), .Cells(48, 2)).Font.Bold = True
      .Range(.Cells(38, 1), .Cells(38, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(42, 1), .Cells(42, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(46, 1), .Cells(46, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(48, 1), .Cells(48, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      '.Range(.Cells(37, 1), .Cells(48, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(50, 1), .Cells(50, 2)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(50, 1), .Cells(50, 2)).Font.Bold = True
            
      .Range(.Cells(53, 1), .Cells(55, 1)).HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(53, 1), .Cells(53, 3)).Merge
      .Range(.Cells(54, 1), .Cells(54, 3)).Merge
      .Range(.Cells(55, 1), .Cells(55, 3)).Merge
      
      .Cells(53, 1) = "   ________________________      ____________________________       _______________________"
      .Cells(54, 1) = "Sr. Roberto Baba Yamamoto          C.P.C. Rossana Meza Bustamante       Sr. Javier Delgado Blanco"
      .Cells(55, 1) = "   Gerente General                                 Contador General                           Unidad de Riesgos"
      
      .Range(.Cells(57, 1), .Cells(57, 3)).Merge
      .Range(.Cells(58, 1), .Cells(58, 3)).Merge
      .Range(.Cells(59, 1), .Cells(59, 3)).Merge
      .Range(.Cells(60, 1), .Cells(60, 3)).Merge
      .Range(.Cells(62, 1), .Cells(62, 3)).Merge
      
      .Cells(57, 1) = "4 Activo Ponderado por Riesgo (APR) multiplicado por el límite global que establece la Ley General en el primer párrafo del artículo 199º y la Vigésima Cuarta Disposición Transitoria."
      .Cells(58, 1) = "5 El patrimonio efectivo de los niveles deberá ser asignado a cubrir riesgo de crédito, de mercado y operacional, de acuerdo con lo establecido en los numerales 1 y 3 del artículo 185º de la Ley General."
      .Cells(59, 1) = "6 De acuerdo con lo establecido en el Reporte Nº3 ""Patrimonio Efectivo"" del Manual de Contabilidad para las Empresas del Sistemas Financiero."
      .Cells(60, 1) = "7 v = Límite global que establece la Ley Generalen el primer párrafo del artículo 199º y la Vigésima Cuarta Disposición Transitoria."
      .Cells(62, 1) = "(*) Fe de Erratas publicada el19 de junio de 2009 en el portal electrónico institucional (www.sbs.gob.pe)"
      
      .Range(.Cells(57, 1), .Cells(59, 1)).RowHeight = 30
      
      .Cells(57, 1).WrapText = True
      .Cells(58, 1).WrapText = True
      .Cells(59, 1).WrapText = True
                                              
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
      .Range(.Cells(8, 1), .Cells(50, 1)).HorizontalAlignment = xlHAlignLeft
            
      .Columns("A").ColumnWidth = 70
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 15
            
      .Columns("B:C").NumberFormat = "###,###,##0.00"
      
      r_int_ConAux = 0
      
      For r_int_Contad = 12 To 50 Step 1
         If r_int_Contad < 35 Then
            If r_int_Contad <> 13 And r_int_Contad <> 14 And r_int_Contad <> 15 And r_int_Contad <> 26 And r_int_Contad <> 27 And r_int_Contad <> 28 And r_int_Contad <> 33 Then
               .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux)
               .Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConAux + 1)
               r_int_ConAux = r_int_ConAux + 2
            End If
         Else
            If r_int_Contad <> 35 And r_int_Contad <> 36 And r_int_Contad <> 37 And r_int_Contad <> 49 Then
               .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux + 1)
               r_int_ConAux = r_int_ConAux + 2
            End If
         End If

      Next
                         
   End With

   
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
      
   r_str_NomRes = "C:\50" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".202"
   
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
   
   Print #r_int_NumRes, Format(202, "0000") & Format(50, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
      
   r_str_Cadena = ""
   
   For r_int_ConTem = 0 To 1 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 18)
   Next
   
   Print #r_int_NumRes, Format(100, "000000") & r_str_Cadena
   
   r_str_Cadena = ""
      
   For r_int_ConTem = 0 To 1 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConTem), "########0.00"), 1, "0", 18)
   Next
      
   Print #r_int_NumRes, Format(120, "000000") & r_str_Cadena
   
   r_str_Cadena = ""
   
   For r_int_ConTem = 0 To 1 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 18)
   Next
   
   Print #r_int_NumRes, Format(140, "000000") & r_str_Cadena
   
   r_int_ConAux = 0
   
   For r_int_Contad = 200 To 3200 Step 100
      r_str_Cadena = ""
      
      If r_int_Contad <> 300 And r_int_Contad <> 1400 And r_int_Contad <> 2000 Then
         For r_int_ConTem = 0 To 1 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
            r_int_ConAux = r_int_ConAux + 1
         Next
      Else
         For r_int_ConTem = 0 To 1 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 18)
         Next
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "000000") & r_str_Cadena
      
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

   g_str_Parame = "DELETE FROM HIS_LIMGLO WHERE "
   g_str_Parame = g_str_Parame & "LIMGLO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "LIMGLO_PERANO = " & r_int_PerAno & " "
   'g_str_Parame = g_str_Parame & "LIMGLO_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = 0
   
   For r_int_Contad = 0 To 27 Step 1
            
      r_str_Cadena = "USP_HIS_LIMGLO ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_Contad) & "', "
      
      If r_int_ConAux < 32 Then
      
         For r_int_ConTem = 0 To 1 Step 1
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux)
            r_int_ConAux = r_int_ConAux + 1
         Next
      
      Else
      
         r_int_ConAux = r_int_ConAux + 1
         
         For r_int_ConTem = 0 To 1 Step 1
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux)
            r_int_ConAux = r_int_ConAux - 1
         Next
         
         r_int_ConAux = r_int_ConAux + 3
      
      End If
      
      r_str_Cadena = r_str_Cadena & ")"
      
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_LIMGLO.", vbCritical, modgen_g_str_NomPlt
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
   Call fs_GeneDB
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
                  
      .Cells(2, 1) = "CÁLCULO DEL LÍMITE GLOBAL A QUE SE REFIERE EL PRIMER PÁRRAFO DEL ARTÍCULO 199º Y LA VIGÉSIMA"
      .Cells(3, 1) = "(Artículos 186º AL 194º y Vigésima Cuarta Dispocisión Transitoria de la Ley General Nº26702)"
      .Cells(4, 1) = "(En Miles de Nuevos Soles)"
      .Cells(6, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      .Cells(8, 1) = "EMPRESA: Edpyme MiCasita S.A."
      .Range(.Cells(2, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(6, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 1), .Cells(2, 3)).Merge
      .Range(.Cells(3, 1), .Cells(3, 3)).Merge
      .Range(.Cells(4, 1), .Cells(4, 3)).Merge
      .Range(.Cells(6, 1), .Cells(6, 3)).Merge
      .Range(.Cells(2, 1), .Cells(6, 1)).WrapText = True

      .Cells(10, 1) = "(I)Requerimiento de Patrimonio Efectivo por Riesgo de Crédito:"
      .Cells(11, 2) = "APR"
      .Cells(11, 3) = "Total4"
      .Cells(12, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Crédito"
      
      .Cells(14, 1) = "(II)Requerimiento de Patrimonio Efectivo por Riesgo de Mercado:"
      .Cells(15, 2) = "APR"
      .Cells(15, 3) = "Total4"
      .Cells(16, 1) = "Método Estándar"
      .Cells(17, 1) = "Riesgo de Tasa de Interés"
      .Cells(18, 1) = "Riesgo de Precio"
      .Cells(19, 1) = "Riesgo Cambiario"
      .Cells(20, 1) = "Riesgo de Commodities"
      .Cells(21, 1) = "Método de Modelos Internos"
      .Cells(22, 1) = "VAR Total"
      .Cells(23, 1) = "Promedio VAR últimos 60 días útiles"
      .Cells(24, 1) = "Factor"
      .Cells(25, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Mercado"
      
      .Range(.Cells(14, 1), .Cells(14, 3)).Font.Bold = True
      .Range(.Cells(15, 1), .Cells(15, 3)).Font.Bold = True
      .Range(.Cells(16, 1), .Cells(16, 3)).Font.Bold = True
      .Range(.Cells(21, 1), .Cells(21, 3)).Font.Bold = True
      .Range(.Cells(25, 1), .Cells(25, 3)).Font.Bold = True
      
      .Cells(27, 1) = "(III)Requerimiento de Patrimonio Efectivo por Riesgo Operacional:"
      .Cells(28, 2) = "APR"
      .Cells(28, 3) = "Total4"
      .Cells(29, 1) = "Método del Indicador Básico"
      .Cells(30, 1) = "Método Estándar Alternativo"
      .Cells(31, 1) = "Métodos Avanzados"
      .Cells(32, 1) = "Total Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
      
      .Range(.Cells(27, 1), .Cells(32, 3)).Font.Bold = True
      
      .Cells(34, 1) = "(IV)Requerimiento de Patrimonio Efectivo Total:"
      
      .Cells(36, 1) = "(V)Patrimonio Efectivo5"
      .Cells(37, 2) = "Total"
      .Cells(38, 1) = "Total Patrimonio Efectivo de Nivel 1"
      .Cells(39, 1) = "Asignado a cubrir riesgo de crédito"
      .Cells(40, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(41, 1) = "Asignado a cubrir riesgo operacional"
      .Cells(42, 1) = "Total Patrimonio Efectivo de Nivel 2"
      .Cells(43, 1) = "Asignado a cubrir riesgo de crédito"
      .Cells(44, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(45, 1) = "Asignado a cubrir riesgo operacional"
      .Cells(46, 1) = "Total Patrimonio Efectivo de Nivel 3"
      .Cells(47, 1) = "Asignado a cubrir riesgo de mercado"
      .Cells(48, 1) = "Total Patrimonio Efectivo6"
      .Cells(50, 1) = "(VI)Ratio de Capital Global(%)"
      
      .Cells(50, 3) = "(*)"
      .Cells(50, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(51, 1) = "(*)"
      
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(12, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(11, 1), .Cells(11, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 1), .Cells(12, 3)).VerticalAlignment = xlVAlignCenter
            
      .Range(.Cells(10, 1), .Cells(12, 3)).Font.Bold = True
      
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(25, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(15, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(15, 1), .Cells(25, 3)).VerticalAlignment = xlVAlignCenter
      '.Range(.Cells(11, 1), .Cells(12, 3)).WrapText = True
      '.Range(.Cells(14, 1), .Cells(25, 3)).Font.Bold = True
      
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(32, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(28, 1), .Cells(28, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(28, 1), .Cells(32, 3)).VerticalAlignment = xlVAlignCenter
      '.Range(.Cells(11, 1), .Cells(12, 3)).WrapText = True
      '.Range(.Cells(14, 1), .Cells(25, 3)).Font.Bold = True
      
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(34, 1), .Cells(34, 3)).Font.Bold = True
      
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(48, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      '.Range(.Cells(37, 1), .Cells(48, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(37, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(37, 1), .Cells(48, 2)).VerticalAlignment = xlVAlignCenter
      
      .Range(.Cells(36, 1), .Cells(38, 2)).Font.Bold = True
      .Range(.Cells(42, 1), .Cells(42, 2)).Font.Bold = True
      .Range(.Cells(46, 1), .Cells(46, 2)).Font.Bold = True
      .Range(.Cells(48, 1), .Cells(48, 2)).Font.Bold = True
      .Range(.Cells(38, 1), .Cells(38, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(42, 1), .Cells(42, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(46, 1), .Cells(46, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(48, 1), .Cells(48, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      '.Range(.Cells(37, 1), .Cells(48, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(50, 1), .Cells(50, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(50, 1), .Cells(50, 2)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(50, 1), .Cells(50, 2)).Font.Bold = True
            
      .Range(.Cells(53, 1), .Cells(55, 1)).HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(53, 1), .Cells(53, 3)).Merge
      .Range(.Cells(54, 1), .Cells(54, 3)).Merge
      .Range(.Cells(55, 1), .Cells(55, 3)).Merge
      
      .Cells(53, 1) = "   ________________________      ____________________________       _______________________"
      .Cells(54, 1) = "Sr. Roberto Baba Yamamoto          C.P.C. Rossana Meza Bustamante       Sr. Javier Delgado Blanco"
      .Cells(55, 1) = "   Gerente General                                 Contador General                           Unidad de Riesgos"
      
      .Range(.Cells(57, 1), .Cells(57, 3)).Merge
      .Range(.Cells(58, 1), .Cells(58, 3)).Merge
      .Range(.Cells(59, 1), .Cells(59, 3)).Merge
      .Range(.Cells(60, 1), .Cells(60, 3)).Merge
      .Range(.Cells(62, 1), .Cells(62, 3)).Merge
      
      .Cells(57, 1) = "4 Activo Ponderado por Riesgo (APR) multiplicado por el límite global que establece la Ley General en el primer párrafo del artículo 199º y la Vigésima Cuarta Disposición Transitoria."
      .Cells(58, 1) = "5 El patrimonio efectivo de los niveles deberá ser asignado a cubrir riesgo de crédito, de mercado y operacional, de acuerdo con lo establecido en los numerales 1 y 3 del artículo 185º de la Ley General."
      .Cells(59, 1) = "6 De acuerdo con lo establecido en el Reporte Nº3 ""Patrimonio Efectivo"" del Manual de Contabilidad para las Empresas del Sistemas Financiero."
      .Cells(60, 1) = "7 v = Límite global que establece la Ley Generalen el primer párrafo del artículo 199º y la Vigésima Cuarta Disposición Transitoria."
      .Cells(62, 1) = "(*) Fe de Erratas publicada el19 de junio de 2009 en el portal electrónico institucional (www.sbs.gob.pe)"
      
      .Range(.Cells(57, 1), .Cells(59, 1)).RowHeight = 30
      
      .Cells(57, 1).WrapText = True
      .Cells(58, 1).WrapText = True
      .Cells(59, 1).WrapText = True
                                              
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
      .Range(.Cells(8, 1), .Cells(50, 1)).HorizontalAlignment = xlHAlignLeft
            
      .Columns("A").ColumnWidth = 70
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 15
            
      .Columns("B:C").NumberFormat = "###,###,##0.00"
      
      r_int_ConAux = 0
      
      For r_int_Contad = 12 To 50 Step 1
         If r_int_Contad < 35 Then
            If r_int_Contad <> 13 And r_int_Contad <> 14 And r_int_Contad <> 15 And r_int_Contad <> 26 And r_int_Contad <> 27 And r_int_Contad <> 28 And r_int_Contad <> 33 Then
               .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux)
               .Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConAux + 1)
               r_int_ConAux = r_int_ConAux + 2
            End If
         Else
            If r_int_Contad <> 35 And r_int_Contad <> 36 And r_int_Contad <> 37 And r_int_Contad <> 49 Then
               .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux + 1)
               r_int_ConAux = r_int_ConAux + 2
            End If
         End If

      Next
                         
   End With
   
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

   r_str_Denomi(0) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Crédito"
   r_str_Denomi(1) = "Método Estándar"
   r_str_Denomi(2) = "Riesgo de Tasa de Interés"
   r_str_Denomi(3) = "Riesgo de Precio"
   r_str_Denomi(4) = "Riesgo Cambiario"
   r_str_Denomi(5) = "Riesgo de Commodities"
   r_str_Denomi(6) = "Método de Modelos Internos"
   r_str_Denomi(7) = "VAR Total"
   r_str_Denomi(8) = "Promedio VAR últimos 60 días útiles"
   r_str_Denomi(9) = "Factor"
   r_str_Denomi(10) = "Total Requerimiento de Patrimonio Efectivo por Riesgo de Mercado"
   r_str_Denomi(11) = "Método del Indicador Básico"
   r_str_Denomi(12) = "Método Estándar Alternativo"
   r_str_Denomi(13) = "Métodos Avanzados"
   r_str_Denomi(14) = "Total Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
   r_str_Denomi(15) = "Requerimiento de Patrimonio Efectivo Total"
   r_str_Denomi(16) = "Total Patrimonio Efectivo de Nivel 1"
   r_str_Denomi(17) = "Asignado a cubrir riesgo de crédito"
   r_str_Denomi(18) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(19) = "Asignado a cubrir riesgo operacional"
   r_str_Denomi(20) = "Total Patrimonio Efectivo de Nivel 2"
   r_str_Denomi(21) = "Asignado a cubrir riesgo de crédito"
   r_str_Denomi(22) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(23) = "Asignado a cubrir riesgo operacional"
   r_str_Denomi(24) = "Total Patrimonio Efectivo de Nivel 3"
   r_str_Denomi(25) = "Asignado a cubrir riesgo de mercado"
   r_str_Denomi(26) = "Total Patrimonio Efectivo"
   r_str_Denomi(27) = "Ratio de Capital Global(%)"

   Erase r_dbl_Evalua()
   
   g_str_Parame = "SELECT * FROM HIS_LIMGLO WHERE "
   g_str_Parame = g_str_Parame & "LIMGLO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "LIMGLO_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY LIMGLO_NUMITE ASC "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_ConAux = -1
   
      Do While Not g_rst_Princi.EOF
         
         r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!LIMGLO_MTOAPR
         r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!LIMGLO_MTOTOT
         
         r_int_ConAux = r_int_ConAux + 2
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
End Sub



