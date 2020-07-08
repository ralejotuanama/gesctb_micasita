VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_30 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   1425
   ClientTop       =   2175
   ClientWidth     =   14025
   Icon            =   "GesCtb_frm_855.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel5 
      Height          =   10095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   14085
      _Version        =   65536
      _ExtentX        =   24844
      _ExtentY        =   17806
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   7650
         Left            =   30
         TabIndex        =   7
         Top             =   2280
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
         _ExtentY        =   13494
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisEEBG 
            Height          =   7515
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   13256
            _Version        =   393216
            Rows            =   21
            Cols            =   30
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            Redraw          =   -1  'True
            MergeCells      =   1
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   13970
         _Version        =   65536
         _ExtentX        =   24642
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
            TabIndex        =   9
            Top             =   180
            Width           =   6735
            _Version        =   65536
            _ExtentX        =   11880
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Origen y Aplicación de Balance General"
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
            Picture         =   "GesCtb_frm_855.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   730
         Width           =   13970
         _Version        =   65536
         _ExtentX        =   24642
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
            Left            =   45
            Picture         =   "GesCtb_frm_855.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   645
            Picture         =   "GesCtb_frm_855.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13320
            Picture         =   "GesCtb_frm_855.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   11
         Top             =   1420
         Width           =   13970
         _Version        =   65536
         _ExtentX        =   24642
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
            Width           =   3795
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
            TabIndex        =   13
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   450
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
   
   Screen.MousePointer = 11
   Call fs_CargaDatos_Balance
   Call fs_CargaDatos_Variacion
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
   Call fs_GenExcRes
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
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   ipp_PerAno.Text = Year(date)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call gs_LimpiaGrid(grd_LisEEBG)
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExcRes.Enabled = estado
End Sub

Private Sub fs_CargaDatos_Balance()
Dim r_str_PerMes    As String
Dim r_str_PerAno    As String
    
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   grd_LisEEBG.Redraw = False
   Call fs_Setea_Columnas
   
   'Prepara SP
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_CUR_GEN_EEBG ("
   g_str_Parame = g_str_Parame & CInt(r_str_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) & ", 1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "') "

   'Ejecuta consulta
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEBG.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_LisEEBG.Redraw = False
   Call gs_LimpiaGrid(grd_LisEEBG)
   
   'Primera Linea
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.Row = 0:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""
   grd_LisEEBG.Col = 3:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "EJERCICIO"
   grd_LisEEBG.Col = 4:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'ENERO
   grd_LisEEBG.Col = 5:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'FEBRERO
   grd_LisEEBG.Col = 6:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'MARZO
   grd_LisEEBG.Col = 7:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'ABRIL
   grd_LisEEBG.Col = 8:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'MAYO
   grd_LisEEBG.Col = 9:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'JUNIO
   grd_LisEEBG.Col = 10:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'JULIO
   grd_LisEEBG.Col = 11:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'AGOSTO
   grd_LisEEBG.Col = 12:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'SETIEMBRE
   grd_LisEEBG.Col = 13:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'OCTUBRE
   grd_LisEEBG.Col = 14:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'NOVIEMBRE
   grd_LisEEBG.Col = 15:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'DICIEMBRE
   grd_LisEEBG.Col = 20:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "FEB-ENE"
   grd_LisEEBG.Col = 21:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "FEB-ENE"
   grd_LisEEBG.Col = 22:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "MAR-FEB"
   grd_LisEEBG.Col = 23:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "MAR-FEB"
   grd_LisEEBG.Col = 24:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ABR-MAR"
   grd_LisEEBG.Col = 25:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ABR-MAR"
   grd_LisEEBG.Col = 26:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "MAY-ABR"
   grd_LisEEBG.Col = 27:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "MAY-ABR"
   grd_LisEEBG.Col = 28:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "JUN-MAY"
   grd_LisEEBG.Col = 29:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "JUN-MAY"
   grd_LisEEBG.Col = 30:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "JUL-JUN"
   grd_LisEEBG.Col = 31:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "JUL-JUN"
   grd_LisEEBG.Col = 32:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "AGO-JUL"
   grd_LisEEBG.Col = 33:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "AGO-JUL"
   grd_LisEEBG.Col = 34:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "SET-AGO"
   grd_LisEEBG.Col = 35:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "SET-AGO"
   grd_LisEEBG.Col = 36:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "OCT-SET"
   grd_LisEEBG.Col = 37:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "OCT-SET"
   grd_LisEEBG.Col = 38:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "NOV-OCT"
   grd_LisEEBG.Col = 39:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "NOV-OCT"
   grd_LisEEBG.Col = 40:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "DIC-NOV"
   grd_LisEEBG.Col = 41:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "DIC-NOV"
   
   'Segunda linea
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.Col = 4:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'ENERO
   grd_LisEEBG.Col = 5:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'FEBRERO
   grd_LisEEBG.Col = 6:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'MARZO
   grd_LisEEBG.Col = 7:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'ABRIL
   grd_LisEEBG.Col = 8:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'MAYO
   grd_LisEEBG.Col = 9:   grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'JUNIO
   grd_LisEEBG.Col = 10:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'JULIO
   grd_LisEEBG.Col = 11:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'AGOSTO
   grd_LisEEBG.Col = 12:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'SETIEMBRE
   grd_LisEEBG.Col = 13:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'OCTUBRE
   grd_LisEEBG.Col = 14:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'NOVIEMBRE
   grd_LisEEBG.Col = 15:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = ""   'DICIEMBRE
   grd_LisEEBG.Col = 20:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 21:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 22:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 23:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 24:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 25:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 26:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 27:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 28:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 29:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 30:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 31:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 32:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 33:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 34:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 35:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 36:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 37:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 38:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 39:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   grd_LisEEBG.Col = 40:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "ORIGEN"
   grd_LisEEBG.Col = 41:  grd_LisEEBG.CellAlignment = flexAlignCenterCenter:  grd_LisEEBG.Text = "APLICACION"
   
   'Tercera fila
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   
   With grd_LisEEBG
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 4
      .FixedRows = 2
   End With
   
   'Detalle
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.Col = 2
   grd_LisEEBG.Text = "T"
 
   'Titulo
   grd_LisEEBG.Col = 3
   grd_LisEEBG.CellFontName = "Arial"
   grd_LisEEBG.CellFontSize = 8
   grd_LisEEBG.Text = "ACTIVO"
   
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         If Trim(g_rst_Princi!INDTIPO) <> "L" Then
            If Trim(g_rst_Princi!INDTIPO) = "B" Then GoTo SALTO
            grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
            grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
            
            grd_LisEEBG.Col = 0
            grd_LisEEBG.Text = Trim(g_rst_Princi!GRUPO)
            
            grd_LisEEBG.Col = 1
            grd_LisEEBG.Text = Trim(g_rst_Princi!SUBGRP)
            
            grd_LisEEBG.Col = 2
            grd_LisEEBG.Text = Trim(g_rst_Princi!INDTIPO)
            
            If Trim(g_rst_Princi!INDTIPO) = "S" Or Trim(g_rst_Princi!INDTIPO) = "N" Or Trim(g_rst_Princi!INDTIPO) = "R" Then
               grd_LisEEBG.Col = 3
               grd_LisEEBG.Text = Space(5) & Trim(g_rst_Princi!NOMSUBGRP)
            ElseIf Trim(g_rst_Princi!INDTIPO) = "L" Then
               grd_LisEEBG.Col = 3
               grd_LisEEBG.Text = ""
            ElseIf Trim(g_rst_Princi!INDTIPO) = "D" Then
               grd_LisEEBG.Col = 3
               grd_LisEEBG.Text = Trim(g_rst_Princi!NOMGRUPO)
            ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "T" Or Trim(g_rst_Princi!INDTIPO) = "A" Then
               grd_LisEEBG.Col = 3
               grd_LisEEBG.Text = Trim(g_rst_Princi!NOMGRUPO)
            End If
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 4
            grd_LisEEBG.Text = Format(g_rst_Princi!MES01, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 5
            grd_LisEEBG.Text = Format(g_rst_Princi!MES02, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 6
            grd_LisEEBG.Text = Format(g_rst_Princi!MES03, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 7
            grd_LisEEBG.Text = Format(g_rst_Princi!MES04, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 8
            grd_LisEEBG.Text = Format(g_rst_Princi!MES05, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 9
            grd_LisEEBG.Text = Format(g_rst_Princi!MES06, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 10
            grd_LisEEBG.Text = Format(g_rst_Princi!MES07, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 11
            grd_LisEEBG.Text = Format(g_rst_Princi!MES08, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 12
            grd_LisEEBG.Text = Format(g_rst_Princi!MES09, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 13
            grd_LisEEBG.Text = Format(g_rst_Princi!MES10, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 14
            grd_LisEEBG.Text = Format(g_rst_Princi!MES11, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 15
            grd_LisEEBG.Text = Format(g_rst_Princi!MES12, "###,###,###,##0.00")
            grd_LisEEBG.CellAlignment = flexAlignRightCenter
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
'            grd_LisEEBG.Col = 16
'            grd_LisEEBG.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0.00")
'            grd_LisEEBG.CellAlignment = flexAlignRightCenter
'            grd_LisEEBG.CellFontName = "Arial"
'            grd_LisEEBG.CellFontSize = 8
'            grd_LisEEBG.Col = 17
'            grd_LisEEBG.Text = Trim(g_rst_Princi!NOMGRUPO)
'            grd_LisEEBG.CellFontName = "Arial"
'            grd_LisEEBG.CellFontSize = 8
'            grd_LisEEBG.Col = 18
'            grd_LisEEBG.Text = Trim(g_rst_Princi!NOMSUBGRP & "")
'            grd_LisEEBG.CellFontName = "Arial"
'            grd_LisEEBG.CellFontSize = 8
            
            grd_LisEEBG.Col = 19
            grd_LisEEBG.Text = Trim(g_rst_Princi!CTATIPO & "")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
            
            If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
               grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
               grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
               
               grd_LisEEBG.Col = 3
               grd_LisEEBG.CellFontName = "Arial"
               grd_LisEEBG.CellFontSize = 8
               grd_LisEEBG.Text = "PASIVO"
            End If
            
         Else
            grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
            grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
         End If

SALTO:
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_LisEEBG.Redraw = True
   If grd_LisEEBG.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEBG)
      Call fs_Activa(True)
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub fs_Setea_Columnas()
   'Ancho de columnas
   grd_LisEEBG.Cols = 42
   grd_LisEEBG.ColWidth(0) = 0       ' GRUPO
   grd_LisEEBG.ColWidth(1) = 0       ' COD SUBGRUPO
   grd_LisEEBG.ColWidth(2) = 0       ' INDICA TIPO
   grd_LisEEBG.ColWidth(3) = 4200    ' DESCRIPCION
   grd_LisEEBG.ColWidth(4) = 0       ' MES 1
   grd_LisEEBG.ColWidth(5) = 0       ' MES 2
   grd_LisEEBG.ColWidth(6) = 0       ' MES 3
   grd_LisEEBG.ColWidth(7) = 0       ' MES 4
   grd_LisEEBG.ColWidth(8) = 0       ' MES 5
   grd_LisEEBG.ColWidth(9) = 0       ' MES 6
   grd_LisEEBG.ColWidth(10) = 0      ' MES 7
   grd_LisEEBG.ColWidth(11) = 0      ' MES 8
   grd_LisEEBG.ColWidth(12) = 0      ' MES 9
   grd_LisEEBG.ColWidth(13) = 0      ' MES 10
   grd_LisEEBG.ColWidth(14) = 0      ' MES 11
   grd_LisEEBG.ColWidth(15) = 0      ' MES 12
   grd_LisEEBG.ColWidth(16) = 0      ' ACUMULADO 930
   grd_LisEEBG.ColWidth(17) = 0      ' NOMBRE GRUPO
   grd_LisEEBG.ColWidth(18) = 0      ' NOMBRE SUBGRUPO
   grd_LisEEBG.ColWidth(19) = 0      ' CUENTA TIPO (D,H)
   grd_LisEEBG.ColWidth(20) = 1290   ' VARIACION ENE-FEB - ORIGEN
   grd_LisEEBG.ColWidth(21) = 1290   ' VARIACION ENE-FEB - VARIACION
   grd_LisEEBG.ColWidth(22) = 1290   ' VARIACION FEB-MAR - ORIGEN
   grd_LisEEBG.ColWidth(23) = 1290   ' VARIACION FEB-MAR - VARIACION
   grd_LisEEBG.ColWidth(24) = 1290   ' VARIACION MAR-ABR - ORIGEN
   grd_LisEEBG.ColWidth(25) = 1290   ' VARIACION MAR-ABR - VARIACION
   grd_LisEEBG.ColWidth(26) = 1290   ' VARIACION ABR-MAY - ORIGEN
   grd_LisEEBG.ColWidth(27) = 1290   ' VARIACION ABR-MAY - VARIACION
   grd_LisEEBG.ColWidth(28) = 1290   ' VARIACION MAY-JUN - ORIGEN
   grd_LisEEBG.ColWidth(29) = 1290   ' VARIACION MAY-JUN - VARIACION
   grd_LisEEBG.ColWidth(30) = 1290   ' VARIACION JUN-JUL - ORIGEN
   grd_LisEEBG.ColWidth(31) = 1290   ' VARIACION JUN-JUL - VARIACION
   grd_LisEEBG.ColWidth(32) = 1290   ' VARIACION JUL-AGO - ORIGEN
   grd_LisEEBG.ColWidth(33) = 1290   ' VARIACION JUL-AGO - VARIACION
   grd_LisEEBG.ColWidth(34) = 1290   ' VARIACION AGO-SET - ORIGEN
   grd_LisEEBG.ColWidth(35) = 1290   ' VARIACION AGO-SET - VARIACION
   grd_LisEEBG.ColWidth(36) = 1290   ' VARIACION SET-OCT - ORIGEN
   grd_LisEEBG.ColWidth(37) = 1290   ' VARIACION SET-OCT - VARIACION
   grd_LisEEBG.ColWidth(38) = 1290   ' VARIACION OCT-NOV - ORIGEN
   grd_LisEEBG.ColWidth(39) = 1290   ' VARIACION OCT-NOV - VARIACION
   grd_LisEEBG.ColWidth(40) = 1290   ' VARIACION NOV-DIC - ORIGEN
   grd_LisEEBG.ColWidth(41) = 1290   ' VARIACION NOV-DIC - VARIACION
   
   grd_LisEEBG.ColAlignment(3) = flexAlignLeftCenter
   grd_LisEEBG.ColAlignment(4) = flexAlignCenterCenter
   grd_LisEEBG.ColAlignment(5) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(6) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(7) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(8) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(9) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(10) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(11) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(12) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(13) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(14) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(15) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(20) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(21) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(22) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(23) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(24) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(25) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(26) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(27) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(28) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(29) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(30) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(31) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(32) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(33) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(34) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(35) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(36) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(37) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(38) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(39) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(40) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(41) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisEEBG)
End Sub

Private Sub fs_CargaDatos_Variacion()
Dim r_int_PerMes        As Integer
Dim r_int_Conta1        As Integer
Dim r_int_Conta2        As Integer
Dim r_int_ColBal        As Integer
Dim r_int_ColVar        As Integer
Dim r_int_NumReg        As Integer
Dim r_dbl_SumOri        As Double
Dim r_dbl_SumApl        As Double

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   If r_int_PerMes = 1 Then
      Exit Sub
   End If
   
   r_int_ColBal = 4
   r_int_ColVar = 20
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   
   'Recorre columnas de meses empezando en febrero
   For r_int_Conta1 = 2 To r_int_PerMes
      'Verifica que se haya procesado el cierre para el mes
      r_int_NumReg = 0
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & CStr(r_int_Conta1) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & CStr(ipp_PerAno) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_int_NumReg = g_rst_Princi!CONTADOR
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_int_NumReg > 0 Then
         r_dbl_SumOri = 0
         r_dbl_SumApl = 0
         
         'Recorre filas de cada mes seleccionado
         For r_int_Conta2 = 1 To grd_LisEEBG.Rows - 1
            
            'Determina si el fila es para evaluar ("S")
            If UCase(grd_LisEEBG.TextMatrix(r_int_Conta2, 2)) = "N" Or UCase(grd_LisEEBG.TextMatrix(r_int_Conta2, 2)) = "R" Or UCase(grd_LisEEBG.TextMatrix(r_int_Conta2, 2)) = "S" Then
               
               'Determina tipo de cuenta
               If UCase(grd_LisEEBG.TextMatrix(r_int_Conta2, 19)) = "D" Then
                  '----DEBE
                  'Columna Origen
                  If CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) > CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) Then
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar) = Format(CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) - CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)), "###,###,##0.00")
                  Else
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar) = Format(0, "###,###,##0.00")
                  End If
                  r_dbl_SumOri = r_dbl_SumOri + CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar))
                  'Columna Aplicacion
                  If CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) < CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) Then
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1) = Format(CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) - CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)), "###,###,##0.00")
                  Else
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1) = Format(0, "###,###,##0.00")
                  End If
                  r_dbl_SumApl = r_dbl_SumApl + CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1))
               Else
                  '----HABER
                  'Columna Origen
                  If CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) < CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) Then
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar) = Format(CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) - CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)), "###,###,##0.00")
                  Else
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar) = Format(0, "###,###,##0.00")
                  End If
                  r_dbl_SumOri = r_dbl_SumOri + CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar))
                  'Columna Aplicacion
                  If CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) > CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)) Then
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1) = Format(CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal + 1)) - CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColBal)), "###,###,##0.00")
                  Else
                     grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1) = Format(0, "###,###,##0.00")
                  End If
                  r_dbl_SumApl = r_dbl_SumApl + CDbl(grd_LisEEBG.TextMatrix(r_int_Conta2, r_int_ColVar + 1))
               End If
            End If
         Next
         
         grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, r_int_ColVar) = Format(r_dbl_SumOri, "###,###,##0.00")
         grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, r_int_ColVar + 1) = Format(r_dbl_SumApl, "###,###,##0.00")
      End If
      
      r_int_ColBal = r_int_ColBal + 1
      r_int_ColVar = r_int_ColVar + 2
   Next
   
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
   
    r_int_nrofil = 5
    r_int_NoFlLi = 2
    r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
    r_int_PerAno = CInt(ipp_PerAno.Text)
    r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
        'Seteo columnas
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 42
        .Columns("D").ColumnWidth = 0
        .Columns("D").HorizontalAlignment = xlHAlignRight
        .Columns("D").NumberFormat = "###,###,##0.00"
        .Columns("E").ColumnWidth = 0
        .Columns("E").NumberFormat = "###,###,##0.00"
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("F").ColumnWidth = 0
        .Columns("F").NumberFormat = "###,###,##0.00"
        .Columns("F").HorizontalAlignment = xlHAlignRight
        .Columns("G").ColumnWidth = 0
        .Columns("G").NumberFormat = "###,###,##0.00"
        .Columns("G").HorizontalAlignment = xlHAlignRight
        .Columns("H").ColumnWidth = 0
        .Columns("H").NumberFormat = "###,###,##0.00"
        .Columns("H").HorizontalAlignment = xlHAlignRight
        .Columns("I").ColumnWidth = 0
        .Columns("I").NumberFormat = "###,###,##0.00"
        .Columns("I").HorizontalAlignment = xlHAlignRight
        .Columns("J").ColumnWidth = 0
        .Columns("J").NumberFormat = "###,###,##0.00"
        .Columns("J").HorizontalAlignment = xlHAlignRight
        .Columns("K").ColumnWidth = 0
        .Columns("K").NumberFormat = "###,###,##0.00"
        .Columns("K").HorizontalAlignment = xlHAlignRight
        .Columns("L").ColumnWidth = 0
        .Columns("L").NumberFormat = "###,###,##0.00"
        .Columns("L").HorizontalAlignment = xlHAlignRight
        .Columns("M").ColumnWidth = 0
        .Columns("M").NumberFormat = "###,###,##0.00"
        .Columns("M").HorizontalAlignment = xlHAlignRight
        .Columns("N").ColumnWidth = 0
        .Columns("N").NumberFormat = "###,###,##0.00"
        .Columns("N").HorizontalAlignment = xlHAlignRight
        .Columns("O").ColumnWidth = 0
        .Columns("O").NumberFormat = "###,###,##0.00"
        .Columns("O").HorizontalAlignment = xlHAlignRight
        .Columns("P").ColumnWidth = 13.5
        .Columns("P").NumberFormat = "###,###,##0.00"
        .Columns("P").HorizontalAlignment = xlHAlignRight
        .Columns("Q").ColumnWidth = 13.5
        .Columns("Q").NumberFormat = "###,###,##0.00"
        .Columns("Q").HorizontalAlignment = xlHAlignRight
        .Columns("R").ColumnWidth = 13.5
        .Columns("R").NumberFormat = "###,###,##0.00"
        .Columns("R").HorizontalAlignment = xlHAlignRight
        .Columns("S").ColumnWidth = 13.5
        .Columns("S").NumberFormat = "###,###,##0.00"
        .Columns("S").HorizontalAlignment = xlHAlignRight
        .Columns("T").ColumnWidth = 13.5
        .Columns("T").NumberFormat = "###,###,##0.00"
        .Columns("T").HorizontalAlignment = xlHAlignRight
        .Columns("U").ColumnWidth = 13.5
        .Columns("U").NumberFormat = "###,###,##0.00"
        .Columns("U").HorizontalAlignment = xlHAlignRight
        .Columns("V").ColumnWidth = 13.5
        .Columns("V").NumberFormat = "###,###,##0.00"
        .Columns("V").HorizontalAlignment = xlHAlignRight
        .Columns("W").ColumnWidth = 13.5
        .Columns("W").NumberFormat = "###,###,##0.00"
        .Columns("W").HorizontalAlignment = xlHAlignRight
        .Columns("X").ColumnWidth = 13.5
        .Columns("X").NumberFormat = "###,###,##0.00"
        .Columns("X").HorizontalAlignment = xlHAlignRight
        .Columns("Y").ColumnWidth = 13.5
        .Columns("Y").NumberFormat = "###,###,##0.00"
        .Columns("Y").HorizontalAlignment = xlHAlignRight
        .Columns("Z").ColumnWidth = 13.5
        .Columns("Z").NumberFormat = "###,###,##0.00"
        .Columns("Z").HorizontalAlignment = xlHAlignRight
        .Columns("AA").ColumnWidth = 13.5
        .Columns("AA").NumberFormat = "###,###,##0.00"
        .Columns("AA").HorizontalAlignment = xlHAlignRight
        .Columns("AB").ColumnWidth = 13.5
        .Columns("AB").NumberFormat = "###,###,##0.00"
        .Columns("AB").HorizontalAlignment = xlHAlignRight
        .Columns("AC").ColumnWidth = 13.5
        .Columns("AC").NumberFormat = "###,###,##0.00"
        .Columns("AC").HorizontalAlignment = xlHAlignRight
        .Columns("AD").ColumnWidth = 13.5
        .Columns("AD").NumberFormat = "###,###,##0.00"
        .Columns("AD").HorizontalAlignment = xlHAlignRight
        .Columns("AE").ColumnWidth = 13.5
        .Columns("AE").NumberFormat = "###,###,##0.00"
        .Columns("AE").HorizontalAlignment = xlHAlignRight
        .Columns("AF").ColumnWidth = 13.5
        .Columns("AF").NumberFormat = "###,###,##0.00"
        .Columns("AF").HorizontalAlignment = xlHAlignRight
        .Columns("AG").ColumnWidth = 13.5
        .Columns("AG").NumberFormat = "###,###,##0.00"
        .Columns("AG").HorizontalAlignment = xlHAlignRight
        .Columns("AH").ColumnWidth = 13.5
        .Columns("AH").NumberFormat = "###,###,##0.00"
        .Columns("AH").HorizontalAlignment = xlHAlignRight
        .Columns("AI").ColumnWidth = 13.5
        .Columns("AI").NumberFormat = "###,###,##0.00"
        .Columns("AI").HorizontalAlignment = xlHAlignRight
        .Columns("AJ").ColumnWidth = 13.5
        .Columns("AJ").NumberFormat = "###,###,##0.00"
        .Columns("AJ").HorizontalAlignment = xlHAlignRight
        .Columns("AK").ColumnWidth = 13.5
        .Columns("AK").NumberFormat = "###,###,##0.00"
        .Columns("AK").HorizontalAlignment = xlHAlignRight
        .Range(.Cells(1, 1), .Cells(99, 37)).Font.Name = "Calibri"
        .Range(.Cells(1, 1), .Cells(99, 37)).Font.Size = 11
        
        'Carga cabecera
        .Cells(1, 2) = "REPORTE DE ORIGEN Y APLICACION DEL BALANCE GENERAL"
        .Range(.Cells(1, 2), .Cells(1, 3)).Merge
        .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
        .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
        .Range(.Cells(2, 2), .Cells(2, 3)).Merge
        .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
        .Cells(3, 2) = "( En Soles )"
        
        .Cells(r_int_nrofil, 16) = "'" & "FEB-ENE / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 16), .Cells(r_int_nrofil, 17)).Merge
        .Cells(r_int_nrofil, 18) = "'" & "MAR-FEB / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 18), .Cells(r_int_nrofil, 19)).Merge
        .Cells(r_int_nrofil, 20) = "'" & "ABR-MAR / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 20), .Cells(r_int_nrofil, 21)).Merge
        .Cells(r_int_nrofil, 22) = "'" & "MAY-ABR / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 22), .Cells(r_int_nrofil, 23)).Merge
        .Cells(r_int_nrofil, 24) = "'" & "JUN-MAY / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 24), .Cells(r_int_nrofil, 25)).Merge
        .Cells(r_int_nrofil, 26) = "'" & "JUL-JUN / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 26), .Cells(r_int_nrofil, 27)).Merge
        .Cells(r_int_nrofil, 28) = "'" & "AGO-JUL / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 28), .Cells(r_int_nrofil, 29)).Merge
        .Cells(r_int_nrofil, 30) = "'" & "SET-AGO / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 30), .Cells(r_int_nrofil, 31)).Merge
        .Cells(r_int_nrofil, 32) = "'" & "OCT-SET / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 32), .Cells(r_int_nrofil, 33)).Merge
        .Cells(r_int_nrofil, 34) = "'" & "NOV-OCT / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 34), .Cells(r_int_nrofil, 35)).Merge
        .Cells(r_int_nrofil, 36) = "'" & "DIC-NOV / " & CStr(r_int_PerAno)
        .Range(.Cells(r_int_nrofil, 36), .Cells(r_int_nrofil, 37)).Merge
        
        r_int_nrofil = r_int_nrofil + 1
        .Cells(r_int_nrofil, 4) = "'" & "ENE / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 5) = "'" & "FEB / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 6) = "'" & "MAR / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 7) = "'" & "ABR / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 8) = "'" & "MAY / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 9) = "'" & "JUN / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 10) = "'" & "JUL / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 11) = "'" & "AGO / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 12) = "'" & "SET / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 13) = "'" & "OCT / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 14) = "'" & "NOV / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 15) = "'" & "DIC / " & Right(r_int_PerAno, 2)
        .Cells(r_int_nrofil, 16) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 17) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 18) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 19) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 20) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 21) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 22) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 23) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 24) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 25) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 26) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 27) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 28) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 29) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 30) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 31) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 32) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 33) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 34) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 35) = "'" & "APLICACION"
        .Cells(r_int_nrofil, 36) = "'" & "ORIGEN"
        .Cells(r_int_nrofil, 37) = "'" & "APLICACION"
        
        .Range(.Cells(5, 2), .Cells(6, 37)).HorizontalAlignment = xlHAlignCenter
        .Range(.Cells(5, 2), .Cells(6, 37)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(5, 2), .Cells(6, 37)).Font.Bold = True
                 
        r_int_nrofil = r_int_nrofil + 1
        .Range(.Cells(4, 4), .Cells(4, 37)).HorizontalAlignment = xlHAlignCenter
         
        For r_int_NoFlLi = 2 To grd_LisEEBG.Rows - 1
            If Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "G" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "F" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "A" _
                 Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "T" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "X" Then
                'TITULO
                .Cells(r_int_nrofil, 2) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 3)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 37)).Interior.Color = RGB(146, 208, 80)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 37)).Font.Bold = True
            Else
                .Cells(r_int_nrofil, 3) = Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 3))
            End If
             
            .Cells(r_int_nrofil, 4) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 4)
            .Cells(r_int_nrofil, 5) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 5)
            .Cells(r_int_nrofil, 6) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 6)
            .Cells(r_int_nrofil, 7) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 7)
            .Cells(r_int_nrofil, 8) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 8)
            .Cells(r_int_nrofil, 9) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 9)
            .Cells(r_int_nrofil, 10) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 10)
            .Cells(r_int_nrofil, 11) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 11)
            .Cells(r_int_nrofil, 12) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 12)
            .Cells(r_int_nrofil, 13) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 13)
            .Cells(r_int_nrofil, 14) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 14)
            .Cells(r_int_nrofil, 15) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 15)
            .Cells(r_int_nrofil, 16) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 20)
            .Cells(r_int_nrofil, 17) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 21)
            .Cells(r_int_nrofil, 18) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 22)
            .Cells(r_int_nrofil, 19) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 23)
            .Cells(r_int_nrofil, 20) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 24)
            .Cells(r_int_nrofil, 21) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 25)
            .Cells(r_int_nrofil, 22) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 26)
            .Cells(r_int_nrofil, 23) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 27)
            .Cells(r_int_nrofil, 24) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 28)
            .Cells(r_int_nrofil, 25) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 29)
            .Cells(r_int_nrofil, 26) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 30)
            .Cells(r_int_nrofil, 27) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 31)
            .Cells(r_int_nrofil, 28) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 32)
            .Cells(r_int_nrofil, 29) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 33)
            .Cells(r_int_nrofil, 30) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 34)
            .Cells(r_int_nrofil, 31) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 35)
            .Cells(r_int_nrofil, 32) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 36)
            .Cells(r_int_nrofil, 33) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 37)
            .Cells(r_int_nrofil, 34) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 38)
            .Cells(r_int_nrofil, 35) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 39)
            .Cells(r_int_nrofil, 36) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 40)
            .Cells(r_int_nrofil, 37) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 41)
            
            r_int_nrofil = r_int_nrofil + 1
        Next r_int_NoFlLi
   End With
   
   r_obj_Excel.Visible = True
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub


