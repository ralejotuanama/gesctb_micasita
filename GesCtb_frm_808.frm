VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   10155
   ClientTop       =   3555
   ClientWidth     =   5925
   Icon            =   "GesCtb_frm_808.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3405
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   6075
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   6006
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
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
            Left            =   630
            TabIndex        =   9
            Top             =   150
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Balances"
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
            Picture         =   "GesCtb_frm_808.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   780
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_808.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_808.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5220
            Picture         =   "GesCtb_frm_808.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2820
            Top             =   120
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   1245
         Left            =   30
         TabIndex        =   11
         Top             =   2100
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   2196
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
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   450
            Width           =   2265
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   3945
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
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
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   825
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   555
         Left            =   30
         TabIndex        =   15
         Top             =   1500
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   979
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   3945
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   150
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_FecImp     As String
Dim l_str_HorImp     As String
Dim l_str_PerAno     As String
Dim l_str_PerMes     As String
Dim l_lng_NumReg     As Long
Dim l_lng_TotReg     As Long
Dim l_str_CodEmp     As String
Dim l_str_NomEmp     As String
Dim l_str_CodSbs     As String
Dim l_str_rutatx     As String
Dim l_Mar_Izq        As Double
Dim l_Mar_Der        As Double
Dim l_Mar_Sup        As Double
Dim l_Mar_Inf        As Double
Dim l_str_CodAux     As String
    
Private Type g_tpo_Bcient
   Bcient_FecMov     As String
   Bcient_Entdad     As String
   Bcient_Cuenta     As String
   Bcient_SldIni     As String
   Bcient_Credit     As String
   Bcient_Debito     As String
   Bcient_SldFin     As String
   Bcient_Filler     As String
End Type

Private Type g_tpo_BCR
   BCR_Codigo        As String
   BCR_Descri        As String
   BCR_PlcSbs        As String
   BCR_SldAju        As String
   BCR_SldMon        As String
   BCR_SldEqu        As String
   BCR_AjuDif        As String
   BCR_SldIni        As String
   BCR_SldFin        As String
   BCR_CodSec        As String
   BCR_Cuenta        As String
   BCR_Opcion        As String
End Type

Private Type g_tpo_CNTABSI
   CNTABSI_Cuenta  As String
   CNTABSI_Nombre  As String
   CNTABSI_Moneda  As Integer
   CNTABSI_SdoIni  As Double
   CNTABSI_SdoFin  As Double
End Type
   
Private Sub cmd_ExpExc_Click()
   l_str_CodAux = ""

   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      If cmb_TipMon.ListIndex = -1 Then
         MsgBox "Debe seleccionar Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipMon)
         Exit Sub
      End If
   End If
   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de Exportar el reporte de " & Me.cmb_TipRep.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   l_str_PerAno = Format(ipp_PerAno.Text, "0000")
   l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "  WHERE EMPGRP_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_str_CodEmp = g_rst_Princi!EMPGRP_CODIGO
      l_str_CodSbs = Trim(g_rst_Princi!EMPGRP_CODSBS)
      l_str_NomEmp = Trim(g_rst_Princi!EMPGRP_RAZSOC)
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
   
      'Procesa infomación
      If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 0 Or cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 1 Or cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 2 Or cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 3 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " USP_BALANCE_COMPROBACION_2("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
         g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
         g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
         g_str_Parame = g_str_Parame & "'" & "01/" & l_str_PerMes & "/" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & "'" & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "')"
          
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
            Screen.MousePointer = 0
            Exit Sub
         End If
         'Genera el Excel
         Call fs_GenExc
      Else
         MsgBox "Esta opción se encuentra aún en desarrollo.", vbCritical, modgen_g_str_NomPlt
      End If
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
      
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      
      'Procesa infomación
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_SUCAVE_BCIENT_2( "
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
      g_str_Parame = g_str_Parame & "'" & l_str_CodSbs & "' , "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & "REPORTE " & UCase(cmb_TipRep.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & "01/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "') "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Genera el Excel
      Call fs_GenExc_BCient(l_str_CodSbs)
            
      If (Len(Trim(l_str_CodAux)) > 0) Then
         MsgBox "Proceso Terminado. El Archivo Plano BCIENT se generó en " & moddat_g_str_RutLoc & Chr(13) & "En el Archivo BSI faltan codigo sectorial: " & l_str_CodAux, vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "Proceso Terminado. El Archivo Plano BCIENT se generó en " & moddat_g_str_RutLoc, vbInformation, modgen_g_str_NomPlt
      End If
      
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      Dim as_chk_origen As String
      Dim as_cod_origen As String
      as_chk_origen = "N"
      as_cod_origen = ""
      
      'Procesa infomación
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_BALANCE_SECTORIZADO_MAIN_2( "
      g_str_Parame = g_str_Parame & "'" & "01/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & "REPORTE " & UCase(cmb_TipRep.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If

      'Genera el Excel
      Call fs_GenExc_BCR
      MsgBox "Proceso Terminado. El Archivo Plano se generó en " & moddat_g_str_RutLoc, vbInformation, modgen_g_str_NomPlt
                    
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
            
      'Procesa infomación
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_BALANCE_SITFIN_ESTRES("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
      g_str_Parame = g_str_Parame & "'" & l_str_CodSbs & "' , "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'ESTADO DE SITUACION FINANCIERA', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & "01/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno & "', 1)"
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Genera el Excel
      Call fs_GenExc_SitFin
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
    
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_RPT_TPR_CARFIA ("
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'REPORTE DE LIMITE' , "
      g_str_Parame = g_str_Parame & Format(ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno, "YYYYMMDD") & " , "
      g_str_Parame = g_str_Parame & "0 , "
      g_str_Parame = g_str_Parame & "2 , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CONCILIACION_CTACNTB("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & UCase(cmb_TipRep.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & "01/" & l_str_PerMes & "/" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & "'" & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno & "', 1)"
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
        
      'Genera el Excel
      Call fs_GenExc_ConCtaCtb
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 6 Then
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_NOTAS_EEFF("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & UCase(cmb_TipRep.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "')"
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
        
      'Genera el Excel
      Call fs_GenExc_NotasEEFF(g_rst_Princi)
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 7 Then
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_VALIDACION_CTACNTB("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "' , "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & "REPORTE DE " & UCase(cmb_TipRep.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "')"
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Genera el Excel
      Call fs_GenExc_ValidaCtaCntb(g_rst_Princi)
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      If cmb_TipMon.ListIndex = -1 Then
            MsgBox "Debe seleccionar Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipMon)
            Exit Sub
      End If
   End If
   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de Imprimir el reporte de " & cmb_TipRep.Text & "?, asegúrese de haber ejecutado la opción Exportar a Excel.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   l_str_PerAno = Format(ipp_PerAno.Text, "0000")
   l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00")
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
   
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "CTB_BALCOM"
      crp_Imprim.SelectionFormula = "{CTB_BALCOM.BALCOM_TIPBAL} = " & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_BALCOM.BALCOM_PERANO} = " & l_str_PerAno & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_BALCOM.BALCOM_PERMES} = " & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_BALCOM.BALCOM_USUCRE} = '" & modgen_g_str_CodUsu & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_BALCOM.BALCOM_TERCRE} = '" & modgen_g_str_NombPC & "' "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_33.RPT"
      crp_Imprim.Destination = crptToWindow
      crp_Imprim.Action = 1
      
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_TABLA_TEMP"
      crp_Imprim.SelectionFormula = "{RPT_TABLA_TEMP.RPT_DESCRI} = '" & Trim(l_str_CodSbs) & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_PERMES} = " & CInt(l_str_PerMes) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_PERANO} = " & CInt(l_str_PerAno) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_USUCRE} = '" & modgen_g_str_CodUsu & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_NOMBRE} = 'REPORTE BCIENT' AND "
      If l_str_PerMes = 12 Then
         crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) NOT LIKE '6%' AND "
      End If
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '201' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '202' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '301' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '401' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '402' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '501' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '502' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '701' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '801' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & " TRIM({RPT_TABLA_TEMP.RPT_CODIGO}) <> '802' "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ctb_rptsol_43.rpt"
      crp_Imprim.Destination = crptToWindow
      crp_Imprim.Action = 1
      
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_TABLA_TEMP"
      crp_Imprim.DataFiles(1) = "RPT_SUBGRUPO"
      crp_Imprim.SelectionFormula = "{RPT_SUBGRUPO.GRUPO} = 'unico' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_PERMES} = " & CInt(l_str_PerMes) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_PERANO} = " & CInt(l_str_PerAno) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_USUCRE} = '" & modgen_g_str_CodUsu & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_NOMBRE} = 'REPORTE BALANCE SECTORIAL BCR' "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ctb_rptsol_44.rpt"
      crp_Imprim.Destination = crptToWindow
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
   Call fs_Recorset_nc
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipMon)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_Period.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
   Call gs_Carga_TipMon(cmb_TipMon)
   ipp_PerAno = Mid(date, 7, 4)
   
   cmb_TipRep.AddItem "BALANCE COMPROBACION SBS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(1)
   cmb_TipRep.AddItem "BCIENT"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(2)
   cmb_TipRep.AddItem "BALANCE SECTORIAL BCR"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(3)
   cmb_TipRep.AddItem "BALANCE SITUACION Y ESTADO RESULTADOS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(4)
   cmb_TipRep.AddItem "CONCILIACION CON CUENTAS CONTABLES"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(5)
   cmb_TipRep.AddItem "NOTAS A LOS ESTADOS FINANCIEROS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(6)
   cmb_TipRep.AddItem "VALIDACIONES CONTABLES"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(7)
   
End Sub
 
Private Sub gs_Carga_TipMon(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM CTB_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY TIPMON_CODIGO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem CInt(g_rst_Genera!TIPMON_CODIGO) & " - " & Trim$(g_rst_Genera!TIPMON_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!TIPMON_CODIGO)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_GenExc_BCient(ByVal l_str_CodSbs As String)
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_bol_FlgMon     As Boolean
Dim r_arr_Matriz()   As g_tpo_Bcient
Dim r_str_CadIni     As String
Dim r_str_CadFin     As String
   
   ReDim r_arr_Matriz(0)
   
   g_str_Parame = ""
   g_str_Parame = " SELECT RPT_CODIGO CUENTA, RPT_VALNUM01 SDOINI, RPT_VALNUM02 DEBITO, RPT_VALNUM03 CREDITO, RPT_VALNUM04 SDOFIN, RPT_MONEDA CODMON, RPT_VALCAD01 FILLER FROM RPT_TABLA_TEMP WHERE "
   g_str_Parame = g_str_Parame & " RPT_PERMES = " & CInt(l_str_PerMes) & " AND RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & " RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & " RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' AND "
   g_str_Parame = g_str_Parame & " (RPT_VALNUM01 <> 0 OR RPT_VALNUM02 <> 0 OR RPT_VALNUM03 <> 0 OR RPT_VALNUM04 <> 0) AND "
   If l_str_PerMes = 12 Then
      g_str_Parame = g_str_Parame & " TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' AND TRIM(RPT_CODIGO) <> '301' AND " 'AND TRIM(RPT_CODIGO) NOT LIKE '6%'
      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '401' AND TRIM(RPT_CODIGO) <> '402' AND TRIM(RPT_CODIGO) <> '501' AND TRIM(RPT_CODIGO) <> '502' AND "
      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '701' AND TRIM(RPT_CODIGO) <> '801' AND TRIM(RPT_CODIGO) <> '802' "
   Else
      g_str_Parame = g_str_Parame & " TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' AND TRIM(RPT_CODIGO) <> '201' AND TRIM(RPT_CODIGO) <> '202' AND TRIM(RPT_CODIGO) <> '301' AND "
      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '401' AND TRIM(RPT_CODIGO) <> '402' AND TRIM(RPT_CODIGO) <> '501' AND TRIM(RPT_CODIGO) <> '502' AND "
      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '701' AND TRIM(RPT_CODIGO) <> '801' AND TRIM(RPT_CODIGO) <> '802' "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY RPT_MONEDA ASC, RPT_CODIGO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   r_obj_Excel.Sheets(1).Name = "BALANCE DE COMPROBACIÓN SBS"
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 1), .Cells(1, 8)).Merge
      .Range(.Cells(1, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "BALANCE DE COMPROBACIÓN SBS"
      
      .Cells(2, 1) = "Año: " & l_str_PerAno & "      " & "Mes: " & l_str_PerMes
      .Range(.Cells(2, 1), .Cells(2, 8)).Merge
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 11
      
      .Cells(3, 1) = Trim(l_str_NomEmp)
      .Cells(3, 2) = " Código SBS: " & Trim(l_str_CodSbs)
      .Cells(3, 2).Font.Italic = True
                  
      .Columns("A").ColumnWidth = 21
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").ColumnWidth = 60
      .Columns("C").ColumnWidth = 19
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 19
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 19
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 19
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("A").NumberFormat = "@"
      
      .Cells(5, 1) = "CUENTA"
      .Cells(5, 2) = "DESCRIPCIÓN"
      .Cells(5, 3) = "SALDO INICIAL"
      .Cells(5, 4) = "DÉBITO"
      .Cells(5, 5) = "CRÉDITO"
      .Cells(5, 6) = "SALDO FINAL"
      .Cells(6, 6) = "Moneda Nacional"
   
      .Range(.Cells(5, 1), .Cells(6, 6)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(6, 6)).HorizontalAlignment = xlHAlignCenter
      .Cells(3, 1).HorizontalAlignment = xlHAlignCenter
   
      g_rst_Princi.MoveFirst
      r_int_ConVer = 7
   
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!CODMON = 2 And r_bol_FlgMon = False Then
            r_int_ConVer = r_int_ConVer + 3
            
            .Cells(r_int_ConVer - 2, 1) = "CUENTA"
            .Cells(r_int_ConVer - 2, 2) = "DESCRIPCIÓN"
            .Cells(r_int_ConVer - 2, 3) = "SALDO INICIAL"
            .Cells(r_int_ConVer - 2, 4) = "DÉBITO"
            .Cells(r_int_ConVer - 2, 5) = "CRÉDITO"
            .Cells(r_int_ConVer - 2, 6) = "SALDO FINAL"
            .Cells(r_int_ConVer - 1, 6) = "Dólares Americanos"
            .Range(.Cells(r_int_ConVer - 2, 1), .Cells(r_int_ConVer - 1, 6)).Font.Bold = True
            .Range(.Cells(r_int_ConVer - 2, 1), .Cells(r_int_ConVer - 1, 6)).HorizontalAlignment = xlHAlignCenter
            r_bol_FlgMon = True
         End If
           
         .Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!Cuenta)
         .Cells(r_int_ConVer, 2) = modsec_gf_Buscar_NomCta(g_rst_Princi!Cuenta)
         .Cells(r_int_ConVer, 3) = CDbl(Format(IIf(IsNull(g_rst_Princi!SDOINI), 0, g_rst_Princi!SDOINI), "###,###,##0.00"))
         .Cells(r_int_ConVer, 4) = CDbl(Format(IIf(IsNull(g_rst_Princi!DEBITO), 0, g_rst_Princi!DEBITO), "###,###,##0.00"))
         .Cells(r_int_ConVer, 5) = CDbl(Format(IIf(IsNull(g_rst_Princi!CREDITO), 0, g_rst_Princi!CREDITO), "###,###,##0.00"))
         .Cells(r_int_ConVer, 6) = CDbl(Format(IIf(IsNull(g_rst_Princi!SDOFIN), 0, g_rst_Princi!SDOFIN), "###,###,##0.00"))
           
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   'Para mostrar los datos en el Archivo de texto
   g_str_Parame = ""
   g_str_Parame = " SELECT RPT_CODIGO CUENTA, RPT_VALNUM01 SDOINI, RPT_VALNUM02 DEBITO, RPT_VALNUM03 CREDITO, RPT_VALNUM04 SDOFIN, RPT_MONEDA CODMON, RPT_VALCAD01 FILLER FROM RPT_TABLA_TEMP WHERE "
   g_str_Parame = g_str_Parame & " RPT_PERMES = " & CInt(l_str_PerMes) & " AND RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & " RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & " RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' AND "
   g_str_Parame = g_str_Parame & " (RPT_VALNUM01 <> 0 OR RPT_VALNUM02 <> 0 OR RPT_VALNUM03 <> 0 OR RPT_VALNUM04 <> 0) AND "
   g_str_Parame = g_str_Parame & " TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "'"
   
   g_str_Parame = g_str_Parame & " ORDER BY RPT_MONEDA ASC, RPT_CODIGO ASC "
   
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
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!CODMON = 1 And Len(Trim(g_rst_Princi!Cuenta)) <= 10 And modsec_gf_Buscar_NomCtaHab(g_rst_Princi!Cuenta) = True Then
         ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_FecMov = l_str_PerAno & l_str_PerMes
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_Entdad = Trim(l_str_CodSbs)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_Cuenta = Trim(g_rst_Princi!Cuenta)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_SldIni = IIf(IsNull(g_rst_Princi!SDOINI), 0, g_rst_Princi!SDOINI)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_Debito = IIf(IsNull(g_rst_Princi!DEBITO), 0, g_rst_Princi!DEBITO)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_Credit = IIf(IsNull(g_rst_Princi!CREDITO), 0, g_rst_Princi!CREDITO)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_SldFin = IIf(IsNull(g_rst_Princi!SDOFIN), 0, g_rst_Princi!SDOFIN)
         r_arr_Matriz(UBound(r_arr_Matriz)).Bcient_Filler = IIf(IsNull(g_rst_Princi!Filler), 0, g_rst_Princi!Filler)
      End If
   
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   Call fs_GeneraArchivo_BCient(r_arr_Matriz)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '-------------GENERACION REPORTE BSI-----------------------
   '----------------------------------------------------------
   Dim r_arr_MtzBsi()   As g_tpo_BCR
   Dim r_arr_MtzCta()   As g_tpo_CNTABSI
   Dim r_str_Cadena     As String
   
   ReDim r_arr_MtzBsi(0)
   ReDim r_arr_MtzCta(0)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT RPAD(RPT_CODIGO,20,'0') CUENTA, RPT_MONEDA CODMON, RPT_CODIGO, RPT_VALNUM01 SALDOINICIAL, "
   g_str_Parame = g_str_Parame & "       RPT_VALNUM04 SALDOFINAL, NVL(CODSEC,'XXXXXXXXXXXX') CODSEC, OPCION, ORDEN, RPT_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM ("
   g_str_Parame = g_str_Parame & "             SELECT RPT_CODIGO, RPT_VALNUM01, RPT_VALNUM04, RPT_MONEDA, TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION,"
   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC AS CODSEC,NVL(C.BSISECT_ORDEN,'A') AS ORDEN, RPT_DESCRI"
   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(B.BSICNTA_CUENTA) AND C.BSISECT_SITUAC = 1"
   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' "
   g_str_Parame = g_str_Parame & "        UNION "
   g_str_Parame = g_str_Parame & "             SELECT CASE TRIM(B.BSICNTA_CUENTA) WHEN '11M70900000000' THEN SUBSTR(RPT_CODIGO,1,6) ELSE RPT_CODIGO END AS RPT_CODIGO,"
   g_str_Parame = g_str_Parame & "                    RPT_VALNUM01, RPT_VALNUM04, RPT_MONEDA, TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION,"
   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC CODSEC, NVL(C.BSISECT_ORDEN,'A') ORDEN, RPT_DESCRI"
   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,3),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
   g_str_Parame = g_str_Parame & "              INNER JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(SUBSTR(RPT_CODIGO,1,2)||'M'||SUBSTR(RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1))"
   g_str_Parame = g_str_Parame & "                    AND C.BSISECT_SITUAC = 2"
   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' "
   g_str_Parame = g_str_Parame & "        UNION "
   g_str_Parame = g_str_Parame & "             SELECT RPT_CODIGO, DECODE(BSISECT_SLDINI,0,RPT_VALNUM01,BSISECT_SLDINI) RPT_VALNUM01, "
   g_str_Parame = g_str_Parame & "                    DECODE(BSISECT_SLDFIN,0,RPT_VALNUM04,BSISECT_SLDFIN) RPT_VALNUM04, RPT_MONEDA, "
   g_str_Parame = g_str_Parame & "                    TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION, "
   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC CODSEC, NVL(C.BSISECT_ORDEN,'A') ORDEN, RPT_DESCRI "
   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,3),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
   g_str_Parame = g_str_Parame & "              INNER JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(SUBSTR(RPT_CODIGO,1,2)||'M'||SUBSTR(RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1))"
   g_str_Parame = g_str_Parame & "                    AND C.BSISECT_SITUAC = 3"
   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "')"
   g_str_Parame = g_str_Parame & " WHERE SUBSTR(RPT_CODIGO,1,1) NOT IN (7,8) "
   g_str_Parame = g_str_Parame & "     ORDER BY RPT_MONEDA ASC, RPT_CODIGO ASC, ORDEN ASC, CODSEC "

'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT RPAD(RPT_CODIGO,20,'0') CUENTA, RPT_MONEDA CODMON, RPT_CODIGO, RPT_VALNUM01 SALDOINICIAL, "
'   g_str_Parame = g_str_Parame & "       RPT_VALNUM04 SALDOFINAL, NVL(CODSEC,'XXXXXXXXXXXX') CODSEC, OPCION, ORDEN, RPT_DESCRI "
'   g_str_Parame = g_str_Parame & "  FROM ("
'   g_str_Parame = g_str_Parame & "             SELECT RPT_CODIGO, RPT_VALNUM01, RPT_VALNUM04, RPT_MONEDA, TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION,"
'   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC AS CODSEC,NVL(C.BSISECT_ORDEN,'A') AS ORDEN, RPT_DESCRI"
'   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
'   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
'   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
'   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(B.BSICNTA_CUENTA) AND C.BSISECT_SITUAC = 1"
'   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
'   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
'   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
'   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
'   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
'   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' "
'   g_str_Parame = g_str_Parame & "        UNION "
'   g_str_Parame = g_str_Parame & "             SELECT CASE TRIM(B.BSICNTA_CUENTA) WHEN '11M70900000000' THEN SUBSTR(RPT_CODIGO,1,6) ELSE RPT_CODIGO END AS RPT_CODIGO,"
'   g_str_Parame = g_str_Parame & "                    RPT_VALNUM01, RPT_VALNUM04, RPT_MONEDA, TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION,"
'   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC CODSEC, NVL(C.BSISECT_ORDEN,'A') ORDEN, RPT_DESCRI"
'   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
'   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,3),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
'   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
'   g_str_Parame = g_str_Parame & "              INNER JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(SUBSTR(RPT_CODIGO,1,2)||'M'||SUBSTR(RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1))"
'   g_str_Parame = g_str_Parame & "                    AND C.BSISECT_SITUAC = 2"
'   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
'   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
'   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
'   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
'   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
'   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' "
'   g_str_Parame = g_str_Parame & "        UNION "
'   g_str_Parame = g_str_Parame & "             SELECT RPT_CODIGO, DECODE(BSISECT_SALDO,0,RPT_VALNUM01,BSISECT_SALDO) RPT_VALNUM01,"
'   g_str_Parame = g_str_Parame & "                    DECODE(BSISECT_SALDO,0,RPT_VALNUM04,BSISECT_SALDO) RPT_VALNUM04, RPT_MONEDA,"
'   g_str_Parame = g_str_Parame & "                    TRIM(B.BSICNTA_CUENTA) BCR_CUENTA, B.BSICNTA_OPCION OPCION,"
'   g_str_Parame = g_str_Parame & "                    C.BSISECT_CODSEC CODSEC, NVL(C.BSISECT_ORDEN,'A') ORDEN, RPT_DESCRI"
'   g_str_Parame = g_str_Parame & "               FROM RPT_TABLA_TEMP A "
'   g_str_Parame = g_str_Parame & "               LEFT JOIN CNTBL_BSICNTA B ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,3),14,'0')) = TRIM(B.BSICNTA_CUENTA)"
'   g_str_Parame = g_str_Parame & "                    AND TRIM(B.BSICNTA_OPCION) = '1'"
'   g_str_Parame = g_str_Parame & "              INNER JOIN CNTBL_BSISECT C ON TRIM(C.BSISECT_CODCTA) = TRIM(SUBSTR(RPT_CODIGO,1,2)||'M'||SUBSTR(RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1))"
'   g_str_Parame = g_str_Parame & "                    AND C.BSISECT_SITUAC = 3"
'   g_str_Parame = g_str_Parame & "              WHERE RPT_PERMES = " & CInt(l_str_PerMes)
'   g_str_Parame = g_str_Parame & "                AND RPT_PERANO = " & CInt(l_str_PerAno)
'   g_str_Parame = g_str_Parame & "                AND RPT_MONEDA = 1"
'   g_str_Parame = g_str_Parame & "                AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
'   g_str_Parame = g_str_Parame & "                AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
'   g_str_Parame = g_str_Parame & "                AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0)"
'   g_str_Parame = g_str_Parame & "                AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "')"
'   g_str_Parame = g_str_Parame & " WHERE SUBSTR(RPT_CODIGO,1,1) NOT IN (7,8) "
'
''   '*******
''   g_str_Parame = g_str_Parame & "      WHERE "
''   If l_str_PerMes = 12 Then
''      g_str_Parame = g_str_Parame & " TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' AND TRIM(RPT_CODIGO) <> '301' AND " 'AND TRIM(RPT_CODIGO) NOT LIKE '6%'
''      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '401' AND TRIM(RPT_CODIGO) <> '402' AND TRIM(RPT_CODIGO) <> '501' AND TRIM(RPT_CODIGO) <> '502' AND "
''      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '701' AND TRIM(RPT_CODIGO) <> '801' AND TRIM(RPT_CODIGO) <> '802' "
''   Else
''      g_str_Parame = g_str_Parame & " TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' AND TRIM(RPT_CODIGO) <> '201' AND TRIM(RPT_CODIGO) <> '202' AND TRIM(RPT_CODIGO) <> '301' AND "
''      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '401' AND TRIM(RPT_CODIGO) <> '402' AND TRIM(RPT_CODIGO) <> '501' AND TRIM(RPT_CODIGO) <> '502' AND "
''      g_str_Parame = g_str_Parame & " TRIM(RPT_CODIGO) <> '701' AND TRIM(RPT_CODIGO) <> '801' AND TRIM(RPT_CODIGO) <> '802' "
''   End If
''   '*******
'   g_str_Parame = g_str_Parame & "     ORDER BY RPT_MONEDA ASC, RPT_CODIGO ASC, ORDEN asc, CODSEC "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      MsgBox "No se encontraron movimientos para el BSI.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_obj_Excel.Sheets(2).Name = "REPORTE BSI PARA BCR"
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(1, 1), .Cells(1, 7)).Merge
      .Range(.Cells(1, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "REPORTE BSI PARA BCR"
      
      .Cells(2, 1) = "Año: " & l_str_PerAno & "      " & "Mes: " & l_str_PerMes
      .Range(.Cells(2, 1), .Cells(2, 7)).Merge
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 11
      
      .Columns("A").ColumnWidth = 7
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 21
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("B").NumberFormat = "@"
      .Columns("C").ColumnWidth = 21
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("C").NumberFormat = "@"
      .Columns("D").ColumnWidth = 19
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 19
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 19
      .Columns("F").NumberFormat = "@"
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 12
      .Columns("G").NumberFormat = "@"
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 1) = "ITEM"
      .Cells(4, 2) = "CUENTA BCR"
      .Cells(4, 3) = "CUENTA MICASITA"
      .Cells(4, 4) = "SALDO INICIAL"
      .Cells(4, 5) = "SALDO FINAL"
      .Cells(4, 6) = "CODIGO SECTOR"
      .Cells(4, 7) = "OPCION"
   
      .Range(.Cells(4, 1), .Cells(4, 7)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 7)).HorizontalAlignment = xlHAlignCenter
      .Cells(4, 1).HorizontalAlignment = xlHAlignCenter
        
      r_int_ConVer = 5
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 4
         .Cells(r_int_ConVer, 2) = Trim(g_rst_Genera!Cuenta)
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Genera!RPT_CODIGO)
         .Cells(r_int_ConVer, 4) = CDbl(Format(IIf(IsNull(g_rst_Genera!SALDOINICIAL), 0, g_rst_Genera!SALDOINICIAL), "###,###,##0.00"))
         .Cells(r_int_ConVer, 5) = CDbl(Format(IIf(IsNull(g_rst_Genera!SALDOFINAL), 0, g_rst_Genera!SALDOFINAL), "###,###,##0.00"))
         .Cells(r_int_ConVer, 6) = Trim(g_rst_Genera!CODSEC)
         .Cells(r_int_ConVer, 7) = Trim(g_rst_Genera!Opcion)
           
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Genera.MoveNext
         DoEvents
      Loop
   End With
      
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      ReDim Preserve r_arr_MtzBsi(UBound(r_arr_MtzBsi) + 1)
      r_str_CadIni = "": r_str_CadFin = ""
      r_str_CadIni = IIf(g_rst_Genera!SALDOINICIAL < 0, "-" & Format(-(g_rst_Genera!SALDOINICIAL), "000000000000000.00"), _
                         Format(g_rst_Genera!SALDOINICIAL, "0000000000000000.00"))
      r_str_CadFin = IIf(g_rst_Genera!SALDOFINAL < 0, "-" & Format(-(g_rst_Genera!SALDOFINAL), "000000000000000.00"), _
                         Format(g_rst_Genera!SALDOFINAL, "0000000000000000.00"))
                      
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_Cuenta = Trim(g_rst_Genera!RPT_CODIGO & "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_Codigo = Trim(g_rst_Genera!Cuenta & "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_CodSec = Trim(g_rst_Genera!CODSEC & "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_SldIni = Replace$(r_str_CadIni, ".", "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_SldFin = Replace$(r_str_CadFin, ".", "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_Opcion = Trim(g_rst_Genera!Opcion & "")
      r_arr_MtzBsi(UBound(r_arr_MtzBsi)).BCR_SldMon = CStr(g_rst_Genera!CODMON & "")
      g_rst_Genera.MoveNext
      DoEvents
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   Call fs_GeneraArchivo_BCR_2(r_arr_MtzBsi, r_arr_MtzCta)
   
   'VALIDACION DE CUENTAS QUE FALTAN FACTORIZAR
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT RPT_CODIGO, RPT_MONEDA, B.BSICNTA_CUENTA, B.BSICNTA_OPCION, C.BSISECT_CODSEC "
   g_str_Parame = g_str_Parame & "  FROM RPT_TABLA_TEMP A LEFT JOIN CNTBL_BSICNTA B "
   g_str_Parame = g_str_Parame & "    ON TRIM(RPAD(SUBSTR(A.RPT_CODIGO,1,2)||'M'||SUBSTR(A.RPT_CODIGO,4,LENGTH(TRIM(RPT_CODIGO))-1),14,'0')) = TRIM(B.BSICNTA_CUENTA) "
   g_str_Parame = g_str_Parame & "   AND TRIM(B.BSICNTA_OPCION) = '1' "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CNTBL_BSISECT C "
   g_str_Parame = g_str_Parame & "    ON TRIM(C.BSISECT_CODCTA) = TRIM(B.BSICNTA_CUENTA) AND C.BSISECT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " Where RPT_PERMES = " & CInt(l_str_PerMes)
   g_str_Parame = g_str_Parame & "   AND RPT_PERANO = " & CInt(l_str_PerAno)
   g_str_Parame = g_str_Parame & "   AND RPT_MONEDA = 1 "
   g_str_Parame = g_str_Parame & "   AND TRIM(B.BSICNTA_OPCION) = '1' "
   g_str_Parame = g_str_Parame & "   AND TRIM(C.BSISECT_CODSEC) IS NULL "
   g_str_Parame = g_str_Parame & "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "   AND RPT_NOMBRE = '" & "REPORTE " & UCase(Me.cmb_TipRep.Text) & "' "
   g_str_Parame = g_str_Parame & "   AND (RPT_VALNUM01 <> 0 OR RPT_VALNUM04 <> 0) "
   g_str_Parame = g_str_Parame & "   AND TRIM(RPT_DESCRI) = '" & Trim(l_str_CodSbs) & "' "
   
   l_str_CodAux = ""
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         l_str_CodAux = l_str_CodAux + Chr(13) + Trim(g_rst_Genera!BSICNTA_CUENTA) + " - " + Trim(g_rst_Genera!RPT_CODIGO)
           
         g_rst_Genera.MoveNext
      Loop
   End If
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   '----------------------------------------------------------FIN BSI

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_BCR()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConAux     As Integer
Dim r_arr_Matriz()   As g_tpo_BCR
Dim r_dbl_TotAct     As Double
Dim r_db_TotPas      As Double
   
   ReDim r_arr_Matriz(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RG.ORDEN, RS.DESCRIPCION, RS.ORDEN, "
   g_str_Parame = g_str_Parame & "        RS.VISIBLE, RS.NIVEL, RS.CNTA_CTBL_BCR, "
   g_str_Parame = g_str_Parame & "        RS.DESC_BCR, RT.RPT_VALNUM01 SDO_AJUSTE, "
   g_str_Parame = g_str_Parame & "        RT.RPT_VALNUM02 SDO_MN , RT.RPT_VALNUM03 SDO_EQUIME "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP RT, "
   g_str_Parame = g_str_Parame & "        RPT_GRUPO RG, "
   g_str_Parame = g_str_Parame & "        RPT_SUBGRUPO RS "
   g_str_Parame = g_str_Parame & "  WHERE RS.REPORTE = RG.REPORTE "
   g_str_Parame = g_str_Parame & "    AND RS.GRUPO = RG.GRUPO "
   g_str_Parame = g_str_Parame & "    AND RS.SUBGRUPO = RT.RPT_CODIGO "
   g_str_Parame = g_str_Parame & "    AND RG.REPORTE = 'BalSec' "
   g_str_Parame = g_str_Parame & "    AND RS.GRUPO = 'unico' "
   g_str_Parame = g_str_Parame & "    AND RPT_PERMES = " & CInt(l_str_PerMes) & " "
   g_str_Parame = g_str_Parame & "    AND RPT_PERANO = " & CInt(l_str_PerAno) & " "
   g_str_Parame = g_str_Parame & "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'REPORTE BALANCE SECTORIAL BCR' "
   g_str_Parame = g_str_Parame & "  ORDER BY RG.ORDEN ASC, RS.ORDEN ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Range(.Cells(1, 1), .Cells(1, 6)).Merge
      .Range(.Cells(1, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "BALANCE SECTORIAL POR AGENTES ECONÓMICOS"
      
      .Cells(2, 1) = "Banco Central de Reserva del Perú"
      .Range(.Cells(2, 1), .Cells(2, 6)).Merge
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 11
      
      .Cells(3, 1) = "(Saldos expresados en miles de soles)"
      .Range(.Cells(3, 1), .Cells(3, 6)).Merge
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Size = 10

      .Cells(4, 1) = "Al " & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno
      .Range(.Cells(4, 1), .Cells(4, 6)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Size = 10
                  
      .Columns("A").ColumnWidth = 23
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("A").NumberFormat = "@"
      .Columns("B").ColumnWidth = 58
      .Columns("C").ColumnWidth = 46
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("C").NumberFormat = "@"
      .Columns("D").ColumnWidth = 12
      .Columns("D").NumberFormat = "#,##0"
      .Columns("E").ColumnWidth = 12
      .Columns("E").NumberFormat = "#,##0"
      .Columns("F").ColumnWidth = 12
      .Columns("F").NumberFormat = "#,##0"
      
      .Cells(6, 1) = "CÓDIGO CUENTAS BCR"
      .Range(.Cells(6, 1), .Cells(7, 1)).Merge
      .Range(.Cells(6, 1), .Cells(7, 1)).HorizontalAlignment = xlHAlignCenter
      .Cells(6, 2) = "DESCRIPCIÓN"
      .Range(.Cells(6, 2), .Cells(7, 2)).Merge
      .Range(.Cells(6, 2), .Cells(7, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(6, 3) = "EQUIVALENCIA SBS"
      .Range(.Cells(6, 3), .Cells(7, 3)).Merge
      .Range(.Cells(6, 3), .Cells(7, 3)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(6, 4) = "MONEDA NACIONAL"
      .Range(.Cells(6, 4), .Cells(6, 5)).Merge
      .Range(.Cells(6, 4), .Cells(6, 5)).HorizontalAlignment = xlHAlignCenter
      
      With .Range(.Cells(6, 1), .Cells(7, 4))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
      End With
      
      .Cells(7, 4) = "AJUSTADO"
      .Cells(7, 5) = "HISTÓRICO"
      .Cells(6, 6) = "MONEDA EXTRANJERA"
      
      With .Range(.Cells(6, 6), .Cells(7, 6))
         .Merge
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
      End With
      
      .Range(.Cells(6, 1), .Cells(7, 6)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(7, 6)).HorizontalAlignment = xlHAlignCenter
      .Cells(3, 1).HorizontalAlignment = xlHAlignCenter
      
      With .Range(.Cells(6, 1), .Cells(7, 6))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         With .Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeTop)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlInsideVertical)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlInsideHorizontal)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
      End With
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 8
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!CNTA_CTBL_BCR)
      For r_int_ConAux = 2 To CInt(Trim(g_rst_Princi!NIVEL))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) & Chr(32) & Chr(32)
      Next r_int_ConAux
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) & Trim(g_rst_Princi!DESCRIPCION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DESC_BCR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SDO_AJUSTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!SDO_MN)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!SDO_EQUIME)
      
      ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
      r_arr_Matriz(UBound(r_arr_Matriz)).BCR_Codigo = Trim(g_rst_Princi!CNTA_CTBL_BCR)
      r_arr_Matriz(UBound(r_arr_Matriz)).BCR_SldAju = Trim(g_rst_Princi!SDO_AJUSTE)
      r_arr_Matriz(UBound(r_arr_Matriz)).BCR_SldMon = Trim(g_rst_Princi!SDO_MN)
      r_arr_Matriz(UBound(r_arr_Matriz)).BCR_SldEqu = Trim(g_rst_Princi!SDO_EQUIME)
      
      If Trim(g_rst_Princi!CNTA_CTBL_BCR) = "1000000000000000000" Then
         r_dbl_TotAct = CDbl(Trim(g_rst_Princi!SDO_MN)) + CDbl(Trim(g_rst_Princi!SDO_EQUIME))
      ElseIf Trim(g_rst_Princi!CNTA_CTBL_BCR) = "2000000000000000000" Then
         r_db_TotPas = CDbl(Trim(g_rst_Princi!SDO_MN)) + CDbl(Trim(g_rst_Princi!SDO_EQUIME))
      End If
      
      If r_dbl_TotAct <> 0 And r_db_TotPas <> 0 Then
         r_arr_Matriz(UBound(r_arr_Matriz)).BCR_AjuDif = r_dbl_TotAct - r_db_TotPas
         r_dbl_TotAct = 0: r_db_TotPas = 0
      End If
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   With r_obj_Excel.ActiveSheet
      With .Range(.Cells(r_int_ConVer - 1, 1), .Cells(r_int_ConVer - 1, 6))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         .Borders(xlEdgeLeft).LineStyle = xlNone
         .Borders(xlEdgeTop).LineStyle = xlNone
         With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
         End With
         .Borders(xlEdgeRight).LineStyle = xlNone
         .Borders(xlInsideVertical).LineStyle = xlNone
         .Borders(xlInsideHorizontal).LineStyle = xlNone
      End With
      
      With .Range(.Cells(8, 6), .Cells(r_int_ConVer - 1, 6))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         .Borders(xlEdgeLeft).LineStyle = xlNone

         With .Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         .Borders(xlInsideVertical).LineStyle = xlNone
         .Borders(xlInsideHorizontal).LineStyle = xlNone
      End With
      
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 6)).Font.Size = 10
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_GeneraArchivo_BCR_1(r_arr_Matriz)

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_SitFin()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_VarAux     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RPT_CODIGO CODIGO, RPT_DESCRI DESCRIPCION, RPT_VALNUM01 MN, RPT_VALNUM02 EQ, RPT_VALNUM03 TOT , RS.FLAG_NEGRITA NEGRITA"
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP RT "
   g_str_Parame = g_str_Parame & "  INNER JOIN RPT_SUBGRUPO RS ON TRIM(RS.SUBGRUPO) = TRIM(RT.RPT_CODIGO) AND RS.REPORTE ='EstBal1 ' AND RS.GRUPO = 'Unico'"
   g_str_Parame = g_str_Parame & "  WHERE RT.RPT_PERMES = " & CInt(l_str_PerMes) & " AND RT.RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "        RT.RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RT.RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "        RT.RPT_NOMBRE = 'ESTADO DE SITUACION FINANCIERA' "
   g_str_Parame = g_str_Parame & " ORDER BY TO_NUMBER (TRIM(RT.RPT_CODIGO), 99) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "1. Situacion Financiera"
   
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 1), .Cells(1, 4)).Merge
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "Institución: EDPYME MICASITA S A"
      
      .Range(.Cells(2, 1), .Cells(2, 4)).Merge
      .Cells(2, 1) = "ESTADO DE SITUACIÓN FINANCIERA"
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 14
      
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Cells(3, 1) = "Al " & ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno)) & " de " & UCase(Left(Me.cmb_Period.Text, 1)) & LCase(Right(cmb_Period.Text, Len(cmb_Period.Text) - 1)) & " de " & ipp_PerAno.Text
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 11
      
      .Range(.Cells(4, 1), .Cells(4, 4)).Merge
      .Cells(4, 1) = "( En soles )"
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Size = 10
                   
      .Range(.Cells(1, 1), .Cells(5, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(4, 4)).Font.Bold = True
                   
      .Columns("A").ColumnWidth = 68
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").NumberFormat = "###,###,##0.00"
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("A").NumberFormat = "@"
      
      .Cells(5, 1) = "ACTIVO"
      .Cells(5, 2) = "Moneda      Nacional"
      .Cells(5, 3) = "Equivalente    en M.E."
      .Cells(5, 4) = "TOTAL"
           
      .Range(.Cells(5, 1), .Cells(6, 6)).Font.Bold = True
      
      With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
      End With
      .Range(.Cells(5, 1), .Cells(6, 1)).Merge
      With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = True
      End With
      With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlLeft
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = True
      End With
      With .Range(.Cells(5, 2), .Cells(6, 4))
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .WrapText = True
           .Orientation = 0
           .AddIndent = False
           .IndentLevel = 0
           .ShrinkToFit = False
           .ReadingOrder = xlContext
           .MergeCells = False
      End With
      With .Range(.Cells(5, 2), .Cells(6, 2))
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .WrapText = True
           .Orientation = 0
           .AddIndent = False
           .IndentLevel = 0
           .ShrinkToFit = False
           .ReadingOrder = xlContext
           .MergeCells = False
      End With
      .Range(.Cells(5, 2), .Cells(6, 2)).Merge
      With .Range(.Cells(5, 3), .Cells(6, 3))
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .WrapText = True
           .Orientation = 0
           .AddIndent = False
           .IndentLevel = 0
           .ShrinkToFit = False
           .ReadingOrder = xlContext
           .MergeCells = False
      End With
      .Range(.Cells(5, 3), .Cells(6, 3)).Merge
      With .Range(.Cells(5, 4), .Cells(6, 4))
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .WrapText = True
           .Orientation = 0
           .AddIndent = False
           .IndentLevel = 0
           .ShrinkToFit = False
           .ReadingOrder = xlContext
           .MergeCells = False
      End With
      .Range(.Cells(5, 4), .Cells(6, 4)).Merge
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 6
   
   Do While Not g_rst_Princi.EOF
      If CInt(Trim(g_rst_Princi!CODIGO)) = 46 Then r_int_ConVer = r_int_ConVer + 2

      If g_rst_Princi!NEGRITA = 1 Then
          r_obj_Excel.Sheets(1).Range(r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 1), r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 4)).Font.Bold = True
      End If
      
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 1) = g_rst_Princi!DESCRIPCION
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 2) = IIf(g_rst_Princi!MN = 0, "", CDbl(Format(g_rst_Princi!MN, "###,###,##0.00")))
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 3) = IIf(g_rst_Princi!Eq = 0, "", CDbl(Format(g_rst_Princi!Eq, "###,###,##0.00")))
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 4) = IIf(g_rst_Princi!TOT = 0, "", CDbl(Format(g_rst_Princi!TOT, "###,###,##0.00")))
      
      For r_int_VarAux = 2 To 4
         Call fs_Formato_NumNeg(r_obj_Excel.Sheets(1).Cells(r_int_ConVer, r_int_VarAux))
      Next r_int_VarAux
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop

   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 4)).Font.Name = "Arial Narrow"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 4)).Font.Size = 8 '9
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 11
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 15
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Size = 11
      .Range(.Cells(5, 1), .Cells(6, 1)).Font.Size = 14
      .Range(.Cells(53, 1), .Cells(53, 1)).Font.Size = 14
      
      'BORDES
      Call fs_Bordes_SitFin(.Range(.Cells(5, 1), .Cells(6, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(7, 1), .Cells(7, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(14, 1), .Cells(15, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(19, 1), .Cells(19, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(22, 1), .Cells(23, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(30, 1), .Cells(32, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(35, 1), .Cells(35, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(38, 1), .Cells(38, 4)))
      Call fs_Bordes_SitFin1(.Range(.Cells(46, 1), .Cells(50, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(53, 1), .Cells(53, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(54, 1), .Cells(54, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(59, 1), .Cells(60, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(64, 1), .Cells(64, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(70, 1), .Cells(73, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(77, 1), .Cells(80, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(81, 1), .Cells(81, 4)))
      
      .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
      .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlDiagonalUp).LineStyle = xlNone
      With .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      .Range(.Cells(5, 1), .Cells(89, 4)).Borders(xlInsideVertical).LineStyle = xlNone
      
      With .Range(.Cells(51, 1), .Cells(52, 4))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         .Borders(xlEdgeLeft).LineStyle = xlNone
         With .Borders(xlEdgeTop)
             .LineStyle = xlDouble
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThick
         End With
         With .Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         .Borders(xlEdgeRight).LineStyle = xlNone
         .Borders(xlInsideVertical).LineStyle = xlNone
         .Borders(xlInsideHorizontal).LineStyle = xlNone
      End With
      Call fs_Bordes_SitFin1(.Range(.Cells(88, 1), .Cells(89, 4)))

      .Columns("B:B").EntireColumn.AutoFit
      .Columns("C:C").EntireColumn.AutoFit
      .Columns("D:D").EntireColumn.AutoFit
   End With
      
   'MARGENES DE IMPRESIÓN
   l_Mar_Izq = 0.3
   l_Mar_Der = 0.3
   l_Mar_Sup = 0.3
   l_Mar_Inf = 0.3
   
   With r_obj_Excel.Sheets(1).PageSetup
      If .Orientation = xlPortrait Then
         .Orientation = xlPortrait
      Else
         .Orientation = xlLandscape
      End If
      
      'Configuración de márgenes:
      .LeftMargin = Application.CentimetersToPoints(l_Mar_Izq)
      .RightMargin = Application.CentimetersToPoints(l_Mar_Der)
      .TopMargin = Application.CentimetersToPoints(l_Mar_Sup)
      .BottomMargin = Application.CentimetersToPoints(l_Mar_Inf)
      .CenterHorizontally = True
      .CenterVertically = True
   End With
   
   'CABECERA
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(52, 1), .Cells(57, 1)).EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
      'PASIVO
      .Range(.Cells(1, 1), .Cells(6, 4)).Copy
      .Range("A58").Insert Shift:=xlDown
      .Cells(62, 1) = "PASIVO Y PATRIMONIO"
      .Rows("64:65").Delete Shift:=xlUp
      .Rows("5:5").EntireRow.AutoFit
      .Range("A1").Select
   End With
   
   'PIE DE PÁGINA
    With r_obj_Excel.Sheets(1)
      .Range(.Cells(54, 1), .Cells(54, 4)).Merge
      .Cells(54, 1) = "________________________      ________________________     ________________________     ________________________"
      .Cells(55, 1) = "                 Director                                                 Director                                 Gerente General                               Contador"
      .Range(.Cells(54, 1), .Cells(55, 4)).Font.Name = "Calibri"
      .Range(.Cells(54, 1), .Cells(55, 4)).Font.Size = 11

      .Range(.Cells(109, 1), .Cells(109, 4)).Merge
      .Cells(109, 1) = "________________________      ________________________     ________________________     ________________________"
      .Cells(110, 1) = "                 Director                                                 Director                                 Gerente General                               Contador"
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   Call fs_GenExc_EstRes(r_obj_Excel)
End Sub

Private Sub fs_GenExc_ConCtaCtb()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConAux     As Integer
Dim r_bol_FlgGrp     As Boolean
   
   g_str_Parame = ""
   g_str_Parame = " SELECT TRIM(A.RPT_CODIGO) CUENTA, A.RPT_DESCRI DESCRIPCION, A.RPT_VALNUM01 SALDO_ACTUAL, A.RPT_VALNUM02 PADRON"
   g_str_Parame = g_str_Parame & "  FROM RPT_TABLA_TEMP A   "
   g_str_Parame = g_str_Parame & " WHERE A.RPT_PERMES = " & CInt(l_str_PerMes) & " AND A.RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "       A.RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND A.RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "       A.RPT_NOMBRE = '" & UCase(Me.cmb_TipRep.Text) & 2 & "' "
   g_str_Parame = g_str_Parame & " ORDER BY A.RPT_MONEDA ASC, A.RPT_CODIGO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 1), .Cells(1, 5)).Merge
      .Range(.Cells(1, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "CONCILIACION CON CUENTAS CONTABLES"
      
      .Range(.Cells(2, 1), .Cells(2, 5)).Merge
      .Cells(2, 1) = "Del " & "01/" & l_str_PerMes & "/" & l_str_PerAno & " Al " & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Bold = False
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 10
       
      .Cells(4, 1) = "CUENTA"
      .Cells(4, 2) = "DESCRIPCIÓN"
      .Cells(4, 3) = "SAL. ACT."
      .Cells(4, 4) = "PADRON"
      .Cells(4, 5) = "DIFERENCIAS"
        
      .Range(.Cells(4, 1), .Cells(4, 5)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
                  
      .Columns("A").ColumnWidth = 14.15
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("A").NumberFormat = "@"
      .Columns("B").ColumnWidth = 60
      .Columns("C").ColumnWidth = 15
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 15
      .Columns("E").NumberFormat = "###,###,##0.00"
      
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 5
   Do While Not g_rst_Princi.EOF
   
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!Cuenta)
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!DESCRIPCION)
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 3) = CDbl(Format(IIf(IsNull(g_rst_Princi!SALDO_ACTUAL), 0, g_rst_Princi!SALDO_ACTUAL), "###,###,##0.00"))
      If Not IsNull(g_rst_Princi!PADRON) Then
            r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 4) = CDbl(Format(IIf(IsNull(g_rst_Princi!PADRON), 0, g_rst_Princi!PADRON), "###,###,##0.00"))
      End If
      r_obj_Excel.Sheets(1).Cells(r_int_ConVer, 5) = CDbl(Format((IIf(IsNull(g_rst_Princi!PADRON), 0, g_rst_Princi!PADRON) - IIf(IsNull(g_rst_Princi!SALDO_ACTUAL), 0, g_rst_Princi!SALDO_ACTUAL)), "###,###,##0.00"))
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(5, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5)).Font.Size = 10
   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 5 & "]C:R[-1]C)"
   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5).Font.Bold = True
   r_obj_Excel.ActiveSheet.Cells(5, 1).Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   
   'BORDES
   With r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5))
        .Select
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5)
        .Select
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
   End With
   With r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(4, 1), r_obj_Excel.ActiveSheet.Cells(4, 5))
        .Select
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With r_obj_Excel.ActiveSheet.Cells(4, 5)
        .Select
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
   r_obj_Excel.ActiveWindow.ScrollRow = 5
   r_obj_Excel.ActiveSheet.Cells(4, 1).Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_NotasEEFF(ByVal g_rst_Princi As ADODB.Recordset)
Dim r_obj_Excel         As Excel.Application
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_dbl_CosAct        As Double
Dim r_dbl_CosAnt        As Double
Dim r_dbl_TotAdi        As Double
Dim r_dbl_TotAju        As Double
Dim r_dbl_TotDep        As Double
Dim r_dbl_DepAcu        As Double
Dim r_dbl_SalAcu        As Double
Dim r_dbl_CreFis        As Double
Dim r_dbl_MtoIgv        As Double
Dim r_int_ConAux        As Integer
Dim r_int_ConDif        As Integer
Dim r_int_ConIgv        As Integer

Dim r_int_ConIn1        As Integer
Dim r_int_ConIn2        As Integer
Dim r_int_ConIn3        As Integer
Dim r_int_ConIn4        As Integer

   r_int_Contad = 6
   r_int_PerMes = CInt(cmb_Period.ItemData(cmb_Period.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RT.RPT_CODIGO CODIGO, RT.RPT_DESCRI DESCRIPCION, RT.RPT_VALNUM01 SALDO_INI_DEBE, RT.RPT_VALNUM02 SALDO_INI_HABER, "
   g_str_Parame = g_str_Parame & "        RT.RPT_VALNUM03 MOV_MES_DEBE, RT.RPT_VALNUM04 MOV_MES_HABER, RT.RPT_VALNUM05 SALDO_FIN_DEBE, RT.RPT_VALNUM06 SALDO_FIN_HABER "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP RT "
   g_str_Parame = g_str_Parame & "  WHERE RT.RPT_PERMES = " & CInt(l_str_PerMes) & " AND RT.RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "        RT.RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RT.RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "        TRIM(RT.RPT_NOMBRE) = '" & Trim(cmb_TipRep.Text) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "NOTA"
   'r_obj_Excel.Visible = True
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "EDPYME MICASITA S.A."
      .Range(.Cells(1, 1), .Cells(1, 4)).Merge
      .Range(.Cells(1, 1), .Cells(1, 4)).Font.Bold = True
            
      .Cells(2, 1) = "NOTAS A LOS ESTADOS FINANCIEROS"
      .Range(.Cells(2, 1), .Cells(2, 4)).Merge
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Bold = True
      
      .Cells(3, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_Period.Text, 1) & LCase(Mid(cmb_Period.Text, 2, Len(cmb_Period.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge

      .Range(.Cells(4, 1), .Cells(4, 4)).Merge
      .Cells(4, 1) = "( Expresado En Soles )"
      .Range(.Cells(5, 1), .Cells(5, 4)).Merge
     
      With .Range(.Cells(1, 1), .Cells(4, 4))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
      End With
      
      .Cells(r_int_Contad, 4) = "'" & Left(cmb_Period.Text, 3) & "-" & Right(Format(r_int_PerAno, "0000"), 2)
              
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
      .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 4)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 55
      .Columns("D").ColumnWidth = 13.5
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0.00"
      .Columns("E").NumberFormat = "###,###,###,##0.00"
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
        
      r_int_Contad = r_int_Contad + 2
      .Cells(r_int_Contad, 3) = "ACTIVO"
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
      r_int_Contad = r_int_Contad + 1
      
      Do While Not g_rst_Princi.EOF
        
            r_int_Contad = r_int_Contad + 1
            
            If g_rst_Princi!GRUPO = 4 Then  'AÑADE INFORMACIÓN DE COSTOS Y DEPRECIACIÓN DE INMUEBLES, MOBILIARIO Y EQUIPO
                g_rst_GenAux.AddNew
                g_rst_GenAux!GRUPO = g_rst_Princi!GRUPO
                g_rst_GenAux!NOMGRUPO = Trim(g_rst_Princi!NOMGRUPO)
                g_rst_GenAux!SUBGRP = g_rst_Princi!SUBGRP
                g_rst_GenAux!NOMSUBGRP = Trim(g_rst_Princi!NOMSUBGRP)
                g_rst_GenAux!CNTACTBLE = Trim(g_rst_Princi!CNTACTBLE)
                g_rst_GenAux!NOMCTA = Trim(g_rst_Princi!NOMCTA)
                g_rst_GenAux!Mes = g_rst_Princi!Mes
                g_rst_GenAux!INDTIPO = Trim(g_rst_Princi!INDTIPO)
                If Trim(g_rst_Princi!CNTACTBLE) = "181301010101" Or g_rst_Princi!CNTACTBLE = "181302010101" Or g_rst_Princi!CNTACTBLE = "181309010101" Or g_rst_Princi!CNTACTBLE = "181701010101" Or g_rst_Princi!CNTACTBLE = "181401010101" Then
                    g_rst_GenAux!TIPO = "C"
                ElseIf Trim(g_rst_Princi!CNTACTBLE) = "181903010101" Or g_rst_Princi!CNTACTBLE = "181903010102" Or g_rst_Princi!CNTACTBLE = "181903010103" Or g_rst_Princi!CNTACTBLE = "181907010101" Or g_rst_Princi!CNTACTBLE = "181904010101" Then
                    g_rst_GenAux!TIPO = "D"
                End If
                g_rst_GenAux.Update
                
            Else
                If g_rst_GenAux.RecordCount > 0 Then
                   g_rst_GenAux.MoveFirst
                   r_int_Contad = r_int_Contad - g_rst_GenAux.RecordCount - 2
                   
                   .Rows(r_int_Contad).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                   .Rows(r_int_Contad).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                   
                   .Cells(r_int_Contad + 1, 4) = "Saldo al 01/01/" & l_str_PerAno
                   .Range(.Cells(r_int_Contad + 1, 4), .Cells(r_int_Contad + 2, 4)).Merge
                   
                   .Cells(r_int_Contad + 1, 5) = "Adiciones"
                   .Range(.Cells(r_int_Contad + 1, 5), .Cells(r_int_Contad + 2, 5)).Merge
                   
                   .Cells(r_int_Contad + 1, 6) = "Ajustes/Retiros"
                   .Range(.Cells(r_int_Contad + 1, 6), .Cells(r_int_Contad + 2, 6)).Merge
                   
                   .Cells(r_int_Contad + 1, 7) = "Saldo al " & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno
                   .Range(.Cells(r_int_Contad + 1, 7), .Cells(r_int_Contad + 2, 7)).Merge
                   
                   With .Range(.Cells(r_int_Contad + 1, 4), .Cells(r_int_Contad + 2, 7))
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .Interior.Color = RGB(146, 208, 80)
                        .Font.Bold = True
                   End With
                   With .Range(.Cells(r_int_Contad + 3, 4), .Cells(r_int_Contad + 3, 7))
                        .Interior.Color = RGB(146, 208, 80)
                        .Font.Bold = True
                   End With
                   r_int_Contad = r_int_Contad + 4
                   .Cells(r_int_Contad, 3) = "Inmuebles, Mobiliario y Equipo"
                   .Cells(r_int_Contad, 3).Font.Bold = True
                   
                   .Cells(r_int_Contad - 1, 7) = .Cells(r_int_Contad - 1, 4)
                   .Cells(r_int_Contad - 1, 4) = ""
                   
                   
                   r_int_Contad = r_int_Contad + 1
                   Do Until g_rst_GenAux.EOF
                       g_rst_GenAux.Find " TIPO = 'C'"
                       If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then
                            .Cells(r_int_Contad, 2) = g_rst_GenAux!CNTACTBLE
                            .Cells(r_int_Contad, 3) = Trim(g_rst_GenAux!NOMCTA)
                            
                            If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
                                g_rst_Genera.MoveFirst
                                g_rst_Genera.Find " CODIGO = '" & Trim(g_rst_GenAux!CNTACTBLE) & "'"
                                
                                'nueva validacion 17-04-2017
                                If g_rst_Genera.BOF = False And Not g_rst_Genera.EOF Then
                                'If Not g_rst_Genera.BOF Then 'And Not g_rst_Genera.EOF
                                    .Cells(r_int_Contad, 4) = g_rst_Genera!SALDO_INI_DEBE
                                    If g_rst_Genera!MOV_MES_DEBE <> 0 Then .Cells(r_int_Contad, 5) = g_rst_Genera!MOV_MES_DEBE
                                    If g_rst_Genera!MOV_MES_HABER <> 0 Then .Cells(r_int_Contad, 6) = g_rst_Genera!MOV_MES_HABER
                                Else
                                    .Cells(r_int_Contad, 7) = 0
                                End If
                                r_dbl_CosAnt = r_dbl_CosAnt + .Cells(r_int_Contad, 4)
                                r_dbl_TotAdi = r_dbl_TotAdi + .Cells(r_int_Contad, 5)
                            End If
                            .Cells(r_int_Contad, 7) = g_rst_GenAux!Mes
                            r_dbl_CosAct = r_dbl_CosAct + g_rst_GenAux!Mes
                            r_int_ConAux = r_int_ConAux + 1
                            r_int_Contad = r_int_Contad + 1
                           g_rst_GenAux.MoveNext
                       End If
                   Loop
                   
                   .Cells(r_int_Contad - r_int_ConAux - 1, 7) = r_dbl_CosAct
                   If r_dbl_TotAdi <> 0 Then .Cells(r_int_Contad - r_int_ConAux - 1, 5) = r_dbl_TotAdi
                   .Cells(r_int_Contad - r_int_ConAux - 1, 4) = r_dbl_CosAnt
                   
                   .Range(.Cells(r_int_Contad - r_int_ConAux - 1, 4), .Cells(r_int_Contad - r_int_ConAux - 1, 7)).Font.Bold = True
                  'r_int_ConAux = 0
                   r_int_Contad = r_int_Contad + 1
                   .Cells(r_int_Contad, 4) = "Depreciación Acumulada Inicial"
                   .Range(.Cells(r_int_Contad, 4), .Cells(r_int_Contad + 1, 4)).Merge
                   .Cells(r_int_Contad, 5) = "Depreciación"
                   .Range(.Cells(r_int_Contad, 5), .Cells(r_int_Contad + 1, 5)).Merge
                   .Cells(r_int_Contad, 6) = "Ajustes/Retiros"
                   .Range(.Cells(r_int_Contad, 6), .Cells(r_int_Contad + 1, 6)).Merge
                   .Cells(r_int_Contad, 7) = "Saldo Acumulado Final"
                   .Range(.Cells(r_int_Contad, 7), .Cells(r_int_Contad + 1, 7)).Merge
                   With .Range(.Cells(r_int_Contad, 4), .Cells(r_int_Contad + 1, 7))
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .Interior.Color = RGB(146, 208, 80)
                        .Font.Bold = True
                   End With
                   With .Range(.Cells(r_int_Contad + 2, 4), .Cells(r_int_Contad + 2, 7))
                        .Font.Bold = True
                        .Interior.Color = RGB(146, 208, 80)
                   End With
                   r_int_Contad = r_int_Contad + 2
                   .Cells(r_int_Contad, 3) = "INMUEBLES, MOBILIARIO Y EQUIPO(Neto)"
                   .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
                   .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
                   r_int_Contad = r_int_Contad + 1
                   .Cells(r_int_Contad, 3) = "Inmuebles, Mobiliario y Equipo"
                   .Cells(r_int_Contad, 3).Font.Bold = True
                   r_int_Contad = r_int_Contad + 1
                   
                   g_rst_GenAux.MoveFirst
                   Do Until g_rst_GenAux.EOF
                       g_rst_GenAux.Find " TIPO = 'D'"
                       If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then
                            If Trim(g_rst_GenAux!INDTIPO) <> "S" And g_rst_GenAux!Mes <> 0 Then
                                .Cells(r_int_Contad, 2) = g_rst_GenAux!CNTACTBLE
                                .Cells(r_int_Contad, 3) = Trim(g_rst_GenAux!NOMCTA)
                                If Not g_rst_Genera.BOF Then 'And Not g_rst_Genera.EOF
                                    g_rst_Genera.MoveFirst
                                    g_rst_Genera.Find " CODIGO = '" & Trim(g_rst_GenAux!CNTACTBLE) & "'"
                                    If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
                                       .Cells(r_int_Contad, 4) = (g_rst_Genera!SALDO_INI_DEBE - g_rst_Genera!SALDO_INI_HABER)
                                        r_dbl_DepAcu = r_dbl_DepAcu + .Cells(r_int_Contad, 4)
                                        If g_rst_Genera!MOV_MES_HABER <> 0 Then
                                             .Cells(r_int_Contad, 5) = (g_rst_Genera!MOV_MES_DEBE - g_rst_Genera!MOV_MES_HABER)
                                             r_dbl_TotDep = r_dbl_TotDep + .Cells(r_int_Contad, 5)
                                        End If
                                    End If
                                End If
                                .Cells(r_int_Contad, 7) = g_rst_GenAux!Mes
                                r_dbl_SalAcu = r_dbl_SalAcu + g_rst_GenAux!Mes
                                r_int_ConAux = r_int_ConAux + 1
                                r_int_Contad = r_int_Contad + 1
                            End If
                       End If
                       If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then g_rst_GenAux.MoveNext
                   Loop
                   .Cells(r_int_Contad - r_int_ConAux + 3, 4) = r_dbl_DepAcu
                   .Cells(r_int_Contad - r_int_ConAux + 4, 4) = r_dbl_DepAcu
                   .Cells(r_int_Contad - r_int_ConAux + 3, 5) = r_dbl_TotDep
                   .Cells(r_int_Contad - r_int_ConAux + 4, 5) = r_dbl_TotDep
                   .Cells(r_int_Contad - r_int_ConAux + 3, 7) = r_dbl_SalAcu
                   .Cells(r_int_Contad - r_int_ConAux + 4, 7) = r_dbl_SalAcu
                   .Range(.Cells(r_int_Contad - r_int_ConAux + 4, 4), .Cells(r_int_Contad - r_int_ConAux + 4, 7)).Font.Bold = True   '.Range(.Cells(r_int_Contad - 6, 4), .Cells(r_int_Contad - 6, 7)).Font.Bold = True
                   .Cells(r_int_Contad - r_int_ConAux - 7, 4) = r_dbl_DepAcu + r_dbl_CosAnt
                   .Cells(r_int_Contad - r_int_ConAux - 7, 5) = r_dbl_TotAdi + r_dbl_TotDep
                   If r_dbl_TotAju <> 0 Then .Cells(r_int_Contad - r_int_ConAux - 7, 6) = r_dbl_TotAju
                   
                   r_int_ConAux = 0
                   .Cells(r_int_Contad + 1, 4) = "AL " & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno
                   With .Range(.Cells(r_int_Contad + 1, 4), .Cells(r_int_Contad + 1, 6))
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .ReadingOrder = xlContext
                        .Interior.Color = RGB(146, 208, 80)
                        .Font.Bold = True
                   End With
                   r_int_Contad = r_int_Contad + 2
                   .Cells(r_int_Contad, 3) = "RUBRO"
                   .Cells(r_int_Contad, 4) = "Costo"
                   .Cells(r_int_Contad, 5) = "Depreciación"
                   .Cells(r_int_Contad, 6) = "Valor Neto"
                    With .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 6))
                        .HorizontalAlignment = xlCenter
                        .Interior.Color = RGB(146, 208, 80)
                        .Font.Bold = True
                   End With

                   r_int_Contad = r_int_Contad + 1
                   r_int_ConAux = r_int_Contad
                   .Cells(r_int_Contad, 3) = "Mobiliario y Equipo"
                   .Cells(r_int_Contad + 1, 3) = "Equipos de Cómputo"
                   .Cells(r_int_Contad + 2, 3) = "Equipos Diversos"
                   .Cells(r_int_Contad + 3, 3) = "Instalaciones y mejoras"
                   .Cells(r_int_Contad + 4, 3) = "Unidad Transporte"

                   g_rst_GenAux.MoveFirst
                   Do Until g_rst_GenAux.EOF
                      g_rst_GenAux.Find " TIPO = 'C'"
                      If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then
                        .Cells(r_int_Contad, 4) = g_rst_GenAux!Mes
                      End If
                      r_int_Contad = r_int_Contad + 1
                      If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then g_rst_GenAux.MoveNext
                   Loop
                   
                   g_rst_GenAux.MoveFirst
                   Do Until g_rst_GenAux.EOF
                      g_rst_GenAux.Find " TIPO = 'D'"
                      If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then
                        .Cells(r_int_ConAux, 5) = g_rst_GenAux!Mes
                        .Cells(r_int_ConAux, 6).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
                      End If
                      
                      r_int_ConAux = r_int_ConAux + 1
                      If Not g_rst_GenAux.BOF And Not g_rst_GenAux.EOF Then g_rst_GenAux.MoveNext
                   Loop
                   
                   .Range(.Cells(r_int_ConAux - 1, 4), .Cells(r_int_ConAux - 1, 6)).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
                   .Range(.Cells(r_int_ConAux - 1, 4), .Cells(r_int_ConAux - 1, 6)).Interior.Color = RGB(146, 208, 80)
                   .Range(.Cells(r_int_ConAux - 1, 4), .Cells(r_int_ConAux - 1, 6)).Font.Bold = True
                   Call fs_Recorset_nc
                   r_int_Contad = r_int_Contad + 2
                   
                   GoTo Seguir
                Else
Seguir:
                   
                   If g_rst_Princi!Mes = 0 And Trim(g_rst_Princi!INDTIPO) <> "L" Then
                       r_int_Contad = r_int_Contad - 1
                       GoTo Saltar
                   End If
                
                   If Trim(g_rst_Princi!INDTIPO) = "L" Then
                      If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
                          r_int_Contad = r_int_Contad + 1
                          .Cells(r_int_Contad, 3) = "PASIVO"
                          .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
                          .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
                          r_int_Contad = r_int_Contad + 1
                      End If
                      g_rst_Princi.MoveNext
                      r_int_Contad = r_int_Contad + 1
                   End If
                   
                   'nueva validacion 17-04-2017
                   If g_rst_Princi.EOF = True Then
                      Exit Do
                   End If
                   
                   If Trim(g_rst_Princi!NOMGRUPO) = "OTROS PASIVOS" And Trim(g_rst_Princi!INDTIPO) = "G" Then
                        r_int_ConAux = r_int_Contad
                   End If
                   If g_rst_Princi!Mes = 0 Then
                        r_int_Contad = r_int_Contad - 2
                        GoTo Saltar
                   End If
                   If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Or Trim(g_rst_Princi!INDTIPO) = "A" _
                          Or Trim(g_rst_Princi!INDTIPO) = "T" Or Trim(g_rst_Princi!INDTIPO) = "X" Then
                      If g_rst_Princi!GRUPO = 5 Then
                        r_int_Contad = r_int_Contad + 1
                        .Cells(r_int_Contad - 1, 5) = "COSTO"
                        .Cells(r_int_Contad - 1, 6) = "AMORTIZACIÓN"
                        .Cells(r_int_Contad - 1, 7) = "VALOR NETO"
                        .Range(.Cells(r_int_Contad - 1, 5), .Cells(r_int_Contad, 7)).Interior.Color = RGB(146, 208, 80)
                        .Range(.Cells(r_int_Contad - 1, 5), .Cells(r_int_Contad, 7)).Font.Bold = True
                        .Range(.Cells(r_int_Contad - 1, 5), .Cells(r_int_Contad, 7)).ColumnWidth = 15
                        .Range(.Cells(r_int_Contad - 1, 5), .Cells(r_int_Contad - 1, 7)).HorizontalAlignment = xlCenter
                      
                      ElseIf g_rst_Princi!GRUPO = 6 Then
                        .Range(.Cells(r_int_Contad - 6, 5), .Cells(r_int_Contad - 6, 7)).FormulaR1C1 = "=SUM(R[1]C:R[4]C)"
                        .Range(.Cells(r_int_Contad - 6, 5), .Cells(r_int_Contad - 6, 7)).Font.Bold = True
                        .Cells(r_int_Contad - 7, 5) = .Cells(r_int_Contad - 6, 5)
                        .Cells(r_int_Contad - 7, 6) = .Cells(r_int_Contad - 6, 6)
                        .Cells(r_int_Contad - 7, 7) = .Cells(r_int_Contad - 6, 7)
                        r_int_Contad = r_int_Contad + 1
                        .Cells(r_int_Contad, 3) = "IMPUESTOS DIFERIDOS"
                        r_int_ConDif = r_int_Contad
                        .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
                        .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
                        r_int_ConAux = r_int_Contad + 1
                        r_int_Contad = r_int_Contad + 3
                      End If
                      If g_rst_Princi!GRUPO = 14 And Trim(g_rst_Princi!INDTIPO) = "G" Then
                        r_int_Contad = r_int_Contad + 1
                      End If
                        .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMGRUPO)
                        .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
                        .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
                   End If
                   If Trim(g_rst_Princi!INDTIPO) = "S" Or Trim(g_rst_Princi!INDTIPO) = "N" Or Trim(g_rst_Princi!INDTIPO) = "R" Then
                      .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
                      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
                      If Trim(g_rst_Princi!NOMSUBGRP) = "Intangibles" Then
                        .Cells(r_int_Contad, 3) = "Otros Gastos Amortizados -Operativos"
                     End If
                   End If
                   If Trim(g_rst_Princi!INDTIPO) = "D" Then
                      If Trim(g_rst_Princi!CNTACTBLE) = "191302010102" Then                             'IMPUESTO DIFERIDO
                            .Cells(r_int_ConAux, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
                            .Cells(r_int_ConAux, 3) = Trim(g_rst_Princi!NOMCTA & "")
                      ElseIf Trim(g_rst_Princi!CNTACTBLE) = "262602010101" Then
                            .Cells(r_int_Contad - 9, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
                            .Cells(r_int_Contad - 9, 3) = Trim(g_rst_Princi!NOMCTA & "")
                      Else
                            .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
                            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
                      End If
                   End If
                   If g_rst_Princi!GRUPO = 5 And (Trim(g_rst_Princi!CNTACTBLE) = "191404010101" Or Trim(g_rst_Princi!CNTACTBLE) = "191403010101" Or Trim(g_rst_Princi!CNTACTBLE) = "191408010101") Then
                        If Trim(g_rst_Princi!CNTACTBLE) = "191401010101" Then
                            r_int_ConIn1 = r_int_Contad
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191403010101" Then
                            r_int_ConIn2 = r_int_Contad
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191404010101" Then
                            r_int_ConIn3 = r_int_Contad
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191408010101" Then
                            r_int_ConIn4 = r_int_Contad
                        End If
                        .Cells(r_int_Contad, 5) = g_rst_Princi!Mes                                          'INTANGIBLES
                   ElseIf g_rst_Princi!GRUPO = 5 And (Trim(g_rst_Princi!CNTACTBLE) = "191409040101" Or Trim(g_rst_Princi!CNTACTBLE) = "191409030101" Or Trim(g_rst_Princi!CNTACTBLE) = "191409080101") Then
                        If Trim(g_rst_Princi!CNTACTBLE) = "191409010101" Then
                            .Cells(r_int_ConIn1, 6) = g_rst_Princi!Mes                                      'INTANGIBLES
                            .Cells(r_int_ConIn1, 7).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191409030101" Then
                            .Cells(r_int_ConIn2, 6) = g_rst_Princi!Mes
                            .Cells(r_int_ConIn2, 7).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191409040101" Then
                            .Cells(r_int_ConIn3, 6) = g_rst_Princi!Mes
                            .Cells(r_int_ConIn3, 7).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
                        ElseIf Trim(g_rst_Princi!CNTACTBLE) = "191409080101" Then
                            .Cells(r_int_ConIn4, 6) = g_rst_Princi!Mes
                            .Cells(r_int_ConIn4, 7).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
                        End If
                   ElseIf g_rst_Princi!GRUPO = 6 And Trim(g_rst_Princi!CNTACTBLE) = "191302010102" Then    'IMPUESTO DIFERIDO
                        .Cells(r_int_ConAux, 4) = g_rst_Princi!Mes
                        .Cells(r_int_ConAux - 1, 4) = g_rst_Princi!Mes
                        .Cells(r_int_ConAux + 2, 4) = .Cells(r_int_ConAux + 2, 4) - g_rst_Princi!Mes
                        .Cells(r_int_ConAux + 3, 4) = .Cells(r_int_ConAux + 3, 4) - g_rst_Princi!Mes
                   ElseIf g_rst_Princi!GRUPO = 13 And Trim(g_rst_Princi!CNTACTBLE) = "262602010101" Then
                        .Rows(r_int_ConAux - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        .Cells(r_int_Contad - 9, 3) = "Otros Adeudados y Obligaciones del Pais y del exterior"
                        .Cells(r_int_Contad - 9, 3).Font.Bold = True
                        .Cells(r_int_Contad - 9, 4) = g_rst_Princi!Mes
                        .Cells(r_int_Contad - 9, 4).Font.Bold = True
                        .Cells(r_int_Contad - 8, 4) = g_rst_Princi!Mes
                        .Cells(r_int_Contad - 21, 4) = .Cells(r_int_Contad - 21, 4) + g_rst_Princi!Mes
                        .Cells(r_int_Contad - 7, 4) = .Cells(r_int_Contad - 7, 4) - g_rst_Princi!Mes
                        .Cells(r_int_Contad - 6, 4) = .Cells(r_int_Contad - 6, 4) - g_rst_Princi!Mes
                   End If
        
                   If Trim(g_rst_Princi!CNTACTBLE) = "191302010102" Or Trim(g_rst_Princi!CNTACTBLE) = "262602010101" Then
                         If Trim(g_rst_Princi!CNTACTBLE) = "191302010102" Then                              'IMPUESTO DIFERIDO
                            .Cells(r_int_Contad, 4) = ""
                         End If
                         If Trim(g_rst_Princi!CNTACTBLE) = "262602010101" Then
                            .Rows(r_int_ConAux + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                         End If
                   ElseIf (Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "S") And g_rst_Princi!GRUPO = 5 Then
                        .Cells(r_int_Contad, 7) = g_rst_Princi!Mes
                   ElseIf .Cells(r_int_Contad, 3) = "TOTAL PASIVO" And r_int_ConIgv > 0 Then
                        .Cells(r_int_Contad, 4) = g_rst_Princi!Mes - r_dbl_MtoIgv
                   ElseIf .Cells(r_int_Contad, 3) = "TOTAL PASIVO Y PATRIMONIO" And r_int_ConIgv > 0 Then
                        .Cells(r_int_Contad, 4) = g_rst_Princi!Mes - r_dbl_MtoIgv
                        r_dbl_MtoIgv = 0
                   Else
                        .Cells(r_int_Contad, 4) = g_rst_Princi!Mes
                   End If
                    
                   If g_rst_Princi!CNTACTBLE = "191602010101" Then r_dbl_CreFis = g_rst_Princi!Mes
                   If g_rst_Princi!CNTACTBLE = "251703020101" Then r_dbl_MtoIgv = g_rst_Princi!Mes
                   If g_rst_Princi!CNTACTBLE = "191602010101" Then r_int_ConIgv = r_int_Contad
                   If r_dbl_CreFis > r_dbl_MtoIgv And r_dbl_MtoIgv > 0 And r_dbl_CreFis > 0 And r_int_ConIgv > 0 Then
                        If r_int_ConIgv <> 0 Then
                            .Cells(r_int_ConIgv + 1, 2) = .Cells(r_int_Contad, 2)
                            .Cells(r_int_ConIgv + 1, 3) = .Cells(r_int_Contad, 3)
                            .Cells(r_int_ConIgv + 1, 4) = .Cells(r_int_Contad, 4) * -1
                            .Cells(r_int_ConIgv - 1, 4) = .Cells(r_int_ConIgv - 1, 4) + .Cells(r_int_ConIgv + 1, 4)
                            .Cells(r_int_ConIgv - 2, 4) = .Cells(r_int_ConIgv - 1, 4)
                            .Cells(r_int_ConIgv + 2, 4) = .Cells(r_int_ConIgv + 2, 4) - r_dbl_MtoIgv
                            .Cells(r_int_Contad, 2) = ""
                            .Cells(r_int_Contad, 4) = 0
                            .Cells(r_int_Contad - 1, 4) = 0
                            .Cells(r_int_Contad - 2, 4) = 0
                        End If
                        
                        .Rows(r_int_ConIgv + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        r_dbl_CreFis = 0
                        r_int_Contad = r_int_Contad + 1
                   End If
               End If
        End If
Saltar:

            g_rst_Princi.MoveNext
        DoEvents
      Loop
      If r_int_ConDif <> 0 Then
         If .Cells(r_int_ConDif, 4) = "" Then
             .Rows(r_int_ConDif).EntireRow.Delete
             .Rows(r_int_ConDif).EntireRow.Delete
             .Rows(r_int_ConDif).EntireRow.Delete
         End If
      End If
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      r_int_ConAux = 0
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_ValidaCtaCntb(ByVal g_rst_Princi As ADODB.Recordset)
Dim r_obj_Excel         As Excel.Application
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_int_NumHoj        As Integer
Dim r_int_ConAux        As Integer
Dim r_bol_FlgEst        As Boolean

   r_int_Contad = 6
   r_int_PerMes = CInt(cmb_Period.ItemData(cmb_Period.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   Else
      g_rst_Princi.MoveFirst
      Do Until g_rst_Princi.EOF
         r_int_NumHoj = r_int_NumHoj + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = r_int_NumHoj
   r_obj_Excel.Workbooks.Add
   
   g_rst_Princi.MoveFirst
   
   For r_int_ConAux = 1 To r_int_NumHoj
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT RPT_MONEDA, RPT_CODIGO, RPT_DESCRI, RPT_VALNUM01, RPT_VALNUM02, RPT_VALCAD01, RPT_VALCAD02, RPT_VALNUM24, RPT_VALNUM25 FROM RPT_TABLA_TEMP RT "
      g_str_Parame = g_str_Parame & "  WHERE RT.RPT_PERMES = " & CInt(l_str_PerMes) & "   AND RT.RPT_PERANO = " & CInt(l_str_PerAno) & " "
      g_str_Parame = g_str_Parame & "    AND RT.RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RT.RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame & "    AND TRIM(RT.RPT_NOMBRE) =  'REPORTE DE " & Trim(cmb_TipRep.Text) & "' AND RT.RPT_VALNUM25 = " & g_rst_Princi!NUMVAL & " "
      g_str_Parame = g_str_Parame & "  ORDER BY RT.RPT_VALNUM01, RT.RPT_CODIGO "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         MsgBox "Error en la Consulta.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         MsgBox "No se encontraron errores.", vbInformation, modgen_g_str_NomPlt
         r_bol_FlgEst = False
         Exit Sub
      End If
      
      If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
      
         r_bol_FlgEst = True
         r_obj_Excel.Sheets(r_int_ConAux).Name = "VALID" & r_int_ConAux
   
         With r_obj_Excel.Sheets(r_int_ConAux)
            .Cells(1, 1) = "EDPYME MICASITA S.A."
            If g_rst_Princi!NUMVAL = 3 Or g_rst_Princi!NUMVAL = 4 Or g_rst_Princi!NUMVAL = 5 Then
               .Range(.Cells(1, 1), .Cells(1, 5)).Merge
               .Range(.Cells(1, 1), .Cells(1, 5)).Font.Bold = True
            Else
               .Range(.Cells(1, 1), .Cells(1, 4)).Merge
               .Range(.Cells(1, 1), .Cells(1, 4)).Font.Bold = True
            End If
            
            .Cells(2, 1) = "VALIDACIÓN DE CUENTAS CONTABLES"
            If g_rst_Princi!NUMVAL = 3 Or g_rst_Princi!NUMVAL = 4 Or g_rst_Princi!NUMVAL = 5 Then
               .Range(.Cells(2, 1), .Cells(2, 5)).Merge
               .Range(.Cells(2, 1), .Cells(2, 5)).Font.Bold = True
            Else
               .Range(.Cells(2, 1), .Cells(2, 4)).Merge
               .Range(.Cells(2, 1), .Cells(2, 4)).Font.Bold = True
            End If
            
            .Cells(3, 1) = "PERIODO : " & UCase(cmb_Period.Text) & " " & r_int_PerAno
            If g_rst_Princi!NUMVAL = 3 Or g_rst_Princi!NUMVAL = 4 Or g_rst_Princi!NUMVAL = 5 Then
               .Range(.Cells(3, 1), .Cells(3, 5)).Merge
               With .Range(.Cells(1, 1), .Cells(4, 5))
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlBottom
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
               End With
            Else
               .Range(.Cells(3, 1), .Cells(3, 4)).Merge
               With .Range(.Cells(1, 1), .Cells(4, 4))
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlBottom
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
               End With
            End If
   
            r_int_Contad = 7
            
            If g_rst_Princi!NUMVAL = 1 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 13
               .Columns("B").HorizontalAlignment = xlHAlignCenter
               .Columns("C").ColumnWidth = 17
               .Columns("C").HorizontalAlignment = xlHAlignCenter
               .Columns("D").ColumnWidth = 17
               .Columns("D").HorizontalAlignment = xlHAlignCenter
               
               .Cells(r_int_Contad, 2) = "ORIGEN"
               .Cells(r_int_Contad, 3) = "NRO_LIBRO"
               .Cells(r_int_Contad, 4) = "NRO_ASIENTO"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).HorizontalAlignment = xlHAlignCenter
            
            ElseIf g_rst_Princi!NUMVAL = 2 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 21
               .Columns("B").NumberFormat = "###,###,###,##0.00"
               .Columns("C").ColumnWidth = 21
               .Columns("C").NumberFormat = "###,###,###,##0.00"
               
               .Cells(r_int_Contad, 2) = "ACTIVO"
               .Cells(r_int_Contad, 3) = "PASIVO + PATRIMONIO"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 3)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 3)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 3)).HorizontalAlignment = xlHAlignCenter
            
            ElseIf g_rst_Princi!NUMVAL = 3 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 13
               .Columns("B").HorizontalAlignment = xlHAlignCenter
               .Columns("C").ColumnWidth = 33
               .Columns("D").ColumnWidth = 18
               .Columns("D").NumberFormat = "###,###,###,##0.00"
               .Columns("E").ColumnWidth = 18
               .Columns("E").NumberFormat = "###,###,###,##0.00"
               
               .Cells(r_int_Contad, 2) = "MONEDA"
               .Cells(r_int_Contad, 3) = "DESCRIPCION"
               .Cells(r_int_Contad, 4) = "DEBE"
               .Cells(r_int_Contad, 5) = "HABER"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).HorizontalAlignment = xlHAlignCenter
               
            ElseIf g_rst_Princi!NUMVAL = 4 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 13
               .Columns("B").HorizontalAlignment = xlHAlignCenter
               .Columns("C").ColumnWidth = 13
               .Columns("C").HorizontalAlignment = xlHAlignCenter
               .Columns("D").ColumnWidth = 13
               .Columns("D").HorizontalAlignment = xlHAlignCenter
               .Columns("E").ColumnWidth = 13
               .Columns("E").HorizontalAlignment = xlHAlignCenter
               
               .Cells(r_int_Contad, 2) = "ORIGEN"
               .Cells(r_int_Contad, 3) = "NRO_LIBRO"
               .Cells(r_int_Contad, 4) = "NRO_ASIENTO"
               .Cells(r_int_Contad, 5) = "CUENTA"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).HorizontalAlignment = xlHAlignCenter
               
            ElseIf g_rst_Princi!NUMVAL = 5 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 13
               .Columns("B").HorizontalAlignment = xlHAlignCenter
               .Columns("C").ColumnWidth = 13
               .Columns("C").HorizontalAlignment = xlHAlignCenter
               .Columns("D").ColumnWidth = 18
               .Columns("D").NumberFormat = "###,###,###,##0.00"
               .Columns("E").ColumnWidth = 18
               .Columns("E").NumberFormat = "###,###,###,##0.00"
               
               .Cells(r_int_Contad, 2) = "TIPO"
               .Cells(r_int_Contad, 3) = "CUENTA"
               .Cells(r_int_Contad, 4) = "DEBE"
               .Cells(r_int_Contad, 5) = "HABER"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).HorizontalAlignment = xlHAlignCenter
               
            ElseIf g_rst_Princi!NUMVAL = 6 Then
               .Columns("A").ColumnWidth = 4
               .Columns("B").ColumnWidth = 13
               .Columns("B").HorizontalAlignment = xlHAlignCenter
               .Columns("C").ColumnWidth = 18
               .Columns("C").NumberFormat = "###,###,###,##0.00"
               .Columns("D").ColumnWidth = 18
               .Columns("D").NumberFormat = "###,###,###,##0.00"
               
               .Cells(r_int_Contad, 2) = "CUENTA"
               .Cells(r_int_Contad, 3) = "DEBE"
               .Cells(r_int_Contad, 4) = "HABER"
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Interior.Color = RGB(146, 208, 80)
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 4)).HorizontalAlignment = xlHAlignCenter
            End If
            
            .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
            .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
               
            If g_rst_Princi!NUMVAL <> 3 Then
               .Cells(5, 1) = g_rst_Genera!RPT_DESCRI
               If g_rst_Princi!NUMVAL = 4 Or g_rst_Princi!NUMVAL = 5 Then .Range(.Cells(5, 1), .Cells(5, 5)).Merge Else .Range(.Cells(5, 1), .Cells(5, 4)).Merge
               .Cells(5, 1).HorizontalAlignment = xlCenter
            Else
               .Cells(5, 1) = "ERRORES DE SUMATORIA EN CLASE DE CUENTAS"
               .Range(.Cells(5, 1), .Cells(5, 5)).Merge
               .Cells(5, 1).HorizontalAlignment = xlCenter
            End If
            r_int_Contad = 8
            
            Do Until g_rst_Genera.EOF
               If g_rst_Genera!RPT_VALNUM25 = 1 Then
                  .Cells(r_int_Contad, 2) = g_rst_Genera!RPT_VALCAD01   'ORIGEN
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_VALNUM01   'NRO_LIBRO
                  .Cells(r_int_Contad, 4) = g_rst_Genera!RPT_VALNUM24   'NRO_ASIENTO
               ElseIf g_rst_Genera!RPT_VALNUM25 = 2 Then
                  .Cells(r_int_Contad, 2) = g_rst_Genera!RPT_VALNUM01   'ACTIVO
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_VALNUM02   'PASIVO + PATRIMONIO
                  .Cells(r_int_Contad + 1, 1) = "DIFERENCIA ->"
                  .Cells(r_int_Contad + 1, 1).ColumnWidth = 11
                  .Cells(r_int_Contad + 1, 1).Font.Bold = True
                  If .Cells(r_int_Contad, 2) > .Cells(r_int_Contad, 3) Then
                     .Cells(r_int_Contad + 1, 3) = .Cells(r_int_Contad, 2) - .Cells(r_int_Contad, 3)
                  Else
                     .Cells(r_int_Contad + 1, 2) = .Cells(r_int_Contad, 3) - .Cells(r_int_Contad, 2)
                  End If
                  .Cells(r_int_Contad + 2, 2).FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
                  .Cells(r_int_Contad + 2, 3).FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
               
               ElseIf g_rst_Genera!RPT_VALNUM25 = 3 Then
                  .Cells(r_int_Contad, 2) = IIf(g_rst_Genera!RPT_MONEDA = 1, "NACIONAL", IIf(g_rst_Genera!RPT_MONEDA = 2, "EXTRANJERA", "EQUIVALENTE")) 'MONEDA
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_DESCRI     'DESCRIPCION
                  .Cells(r_int_Contad, 4) = g_rst_Genera!RPT_VALNUM01   'DEBE
                  .Cells(r_int_Contad, 5) = g_rst_Genera!RPT_VALNUM02   'HABER
               ElseIf g_rst_Genera!RPT_VALNUM25 = 4 Then
                  .Cells(r_int_Contad, 2) = g_rst_Genera!RPT_VALCAD01   'ORIGEN
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_VALNUM01   'NRO_LIBRO
                  .Cells(r_int_Contad, 4) = g_rst_Genera!RPT_VALNUM02   'NRO_ASIENTO
                  .Cells(r_int_Contad, 5) = g_rst_Genera!RPT_VALCAD02   'CUENTA
               ElseIf g_rst_Genera!RPT_VALNUM25 = 5 Then
                  .Cells(r_int_Contad, 2) = g_rst_Genera!RPT_VALCAD01   'TIPO
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_VALCAD02     'CUENTA
                  .Cells(r_int_Contad, 4) = g_rst_Genera!RPT_VALNUM01   'DEBE
                  .Cells(r_int_Contad, 5) = g_rst_Genera!RPT_VALNUM02   'HABER
               ElseIf g_rst_Genera!RPT_VALNUM25 = 6 Then
                  .Cells(r_int_Contad, 2) = g_rst_Genera!RPT_VALCAD02   'CUENTA
                  .Cells(r_int_Contad, 3) = g_rst_Genera!RPT_VALNUM01   'DEBE
                  .Cells(r_int_Contad, 4) = g_rst_Genera!RPT_VALNUM02   'HABER
               End If
               r_int_Contad = r_int_Contad + 1
               g_rst_Genera.MoveNext
               
            Loop
         g_rst_Princi.MoveNext
         
         End With
      End If
   Next r_int_ConAux
   
   g_rst_Princi.Close
   g_rst_Genera.Close
   Set g_rst_Princi = Nothing
   Set g_rst_Genera = Nothing
   
   If r_bol_FlgEst = True Then
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   Else
      Set r_obj_Excel = Nothing
   End If

End Sub

Private Sub fs_Recorset_nc()
    Set g_rst_GenAux = New ADODB.Recordset
    
    g_rst_GenAux.Fields.Append "GRUPO", adBigInt, 2, adFldFixed
    g_rst_GenAux.Fields.Append "NOMGRUPO", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "SUBGRP", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "NOMSUBGRP", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "CNTACTBLE", adChar, 30, adFldIsNullable
    g_rst_GenAux.Fields.Append "NOMCTA", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "MES", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "INDTIPO", adChar, 5, adFldFixed
    g_rst_GenAux.Fields.Append "TIPO", adChar, 1, adFldFixed
    g_rst_GenAux.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub fs_Formato_NumNeg(ByVal r_obj_rango As Excel.Range)
   If r_obj_rango.Value < 0 Then
      r_obj_rango.Font.Color = vbRed
   End If
End Sub

Private Sub fs_Bordes_SitFin1(ByVal r_obj_rango As Excel.Range)
   r_obj_rango.Borders(xlDiagonalDown).LineStyle = xlNone
   r_obj_rango.Borders(xlDiagonalUp).LineStyle = xlNone
   r_obj_rango.Borders(xlEdgeLeft).LineStyle = xlNone
   With r_obj_rango.Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With r_obj_rango.Borders(xlEdgeBottom)
       .LineStyle = xlDouble
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThick
   End With
   With r_obj_rango.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
   End With
   With r_obj_rango.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
   End With
   r_obj_rango.Borders(xlInsideVertical).LineStyle = xlNone
   r_obj_rango.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Private Sub fs_Bordes_SitFin(ByVal r_obj_rango As Excel.Range)
   r_obj_rango.Borders(xlDiagonalDown).LineStyle = xlNone
   r_obj_rango.Borders(xlDiagonalUp).LineStyle = xlNone
   With r_obj_rango.Borders(xlEdgeLeft)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With r_obj_rango.Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With r_obj_rango.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   With r_obj_rango.Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   r_obj_rango.Borders(xlInsideVertical).LineStyle = xlNone
   r_obj_rango.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Private Sub fs_GenExc_EstRes(ByVal r_obj_Excel As Excel.Application)
Dim r_int_ConVer     As Integer
Dim r_int_VarAux     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RT.RPT_CODIGO CODIGO, RT.RPT_DESCRI DESCRIPCION, RT.RPT_VALNUM01 MN, RT.RPT_VALNUM02 EQ, RT.RPT_VALNUM03 TOT, RS.FLAG_NEGRITA NEGRITA, RS.FLAG_SUBRAYADO SUBRAYADO "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP RT "
   g_str_Parame = g_str_Parame & "  INNER JOIN RPT_SUBGRUPO RS ON TRIM(RS.SUBGRUPO) = TRIM(RT.RPT_CODIGO) AND RS.REPORTE ='EstBal2 ' AND RS.GRUPO='Unico'"
   g_str_Parame = g_str_Parame & " WHERE RT.RPT_PERMES = " & CInt(l_str_PerMes) & " AND RT.RPT_PERANO = " & CInt(l_str_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "       RT.RPT_USUCRE = '" & modgen_g_str_CodUsu & "' AND RT.RPT_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "       RT.RPT_NOMBRE = 'ESTADO DE RESULTADOS' "
   g_str_Parame = g_str_Parame & " ORDER BY TO_NUMBER (TRIM(RT.RPT_CODIGO), 99) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_obj_Excel.Sheets(2).Name = "2.Resultado del ejercicio"
 
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(1, 1), .Cells(1, 4)).Merge
      .Cells(1, 1) = "Institución: EDPYME MICASITA SA"
      
      .Range(.Cells(2, 1), .Cells(2, 4)).Merge
      .Cells(2, 1) = "ESTADO DE RESULTADOS"
      
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Cells(3, 1) = "Al " & ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno)) & " de " & UCase(Left(Me.cmb_Period.Text, 1)) & LCase(Right(cmb_Period.Text, Len(cmb_Period.Text) - 1)) & " de " & ipp_PerAno.Text
      
      .Range(.Cells(4, 1), .Cells(4, 4)).Merge
      .Cells(4, 1) = "( En soles )"
                   
      .Range(.Cells(1, 1), .Cells(5, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(4, 4)).Font.Bold = True
     
      .Columns("A").ColumnWidth = 68
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").NumberFormat = "###,###,##0.00"
      .Columns("C").ColumnWidth = 8.3
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("A").NumberFormat = "@"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 7
   
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!NEGRITA = 1 Then
          r_obj_Excel.Sheets(2).Range(r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 1), r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 4)).Font.Bold = True
      End If
      
      If g_rst_Princi!SUBRAYADO = 1 Then
         With r_obj_Excel.Sheets(2)
            With .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 4))
               .Borders(xlDiagonalDown).LineStyle = xlNone
               .Borders(xlDiagonalUp).LineStyle = xlNone
               .Borders(xlEdgeLeft).LineStyle = xlNone
            With .Borders(xlEdgeTop)
                  .LineStyle = xlContinuous
                  .ColorIndex = 0
                  .TintAndShade = 0
                  .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                  .LineStyle = xlDouble
                  .ColorIndex = 0
                  .TintAndShade = 0
                  .Weight = xlThick
            End With
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With
         End With
      End If
      
      r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 1) = g_rst_Princi!DESCRIPCION
      r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 2) = IIf(g_rst_Princi!MN = 0, "", CDbl(Format(g_rst_Princi!MN, "###,###,##0.00")))
      r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 3) = IIf(g_rst_Princi!Eq = 0, "", CDbl(Format(g_rst_Princi!Eq, "###,###,##0.00")))
      r_obj_Excel.Sheets(2).Cells(r_int_ConVer, 4) = IIf(g_rst_Princi!TOT = 0, "", CDbl(Format(g_rst_Princi!TOT, "###,###,##0.00")))
      
      For r_int_VarAux = 2 To 4
         Call fs_Formato_NumNeg(r_obj_Excel.Sheets(2).Cells(r_int_ConVer, r_int_VarAux))
      Next r_int_VarAux
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 4)).Font.Name = "Arial Narrow"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 4)).Font.Size = 8 '10
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 11
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 15
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Size = 11

      Call fs_Bordes_SitFin(.Range(.Cells(7, 1), .Cells(7, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(17, 1), .Cells(17, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(21, 1), .Cells(21, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(28, 1), .Cells(28, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(30, 1), .Cells(31, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(40, 1), .Cells(40, 4)))
      
      With .Range(.Cells(5, 1), .Cells(r_int_ConVer - 1, 4))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         With .Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeTop)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With .Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThin
         End With
         .Borders(xlInsideVertical).LineStyle = xlNone
      End With
      
      Call fs_Bordes_SitFin1(.Range(.Cells(r_int_ConVer - 1, 1), .Cells(r_int_ConVer - 1, 4)))
      .Columns("B:B").EntireColumn.AutoFit
      '.Columns("C:C").EntireColumn.AutoFit
      .Columns("D:D").EntireColumn.AutoFit
   End With
   
   'CABECERA
   With r_obj_Excel.Sheets(2)
      .Cells(5, 1) = ""
      .Cells(5, 2) = "Moneda Nacional"
      .Cells(5, 3) = "Equivalente   en M.E."
      .Cells(5, 4) = "TOTAL"

      .Range(.Cells(5, 1), .Cells(6, 6)).Font.Bold = True
       With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
      End With
      .Range(.Cells(5, 1), .Cells(6, 1)).Merge
      With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = True
      End With
      With .Range(.Cells(5, 1), .Cells(6, 1))
         .HorizontalAlignment = xlLeft
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = True
      End With
      With .Range(.Cells(5, 2), .Cells(6, 4))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
      End With
      With .Range(.Cells(5, 2), .Cells(6, 2))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
      End With
      .Range(.Cells(5, 2), .Cells(6, 2)).Merge
      With .Range(.Cells(5, 3), .Cells(6, 3))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
      End With
      .Range(.Cells(5, 3), .Cells(6, 3)).Merge
      With .Range(.Cells(5, 4), .Cells(6, 4))
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .WrapText = True
           .Orientation = 0
           .AddIndent = False
           .IndentLevel = 0
           .ShrinkToFit = False
           .ReadingOrder = xlContext
           .MergeCells = False
      End With
      .Range(.Cells(5, 4), .Cells(6, 4)).Merge
      With .Range(.Cells(6, 1), .Cells(6, 4))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         .Borders(xlEdgeLeft).LineStyle = xlNone
         .Borders(xlEdgeTop).LineStyle = xlNone
         With .Borders(xlEdgeBottom)
              .LineStyle = xlContinuous
              .ColorIndex = 0
              .TintAndShade = 0
              .Weight = xlThin
         End With
         .Borders(xlEdgeRight).LineStyle = xlNone
         .Borders(xlInsideVertical).LineStyle = xlNone
         .Borders(xlInsideHorizontal).LineStyle = xlNone
      End With
   End With
   
   'MARGENES DE IMPRESIÓN
   l_Mar_Izq = 0.3
   l_Mar_Der = 0.3
   l_Mar_Sup = 0.3
   l_Mar_Inf = 0.3
   
   With r_obj_Excel.Sheets(2).PageSetup
      If .Orientation = xlPortrait Then
         .Orientation = xlPortrait
      Else
         .Orientation = xlLandscape
      End If
      
      'Configuración de márgenes:
      .LeftMargin = Application.CentimetersToPoints(l_Mar_Izq)
      .RightMargin = Application.CentimetersToPoints(l_Mar_Der)
      .TopMargin = Application.CentimetersToPoints(l_Mar_Sup)
      .BottomMargin = Application.CentimetersToPoints(l_Mar_Inf)
      .CenterHorizontally = True
      .CenterVertically = True
   End With
   
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(46, 1), .Cells(56, 1)).EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
      
      'CABECERA PÁGINA
      .Range(.Cells(1, 1), .Cells(6, 4)).Copy
      .Range("A57").Insert Shift:=xlDown
      
      With .Range(.Cells(46, 1), .Cells(56, 4))
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
         .Borders(xlEdgeLeft).LineStyle = xlNone
         With .Borders(xlEdgeTop)
             .LineStyle = xlDouble
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlThick
         End With
         .Borders(xlEdgeBottom).LineStyle = xlNone
         .Borders(xlEdgeRight).LineStyle = xlNone
         .Borders(xlInsideVertical).LineStyle = xlNone
         .Borders(xlInsideHorizontal).LineStyle = xlNone
      End With
   
      .Range(.Cells(54, 1), .Cells(54, 4)).Merge
      .Cells(54, 1) = "________________________      ________________________     ________________________     ________________________"
      .Cells(55, 1) = "                 Director                                                 Director                                 Gerente General                               Contador"
      .Range(.Cells(54, 1), .Cells(55, 4)).Font.Name = "Calibri"
      .Range(.Cells(54, 1), .Cells(55, 4)).Font.Size = 11
      .Range(.Cells(54, 1), .Cells(55, 4)).Font.Bold = False
      
      .Range(.Cells(109, 1), .Cells(109, 4)).Merge
      .Cells(109, 1) = "________________________      ________________________     ________________________     ________________________"
      .Cells(110, 1) = "                 Director                                                 Director                                 Gerente General                               Contador"
      
      Call fs_Bordes_SitFin(.Range(.Cells(35, 1), .Cells(35, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(61, 1), .Cells(62, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(68, 1), .Cells(69, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(72, 1), .Cells(74, 4)))
      Call fs_Bordes_SitFin(.Range(.Cells(5, 1), .Cells(6, 4)))

      .Rows("5:5").EntireRow.AutoFit
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
 
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function modsec_gf_Buscar_NomCtaHab(ByVal p_CtaCtb As String) As Boolean
   modsec_gf_Buscar_NomCtaHab = False
   
   g_str_Parame = "SELECT FLAG_REQ_SUCAVE FROM CNTBL_CNTA WHERE CNTA_CTBL= '" & Trim(p_CtaCtb) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_Buscar_NomCtaHab = IIf(IsNull(Trim(g_rst_Listas!FLAG_REQ_SUCAVE)), 0, Trim(g_rst_Listas!FLAG_REQ_SUCAVE))
   Else
       If Len(Trim(p_CtaCtb)) = 3 Then
         modsec_gf_Buscar_NomCtaHab = True
       End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GeneraArchivo_BCient(ByRef r_arr_Matriz() As g_tpo_Bcient)
Dim r_str_NomRes  As String
Dim r_int_NumRes  As String
Dim r_str_FecRpt  As String
Dim r_int_Conta   As Integer
Dim r_str_Cadena  As String
   
   'genera archivo texto
   r_int_NumRes = FreeFile
   r_str_NomRes = moddat_g_str_RutLoc & "\" & "Bcient " & UCase(Left(Me.cmb_Period.Text, 1)) & LCase(Mid(Me.cmb_Period.Text, 2, 2)) & Right(Me.ipp_PerAno.Text, 2) & ".100"
   
   Open r_str_NomRes For Output As #r_int_NumRes
   
   For r_int_Conta = 1 To UBound(r_arr_Matriz)
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).Bcient_FecMov, 1, " ", 6)
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).Bcient_Entdad, 1, " ", 3)
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).Bcient_Cuenta, 1, " ", Len(r_arr_Matriz(r_int_Conta).Bcient_Cuenta))
      r_str_Cadena = r_str_Cadena & GeneraLongitud(Format(Replace(r_arr_Matriz(r_int_Conta).Bcient_SldIni, "-", ""), "########0.00"), 1, "0", (37 - Len(r_arr_Matriz(r_int_Conta).Bcient_Cuenta)))
      r_str_Cadena = r_str_Cadena & IIf(InStr(r_arr_Matriz(r_int_Conta).Bcient_SldIni, "-") > 0, "-", "+")
      r_str_Cadena = r_str_Cadena & GeneraLongitud(Format(Replace(r_arr_Matriz(r_int_Conta).Bcient_Debito, "-", ""), "########0.00"), 1, "0", 17)
      r_str_Cadena = r_str_Cadena & IIf(InStr(r_arr_Matriz(r_int_Conta).Bcient_Debito, "-") > 0, "-", "+")
      r_str_Cadena = r_str_Cadena & GeneraLongitud(Format(Replace(r_arr_Matriz(r_int_Conta).Bcient_Credit, "-", ""), "########0.00"), 1, "0", 17)
      r_str_Cadena = r_str_Cadena & IIf(InStr(r_arr_Matriz(r_int_Conta).Bcient_Credit, "-") > 0, "-", "+")
      r_str_Cadena = r_str_Cadena & GeneraLongitud(Format(Replace(r_arr_Matriz(r_int_Conta).Bcient_SldFin, "-", ""), "########0.00"), 1, "0", 17)
      r_str_Cadena = r_str_Cadena & IIf(InStr(r_arr_Matriz(r_int_Conta).Bcient_SldFin, "-") > 0, "-", "+")
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).Bcient_Filler, 1, " ", 18)
      
      Print #r_int_NumRes, r_str_Cadena
   Next r_int_Conta
   
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
End Sub

Private Sub fs_GeneraArchivo_BCR_1(ByRef r_arr_Matriz() As g_tpo_BCR)
Dim r_str_NomRes  As String
Dim r_int_NumRes  As String
Dim r_str_FecRpt  As String
Dim r_int_Conta   As Integer
Dim r_int_CodBCR  As String
Dim r_str_Cadena  As String
   
   '-- Lectura de Codigo BCR de la empresa financiera
   g_str_Parame = "SELECT TRIM(codigo_bcr) CodBCR FROM genparam "
   g_str_Parame = g_str_Parame & "WHERE reckey = '1' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_CodBCR = g_rst_Princi!CodBCR
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   l_str_FecImp = "01" & "/" & l_str_PerMes & "/" & l_str_PerAno
   l_str_FecImp = Mid(l_str_FecImp, 4, 2) + Mid(l_str_FecImp, 9, 2)
   
   'genera archivo texto
   r_int_NumRes = FreeFile
   r_str_NomRes = moddat_g_str_RutLoc & "\" & r_int_CodBCR & l_str_FecImp & ".TXT"
   
   Open r_str_NomRes For Output As #r_int_NumRes
   
   For r_int_Conta = 1 To UBound(r_arr_Matriz)
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).BCR_Codigo, 1, " ", 19)
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).BCR_SldAju, 1, " ", 12)
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).BCR_SldMon, 1, " ", 12)
      r_str_Cadena = r_str_Cadena & GeneraLongitud(r_arr_Matriz(r_int_Conta).BCR_SldEqu, 1, " ", 12)
          
      Print #r_int_NumRes, r_str_Cadena
   Next r_int_Conta
   
   For r_int_Conta = 1 To UBound(r_arr_Matriz)
      If r_arr_Matriz(r_int_Conta).BCR_AjuDif <> "" Then
         Print #r_int_NumRes, GeneraLongitud(r_arr_Matriz(r_int_Conta).BCR_AjuDif, 1, " ", 72)
      End If
   Next r_int_Conta
   
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
End Sub

Private Sub fs_GeneraArchivo_BCR_2(ByRef r_arr_Matriz() As g_tpo_BCR, ByRef r_arr_MtzCta() As g_tpo_CNTABSI)
Dim r_str_NomRes  As String
Dim r_int_NumRes  As String
Dim r_str_FecRpt  As String
Dim r_int_Conta   As Integer
Dim r_int_CodBCR  As String
Dim r_int_CodSBS  As String
Dim r_str_Cadena  As String
   
   '-- Lectura de Codigo BCR de la empresa financiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(codigo_bcr) CODBCR, TRIM(codigo_sbs) CODSBS "
   g_str_Parame = g_str_Parame & "  FROM genparam "
   g_str_Parame = g_str_Parame & " WHERE reckey = '1' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_int_CodSBS = g_rst_Genera!CODSBS
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
      
   l_str_FecImp = "01" & "/" & l_str_PerMes & "/" & l_str_PerAno
   l_str_FecImp = Format(l_str_PerAno, "0000") & Format(l_str_PerMes, "00")
   
   'genera archivo texto - previo
   r_int_NumRes = FreeFile
   r_str_Cadena = ""
   r_str_Cadena = "BSI" & r_int_CodSBS & l_str_FecImp & "P.TXT"
   r_str_NomRes = moddat_g_str_RutLoc & "\" & r_str_Cadena
   
   Open r_str_NomRes For Output As #r_int_NumRes
   r_str_Cadena = GeneraLongitud(Left(r_str_Cadena, 13), 1, "", 13)
   Print #r_int_NumRes, r_str_Cadena
   For r_int_Conta = 1 To UBound(r_arr_Matriz)
       r_str_Cadena = ""
       r_str_Cadena = r_str_Cadena & Trim(l_str_FecImp)
       r_str_Cadena = r_str_Cadena & Trim(r_int_CodSBS)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_Codigo)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_CodSec)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_SldIni)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_SldFin)
       Print #r_int_NumRes, r_str_Cadena
   Next r_int_Conta
   Close #r_int_NumRes
   
   'genera archivo texto - final
   r_int_NumRes = FreeFile
   r_str_Cadena = ""
   r_str_Cadena = "BSI" & r_int_CodSBS & l_str_FecImp & "D.TXT"
   r_str_NomRes = moddat_g_str_RutLoc & "\" & r_str_Cadena
   
   Open r_str_NomRes For Output As #r_int_NumRes
   r_str_Cadena = GeneraLongitud(Left(r_str_Cadena, 13), 1, "", 13)
   Print #r_int_NumRes, r_str_Cadena
   For r_int_Conta = 1 To UBound(r_arr_Matriz)
       r_str_Cadena = ""
       r_str_Cadena = r_str_Cadena & Trim(l_str_FecImp)
       r_str_Cadena = r_str_Cadena & Trim(r_int_CodSBS)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_Codigo)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_CodSec)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_SldIni)
       r_str_Cadena = r_str_Cadena & Trim(r_arr_Matriz(r_int_Conta).BCR_SldFin)
       Print #r_int_NumRes, r_str_Cadena
   Next r_int_Conta
   Close #r_int_NumRes
End Sub

Private Function GeneraLongitud(ByVal p_parametro As String, ByVal p_Opcion As Integer, ByVal p_caracter As String, ByVal p_longitud As Integer, Optional ByVal p_Separador As String) As String
Dim l_cadena      As String
Dim l_cadaux      As String
Dim l_contad      As Integer
Dim l_longit      As Integer
Dim p_caracter2   As String
   
   If p_Separador = "" Then
      p_Separador = "."
   End If
   
   If p_Opcion = 1 Then
      l_longit = Len(p_parametro)
      For l_contad = 1 To l_longit Step 1
         If Mid(p_parametro, l_contad, 1) <> p_Separador Then
            l_cadena = l_cadena & Mid(p_parametro, l_contad, 1)
         End If
         If Mid(p_parametro, l_contad, 1) = "-" Then
            p_caracter2 = "-"
         End If
      Next
      
      If p_caracter = " " Then
         For l_contad = Len(l_cadena) To p_longitud - 1 Step 1
            l_cadena = p_caracter & l_cadena
         Next
         GeneraLongitud = l_cadena
      Else
         If p_caracter2 = "-" Then
            For l_contad = 1 To p_longitud - 1 Step 1
               l_cadaux = l_cadaux & p_caracter
            Next
            GeneraLongitud = Format(l_cadena, l_cadaux)
         Else
            For l_contad = Len(l_cadena) + 1 To p_longitud Step 1
               l_cadaux = l_cadaux & p_caracter
            Next
            GeneraLongitud = l_cadaux & l_cadena
         End If
      End If
   ElseIf p_Opcion = 2 Then
      
      l_longit = Len(p_parametro)
      For l_contad = l_longit To p_longitud - 1 Step 1
         p_parametro = p_parametro & p_caracter
      Next
      GeneraLongitud = p_parametro
   End If
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConAux     As Integer
Dim r_bol_FlgGrp     As Boolean
   
   g_str_Parame = ""
   g_str_Parame = "SELECT BALCOM_CUENTA, BALCOM_DESCRI, BALCOM_SLINDB, BALCOM_SLINHB, BALCOM_IMPDEB, BALCOM_IMPHAB, BALCOM_SLFIDB, BALCOM_SLFIHB, BALCOM_GRUPO FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = '" & l_str_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_USUCRE = '" & modgen_g_str_CodUsu & "' AND "
   g_str_Parame = g_str_Parame & "BALCOM_TERCRE = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & "' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(TRIM(BALCOM_CUENTA),LENGTH(TRIM(BALCOM_CUENTA))-1,2) <> 00 "
   g_str_Parame = g_str_Parame & " AND (BALCOM_SLINDB <> 0 OR BALCOM_SLINHB <> 0 OR BALCOM_IMPDEB <> 0 OR "
   g_str_Parame = g_str_Parame & " BALCOM_IMPHAB <> 0 OR BALCOM_SLFIDB <> 0 OR BALCOM_SLFIHB <> 0 ) "
   g_str_Parame = g_str_Parame & "ORDER BY BALCOM_CUENTA ASC, BALCOM_TIPMON ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Range(.Cells(1, 1), .Cells(1, 8)).Merge
      .Range(.Cells(1, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(2, 1)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
      .Cells(1, 1) = "BALANCE DE COMPROBACIÓN"
      
      .Range(.Cells(2, 1), .Cells(2, 8)).Merge
      .Range(.Cells(2, 1), .Cells(2, 1)).Font.Size = 11
      .Cells(2, 1) = Trim(Mid(Me.cmb_TipMon.Text, InStr(Me.cmb_TipMon.Text, "-") + 1))
      
      .Range(.Cells(3, 1), .Cells(3, 8)).Merge
      .Cells(3, 1) = "Del " & "01/" & l_str_PerMes & "/" & l_str_PerAno & " Al " & ff_Ultimo_Dia_Mes(l_str_PerMes, CInt(l_str_PerAno)) & "/" & l_str_PerMes & "/" & l_str_PerAno
      
      .Cells(5, 1) = "CUENTA"
      .Cells(5, 2) = "DESCRIPCIÓN"
      .Cells(5, 3) = "SAL. INI. DEBE"
      .Cells(5, 4) = "SAL. INI. HABER"
      .Cells(5, 5) = "MOV. MES. DEBE"
      .Cells(5, 6) = "MOV. MES. HABER"
      .Cells(5, 7) = "SAL. ACT. DEBE"
      .Cells(5, 8) = "SAL. ACT. HABER"
   
      .Range(.Cells(5, 1), .Cells(5, 8)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 8)).HorizontalAlignment = xlHAlignCenter
                  
      .Columns("A").ColumnWidth = 14.15
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("A").NumberFormat = "@"
      .Columns("B").ColumnWidth = 60
      .Columns("C").ColumnWidth = 15
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 15
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 6
   Do While Not g_rst_Princi.EOF
   
      If g_rst_Princi!BALCOM_GRUPO = 2 And r_bol_FlgGrp = False Then
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(3, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Size = 10
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Size = 11
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).NumberFormat = "###,###,##0.00"
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 6 & "]C:R[-1]C)"
         r_int_ConVer = r_int_ConVer + 1
         r_int_ConAux = r_int_ConVer
         r_bol_FlgGrp = True
      End If
         
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!BALCOM_CUENTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!BALCOM_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CDbl(Format(IIf(IsNull(g_rst_Princi!BALCOM_SLINDB), 0, g_rst_Princi!BALCOM_SLINDB), "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CDbl(Format(IIf(IsNull(g_rst_Princi!BALCOM_SLINHB), 0, g_rst_Princi!BALCOM_SLINHB), "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDbl(Format(IIf(IsNull(g_rst_Princi!BalCom_ImpDeb), 0, g_rst_Princi!BalCom_ImpDeb), "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDbl(Format(IIf(IsNull(g_rst_Princi!BalCom_ImpHab), 0, g_rst_Princi!BalCom_ImpHab), "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDbl(Format(IIf(IsNull(g_rst_Princi!BALCOM_SLFIDB), 0, g_rst_Princi!BALCOM_SLFIDB), "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDbl(Format(IIf(IsNull(g_rst_Princi!BALCOM_SLFIHB), 0, g_rst_Princi!BALCOM_SLFIHB), "###,###,##0.00"))
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Size = 10
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Size = 11
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).Font.Bold = True
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).NumberFormat = "###,###,##0.00"
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8)).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - r_int_ConAux & "]C:R[-1]C)"
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Period)
   End If
End Sub

Private Sub cmb_TipRep_Click()

'   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Or _
'      cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Or _
'      cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 6 Then
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) <> 1 Then
      cmb_TipMon.Enabled = False
      cmb_TipMon.ListIndex = -1
      cmd_Imprim.Visible = False
      Call gs_SetFocus(cmb_Period)
   Else
      Me.cmb_TipMon.Enabled = True
      cmd_Imprim.Visible = True
      Call gs_SetFocus(cmb_TipMon)
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMon)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
