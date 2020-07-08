VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RepSbs_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   6825
   ClientTop       =   6570
   ClientWidth     =   4905
   Icon            =   "GesCtb_frm_710.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
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
         TabIndex        =   1
         Top             =   60
         Width           =   4845
         _Version        =   65536
         _ExtentX        =   8546
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
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Anexo N° 14"
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
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0114-01 Obligaciones con el Exterior"
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
            Picture         =   "GesCtb_frm_710.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   4845
         _Version        =   65536
         _ExtentX        =   8546
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
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_710.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4230
            Picture         =   "GesCtb_frm_710.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_710.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_710.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_710.frx":11AE
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
         Width           =   4845
         _Version        =   65536
         _ExtentX        =   8546
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
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_dbl_Evalua(100)   As Double
Dim r_str_Denomi(20)    As String
Dim r_int_NumRes        As Integer
Dim r_str_PerMes        As Integer
Dim r_str_PerAno        As Integer
Dim r_int_Contad        As Integer
Dim r_int_ConGen        As Integer
Dim r_int_ConAux        As Integer
Dim r_int_ConTem        As Integer
Dim r_dbl_MonTot        As Double
Dim r_dbl_TipCam        As Double
Dim r_int_CodEmp        As Integer
Dim r_str_Cadena        As String
Dim r_str_NomRes        As String
Dim r_str_FecRpt        As String
Dim r_int_Cantid        As Integer
Dim r_int_FlgRpr        As Integer


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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_OBLEXT", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("¿Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_Genera_ArcPla
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_Genera_ArcPla
End Sub

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
   
  Call fs_GenRpt
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer

   r_int_PerMes = Month(date)
   r_int_PerAno = Year(date)
   
   If Month(date) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If
 
   Call gs_BuscarCombo_Item(cmb_PerMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
End Sub

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_OBLEXT", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

Private Sub fs_Genera_ArcPla()

   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
      
   r_str_NomRes = "C:\01" & Right(r_str_PerAno, 2) & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".114"
      
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
   
   Print #r_int_NumRes, Format(114, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_str_PerAno & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   'r_str_Cadena = Format(1, "0000") & gs_modsec_Genera("BID-FOMIN", 2, " ", 50) & gs_modsec_Genera("OTROS", 2, " ", 11) & gs_modsec_Genera("O", 2, " ", 1) & gs_modsec_Genera(Format(4016, "###########00"), 1, "0", 4) & gs_modsec_Genera(Format(2, "###########00"), 2, " ", 2)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(7.3, "###########0.000"), 1, "0", 7) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15)
   
   'Print #r_int_NumRes, r_str_Cadena
   
   'r_str_Cadena = Format(1000, "0000") & gs_modsec_Genera(" ", 2, " ", 50) & gs_modsec_Genera(" ", 2, " ", 11) & gs_modsec_Genera(" ", 2, " ", 1) & gs_modsec_Genera(" ", 2, " ", 4) & gs_modsec_Genera(" ", 2, " ", 2)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(7.3, "###########0.000"), 1, "0", 7) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot, "###########0.00"), 1, "0", 15)
   
   'Print #r_int_NumRes, r_str_Cadena
   
   'r_str_Cadena = Format(1100, "0000") & gs_modsec_Genera(" ", 2, " ", 50) & gs_modsec_Genera(" ", 2, " ", 11) & gs_modsec_Genera(" ", 2, " ", 1) & gs_modsec_Genera(" ", 2, " ", 4) & gs_modsec_Genera(" ", 2, " ", 2)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MonTot * r_dbl_TipCam, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.000"), 1, "0", 7) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)
   'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 15)

   r_int_ConTem = -1
   r_int_ConAux = -1
   
   For r_int_Contad = 1 To 1100 Step 100

      r_str_Cadena = gs_modsec_Genera(IIf(r_str_Denomi(r_int_ConTem + 1) = "BID-FOMIN", r_str_Denomi(r_int_ConTem + 1), ""), 2, " ", 50) & gs_modsec_Genera(r_str_Denomi(r_int_ConTem + 2), 2, " ", 11) & gs_modsec_Genera(Left(r_str_Denomi(r_int_ConTem + 3), 1), 2, " ", 1) & gs_modsec_Genera(IIf(r_str_Denomi(r_int_ConTem + 4) = "USA", 4016, ""), 2, " ", 4) & gs_modsec_Genera(IIf(r_str_Denomi(r_int_ConTem + 5) = "D.A.", "02", ""), 2, " ", 2)
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 1), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 2), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 3), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 4), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 5), "###########0.00"), 1, "0", 15)
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 6), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 7), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 8), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 9), "###########0.00"), 1, "0", 15)
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 10), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 11), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 12), "###########0.000"), 1, "0", 7) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 13), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 14), "###########0.00"), 1, "0", 15)
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 15), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 16), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 17), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 18), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 19), "###########0.00"), 1, "0", 15)
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 20), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 21), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 22), "###########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 23), "###########0.00"), 1, "0", 15)
   
      Print #r_int_NumRes, Format(r_int_Contad, "0000") & r_str_Cadena
      
      r_int_ConTem = r_int_ConTem + 5
      r_int_ConAux = r_int_ConAux + 23
      
      If r_int_Contad < 1000 Then
         r_int_Contad = r_int_Contad + 899
      End If
   
   Next
    
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt

End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   
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
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(3, 26), .Cells(5, 26)).HorizontalAlignment = xlHAlignRight
      .Cells(3, 26) = "Anexo Nº 14"
      .Cells(5, 26) = "CODIGO S.B.S.: 240"
      .Range(.Cells(3, 26), .Cells(3, 27)).Merge
      .Range(.Cells(5, 26), .Cells(5, 27)).Merge
            
      .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignLeft
      .Cells(3, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
      .Cells(5, 1) = "EMPRESA: MICASITA"
      .Range(.Cells(3, 1), .Cells(3, 2)).Merge
      .Range(.Cells(5, 1), .Cells(5, 2)).Merge
      
      .Range(.Cells(6, 1), .Cells(7, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(6, 1), .Cells(7, 28)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(7, 28)).Font.Underline = xlUnderlineStyleSingle
      .Cells(6, 14) = "Obligaciones con el Extranjero*"
      .Cells(7, 14) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_str_PerAno, "0,000")

               
      .Cells(12, 1) = "ENTIDAD 1/"
      .Cells(12, 2) = "CODIGO 2/"
      .Cells(12, 3) = "TIPO DE ENTIDAD 3/"
      .Cells(12, 4) = "PAIS DE ORIGEN"
      .Cells(12, 5) = "MONEDA"
      .Cells(12, 6) = "MONTO AUTORIZADO (EQUIVALENTE EN US$)"
      .Cells(12, 7) = "OBLIGACIONES DIRECTAS (EQUIVALENTE EN US$)"
      .Cells(12, 18) = "OBLIGACIONES CONTINGENTES (EQUIVALENTE EN US$)"
      .Cells(12, 28) = "TOTAL DE OBLIGACIONES CON EL EXTERIOR (A)+(B)(EQUIVALENTE EN US$)"
      
      .Cells(13, 7) = "POR DESTINO:"
      .Cells(13, 10) = "POR VENCER A:"
      .Cells(13, 16) = "TOTAL (A) 4/"
      .Cells(13, 17) = "TASA DE INTERES PROMEDIO 5/"
      .Cells(13, 18) = "POR DESTINO:"
      .Cells(13, 21) = "POR VENCER A:"
      .Cells(13, 27) = "TOTAL (B) 4/"
      
      .Cells(14, 7) = "EXPORTACION"
      .Cells(14, 8) = "IMPORTACION"
      .Cells(14, 9) = "CAPITAL DE TRABAJO"
      .Cells(14, 10) = "0-30 DIAS"
      .Cells(14, 11) = "31-90 DIAS"
      .Cells(14, 12) = "91-180 DIAS"
      .Cells(14, 13) = "181-270 DIAS"
      .Cells(14, 14) = "271-360 DIAS"
      .Cells(14, 15) = "MAS DE 360 DIAS"
      
      .Cells(14, 18) = "EXPORTACION"
      .Cells(14, 19) = "IMPORTACION"
      .Cells(14, 20) = "CAPITAL DE TRABAJO"
      .Cells(14, 21) = "0-30 DIAS"
      .Cells(14, 22) = "31-90 DIAS"
      .Cells(14, 23) = "91-180 DIAS"
      .Cells(14, 24) = "181-270 DIAS"
      .Cells(14, 25) = "271-360 DIAS"
      .Cells(14, 26) = "MAS DE 360 DIAS"
      
      .Cells(15, 1).RowHeight = 30
   
      .Range(.Cells(12, 1), .Cells(15, 28)).Font.Bold = True
      '.Range(.Cells(12, 1), .Cells(15, 28)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
      For r_int_Contad = 7 To 18 Step 11
         If r_int_Contad = 18 Then
            .Range(.Cells(12, r_int_Contad), .Cells(12, r_int_Contad + 9)).Merge
         Else
            .Range(.Cells(12, r_int_Contad), .Cells(12, r_int_Contad + 10)).Merge
         End If
         .Range(.Cells(13, r_int_Contad), .Cells(13, r_int_Contad + 2)).Merge
         .Range(.Cells(13, r_int_Contad + 3), .Cells(13, r_int_Contad + 8)).Merge
      Next
            
      For r_int_Contad = 1 To 28 Step 1
         If r_int_Contad < 7 Then
            .Range(.Cells(12, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 16 Then
            .Range(.Cells(14, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 18 Then
            .Range(.Cells(13, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 27 Then
            .Range(.Cells(14, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad = 27 Then
            .Range(.Cells(13, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad = 28 Then
            .Range(.Cells(12, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         End If
      Next
      
      For r_int_Contad = 1 To 28 Step 1
         .Range(.Cells(16, r_int_Contad), .Cells(21, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(16, r_int_Contad), .Cells(21, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 20 To 21 Step 1
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Borders(xlEdgeTop).LineStyle = xlContinuous
      Next
           
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 10
      
      .Range(.Cells(12, 1), .Cells(15, 28)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(12, 1), .Cells(15, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(12, 1), .Cells(15, 28)).VerticalAlignment = xlHAlignFill
              
      .Columns("A").ColumnWidth = 30
      '.Columns("A").NumberFormat = "@"
      
      .Columns("B").ColumnWidth = 15
      '.Columns("B").NumberFormat = "@"
            
      .Columns("C").ColumnWidth = 15
      '.Columns("C").NumberFormat = "@"
            
      .Columns("D").ColumnWidth = 15
      '.Columns("D").NumberFormat = "@"
            
      .Columns("E").ColumnWidth = 15
      '.Columns("E").NumberFormat = "@"
      
      .Columns("F").ColumnWidth = 15
      '.Columns("F").NumberFormat = "@"
      
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,##0.00"
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,##0.00"
      
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,##0.00"
      
      .Columns("N").ColumnWidth = 15
      .Columns("N").NumberFormat = "###,###,##0.00"
      
      .Columns("O").ColumnWidth = 15
      .Columns("O").NumberFormat = "###,###,##0.00"
      
      .Columns("P").ColumnWidth = 15
      .Columns("P").NumberFormat = "###,###,##0.00"
      
      .Columns("Q").ColumnWidth = 15
      .Columns("Q").NumberFormat = "###,###,##0.00"
      
      .Columns("R").ColumnWidth = 15
      .Columns("R").NumberFormat = "###,###,##0.00"
      
      .Columns("S").ColumnWidth = 15
      .Columns("S").NumberFormat = "###,###,##0.00"
      
      .Columns("T").ColumnWidth = 15
      .Columns("T").NumberFormat = "###,###,##0.00"
      
      .Columns("U").ColumnWidth = 15
      .Columns("U").NumberFormat = "###,###,##0.00"
      
      .Columns("V").ColumnWidth = 15
      .Columns("V").NumberFormat = "###,###,##0.00"
      
      .Columns("W").ColumnWidth = 15
      .Columns("W").NumberFormat = "###,###,##0.00"
      
      .Columns("X").ColumnWidth = 15
      .Columns("X").NumberFormat = "###,###,##0.00"
      
      .Columns("Y").ColumnWidth = 15
      .Columns("Y").NumberFormat = "###,###,##0.00"
      
      .Columns("Z").ColumnWidth = 15
      .Columns("Z").NumberFormat = "###,###,##0.00"
      
      .Columns("AA").ColumnWidth = 15
      .Columns("AA").NumberFormat = "###,###,##0.00"
      
      .Columns("AB").ColumnWidth = 15
      .Columns("AB").NumberFormat = "###,###,##0.00"
                
      .Range(.Cells(20, 2), .Cells(20, 5)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(21, 2), .Cells(21, 15)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(21, 17), .Cells(21, 28)).Interior.Color = RGB(0, 0, 0)
                
   End With
      
   r_obj_Excel.ActiveSheet.Cells(16, 1) = "BID-FOMIN"
   r_obj_Excel.ActiveSheet.Cells(16, 2) = "OTROS"
   r_obj_Excel.ActiveSheet.Cells(16, 3) = "ORG. INTERNAC."
   r_obj_Excel.ActiveSheet.Cells(16, 4) = "USA"
   r_obj_Excel.ActiveSheet.Cells(16, 5) = "D.A."
   
   r_int_ConAux = 0
   
   For r_int_Contad = 16 To 21 Step 1
      If r_int_Contad = 16 Or r_int_Contad = 20 Or r_int_Contad = 21 Then
         For r_int_ConGen = 6 To 28 Step 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad, r_int_ConGen) = r_dbl_Evalua(r_int_ConAux)
            r_int_ConAux = r_int_ConAux + 1
         Next
      End If
   Next
   
   r_obj_Excel.ActiveSheet.Cells(20, 1) = "TOTAL (EN US$)"
   r_obj_Excel.ActiveSheet.Cells(21, 1) = "TOTAL (EN NUEVOS SOLES) 4/ 6/"
   'r_obj_Excel.ActiveSheet.Cells(21, 16) = r_dbl_MonTot * r_dbl_TipCam
      
   
   'r_obj_Excel.ActiveSheet.Cells(24, 1) = "NOTAS"
   
   'r_obj_Excel.ActiveSheet.Cells(25, 1) = "1. La información correspondiente a las cuentas contables establecidas en el articulo 5º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   'r_obj_Excel.ActiveSheet.Cells(26, 1) = "2. El requerimiento de patrimonio efectivo por riesgo operacional será el promedio de los valores positivos del margen operacional bruto multiplicado por 15%, según lo indicado en el artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   'r_obj_Excel.ActiveSheet.Cells(28, 1) = "3. El APR por riesgo operacional se halla multiplicado el requerimiento de patrimonio efectivo por riesgo operacional por la inversa del limite global que establece la Ley General en el artículo 199º y la Vigésima Cuarta Disposición Transitoria y por los factores de ajsute que se consignan al final del artículo 3º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
         
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenDat()

   Erase r_str_Denomi

   r_str_Denomi(0) = "BID-FOMIN"
   r_str_Denomi(1) = "OTROS"
   r_str_Denomi(2) = "ORG. INTERNAC."
   r_str_Denomi(3) = "USA"
   r_str_Denomi(4) = "D.A."
   r_str_Denomi(5) = "TOTAL (EN US$)"
   r_str_Denomi(10) = "TOTAL (EN NUEVOS SOLES)"
   
   Erase r_dbl_Evalua
   r_dbl_MonTot = 0
           
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & r_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & r_str_PerAno & " "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_dbl_TipCam = g_rst_Princi!HIPCIE_TIPCAM
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   g_str_Parame = "SELECT * FROM CTB_MNTREP WHERE "
   g_str_Parame = g_str_Parame & "MNTREP_NUMFOR = 0114 AND "
   g_str_Parame = g_str_Parame & "MNTREP_NUMANX = 001 "
   g_str_Parame = g_str_Parame & "ORDER BY MNTREP_NUMCTA ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
        
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
                
      Do While Not g_rst_Princi.EOF
     
         g_str_Parame = "SELECT * FROM CNTBL_ASIENTO_DET "
         g_str_Parame = g_str_Parame & "WHERE CNTA_CTBL = '" & Trim(g_rst_Princi!MNTREP_NUMCTA) & "' AND "
         g_str_Parame = g_str_Parame & "MES = 1 AND "
         g_str_Parame = g_str_Parame & "ANO = 2010 "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
            Exit Sub
         End If
              
         If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
            g_rst_GenAux.MoveFirst
                
            Do While Not g_rst_GenAux.EOF
            
               If IsNull(g_rst_GenAux!IMP_MOVDOL) Then
                  r_dbl_MonTot = r_dbl_MonTot + 0
               Else
                  r_dbl_MonTot = r_dbl_MonTot + g_rst_GenAux!IMP_MOVDOL
               End If
               
               g_rst_GenAux.MoveNext
               DoEvents
                           
            Loop
            
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
            
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
   End If
               
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_Evalua(0) = r_dbl_MonTot
   r_dbl_Evalua(3) = r_dbl_MonTot
   r_dbl_Evalua(9) = r_dbl_MonTot
   r_dbl_Evalua(10) = r_dbl_MonTot
   r_dbl_Evalua(11) = 7.3
   r_dbl_Evalua(22) = r_dbl_MonTot
   r_dbl_Evalua(23) = r_dbl_MonTot
   r_dbl_Evalua(26) = r_dbl_MonTot
   r_dbl_Evalua(32) = r_dbl_MonTot
   r_dbl_Evalua(33) = r_dbl_MonTot
   r_dbl_Evalua(34) = 7.3
   r_dbl_Evalua(45) = r_dbl_MonTot
   r_dbl_Evalua(56) = r_dbl_MonTot * r_dbl_TipCam

End Sub

Private Sub fs_GeneDB()

   If (r_str_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_str_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM HIS_OBLEXT WHERE "
   g_str_Parame = g_str_Parame & "OBLEXT_PERMES = " & r_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "OBLEXT_PERANO = " & r_str_PerAno & ""
   'g_str_Parame = g_str_Parame & "OBLEXT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = 0
   r_int_ConTem = 0
   
   For r_int_Contad = 0 To 2 Step 1
   
      r_str_Cadena = "USP_HIS_OBLEXT ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_str_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_str_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_ConAux) & "', "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_ConAux + 1) & "', "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_ConAux + 2) & "', "
      'r_str_Cadena = r_str_Cadena & "'" & Left(r_str_Denomi(r_int_ConAux + 2), 1) & "', "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_ConAux + 3) & "', "
      'r_str_Cadena = r_str_Cadena & "'" & IIf(r_str_Denomi(r_int_ConAux + 3) = "USA", 4016, "") & "', "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_ConAux + 4) & "', "
      
      For r_int_ConGen = 0 To 22 Step 1
         If r_int_ConGen = 22 Then
            r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & " "
         Else
            r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & ", "
         End If
         r_int_ConTem = r_int_ConTem + 1
      Next
      
      r_str_Cadena = r_str_Cadena & ") "
      
      r_int_ConAux = r_int_ConAux + 5
      
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_OBLEXT.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   
   Next

End Sub


Private Sub fs_GenRpt()

   Dim r_obj_Excel      As Excel.Application
   
   Call fs_GenDat
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(3, 26), .Cells(5, 26)).HorizontalAlignment = xlHAlignRight
      .Cells(3, 26) = "Anexo Nº 14"
      .Cells(5, 26) = "CODIGO S.B.S.: 240"
      .Range(.Cells(3, 26), .Cells(3, 27)).Merge
      .Range(.Cells(5, 26), .Cells(5, 27)).Merge
            
      .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignLeft
      .Cells(3, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
      .Cells(5, 1) = "EMPRESA: MICASITA"
      .Range(.Cells(3, 1), .Cells(3, 2)).Merge
      .Range(.Cells(5, 1), .Cells(5, 2)).Merge
      
      .Range(.Cells(6, 1), .Cells(7, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(6, 1), .Cells(7, 28)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(7, 28)).Font.Underline = xlUnderlineStyleSingle
      .Cells(6, 14) = "Obligaciones con el Extranjero*"
      .Cells(7, 14) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_str_PerAno, "0,000")

               
      .Cells(12, 1) = "ENTIDAD 1/"
      .Cells(12, 2) = "CODIGO 2/"
      .Cells(12, 3) = "TIPO DE ENTIDAD 3/"
      .Cells(12, 4) = "PAIS DE ORIGEN"
      .Cells(12, 5) = "MONEDA"
      .Cells(12, 6) = "MONTO AUTORIZADO (EQUIVALENTE EN US$)"
      .Cells(12, 7) = "OBLIGACIONES DIRECTAS (EQUIVALENTE EN US$)"
      .Cells(12, 18) = "OBLIGACIONES CONTINGENTES (EQUIVALENTE EN US$)"
      .Cells(12, 28) = "TOTAL DE OBLIGACIONES CON EL EXTERIOR (A)+(B)(EQUIVALENTE EN US$)"
      
      .Cells(13, 7) = "POR DESTINO:"
      .Cells(13, 10) = "POR VENCER A:"
      .Cells(13, 16) = "TOTAL (A) 4/"
      .Cells(13, 17) = "TASA DE INTERES PROMEDIO 5/"
      .Cells(13, 18) = "POR DESTINO:"
      .Cells(13, 21) = "POR VENCER A:"
      .Cells(13, 27) = "TOTAL (B) 4/"
      
      .Cells(14, 7) = "EXPORTACION"
      .Cells(14, 8) = "IMPORTACION"
      .Cells(14, 9) = "CAPITAL DE TRABAJO"
      .Cells(14, 10) = "0-30 DIAS"
      .Cells(14, 11) = "31-90 DIAS"
      .Cells(14, 12) = "91-180 DIAS"
      .Cells(14, 13) = "181-270 DIAS"
      .Cells(14, 14) = "271-360 DIAS"
      .Cells(14, 15) = "MAS DE 360 DIAS"
      
      .Cells(14, 18) = "EXPORTACION"
      .Cells(14, 19) = "IMPORTACION"
      .Cells(14, 20) = "CAPITAL DE TRABAJO"
      .Cells(14, 21) = "0-30 DIAS"
      .Cells(14, 22) = "31-90 DIAS"
      .Cells(14, 23) = "91-180 DIAS"
      .Cells(14, 24) = "181-270 DIAS"
      .Cells(14, 25) = "271-360 DIAS"
      .Cells(14, 26) = "MAS DE 360 DIAS"
      
      .Cells(15, 1).RowHeight = 30
   
      .Range(.Cells(12, 1), .Cells(15, 28)).Font.Bold = True
      '.Range(.Cells(12, 1), .Cells(15, 28)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(15, 28)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
      For r_int_Contad = 7 To 18 Step 11
         If r_int_Contad = 18 Then
            .Range(.Cells(12, r_int_Contad), .Cells(12, r_int_Contad + 9)).Merge
         Else
            .Range(.Cells(12, r_int_Contad), .Cells(12, r_int_Contad + 10)).Merge
         End If
         .Range(.Cells(13, r_int_Contad), .Cells(13, r_int_Contad + 2)).Merge
         .Range(.Cells(13, r_int_Contad + 3), .Cells(13, r_int_Contad + 8)).Merge
      Next
            
      For r_int_Contad = 1 To 28 Step 1
         If r_int_Contad < 7 Then
            .Range(.Cells(12, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 16 Then
            .Range(.Cells(14, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 18 Then
            .Range(.Cells(13, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad < 27 Then
            .Range(.Cells(14, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad = 27 Then
            .Range(.Cells(13, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         ElseIf r_int_Contad = 28 Then
            .Range(.Cells(12, r_int_Contad), .Cells(15, r_int_Contad)).Merge
         End If
      Next
      
      For r_int_Contad = 1 To 28 Step 1
         .Range(.Cells(16, r_int_Contad), .Cells(21, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(16, r_int_Contad), .Cells(21, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 20 To 21 Step 1
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 28)).Borders(xlEdgeTop).LineStyle = xlContinuous
      Next
           
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      '.Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 10
      
      .Range(.Cells(12, 1), .Cells(15, 28)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(12, 1), .Cells(15, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(12, 1), .Cells(15, 28)).VerticalAlignment = xlHAlignFill
              
      .Columns("A").ColumnWidth = 30
      '.Columns("A").NumberFormat = "@"
      
      .Columns("B").ColumnWidth = 15
      '.Columns("B").NumberFormat = "@"
            
      .Columns("C").ColumnWidth = 15
      '.Columns("C").NumberFormat = "@"
            
      .Columns("D").ColumnWidth = 15
      '.Columns("D").NumberFormat = "@"
            
      .Columns("E").ColumnWidth = 15
      '.Columns("E").NumberFormat = "@"
      
      .Columns("F").ColumnWidth = 15
      '.Columns("F").NumberFormat = "@"
      
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,##0.00"
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,##0.00"
      
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,##0.00"
      
      .Columns("N").ColumnWidth = 15
      .Columns("N").NumberFormat = "###,###,##0.00"
      
      .Columns("O").ColumnWidth = 15
      .Columns("O").NumberFormat = "###,###,##0.00"
      
      .Columns("P").ColumnWidth = 15
      .Columns("P").NumberFormat = "###,###,##0.00"
      
      .Columns("Q").ColumnWidth = 15
      .Columns("Q").NumberFormat = "###,###,##0.00"
      
      .Columns("R").ColumnWidth = 15
      .Columns("R").NumberFormat = "###,###,##0.00"
      
      .Columns("S").ColumnWidth = 15
      .Columns("S").NumberFormat = "###,###,##0.00"
      
      .Columns("T").ColumnWidth = 15
      .Columns("T").NumberFormat = "###,###,##0.00"
      
      .Columns("U").ColumnWidth = 15
      .Columns("U").NumberFormat = "###,###,##0.00"
      
      .Columns("V").ColumnWidth = 15
      .Columns("V").NumberFormat = "###,###,##0.00"
      
      .Columns("W").ColumnWidth = 15
      .Columns("W").NumberFormat = "###,###,##0.00"
      
      .Columns("X").ColumnWidth = 15
      .Columns("X").NumberFormat = "###,###,##0.00"
      
      .Columns("Y").ColumnWidth = 15
      .Columns("Y").NumberFormat = "###,###,##0.00"
      
      .Columns("Z").ColumnWidth = 15
      .Columns("Z").NumberFormat = "###,###,##0.00"
      
      .Columns("AA").ColumnWidth = 15
      .Columns("AA").NumberFormat = "###,###,##0.00"
      
      .Columns("AB").ColumnWidth = 15
      .Columns("AB").NumberFormat = "###,###,##0.00"
                
      .Range(.Cells(20, 2), .Cells(20, 5)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(21, 2), .Cells(21, 15)).Interior.Color = RGB(0, 0, 0)
      .Range(.Cells(21, 17), .Cells(21, 28)).Interior.Color = RGB(0, 0, 0)
                
   End With
      
   r_obj_Excel.ActiveSheet.Cells(16, 1) = "BID-FOMIN"
   r_obj_Excel.ActiveSheet.Cells(16, 2) = "OTROS"
   r_obj_Excel.ActiveSheet.Cells(16, 3) = "ORG. INTERNAC."
   r_obj_Excel.ActiveSheet.Cells(16, 4) = "USA"
   r_obj_Excel.ActiveSheet.Cells(16, 5) = "D.A."
   
   r_int_ConAux = 0
   
   For r_int_Contad = 16 To 21 Step 1
      If r_int_Contad = 16 Or r_int_Contad = 20 Or r_int_Contad = 21 Then
         For r_int_ConGen = 6 To 28 Step 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad, r_int_ConGen) = r_dbl_Evalua(r_int_ConAux)
            r_int_ConAux = r_int_ConAux + 1
         Next
      End If
   Next
   
   r_obj_Excel.ActiveSheet.Cells(20, 1) = "TOTAL (EN US$)"
   r_obj_Excel.ActiveSheet.Cells(21, 1) = "TOTAL (EN NUEVOS SOLES) 4/ 6/"
   'r_obj_Excel.ActiveSheet.Cells(21, 16) = r_dbl_MonTot * r_dbl_TipCam
      
   
   'r_obj_Excel.ActiveSheet.Cells(24, 1) = "NOTAS"
   
   'r_obj_Excel.ActiveSheet.Cells(25, 1) = "1. La información correspondiente a las cuentas contables establecidas en el articulo 5º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   'r_obj_Excel.ActiveSheet.Cells(26, 1) = "2. El requerimiento de patrimonio efectivo por riesgo operacional será el promedio de los valores positivos del margen operacional bruto multiplicado por 15%, según lo indicado en el artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   'r_obj_Excel.ActiveSheet.Cells(28, 1) = "3. El APR por riesgo operacional se halla multiplicado el requerimiento de patrimonio efectivo por riesgo operacional por la inversa del limite global que establece la Ley General en el artículo 199º y la Vigésima Cuarta Disposición Transitoria y por los factores de ajsute que se consignan al final del artículo 3º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
      
   Call fs_GeneDB
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
      
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

End Sub



Private Sub fs_GenDat_DB()

   Erase r_str_Denomi

   'r_str_Denomi(0) = "BID-FOMIN"
   'r_str_Denomi(1) = "OTROS"
   'r_str_Denomi(2) = "ORG. INTERNAC."
   'r_str_Denomi(3) = "USA"
   'r_str_Denomi(4) = "D.A."
   'r_str_Denomi(5) = "TOTAL (EN US$)"
   'r_str_Denomi(10) = "TOTAL (EN NUEVOS SOLES)"
   
   Erase r_dbl_Evalua
   'r_dbl_MonTot = 0
           
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   
   g_str_Parame = "SELECT * FROM HIS_OBLEXT WHERE "
   g_str_Parame = g_str_Parame & "OBLEXT_PERMES = " & r_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "OBLEXT_PERANO = " & r_str_PerAno & ""
   g_str_Parame = g_str_Parame & "ORDER BY OBLEXT_NUMITE ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = -1
   r_int_ConTem = -1
        
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
                
      Do While Not g_rst_Princi.EOF
      
         r_str_Denomi(r_int_ConTem + 1) = IIf(IsNull(Trim(g_rst_Princi!OBLEXT_DESCRI)) = True, "", Trim(g_rst_Princi!OBLEXT_DESCRI))
         r_str_Denomi(r_int_ConTem + 2) = IIf(IsNull(Trim(g_rst_Princi!OBLEXT_CODIGO)) = True, "", Trim(g_rst_Princi!OBLEXT_CODIGO))
         r_str_Denomi(r_int_ConTem + 3) = IIf(IsNull(Trim(g_rst_Princi!OBLEXT_TIPENT)) = True, "", Trim(g_rst_Princi!OBLEXT_TIPENT))
         r_str_Denomi(r_int_ConTem + 4) = IIf(IsNull(Trim(g_rst_Princi!OBLEXT_PAIORG)) = True, "", Trim(g_rst_Princi!OBLEXT_PAIORG))
         r_str_Denomi(r_int_ConTem + 5) = IIf(IsNull(Trim(g_rst_Princi!OBLEXT_MONEDA)) = True, "", Trim(g_rst_Princi!OBLEXT_MONEDA))
         
         r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!OBLEXT_MOAUEQ
         r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!OBLEXT_ODDEEX
         r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!OBLEXT_ODDEIM
         r_dbl_Evalua(r_int_ConAux + 4) = g_rst_Princi!OBLEXT_ODDECA
         r_dbl_Evalua(r_int_ConAux + 5) = g_rst_Princi!OBLEXT_ODV000
         r_dbl_Evalua(r_int_ConAux + 6) = g_rst_Princi!OBLEXT_ODV031
         r_dbl_Evalua(r_int_ConAux + 7) = g_rst_Princi!OBLEXT_ODV091
         r_dbl_Evalua(r_int_ConAux + 8) = g_rst_Princi!OBLEXT_ODV181
         r_dbl_Evalua(r_int_ConAux + 9) = g_rst_Princi!OBLEXT_ODV271
         r_dbl_Evalua(r_int_ConAux + 10) = g_rst_Princi!OBLEXT_ODV360
         r_dbl_Evalua(r_int_ConAux + 11) = g_rst_Princi!OBLEXT_ODTOTA
         r_dbl_Evalua(r_int_ConAux + 12) = g_rst_Princi!OBLEXT_ODTAIN
         r_dbl_Evalua(r_int_ConAux + 13) = g_rst_Princi!OBLEXT_OCDEEX
         r_dbl_Evalua(r_int_ConAux + 14) = g_rst_Princi!OBLEXT_OCDEIM
         r_dbl_Evalua(r_int_ConAux + 15) = g_rst_Princi!OBLEXT_OCDECA
         r_dbl_Evalua(r_int_ConAux + 16) = g_rst_Princi!OBLEXT_OCV000
         r_dbl_Evalua(r_int_ConAux + 17) = g_rst_Princi!OBLEXT_OCV031
         r_dbl_Evalua(r_int_ConAux + 18) = g_rst_Princi!OBLEXT_OCV091
         r_dbl_Evalua(r_int_ConAux + 19) = g_rst_Princi!OBLEXT_OCV181
         r_dbl_Evalua(r_int_ConAux + 20) = g_rst_Princi!OBLEXT_OCV271
         r_dbl_Evalua(r_int_ConAux + 21) = g_rst_Princi!OBLEXT_OCV360
         r_dbl_Evalua(r_int_ConAux + 22) = g_rst_Princi!OBLEXT_OCTOTA
         r_dbl_Evalua(r_int_ConAux + 23) = g_rst_Princi!OBLEXT_OCTOOB
                  
         r_int_ConAux = r_int_ConAux + 23
         r_int_ConTem = r_int_ConTem + 5
                  
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
   End If
               
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'r_dbl_Evalua(0) = r_dbl_MonTot
   'r_dbl_Evalua(3) = r_dbl_MonTot
   'r_dbl_Evalua(9) = r_dbl_MonTot
   'r_dbl_Evalua(10) = r_dbl_MonTot
   'r_dbl_Evalua(11) = 7.3
   'r_dbl_Evalua(22) = r_dbl_MonTot
   'r_dbl_Evalua(23) = r_dbl_MonTot
   'r_dbl_Evalua(26) = r_dbl_MonTot
   'r_dbl_Evalua(32) = r_dbl_MonTot
   'r_dbl_Evalua(33) = r_dbl_MonTot
   'r_dbl_Evalua(34) = 7.3
   'r_dbl_Evalua(45) = r_dbl_MonTot
   'r_dbl_Evalua(56) = r_dbl_MonTot * r_dbl_TipCam

End Sub



