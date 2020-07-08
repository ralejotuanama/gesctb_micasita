VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_12 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   2385
   ClientTop       =   5145
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   7155
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
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
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Anexo N° 7"
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
            Width           =   6285
            _Version        =   65536
            _ExtentX        =   11086
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0107-01 Medición del Riesgo de Tasa de Interés en Moneda Nacional"
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
            Picture         =   "GesCtb_frm_713.frx":0000
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   750
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
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_713.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_713.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_713.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "GesCtb_frm_713.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_713.frx":11A2
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
         Top             =   1440
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
Attribute VB_Name = "frm_RepSbs_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim r_dbl_Evalua(1600)  As Double
   Dim r_str_Denomi(100)   As String
   Dim r_str_Fechas(26)    As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_TemAux        As Integer
   Dim r_int_AuxTem        As Integer
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String
   Dim r_str_Cadena        As String
   Dim r_dbl_TipCam        As Double
   Dim r_dbl_MulUso        As Double
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

Private Sub cmd_Proces_Click()

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
   
   Screen.MousePointer = 11
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   Call fs_GenDat
   Call fs_GeneDB
   
   MsgBox "Proceso Terminado", vbExclamation, modgen_g_str_NomPlt
   
   Screen.MousePointer = 0
   
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
   Screen.MousePointer = 0
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_MERITA", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_MERITA", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("¿Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_Genera_Arc001
         Call fs_Genera_Arc002
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_Genera_Arc001
  Call fs_Genera_Arc002
  
End Sub

Private Sub fs_Genera_Arc001()

   Dim r_int_NumRes     As Integer
   Dim r_int_CodEmp     As Integer
     
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   Dim r_dbl_MulUso     As Double
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat_Exc
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
      
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\01" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".107"
   
   Screen.MousePointer = 11
   
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
   
   Print #r_int_NumRes, Format(107, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_int_ConAux = 0
      
   For r_int_Contad = 100 To 3800 Step 100
      r_str_Cadena = ""
      
      If r_int_Contad <> 100 And r_int_Contad <> 1500 Then
         For r_int_ConTem = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                         
      Else
         For r_int_ConTem = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                     
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "0000") & r_str_Cadena
            
   Next
          
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo 1 creado.", vbInformation, modgen_g_str_NomPlt
  
End Sub


Private Sub fs_Genera_Arc002()

   Dim r_int_NumRes     As Integer
   Dim r_int_CodEmp     As Integer
     
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   Dim r_dbl_MulUso     As Double
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat_Exc
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
      
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\02" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".107"
   
   Screen.MousePointer = 11
   
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
   
   Print #r_int_NumRes, Format(107, "0000") & Format(2, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_int_ConAux = 800
   
   
   For r_int_Contad = 100 To 3800 Step 100
      r_str_Cadena = ""
      
      If r_int_Contad <> 100 And r_int_Contad <> 1500 Then
         For r_int_ConTem = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                         
      Else
         For r_int_ConTem = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                     
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "0000") & r_str_Cadena
            
   Next
          
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo 2 creado.", vbInformation, modgen_g_str_NomPlt

  
End Sub


Private Sub fs_GenExc()

   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_lng_TotReg     As Long
   Dim r_lng_RegAct     As Long
   Dim r_int_ValNeg     As Integer
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   Dim r_str_FecRpt     As String
      
   Screen.MousePointer = 11
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat_Exc
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If

   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   For r_int_TemAux = 1 To 2 Step 1
   
      Select Case r_int_TemAux
         Case 1:
            r_obj_Excel.Sheets(r_int_TemAux).Name = "01"
         Case 2:
            r_obj_Excel.Sheets(r_int_TemAux).Name = "02"
      End Select
      
      With r_obj_Excel.Sheets(r_int_TemAux)
      
         '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
         '.DrawingObjects(1).Left = 20
         '.DrawingObjects(1).Top = 20
         
         .Range(.Cells(3, 16), .Cells(5, 16)).HorizontalAlignment = xlHAlignRight
         .Cells(3, 16) = "Anexo Nº 7"
         .Cells(5, 16) = "CODIGO S.B.S.: 240"
         .Range(.Cells(3, 16), .Cells(3, 17)).Merge
         .Range(.Cells(5, 16), .Cells(5, 17)).Merge
               
         .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignLeft
         .Cells(3, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
         .Cells(5, 1) = "EMPRESA: MI CASITA"
         .Range(.Cells(3, 1), .Cells(3, 2)).Merge
         .Range(.Cells(5, 1), .Cells(5, 2)).Merge
   
          
         .Range(.Cells(6, 8), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(6, 8), .Cells(7, 1)).Font.Bold = True
          
         .Range(.Cells(6, 1), .Cells(6, 14)).Merge
         .Range(.Cells(7, 1), .Cells(7, 14)).Merge
         .Range(.Cells(8, 1), .Cells(8, 14)).Merge
         
         '.Range(.Cells(6, 8), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
         If r_int_TemAux = 1 Then
            .Cells(6, 1) = "A. MEDICIÓN DEL RIESGO DE TASA DE INTERES EN MONEDA NACIONAL"
            .Cells(8, 1) = "(Expresado en Nuevos Soles)"
         Else
            .Cells(6, 1) = "B. MEDICIÓN DEL RIESGO DE TASA DE INTERES EN MONEDA EXTRANJERA"
            .Cells(8, 1) = "(Expresado en Dólares Americanos)"
         End If
         
         .Cells(7, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
         '.Cells(8, 1) = "(En Nuevos Soles y de Dólares Americanos)"
                              
         .Cells(13, 2) = "MONEDA: NUEVOS SOLES"
         .Cells(15, 1) = "CUENTAS "
         .Cells(15, 2) = "DENOMINACION"
         .Cells(13, 3) = "1"
         .Cells(15, 3) = "1-7d"
         .Cells(13, 4) = "2"
         .Cells(15, 4) = "8-15d"
         .Cells(13, 5) = "3"
         .Cells(15, 5) = "16d-1M"
         .Cells(13, 6) = "4"
         .Cells(15, 6) = "2M"
         .Cells(13, 7) = "5"
         .Cells(15, 7) = "3M"
         .Cells(13, 8) = "6"
         .Cells(15, 8) = "4-6M"
         .Cells(13, 9) = "7"
         .Cells(15, 9) = "7M-1A"
         .Cells(13, 10) = "8"
         .Cells(15, 10) = "2A"
         .Cells(13, 11) = "9"
         .Cells(15, 11) = "3A"
         .Cells(13, 12) = "10"
         .Cells(15, 12) = "4A"
         .Cells(13, 13) = "11"
         .Cells(15, 13) = "5A"
         .Cells(13, 14) = "12"
         .Cells(15, 14) = "6-10A"
         .Cells(13, 15) = "13"
         .Cells(15, 15) = "11-20A"
         .Cells(13, 16) = "14"
         .Cells(15, 16) = "+20A"
         .Cells(13, 17) = "0"
         .Cells(15, 17) = "No Recuperable (1)"
         .Cells(13, 18) = "TOTAL"
         .Cells(15, 18) = ".(2)."
                        
         .Cells(16, 2) = "Activos "
         .Cells(17, 1) = "1100"
         .Cells(17, 2) = "Disponible "
         .Cells(18, 1) = "1200"
         .Cells(18, 2) = "Fondos Interbancarios"
         .Cells(19, 1) = "1300"
         .Cells(19, 2) = "Inversiones Negociables y a Vencimiento "
         
         .Cells(21, 1) = "1400"
         .Cells(21, 2) = "Créditos"
         .Cells(22, 1) = "1401"
         .Cells(22, 2) = "Vigentes (*)"
         .Cells(23, 1) = "1403+1404+1405+1406"
         .Cells(23, 2) = "Otras (*)"
         .Cells(24, 1) = "1500"
         .Cells(24, 2) = "Cuentas por Cobrar"
         .Cells(25, 1) = "1700"
         .Cells(25, 2) = "Inversiones Permanentes"
         
         .Cells(27, 1) = "8409.04.01+8409.04.02+8409.04.09"
         .Cells(27, 2) = "Pos. Activas en Derivados de Tasa de Int."
         .Cells(28, 1) = "7106.01+7106.02"
         .Cells(28, 2) = "Pos. Activas en Derivados de M.E."
         .Cells(29, 1) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
         .Cells(29, 2) = "Pos. Activas en Otros Derivados Sensibles"
         .Cells(30, 1) = "1600+1800+1900"
         .Cells(30, 2) = "Otras Cuentas Activas"
         
         .Cells(31, 2) = "Total (I)(3)"
         .Cells(32, 2) = "Pasivos"
         .Cells(33, 1) = "2100"
         .Cells(33, 2) = "Obligaciones con el Público"
         .Cells(34, 1) = "2101"
         .Cells(34, 2) = "Obligaciones a la Vista (*)"
         .Cells(35, 1) = "2102"
         .Cells(35, 2) = "Obligaciones por Cuentas de Ahorro (*)"
         .Cells(36, 1) = "2103"
         .Cells(36, 2) = "Obligaciones por Cuentas a Plazos (*)"
         
         .Cells(38, 1) = "2200"
         .Cells(38, 2) = "Fondos Interbancarios"
         .Cells(39, 1) = "2300"
         .Cells(39, 2) = "Depósitos de Empresas del Sistema Financiero y O.I."
         .Cells(40, 1) = "2400+2600"
         .Cells(40, 2) = "Adeudados y Otras Obligaciones Financieras "
         .Cells(35, 1) = "2500"
         .Cells(35, 2) = "Cuentas por Pagar"
         .Cells(36, 1) = "2800"
         .Cells(36, 2) = "Valores, Titulos y Obligaciones en Circulación "
         
         .Cells(38, 1) = "8409.04.01+8409.04.02+8409.04.09"
         .Cells(38, 2) = "Pos. Activas en Derivados de Tasa de Int."
         .Cells(39, 1) = "7106.01+7106.02"
         .Cells(39, 2) = "Pos. Activas en Derivados de M.E."
         .Cells(40, 1) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
         .Cells(40, 2) = "Pos. Activas en Otros Derivados Sensibles"
         .Cells(41, 1) = "2700+2900"
         .Cells(41, 2) = "Otras Cuentas Pasivas"
         
         .Cells(42, 2) = "Total (II)(3)"
         .Cells(43, 2) = "Monto Delta Neto de Opciones (III)(4)"
         .Cells(44, 2) = "Descalce Marginal en MN (I-II+III)"
         .Cells(45, 2) = "Descalce Marginal/Patrimonio Efectivo (5)"
         .Cells(46, 2) = "Descalce Acumulado en MN"
         .Cells(47, 2) = "Acumulado/Patrimonio Efectivo (5)"
         .Cells(48, 2) = "Descalce Acumulado VAC (6)"
         .Cells(49, 2) = "Total VAC/Patrimonio Efectivo (5)"
         .Cells(50, 2) = "Descalce Acumulado Tasa (7)"
         .Cells(51, 2) = "Tasa/Patrimonio Efectivo (5)"
         
         .Cells(14, 1).RowHeight = 20
         .Cells(15, 1).RowHeight = 20
               
         .Columns("A").ColumnWidth = 30
         .Columns("A").NumberFormat = "@"
               
         .Columns("B").ColumnWidth = 34
         '.Columns("B").NumberFormat = "@"
         '.Columns("B").HorizontalAlignment = xlHAlignCenter
         
         .Range("C:R").NumberFormat = "###,###,##0.00"
         .Range("C:R").ColumnWidth = 12
               
         '.Range(.Cells(13, 1), .Cells(15, 1)).Merge
         '.Range(.Cells(13, 1), .Cells(15, 1)).WrapText = True
         .Range(.Cells(13, 1), .Cells(15, 1)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 1), .Cells(15, 1)).HorizontalAlignment = xlHAlignCenter
         
               
         '.Range(.Cells(13, 2), .Cells(15, 2)).Merge
         '.Range(.Cells(13, 2), .Cells(15, 2)).WrapText = True
         .Range(.Cells(13, 2), .Cells(15, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 2), .Cells(15, 2)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 3), .Cells(15, 3)).Merge
         .Range(.Cells(14, 3), .Cells(15, 3)).WrapText = True
         .Range(.Cells(14, 3), .Cells(15, 3)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 3), .Cells(15, 3)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 4), .Cells(15, 4)).Merge
         .Range(.Cells(14, 4), .Cells(15, 4)).WrapText = True
         .Range(.Cells(14, 4), .Cells(15, 4)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 4), .Cells(15, 4)).HorizontalAlignment = xlHAlignCenter
               
         .Range(.Cells(14, 5), .Cells(15, 5)).Merge
         .Range(.Cells(14, 5), .Cells(15, 5)).WrapText = True
         .Range(.Cells(14, 5), .Cells(15, 5)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 5), .Cells(15, 5)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 6), .Cells(15, 6)).Merge
         .Range(.Cells(14, 6), .Cells(15, 6)).WrapText = True
         .Range(.Cells(14, 6), .Cells(15, 6)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 6), .Cells(15, 6)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 7), .Cells(15, 7)).Merge
         .Range(.Cells(14, 7), .Cells(15, 7)).WrapText = True
         .Range(.Cells(14, 7), .Cells(15, 7)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 7), .Cells(15, 7)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 8), .Cells(15, 8)).Merge
         .Range(.Cells(14, 8), .Cells(15, 8)).WrapText = True
         .Range(.Cells(14, 8), .Cells(15, 8)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 8), .Cells(15, 8)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 9), .Cells(15, 9)).Merge
         .Range(.Cells(14, 9), .Cells(15, 9)).WrapText = True
         .Range(.Cells(14, 9), .Cells(15, 9)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 9), .Cells(15, 9)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 10), .Cells(15, 10)).Merge
         .Range(.Cells(14, 10), .Cells(15, 10)).WrapText = True
         .Range(.Cells(14, 10), .Cells(15, 10)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 10), .Cells(15, 10)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 11), .Cells(15, 11)).Merge
         .Range(.Cells(14, 11), .Cells(15, 11)).WrapText = True
         .Range(.Cells(14, 11), .Cells(15, 11)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 11), .Cells(15, 11)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 12), .Cells(15, 12)).Merge
         .Range(.Cells(14, 12), .Cells(15, 12)).WrapText = True
         .Range(.Cells(14, 12), .Cells(15, 12)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 12), .Cells(15, 12)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 13), .Cells(15, 13)).Merge
         .Range(.Cells(14, 13), .Cells(15, 13)).WrapText = True
         .Range(.Cells(14, 13), .Cells(15, 13)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 13), .Cells(15, 13)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 14), .Cells(15, 14)).Merge
         .Range(.Cells(14, 14), .Cells(15, 14)).WrapText = True
         .Range(.Cells(14, 14), .Cells(15, 14)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 14), .Cells(15, 14)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 15), .Cells(15, 15)).Merge
         .Range(.Cells(14, 15), .Cells(15, 15)).WrapText = True
         .Range(.Cells(14, 15), .Cells(15, 15)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 15), .Cells(15, 15)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 16), .Cells(15, 16)).Merge
         .Range(.Cells(14, 16), .Cells(15, 16)).WrapText = True
         .Range(.Cells(14, 16), .Cells(15, 16)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 16), .Cells(15, 16)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 17), .Cells(15, 17)).Merge
         .Range(.Cells(14, 17), .Cells(15, 17)).WrapText = True
         .Range(.Cells(14, 17), .Cells(15, 17)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 17), .Cells(15, 17)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 18), .Cells(15, 18)).Merge
         .Range(.Cells(14, 18), .Cells(15, 18)).WrapText = True
         .Range(.Cells(14, 18), .Cells(15, 18)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 18), .Cells(15, 18)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(13, 3), .Cells(15, 18)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         .Range(.Cells(13, 3), .Cells(13, 18)).NumberFormat = "###"
         '.Range(.Cells(14, 3), .Cells(14, 18)).NumberFormat = "@"
                                
         .Range(.Cells(16, 2), .Cells(16, 2)).Font.Bold = True
         .Range(.Cells(16, 2), .Cells(16, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(16, 2), .Cells(16, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(16, 2), .Cells(16, 2)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(31, 1), .Cells(31, 18)).Font.Bold = True
         .Range(.Cells(31, 1), .Cells(31, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 18)).VerticalAlignment = xlVAlignCenter
         
         .Range(.Cells(32, 2), .Cells(32, 2)).Font.Bold = True
         .Range(.Cells(32, 2), .Cells(32, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(32, 2), .Cells(32, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(32, 2), .Cells(32, 2)).HorizontalAlignment = xlHAlignCenter
                    
         Dim r_int_Contad As Integer
         
         For r_int_Contad = 1 To 18 Step 1
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         For r_int_Contad = 42 To 51 Step 1
            If r_int_Contad <> 50 And r_int_Contad <> 48 And r_int_Contad <> 46 And r_int_Contad <> 44 Then
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If
         Next
         
         .Range(.Cells(13, 1), .Cells(15, 18)).Font.Bold = True
         .Range(.Cells(13, 1), .Cells(15, 18)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 1), .Cells(15, 18)).HorizontalAlignment = xlHAlignCenter
         '.Range(.Cells(13, 1), .Cells(15, 16)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         
         For r_int_Contad = 29 To 40 Step 11
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).RowHeight = 30
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).WrapText = True
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).VerticalAlignment = xlVAlignCenter
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).HorizontalAlignment = xlHAlignLeft
         Next
         
         For r_int_Contad = 44 To 52 Step 1
            If r_int_Contad <> 49 Then
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 2)).Merge
            End If
         Next
         
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 8
      
         If r_int_TemAux = 1 Then
            r_int_ConTem = 0
         Else
            r_int_ConTem = 800
         End If
         
         For r_int_ConAux = 17 To 51 Step 1
            If r_int_ConAux <> 20 And r_int_ConAux <> 32 And r_int_ConAux <> 37 Then
               For r_int_Contad = 3 To 18 Step 1
                  .Cells(r_int_ConAux, r_int_Contad) = r_dbl_Evalua(r_int_ConTem)
                  r_int_ConTem = r_int_ConTem + 1
               Next
            End If
         Next
      
      End With
   
   Next
   
   Screen.MousePointer = 0
  
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Function GenFec(ByVal p_Fecha As String) As String
   
   GenFec = Right(p_Fecha, 4) & Mid(p_Fecha, 4, 2) & Left(p_Fecha, 2)

End Function

Private Sub fs_GenDat()

   Erase r_str_Fechas

   Dim r_str_FecRpt        As String
   Dim l_int_mes           As Integer
   Dim l_int_ano           As Integer
   Dim l_str_fec           As String
   Dim l_str_aux           As String
   Dim l_dat_fec           As Date
   Dim l_dat_aux           As Date
   Dim l_int_con           As Integer
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   l_int_mes = cmb_PerMes.ListIndex + 1
   l_int_ano = CInt(ipp_PerAno.Text)
   l_str_fec = "01/" & l_int_mes & "/" & l_int_ano
   l_dat_fec = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   l_str_aux = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, GenFec(l_dat_fec), 2)
   
     r_str_Fechas(0) = CDate(l_dat_fec + 1)
     r_str_Fechas(1) = CDate(l_dat_fec + 7)
     r_str_Fechas(2) = CDate(l_dat_fec + 8)
     r_str_Fechas(3) = CDate(l_dat_fec + 15)
     r_str_Fechas(4) = CDate(l_dat_fec + 16)
     l_str_fec = l_dat_fec + 1
     r_str_Fechas(5) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4))))
     r_str_Fechas(6) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
     r_str_Fechas(7) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
     r_str_Fechas(8) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     r_str_Fechas(9) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
     r_str_Fechas(10) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     r_str_Fechas(11) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
     r_str_Fechas(12) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
     r_str_Fechas(13) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
     r_str_Fechas(14) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     l_dat_fec = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
        r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
     Else
        r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     End If
           
     If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
        r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     Else
        r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
     End If
     l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     r_str_Fechas(17) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
                 
     r_str_Fechas(18) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     r_str_Fechas(19) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
           
     r_str_Fechas(20) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     r_str_Fechas(21) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
     r_str_Fechas(22) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     
     For l_int_con = 1 To 5
        l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     Next l_int_con
     
     If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
        r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
     Else
        r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     End If
           
     If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
        r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     Else
        r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
     End If
     
     For l_int_con = 1 To 10
        l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
     Next l_int_con
     
     r_str_Fechas(25) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
           
     r_str_Fechas(26) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
          
         
   '  For l_int_con = 1 To 20
   '     l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   '  Next l_int_con
     
   '  .Cells(3, 15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   
   
   '**********************************************************************************************************************************************************
   
   
   g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & GenFec(l_dat_fec) & " "
   g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
      
         If Trim(g_rst_Princi!HIPMAE_CODPRD) = "001" Then
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "002" Then
            r_dbl_Evalua(15) = r_dbl_Evalua(15) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "003" Then
            r_dbl_Evalua(30) = r_dbl_Evalua(30) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "004" Then
            r_dbl_Evalua(45) = r_dbl_Evalua(45) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "006" Then
            r_dbl_Evalua(60) = r_dbl_Evalua(60) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "007" Then
            r_dbl_Evalua(75) = r_dbl_Evalua(75) + g_rst_Princi!CAPITAL
            
         End If

         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
         
   
   
   For r_int_Contad = 0 To 26 Step 2
   
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
            
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
         g_rst_Princi.MoveFirst
               
         Do While Not g_rst_Princi.EOF
         
            If Trim(g_rst_Princi!HIPMAE_CODPRD) = "001" Then
               r_dbl_Evalua((r_int_Contad / 2) + 1) = r_dbl_Evalua((r_int_Contad / 2) + 1) + g_rst_Princi!CAPITAL
              
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "002" Then
               r_dbl_Evalua((r_int_Contad / 2) + 16) = r_dbl_Evalua((r_int_Contad / 2) + 16) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "003" Then
               r_dbl_Evalua((r_int_Contad / 2) + 31) = r_dbl_Evalua((r_int_Contad / 2) + 31) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 46) = r_dbl_Evalua((r_int_Contad / 2) + 46) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "006" Then
               r_dbl_Evalua((r_int_Contad / 2) + 61) = r_dbl_Evalua((r_int_Contad / 2) + 61) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 76) = r_dbl_Evalua((r_int_Contad / 2) + 76) + g_rst_Princi!CAPITAL
               
            End If
   
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
   Next
   
   
   g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & GenFec(l_dat_fec) & " AND "
   g_str_Parame = g_str_Parame & "(HIPMAE_NUMOPE <> '0040700001' AND HIPMAE_NUMOPE <> '0040700002' AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE <> '0040700003' AND HIPMAE_NUMOPE <> '0040700004') "
   g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "

   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
      
         If Trim(g_rst_Princi!HIPMAE_CODPRD) = "001" Then
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "002" Then
            r_dbl_Evalua(15) = r_dbl_Evalua(15) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "003" Then
            r_dbl_Evalua(30) = r_dbl_Evalua(30) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "004" Then
            r_dbl_Evalua(45) = r_dbl_Evalua(45) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "006" Then
            'r_dbl_Evalua(60) = r_dbl_Evalua(60) + g_rst_Princi!CAPITAL
            
         ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "007" Then
            r_dbl_Evalua(75) = r_dbl_Evalua(75) + g_rst_Princi!CAPITAL
            
         End If

         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   
   For r_int_Contad = 0 To 26 Step 2
   
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
      
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " AND "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & GenFec(r_str_Fechas(r_int_Contad + 1)) & " AND "
      End If
      
      g_str_Parame = g_str_Parame & "(HIPMAE_NUMOPE <> '0040700001' AND HIPMAE_NUMOPE <> '0040700002' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE <> '0040700003' AND HIPMAE_NUMOPE <> '0040700004') "
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
         g_rst_Princi.MoveFirst
               
         Do While Not g_rst_Princi.EOF
         
            If Trim(g_rst_Princi!HIPMAE_CODPRD) = "001" Then
               r_dbl_Evalua((r_int_Contad / 2) + 1) = r_dbl_Evalua((r_int_Contad / 2) + 1) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "002" Then
               r_dbl_Evalua((r_int_Contad / 2) + 16) = r_dbl_Evalua((r_int_Contad / 2) + 16) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "003" Then
               r_dbl_Evalua((r_int_Contad / 2) + 31) = r_dbl_Evalua((r_int_Contad / 2) + 31) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 46) = r_dbl_Evalua((r_int_Contad / 2) + 46) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "006" Then
               'r_dbl_Evalua((r_int_Contad / 2) + 61) = r_dbl_Evalua((r_int_Contad / 2) + 61) + g_rst_Princi!CAPITAL
               
            ElseIf Trim(g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 76) = r_dbl_Evalua((r_int_Contad / 2) + 76) + g_rst_Princi!CAPITAL
               
            End If
            
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
   Next
      
   
   For r_int_Contad = 0 To 26 Step 2
   
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 AND "
      
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " AND "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & GenFec(r_str_Fechas(r_int_Contad + 1)) & " AND "
      End If
      
      g_str_Parame = g_str_Parame & "(hipmae_numope = '0040700001' or hipmae_numope = '0040700002' or "
      g_str_Parame = g_str_Parame & "hipmae_numope = '0040700003' or hipmae_numope = '0040700004') "
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
         g_rst_Princi.MoveFirst
               
         Do While Not g_rst_Princi.EOF
                    
            If Trim(g_rst_Princi!HIPMAE_CODPRD) = "006" Then
               r_dbl_Evalua((r_int_Contad / 2) + 61) = r_dbl_Evalua((r_int_Contad / 2) + 61) + g_rst_Princi!CAPITAL
   
            End If
            
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
   Next
   

End Sub

Private Sub fs_GeneDB()

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   For r_int_TemAux = 1 To 2 Step 1

      g_str_Parame = "DELETE FROM HIS_MERITA WHERE "
      g_str_Parame = g_str_Parame & "MERITA_PERMES = " & r_int_PerMes & " AND "
      g_str_Parame = g_str_Parame & "MERITA_PERANO = " & r_int_PerAno & " AND "
      g_str_Parame = g_str_Parame & "MERITA_MONEDA = " & r_int_TemAux & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If r_int_TemAux = 1 Then
         r_int_ConTem = 0
      Else
         r_int_ConTem = 800
      End If
         
      For r_int_Contad = 0 To 31 Step 1
   
         r_str_Cadena = "USP_HIS_MERITA ("
         r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
         r_str_Cadena = r_str_Cadena & 0 & ", "
         r_str_Cadena = r_str_Cadena & 0 & ", "
         r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
         r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
         r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
         r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
         r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
         r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_Contad * 2) & "', "
         r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi((r_int_Contad * 2) + 1) & "', "
   
         For r_int_ConAux = 0 To 15 Step 1
            r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & ", "
            r_int_ConTem = r_int_ConTem + 1
         Next
   
         r_str_Cadena = r_str_Cadena & CInt(r_int_TemAux) & ") "
         
         'r_str_Cadena = Left(r_str_Cadena, Len(r_str_Cadena) - 2) & ") "
   
         If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
            MsgBox "Error al ejecutar el Procedimiento USP_HIS_MERITA.", vbCritical, modgen_g_str_NomPlt
            Exit Sub
         End If
   
      Next
   
   Next

End Sub

Private Sub fs_GenDat_Exc()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi

   r_str_Denomi(0) = "1100"
   r_str_Denomi(1) = "Disponible "
   r_str_Denomi(2) = "1200"
   r_str_Denomi(3) = "Fondos Interbancarios"
   r_str_Denomi(4) = "1300"
   r_str_Denomi(5) = "Inversiones Negociables y a Vencimiento "
   
   r_str_Denomi(6) = "1400"
   r_str_Denomi(7) = "Créditos"
   r_str_Denomi(8) = "1401"
   r_str_Denomi(9) = "Vigentes"
   r_str_Denomi(10) = "1403+1404+1405+1406"
   r_str_Denomi(11) = "Otras"
   r_str_Denomi(12) = "1500"
   r_str_Denomi(13) = "Cuentas por Cobrar"
   r_str_Denomi(14) = "1700"
   r_str_Denomi(15) = "Inversiones Permanentes"
   
   r_str_Denomi(16) = "8409.04.01+8409.04.02+8409.04.09"
   r_str_Denomi(17) = "Pos. Activas en Derivados de Tasa de Int."
   r_str_Denomi(18) = "7106.01+7106.02"
   r_str_Denomi(19) = "Pos. Activas en Derivados de M.E."
   r_str_Denomi(20) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
   r_str_Denomi(21) = "Pos. Activas en Otros Derivados Sensibles"
   r_str_Denomi(22) = "1600+1800+1900"
   r_str_Denomi(23) = "Otras Cuentas Activas"
   
   r_str_Denomi(25) = "Total (I)"
   
   r_str_Denomi(26) = "2100"
   r_str_Denomi(27) = "Obligaciones con el Público"
   r_str_Denomi(28) = "2101"
   r_str_Denomi(29) = "Obligaciones a la Vista"
   r_str_Denomi(30) = "2102"
   r_str_Denomi(31) = "Obligaciones por Cuentas de Ahorro"
   r_str_Denomi(32) = "2103"
   r_str_Denomi(33) = "Obligaciones por Cuentas a Plazos"
   
   r_str_Denomi(34) = "2200"
   r_str_Denomi(35) = "Fondos Interbancarios"
   r_str_Denomi(36) = "2300"
   r_str_Denomi(37) = "Depósitos de Empresas del Sistema Financiero y O.I."
   r_str_Denomi(38) = "2400+2600"
   r_str_Denomi(39) = "Adeudados y Otras Obligaciones Financieras "
   r_str_Denomi(40) = "2500"
   r_str_Denomi(41) = "Cuentas por Pagar"
   r_str_Denomi(42) = "2800"
   r_str_Denomi(43) = "Valores, Titulos y Obligaciones en Circulación "
   
   r_str_Denomi(44) = "8409.04.01+8409.04.02+8409.04.09"
   r_str_Denomi(45) = "Pos. Activas en Derivados de Tasa de Int."
   r_str_Denomi(46) = "7106.01+7106.02"
   r_str_Denomi(47) = "Pos. Activas en Derivados de M.E."
   r_str_Denomi(48) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
   r_str_Denomi(49) = "Pos. Activas en Otros Derivados Sensibles"
   r_str_Denomi(50) = "2700+2900"
   r_str_Denomi(51) = "Otras Cuentas Pasivas"
   
   r_str_Denomi(53) = "Total (II)"
   r_str_Denomi(55) = "Monto Delta Neto de Opciones (III)"
   r_str_Denomi(57) = "Descalce Marginal en MN (I-II+III)"
   r_str_Denomi(59) = "Descalce Marginal/Patrimonio Efectivo"
   r_str_Denomi(61) = "Descalce Acumulado en MN"
   r_str_Denomi(63) = "Acumulado/Patrimonio Efectivo"
   r_str_Denomi(65) = "Descalce Acumulado VAC"
   r_str_Denomi(67) = "Total VAC/Patrimonio Efectivo"
   r_str_Denomi(69) = "Descalce Acumulado Tasa"
   r_str_Denomi(71) = "Tasa/Patrimonio Efectivo"

   Erase r_dbl_Evalua

   g_str_Parame = "SELECT * FROM RPT_ANEXOS WHERE "
   g_str_Parame = g_str_Parame & "ANEXOS_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY ANEXOS_CODPRD ASC "

   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Erase r_dbl_Evalua
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
      
         If Trim(g_rst_Princi!ANEXOS_CODPRD) = 3 Or Trim(g_rst_Princi!ANEXOS_CODPRD) = 4 Or Trim(g_rst_Princi!ANEXOS_CODPRD) = 6 Or Trim(g_rst_Princi!ANEXOS_CODPRD) = 7 Then
                     
            r_dbl_Evalua(401) = r_dbl_Evalua(401) + g_rst_Princi!ANEXOS_FECH01 + g_rst_Princi!ANEXOS_FECH02 + g_rst_Princi!ANEXOS_FECH03
            r_dbl_Evalua(402) = r_dbl_Evalua(402) + g_rst_Princi!ANEXOS_FECH04
            r_dbl_Evalua(403) = r_dbl_Evalua(403) + g_rst_Princi!ANEXOS_FECH05
            r_dbl_Evalua(404) = r_dbl_Evalua(404) + g_rst_Princi!ANEXOS_FECH06
            r_dbl_Evalua(405) = r_dbl_Evalua(405) + g_rst_Princi!ANEXOS_FECH07
            r_dbl_Evalua(406) = r_dbl_Evalua(406) + g_rst_Princi!ANEXOS_FECH08 + g_rst_Princi!ANEXOS_FECH09 + g_rst_Princi!ANEXOS_FECH10
            r_dbl_Evalua(406) = r_dbl_Evalua(406) + g_rst_Princi!ANEXOS_FECH11 + g_rst_Princi!ANEXOS_FECH12 + g_rst_Princi!ANEXOS_FECH13 + g_rst_Princi!ANEXOS_FECH14
                        
         ElseIf Trim(g_rst_Princi!ANEXOS_CODPRD) = 1 Or Trim(g_rst_Princi!ANEXOS_CODPRD) = 2 Then
         
            r_dbl_Evalua(408) = r_dbl_Evalua(408) + g_rst_Princi!ANEXOS_FECH01 + g_rst_Princi!ANEXOS_FECH02 + g_rst_Princi!ANEXOS_FECH03
            r_dbl_Evalua(409) = r_dbl_Evalua(409) + g_rst_Princi!ANEXOS_FECH04
            r_dbl_Evalua(410) = r_dbl_Evalua(410) + g_rst_Princi!ANEXOS_FECH05
            r_dbl_Evalua(411) = r_dbl_Evalua(411) + g_rst_Princi!ANEXOS_FECH06
            r_dbl_Evalua(412) = r_dbl_Evalua(412) + g_rst_Princi!ANEXOS_FECH07
            r_dbl_Evalua(413) = r_dbl_Evalua(413) + g_rst_Princi!ANEXOS_FECH08 + g_rst_Princi!ANEXOS_FECH09 + g_rst_Princi!ANEXOS_FECH10
            r_dbl_Evalua(413) = r_dbl_Evalua(413) + g_rst_Princi!ANEXOS_FECH11 + g_rst_Princi!ANEXOS_FECH12 + g_rst_Princi!ANEXOS_FECH13 + g_rst_Princi!ANEXOS_FECH14
            
         End If

         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      r_dbl_MulUso = 0
      
      For r_int_Contad = 401 To 406 Step 1
         r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_Contad)
      Next
      
      r_dbl_Evalua(400) = r_dbl_MulUso
      
      
      r_dbl_MulUso = 0
      
      For r_int_Contad = 408 To 413 Step 1
         r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_Contad)
      Next
      
      r_dbl_Evalua(407) = r_dbl_MulUso
   
   End If
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub fs_GenRpt()
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_lng_TotReg     As Long
   Dim r_lng_RegAct     As Long
   Dim r_int_ValNeg     As Integer
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   Dim r_str_FecRpt     As String
      
   Screen.MousePointer = 11
   
   Call fs_GenDat_Exc
   Call fs_GeneDB

   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   For r_int_TemAux = 1 To 2 Step 1
   
      Select Case r_int_TemAux
         Case 1:
            r_obj_Excel.Sheets(r_int_TemAux).Name = "01"
         Case 2:
            r_obj_Excel.Sheets(r_int_TemAux).Name = "02"
      End Select
      
      With r_obj_Excel.Sheets(r_int_TemAux)
      
         '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
         '.DrawingObjects(1).Left = 20
         '.DrawingObjects(1).Top = 20
         
         .Range(.Cells(3, 16), .Cells(5, 16)).HorizontalAlignment = xlHAlignRight
         .Cells(3, 16) = "Anexo Nº 7"
         .Cells(5, 16) = "CODIGO S.B.S.: 240"
         .Range(.Cells(3, 16), .Cells(3, 17)).Merge
         .Range(.Cells(5, 16), .Cells(5, 17)).Merge
               
         .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignLeft
         .Cells(3, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
         .Cells(5, 1) = "EMPRESA: MI CASITA"
         .Range(.Cells(3, 1), .Cells(3, 2)).Merge
         .Range(.Cells(5, 1), .Cells(5, 2)).Merge
   
          
         .Range(.Cells(6, 8), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(6, 8), .Cells(7, 1)).Font.Bold = True
          
         .Range(.Cells(6, 1), .Cells(6, 14)).Merge
         .Range(.Cells(7, 1), .Cells(7, 14)).Merge
         .Range(.Cells(8, 1), .Cells(8, 14)).Merge
         
         '.Range(.Cells(6, 8), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
         If r_int_TemAux = 1 Then
            .Cells(6, 1) = "A. MEDICIÓN DEL RIESGO DE TASA DE INTERES EN MONEDA NACIONAL"
            .Cells(8, 1) = "(Expresado en Nuevos Soles)"
         Else
            .Cells(6, 1) = "B. MEDICIÓN DEL RIESGO DE TASA DE INTERES EN MONEDA EXTRANJERA"
            .Cells(8, 1) = "(Expresado en Dólares Americanos)"
         End If
         
         .Cells(7, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
         '.Cells(8, 1) = "(En Nuevos Soles y de Dólares Americanos)"
                              
         .Cells(13, 2) = "MONEDA: NUEVOS SOLES"
         .Cells(15, 1) = "CUENTAS "
         .Cells(15, 2) = "DENOMINACION"
         .Cells(13, 3) = "1"
         .Cells(15, 3) = "1-7d"
         .Cells(13, 4) = "2"
         .Cells(15, 4) = "8-15d"
         .Cells(13, 5) = "3"
         .Cells(15, 5) = "16d-1M"
         .Cells(13, 6) = "4"
         .Cells(15, 6) = "2M"
         .Cells(13, 7) = "5"
         .Cells(15, 7) = "3M"
         .Cells(13, 8) = "6"
         .Cells(15, 8) = "4-6M"
         .Cells(13, 9) = "7"
         .Cells(15, 9) = "7M-1A"
         .Cells(13, 10) = "8"
         .Cells(15, 10) = "2A"
         .Cells(13, 11) = "9"
         .Cells(15, 11) = "3A"
         .Cells(13, 12) = "10"
         .Cells(15, 12) = "4A"
         .Cells(13, 13) = "11"
         .Cells(15, 13) = "5A"
         .Cells(13, 14) = "12"
         .Cells(15, 14) = "6-10A"
         .Cells(13, 15) = "13"
         .Cells(15, 15) = "11-20A"
         .Cells(13, 16) = "14"
         .Cells(15, 16) = "+20A"
         .Cells(13, 17) = "0"
         .Cells(15, 17) = "No Recuperable (1)"
         .Cells(13, 18) = "TOTAL"
         .Cells(15, 18) = ".(2)."
                        
         .Cells(16, 2) = "Activos "
         .Cells(17, 1) = "1100"
         .Cells(17, 2) = "Disponible "
         .Cells(18, 1) = "1200"
         .Cells(18, 2) = "Fondos Interbancarios"
         .Cells(19, 1) = "1300"
         .Cells(19, 2) = "Inversiones Negociables y a Vencimiento "
         
         .Cells(21, 1) = "1400"
         .Cells(21, 2) = "Créditos"
         .Cells(22, 1) = "1401"
         .Cells(22, 2) = "Vigentes (*)"
         .Cells(23, 1) = "1403+1404+1405+1406"
         .Cells(23, 2) = "Otras (*)"
         .Cells(24, 1) = "1500"
         .Cells(24, 2) = "Cuentas por Cobrar"
         .Cells(25, 1) = "1700"
         .Cells(25, 2) = "Inversiones Permanentes"
         
         .Cells(27, 1) = "8409.04.01+8409.04.02+8409.04.09"
         .Cells(27, 2) = "Pos. Activas en Derivados de Tasa de Int."
         .Cells(28, 1) = "7106.01+7106.02"
         .Cells(28, 2) = "Pos. Activas en Derivados de M.E."
         .Cells(29, 1) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
         .Cells(29, 2) = "Pos. Activas en Otros Derivados Sensibles"
         .Cells(30, 1) = "1600+1800+1900"
         .Cells(30, 2) = "Otras Cuentas Activas"
         
         .Cells(31, 2) = "Total (I)(3)"
         .Cells(32, 2) = "Pasivos"
         .Cells(33, 1) = "2100"
         .Cells(33, 2) = "Obligaciones con el Público"
         .Cells(34, 1) = "2101"
         .Cells(34, 2) = "Obligaciones a la Vista (*)"
         .Cells(35, 1) = "2102"
         .Cells(35, 2) = "Obligaciones por Cuentas de Ahorro (*)"
         .Cells(36, 1) = "2103"
         .Cells(36, 2) = "Obligaciones por Cuentas a Plazos (*)"
         
         .Cells(38, 1) = "2200"
         .Cells(38, 2) = "Fondos Interbancarios"
         .Cells(39, 1) = "2300"
         .Cells(39, 2) = "Depósitos de Empresas del Sistema Financiero y O.I."
         .Cells(40, 1) = "2400+2600"
         .Cells(40, 2) = "Adeudados y Otras Obligaciones Financieras "
         .Cells(35, 1) = "2500"
         .Cells(35, 2) = "Cuentas por Pagar"
         .Cells(36, 1) = "2800"
         .Cells(36, 2) = "Valores, Titulos y Obligaciones en Circulación "
         
         .Cells(38, 1) = "8409.04.01+8409.04.02+8409.04.09"
         .Cells(38, 2) = "Pos. Activas en Derivados de Tasa de Int."
         .Cells(39, 1) = "7106.01+7106.02"
         .Cells(39, 2) = "Pos. Activas en Derivados de M.E."
         .Cells(40, 1) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
         .Cells(40, 2) = "Pos. Activas en Otros Derivados Sensibles"
         .Cells(41, 1) = "2700+2900"
         .Cells(41, 2) = "Otras Cuentas Pasivas"
         
         .Cells(42, 2) = "Total (II)(3)"
         .Cells(43, 2) = "Monto Delta Neto de Opciones (III)(4)"
         .Cells(44, 2) = "Descalce Marginal en MN (I-II+III)"
         .Cells(45, 2) = "Descalce Marginal/Patrimonio Efectivo (5)"
         .Cells(46, 2) = "Descalce Acumulado en MN"
         .Cells(47, 2) = "Acumulado/Patrimonio Efectivo (5)"
         .Cells(48, 2) = "Descalce Acumulado VAC (6)"
         .Cells(49, 2) = "Total VAC/Patrimonio Efectivo (5)"
         .Cells(50, 2) = "Descalce Acumulado Tasa (7)"
         .Cells(51, 2) = "Tasa/Patrimonio Efectivo (5)"
         
         .Cells(14, 1).RowHeight = 20
         .Cells(15, 1).RowHeight = 20
               
         .Columns("A").ColumnWidth = 30
         .Columns("A").NumberFormat = "@"
               
         .Columns("B").ColumnWidth = 34
         '.Columns("B").NumberFormat = "@"
         '.Columns("B").HorizontalAlignment = xlHAlignCenter
         
         .Range("C:R").NumberFormat = "###,###,##0.00"
         .Range("C:R").ColumnWidth = 12
               
         '.Range(.Cells(13, 1), .Cells(15, 1)).Merge
         '.Range(.Cells(13, 1), .Cells(15, 1)).WrapText = True
         .Range(.Cells(13, 1), .Cells(15, 1)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 1), .Cells(15, 1)).HorizontalAlignment = xlHAlignCenter
         
               
         '.Range(.Cells(13, 2), .Cells(15, 2)).Merge
         '.Range(.Cells(13, 2), .Cells(15, 2)).WrapText = True
         .Range(.Cells(13, 2), .Cells(15, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 2), .Cells(15, 2)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 3), .Cells(15, 3)).Merge
         .Range(.Cells(14, 3), .Cells(15, 3)).WrapText = True
         .Range(.Cells(14, 3), .Cells(15, 3)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 3), .Cells(15, 3)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 4), .Cells(15, 4)).Merge
         .Range(.Cells(14, 4), .Cells(15, 4)).WrapText = True
         .Range(.Cells(14, 4), .Cells(15, 4)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 4), .Cells(15, 4)).HorizontalAlignment = xlHAlignCenter
               
         .Range(.Cells(14, 5), .Cells(15, 5)).Merge
         .Range(.Cells(14, 5), .Cells(15, 5)).WrapText = True
         .Range(.Cells(14, 5), .Cells(15, 5)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 5), .Cells(15, 5)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 6), .Cells(15, 6)).Merge
         .Range(.Cells(14, 6), .Cells(15, 6)).WrapText = True
         .Range(.Cells(14, 6), .Cells(15, 6)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 6), .Cells(15, 6)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 7), .Cells(15, 7)).Merge
         .Range(.Cells(14, 7), .Cells(15, 7)).WrapText = True
         .Range(.Cells(14, 7), .Cells(15, 7)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 7), .Cells(15, 7)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 8), .Cells(15, 8)).Merge
         .Range(.Cells(14, 8), .Cells(15, 8)).WrapText = True
         .Range(.Cells(14, 8), .Cells(15, 8)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 8), .Cells(15, 8)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 9), .Cells(15, 9)).Merge
         .Range(.Cells(14, 9), .Cells(15, 9)).WrapText = True
         .Range(.Cells(14, 9), .Cells(15, 9)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 9), .Cells(15, 9)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 10), .Cells(15, 10)).Merge
         .Range(.Cells(14, 10), .Cells(15, 10)).WrapText = True
         .Range(.Cells(14, 10), .Cells(15, 10)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 10), .Cells(15, 10)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 11), .Cells(15, 11)).Merge
         .Range(.Cells(14, 11), .Cells(15, 11)).WrapText = True
         .Range(.Cells(14, 11), .Cells(15, 11)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 11), .Cells(15, 11)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 12), .Cells(15, 12)).Merge
         .Range(.Cells(14, 12), .Cells(15, 12)).WrapText = True
         .Range(.Cells(14, 12), .Cells(15, 12)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 12), .Cells(15, 12)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 13), .Cells(15, 13)).Merge
         .Range(.Cells(14, 13), .Cells(15, 13)).WrapText = True
         .Range(.Cells(14, 13), .Cells(15, 13)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 13), .Cells(15, 13)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 14), .Cells(15, 14)).Merge
         .Range(.Cells(14, 14), .Cells(15, 14)).WrapText = True
         .Range(.Cells(14, 14), .Cells(15, 14)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 14), .Cells(15, 14)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 15), .Cells(15, 15)).Merge
         .Range(.Cells(14, 15), .Cells(15, 15)).WrapText = True
         .Range(.Cells(14, 15), .Cells(15, 15)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 15), .Cells(15, 15)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 16), .Cells(15, 16)).Merge
         .Range(.Cells(14, 16), .Cells(15, 16)).WrapText = True
         .Range(.Cells(14, 16), .Cells(15, 16)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 16), .Cells(15, 16)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 17), .Cells(15, 17)).Merge
         .Range(.Cells(14, 17), .Cells(15, 17)).WrapText = True
         .Range(.Cells(14, 17), .Cells(15, 17)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 17), .Cells(15, 17)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(14, 18), .Cells(15, 18)).Merge
         .Range(.Cells(14, 18), .Cells(15, 18)).WrapText = True
         .Range(.Cells(14, 18), .Cells(15, 18)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(14, 18), .Cells(15, 18)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(13, 3), .Cells(15, 18)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         .Range(.Cells(13, 3), .Cells(13, 18)).NumberFormat = "###"
         '.Range(.Cells(14, 3), .Cells(14, 18)).NumberFormat = "@"
                                
         .Range(.Cells(16, 2), .Cells(16, 2)).Font.Bold = True
         .Range(.Cells(16, 2), .Cells(16, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(16, 2), .Cells(16, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(16, 2), .Cells(16, 2)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(31, 1), .Cells(31, 18)).Font.Bold = True
         .Range(.Cells(31, 1), .Cells(31, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 18)).VerticalAlignment = xlVAlignCenter
         
         .Range(.Cells(32, 2), .Cells(32, 2)).Font.Bold = True
         .Range(.Cells(32, 2), .Cells(32, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(32, 2), .Cells(32, 2)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(32, 2), .Cells(32, 2)).HorizontalAlignment = xlHAlignCenter
                    
         Dim r_int_Contad As Integer
         
         For r_int_Contad = 1 To 18 Step 1
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(16, 1), .Cells(51, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         For r_int_Contad = 42 To 51 Step 1
            If r_int_Contad <> 50 And r_int_Contad <> 48 And r_int_Contad <> 46 And r_int_Contad <> 44 Then
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Font.Bold = True
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If
         Next
         
         .Range(.Cells(13, 1), .Cells(15, 18)).Font.Bold = True
         .Range(.Cells(13, 1), .Cells(15, 18)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 1), .Cells(15, 18)).HorizontalAlignment = xlHAlignCenter
         '.Range(.Cells(13, 1), .Cells(15, 16)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(15, 18)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         
         For r_int_Contad = 29 To 40 Step 11
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).RowHeight = 30
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).WrapText = True
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 18)).VerticalAlignment = xlVAlignCenter
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 2)).HorizontalAlignment = xlHAlignLeft
         Next
         
         For r_int_Contad = 44 To 52 Step 1
            If r_int_Contad <> 49 Then
               .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 2)).Merge
            End If
         Next
         
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 8
      
         If r_int_TemAux = 1 Then
            r_int_ConTem = 0
         Else
            r_int_ConTem = 800
         End If
         
         For r_int_ConAux = 17 To 51 Step 1
            If r_int_ConAux <> 20 And r_int_ConAux <> 32 And r_int_ConAux <> 37 Then
               For r_int_Contad = 3 To 18 Step 1
                  .Cells(r_int_ConAux, r_int_Contad) = r_dbl_Evalua(r_int_ConTem)
                  r_int_ConTem = r_int_ConTem + 1
               Next
            End If
         Next
      
      End With
      
      'Bloquear el archivo
      r_obj_Excel.Sheets(r_int_TemAux).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
   
   Next
        
   Screen.MousePointer = 0
  
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

End Sub

Private Sub fs_GenDat_DB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi

   r_str_Denomi(0) = "1100"
   r_str_Denomi(1) = "Disponible "
   r_str_Denomi(2) = "1200"
   r_str_Denomi(3) = "Fondos Interbancarios"
   r_str_Denomi(4) = "1300"
   r_str_Denomi(5) = "Inversiones Negociables y a Vencimiento "
   
   r_str_Denomi(6) = "1400"
   r_str_Denomi(7) = "Créditos"
   r_str_Denomi(8) = "1401"
   r_str_Denomi(9) = "Vigentes"
   r_str_Denomi(10) = "1403+1404+1405+1406"
   r_str_Denomi(11) = "Otras"
   r_str_Denomi(12) = "1500"
   r_str_Denomi(13) = "Cuentas por Cobrar"
   r_str_Denomi(14) = "1700"
   r_str_Denomi(15) = "Inversiones Permanentes"
   
   r_str_Denomi(16) = "8409.04.01+8409.04.02+8409.04.09"
   r_str_Denomi(17) = "Pos. Activas en Derivados de Tasa de Int."
   r_str_Denomi(18) = "7106.01+7106.02"
   r_str_Denomi(19) = "Pos. Activas en Derivados de M.E."
   r_str_Denomi(20) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
   r_str_Denomi(21) = "Pos. Activas en Otros Derivados Sensibles"
   r_str_Denomi(22) = "1600+1800+1900"
   r_str_Denomi(23) = "Otras Cuentas Activas"
   
   r_str_Denomi(25) = "Total (I)"
   
   r_str_Denomi(26) = "2100"
   r_str_Denomi(27) = "Obligaciones con el Público"
   r_str_Denomi(28) = "2101"
   r_str_Denomi(29) = "Obligaciones a la Vista"
   r_str_Denomi(30) = "2102"
   r_str_Denomi(31) = "Obligaciones por Cuentas de Ahorro"
   r_str_Denomi(32) = "2103"
   r_str_Denomi(33) = "Obligaciones por Cuentas a Plazos"
   
   r_str_Denomi(34) = "2200"
   r_str_Denomi(35) = "Fondos Interbancarios"
   r_str_Denomi(36) = "2300"
   r_str_Denomi(37) = "Depósitos de Empresas del Sistema Financiero y O.I."
   r_str_Denomi(38) = "2400+2600"
   r_str_Denomi(39) = "Adeudados y Otras Obligaciones Financieras "
   r_str_Denomi(40) = "2500"
   r_str_Denomi(41) = "Cuentas por Pagar"
   r_str_Denomi(42) = "2800"
   r_str_Denomi(43) = "Valores, Titulos y Obligaciones en Circulación "
   
   r_str_Denomi(44) = "8409.04.01+8409.04.02+8409.04.09"
   r_str_Denomi(45) = "Pos. Activas en Derivados de Tasa de Int."
   r_str_Denomi(46) = "7106.01+7106.02"
   r_str_Denomi(47) = "Pos. Activas en Derivados de M.E."
   r_str_Denomi(48) = "Valor Nominal Registrado en Ctas. Contingentes o de Orden"
   r_str_Denomi(49) = "Pos. Activas en Otros Derivados Sensibles"
   r_str_Denomi(50) = "2700+2900"
   r_str_Denomi(51) = "Otras Cuentas Pasivas"
   
   r_str_Denomi(53) = "Total (II)"
   r_str_Denomi(55) = "Monto Delta Neto de Opciones (III)"
   r_str_Denomi(57) = "Descalce Marginal en MN (I-II+III)"
   r_str_Denomi(59) = "Descalce Marginal/Patrimonio Efectivo"
   r_str_Denomi(61) = "Descalce Acumulado en MN"
   r_str_Denomi(63) = "Acumulado/Patrimonio Efectivo"
   r_str_Denomi(65) = "Descalce Acumulado VAC"
   r_str_Denomi(67) = "Total VAC/Patrimonio Efectivo"
   r_str_Denomi(69) = "Descalce Acumulado Tasa"
   r_str_Denomi(71) = "Tasa/Patrimonio Efectivo"

   Erase r_dbl_Evalua
   
   For r_int_Contad = 1 To 2 Step 1

      g_str_Parame = "SELECT * FROM HIS_MERITA WHERE "
      g_str_Parame = g_str_Parame & "MERITA_PERMES = " & r_int_PerMes & " AND "
      g_str_Parame = g_str_Parame & "MERITA_PERANO = " & r_int_PerAno & " AND "
      g_str_Parame = g_str_Parame & "MERITA_MONEDA = " & r_int_Contad & " "
      g_str_Parame = g_str_Parame & "ORDER BY MERITA_NUMITE ASC "
   
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
                  
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
         g_rst_Princi.MoveFirst
         If r_int_Contad = 1 Then
            r_int_ConAux = -1
         Else
            r_int_ConAux = 799
         End If
               
         Do While Not g_rst_Princi.EOF
                        
            r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!MERITA_MTO001
            r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!MERITA_MTO002
            r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!MERITA_MTO003
            r_dbl_Evalua(r_int_ConAux + 4) = g_rst_Princi!MERITA_MTO004
            r_dbl_Evalua(r_int_ConAux + 5) = g_rst_Princi!MERITA_MTO005
            r_dbl_Evalua(r_int_ConAux + 6) = g_rst_Princi!MERITA_MTO006
            r_dbl_Evalua(r_int_ConAux + 7) = g_rst_Princi!MERITA_MTO007
            r_dbl_Evalua(r_int_ConAux + 8) = g_rst_Princi!MERITA_MTO008
            r_dbl_Evalua(r_int_ConAux + 9) = g_rst_Princi!MERITA_MTO009
            r_dbl_Evalua(r_int_ConAux + 10) = g_rst_Princi!MERITA_MTO010
            r_dbl_Evalua(r_int_ConAux + 11) = g_rst_Princi!MERITA_MTO011
            r_dbl_Evalua(r_int_ConAux + 12) = g_rst_Princi!MERITA_MTO012
            r_dbl_Evalua(r_int_ConAux + 13) = g_rst_Princi!MERITA_MTO013
            r_dbl_Evalua(r_int_ConAux + 14) = g_rst_Princi!MERITA_MTO014
            r_dbl_Evalua(r_int_ConAux + 15) = g_rst_Princi!MERITA_MTONRE
            r_dbl_Evalua(r_int_ConAux + 16) = g_rst_Princi!MERITA_MTOTOT
            
            r_int_ConAux = r_int_ConAux + 16
            
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
   Next

End Sub



