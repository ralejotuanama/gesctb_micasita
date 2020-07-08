VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   7950
   ClientTop       =   4365
   ClientWidth     =   7035
   Icon            =   "GesCtb_frm_719.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      _Version        =   65536
      _ExtentX        =   12462
      _ExtentY        =   4260
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
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
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
            Caption         =   "Reporte N° 2C-1"
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
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0202-41 Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
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
            Picture         =   "GesCtb_frm_719.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
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
            Picture         =   "GesCtb_frm_719.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6360
            Picture         =   "GesCtb_frm_719.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_719.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_719.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_719.frx":11AE
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
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
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
Attribute VB_Name = "frm_RepSbs_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_str_PerMes     As Integer
Dim r_str_PerAno     As Integer
Dim r_int_Contad     As Integer
Dim r_int_ConGen     As Integer
Dim r_int_ConTem     As Integer
Dim r_int_TemAux     As Integer
Dim r_str_Cadena     As String
Dim r_str_FecRpt     As String
Dim r_dbl_Period(50) As Double
Dim r_str_Descri(10) As String
Dim r_int_ConAux     As Integer
Dim r_int_Period     As Integer
Dim r_str_Period     As String
Dim r_dbl_MaOpBr     As Double
Dim r_dbl_RePaEf     As Double
Dim r_int_CodEmp     As Integer
Dim r_dbl_InLiGl     As Double 'Inversa del limite global
Dim r_dbl_FacAju     As Double 'Factor de ajuste

Dim r_int_NumRes     As Integer
Dim r_str_NomRes     As String
Dim r_int_Cantid     As Integer
Dim r_int_FlgRpr     As Integer

Private Sub cmd_Detall_Click()
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
   
  Call fs_GenExc_Detall
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_RPEROP", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()

   Dim r_int_PerMes        As Integer
   Dim r_int_PerAno        As Integer

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
      
   r_int_Cantid = modsec_gf_CanReg("HIS_RPEROP", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
      
   r_str_NomRes = "C:\41" & Right(r_str_PerAno, 2) & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".202"
   
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
   
   Print #r_int_NumRes, Format(202, "0000") & Format(41, "00") & Format(r_int_CodEmp, "00000") & r_str_PerAno & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
      
   r_int_ConAux = -1
      
   For r_int_Contad = 100 To 500 Step 100
   
      r_str_Cadena = ""
      
      If r_int_Contad < 400 Then
                     
         r_str_Cadena = gs_modsec_Genera(Format(r_dbl_Period(r_int_ConAux + 1), "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Period(r_int_ConAux + 2), "###########0.00"), 1, "0", 18)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Period(r_int_ConAux + 3), "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Period(r_int_ConAux + 4), "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Period(r_int_ConAux + 5), "###########0.00"), 1, "0", 18)
      
      ElseIf r_int_Contad = 400 Then
         r_str_Cadena = gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Period(19), "###########0.00"), 1, "0", 18)
         
      ElseIf r_int_Contad = 500 Then
         r_str_Cadena = gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(0, "###########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Period(24), "###########0.00"), 1, "0", 18)
            
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "000000") & r_str_Cadena
      
      r_int_ConAux = r_int_ConAux + 5
   
   Next
    
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
      
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt

End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_ConGen     As Integer
   
   Dim r_str_Cadena     As String
   Dim r_str_FecRpt     As String

   Dim r_int_Period     As Integer
   Dim r_str_Period     As String
   
   Dim r_int_CodEmp     As Integer
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   Screen.MousePointer = 11
   
   r_dbl_InLiGl = 10.5
   r_dbl_FacAju = 0.4
        
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(4, 1), .Cells(10, 6)).Font.Bold = True
      .Cells(4, 1) = "REPORTE Nº 2-C1"
      .Range(.Cells(4, 1), .Cells(4, 6)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(6, 1), .Cells(6, 6)).Merge
      .Range(.Cells(7, 1), .Cells(7, 6)).Merge
      .Range(.Cells(8, 1), .Cells(8, 6)).Merge
      .Cells(6, 1) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO OPERACIONAL"
      .Cells(7, 1) = "METODO DEL INDICADOR BASICO"
      .Cells(8, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_str_PerAno, "0,000")
      .Range(.Cells(8, 1), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
      .Range(.Cells(8, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(10, 1) = "ENTIDAD: EDPYME MICASITA"
      .Range(.Cells(10, 1), .Cells(10, 6)).Merge
               
      .Cells(12, 1) = "PERIODOS DE 12 MESES"
      .Cells(12, 2) = "INGRESOS FINANCIEROS 1"
      .Cells(12, 3) = "INGRESOS POR SERVICIOS 1"
      .Cells(12, 4) = "GASTOS FINANCIEROS 1"
      .Cells(12, 5) = "GASTOS POR SERVICIOS 1"
      .Cells(12, 6) = "MARGEN OPERACIONAL BRUTO"
      
      .Cells(12, 1).RowHeight = 45
      .Cells(13, 1).RowHeight = 0
   
      .Range(.Cells(12, 1), .Cells(12, 6)).Font.Bold = True
      '.Range(.Cells(12, 1), .Cells(12, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(12, 1), .Cells(29, 6)).VerticalAlignment = xlVAlignCenter
      
      .Range(.Cells(17, 1), .Cells(17, 1)).Font.Bold = True
            
      .Cells(17, 1).RowHeight = 30
      .Cells(18, 1).RowHeight = 0
      
      .Range(.Cells(26, 1), .Cells(26, 10)).Merge
            
      .Range(.Cells(27, 1), .Cells(27, 10)).Merge
            
      .Cells(26, 1).RowHeight = 30
      .Cells(27, 1).RowHeight = 30
      
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
      
      .Range(.Cells(12, 1), .Cells(12, 6)).WrapText = True
      .Range(.Cells(17, 1), .Cells(17, 1)).WrapText = True
      .Range(.Cells(26, 1), .Cells(26, 10)).WrapText = True
      .Range(.Cells(27, 1), .Cells(27, 10)).WrapText = True
              
      .Columns("A").ColumnWidth = 40
      .Columns("A").NumberFormat = "@"
      
      .Columns("B").ColumnWidth = 15
      .Columns("B").NumberFormat = "###,###,##0.00"
            
      .Columns("C").ColumnWidth = 15
      .Columns("C").NumberFormat = "###,###,##0.00"
            
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,##0.00"
            
      .Columns("E").ColumnWidth = 15
      .Columns("E").NumberFormat = "###,###,##0.00"
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,##0.00"
                
   End With
   
   r_obj_Excel.ActiveSheet.Cells(14, 1) = r_str_Descri(0)
   r_obj_Excel.ActiveSheet.Cells(15, 1) = r_str_Descri(1)
   r_obj_Excel.ActiveSheet.Cells(16, 1) = r_str_Descri(2)
   r_obj_Excel.ActiveSheet.Cells(17, 1) = r_str_Descri(3) & " 2"
   r_obj_Excel.ActiveSheet.Cells(19, 1) = r_str_Descri(4) & " 3"
   r_obj_Excel.ActiveSheet.Cells(21, 1) = r_str_Descri(5)
   r_obj_Excel.ActiveSheet.Cells(22, 1) = r_str_Descri(6)
   
   r_int_ConAux = -1
   
   For r_int_Contad = 14 To 18 Step 1
   
      r_dbl_MaOpBr = 0
      r_str_Cadena = ""
      
      If r_int_Contad < 17 Then
                  
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 2) = r_dbl_Period(r_int_ConAux + 1)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 3) = r_dbl_Period(r_int_ConAux + 2)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 4) = r_dbl_Period(r_int_ConAux + 3)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 5) = r_dbl_Period(r_int_ConAux + 4)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 6) = r_dbl_Period(r_int_ConAux + 5)
      
      ElseIf r_int_Contad = 17 Then
        
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 2) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 5) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 6) = r_dbl_Period(19)
         
      ElseIf r_int_Contad = 18 Then
         
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 2) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 5) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 6) = r_dbl_Period(24)
            
      End If
      
      r_dbl_RePaEf = r_dbl_RePaEf + r_dbl_MaOpBr
      
      r_int_ConAux = r_int_ConAux + 5
   
   Next
   
   r_obj_Excel.ActiveSheet.Cells(21, 2) = r_dbl_InLiGl
   r_obj_Excel.ActiveSheet.Cells(22, 2) = r_dbl_FacAju
   
   r_obj_Excel.ActiveSheet.Cells(24, 1) = "NOTAS"
   
   r_obj_Excel.ActiveSheet.Cells(25, 1) = "1. La información correspondiente a las cuentas contables establecidas en el articulo 5º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   r_obj_Excel.ActiveSheet.Cells(26, 1) = "2. El requerimiento de patrimonio efectivo por riesgo operacional será el promedio de los valores positivos del margen operacional bruto multiplicado por 15%, según lo indicado en el artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   r_obj_Excel.ActiveSheet.Cells(27, 1) = "3. El APR por riesgo operacional se halla multiplicado el requerimiento de patrimonio efectivo por riesgo operacional por la inversa del limite global que establece la Ley General en el artículo 199º y la Vigésima Cuarta Disposición Transitoria y por los factores de ajsute que se consignan al final del artículo 3º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
      
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub


Private Sub fs_GenDat()

   Erase r_dbl_Period
   Erase r_str_Descri

   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_dbl_InLiGl = 10.5
   r_dbl_FacAju = 0.4
   
   r_str_Descri(0) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 1
   r_str_Descri(1) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno - 1 & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 2
   r_str_Descri(2) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno - 2 & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 3
   r_str_Descri(3) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO OPERACIONAL"
   r_str_Descri(4) = "APR RIESGO OPERACIONAL"
   r_str_Descri(5) = "INVERSA DEL LIMITE GLOBAL"
   r_str_Descri(6) = "FACTOR DE AJUSTE"

   r_int_ConAux = -1
   
   For r_int_Contad = 0 To 2 Step 1
   
      'g_str_Parame = "SELECT * FROM CNTBL_ASIENTO_DET "
      'g_str_Parame = g_str_Parame & "WHERE (ANO = " & (r_str_PerAno - r_int_Contad) & " OR "
      'g_str_Parame = g_str_Parame & "ANO = " & (r_str_PerAno - (r_int_Contad + 1)) & ") "
      'g_str_Parame = g_str_Parame & "ORDER BY ANO DESC, MES DESC"
      
      g_str_Parame = "SELECT * FROM CNTBL_ASIENTO_DET "
      g_str_Parame = g_str_Parame & "WHERE TO_NUMBER(CONCAT(ANO,DECODE(LENGTH(MES),1,CONCAT('0',MES),MES))) <= " & (r_str_PerAno - r_int_Contad) & Format(r_str_PerMes, "00") & " AND "
      g_str_Parame = g_str_Parame & "TO_NUMBER(CONCAT(ANO,DECODE(LENGTH(MES),1,CONCAT('0',MES),MES))) > " & (r_str_PerAno - (r_int_Contad + 1)) & Format(r_str_PerMes, "00") & " "
      g_str_Parame = g_str_Parame & "ORDER BY ANO DESC, MES DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
         
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
                         
         Do While Not g_rst_Princi.EOF
            
            If g_rst_Princi!MES < 10 Then
               r_str_Period = CStr(g_rst_Princi!ANO & Format(g_rst_Princi!MES, "00"))
            Else
               r_str_Period = CStr(g_rst_Princi!ANO & g_rst_Princi!MES)
            End If
            
            'If r_str_Period <= CStr(r_str_PerAno & Format(r_str_PerMes)) And r_str_Period > CStr((r_str_PerAno - 1) & Format(r_str_PerMes)) Then
            'If r_str_Period <= CStr((r_str_PerAno - r_int_Contad) & Format(r_str_PerMes, "00")) And r_str_Period > CStr((r_str_PerAno - (r_int_Contad + 1)) & Format(r_str_PerMes, "00")) Then
            
               If Left(g_rst_Princi!CNTA_CTBL, 2) = "51" And Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
                  'r_dbl_Period(r_int_Contad + r_int_ConAux) = r_dbl_Period(r_int_Contad + r_int_ConAux) + g_rst_Princi!IMP_MOVSOL
                  r_dbl_Period(r_int_ConAux + 1) = r_dbl_Period(r_int_ConAux + 1) + g_rst_Princi!IMP_MOVSOL
                  
               ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "52" And Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
                  'r_dbl_Period(r_int_Contad + r_int_ConAux + 1) = r_dbl_Period(r_int_Contad + r_int_ConAux + 1) + g_rst_Princi!IMP_MOVSOL
                  r_dbl_Period(r_int_ConAux + 2) = r_dbl_Period(r_int_ConAux + 2) + g_rst_Princi!IMP_MOVSOL
                  
               ElseIf Left(g_rst_Princi!CNTA_CTBL, 2) = "41" And Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
                  'r_dbl_Period(r_int_Contad + r_int_ConAux + 2) = r_dbl_Period(r_int_Contad + r_int_ConAux + 2) + g_rst_Princi!IMP_MOVSOL
                  r_dbl_Period(r_int_ConAux + 3) = r_dbl_Period(r_int_ConAux + 3) + g_rst_Princi!IMP_MOVSOL
                  
               ElseIf (Left(g_rst_Princi!CNTA_CTBL, 2) = "42" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "49") And Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
                  'r_dbl_Period(r_int_Contad + r_int_ConAux + 3) = r_dbl_Period(r_int_Contad + r_int_ConAux + 3) + g_rst_Princi!IMP_MOVSOL
                  r_dbl_Period(r_int_ConAux + 4) = r_dbl_Period(r_int_ConAux + 4) + g_rst_Princi!IMP_MOVSOL
                  
               End If
            
            'End If
                  
            g_rst_Princi.MoveNext
            DoEvents
         Loop
       
      End If
                  
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      r_int_ConAux = r_int_ConAux + 5
     
   Next
   
   r_int_ConGen = 0
   r_dbl_RePaEf = 0
   
   For r_int_Contad = 0 To 3 Step 1
      r_dbl_MaOpBr = 0
      
      For r_int_ConAux = 0 To 3 Step 1
         If r_int_ConAux < 2 Then
            r_dbl_MaOpBr = r_dbl_MaOpBr + r_dbl_Period(r_int_ConGen)
         Else
            r_dbl_MaOpBr = r_dbl_MaOpBr - r_dbl_Period(r_int_ConGen)
         End If
         r_int_ConGen = r_int_ConGen + 1
      Next
      
      r_dbl_RePaEf = r_dbl_RePaEf + r_dbl_MaOpBr
      r_dbl_Period(r_int_ConGen) = r_dbl_MaOpBr
      r_int_ConGen = r_int_ConGen + 1
      
   Next

   r_dbl_Period(19) = r_dbl_RePaEf * 0.15 / 3
   r_dbl_Period(24) = r_dbl_RePaEf * 0.15 / 3 * r_dbl_InLiGl * r_dbl_FacAju
   r_dbl_Period(25) = r_dbl_InLiGl
   r_dbl_Period(30) = r_dbl_FacAju
   
End Sub


Private Sub fs_GeneDB()


   If (r_str_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_str_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "DELETE FROM HIS_RPEROP WHERE "
   g_str_Parame = g_str_Parame & "RPEROP_PERMES = " & r_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "RPEROP_PERANO = " & r_str_PerAno & " "
   'g_str_Parame = g_str_Parame & "RPEROP_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
            
   r_int_ConTem = 0
   
   For r_int_Contad = 0 To 6 Step 1
  
      r_str_Cadena = "USP_HIS_RPEROP ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_str_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_str_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Descri(r_int_Contad) & "', "
                                    
      For r_int_ConAux = 0 To 4 Step 1
         If r_int_ConAux = 4 Then
            r_str_Cadena = r_str_Cadena & r_dbl_Period(r_int_ConTem) & " "
         Else
            r_str_Cadena = r_str_Cadena & r_dbl_Period(r_int_ConTem) & ", "
         End If
         
         r_int_ConTem = r_int_ConTem + 1
      Next
      
      r_str_Cadena = r_str_Cadena & ")"
          
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_RPEROP.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If

   Next
   
End Sub

Private Sub fs_GenRpt()

   Dim r_obj_Excel      As Excel.Application
   
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_ConGen     As Integer
   
   Dim r_str_Cadena     As String
   Dim r_str_FecRpt     As String
     
   Dim r_int_ConAux     As Integer
   Dim r_int_Period     As Integer
   
   Dim r_str_Period     As String

   
   Dim r_int_CodEmp     As Integer
   
   Call fs_GenDat
   
   r_dbl_InLiGl = 10.5
   r_dbl_FacAju = 0.4
        
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(4, 1), .Cells(10, 6)).Font.Bold = True
      .Cells(4, 1) = "REPORTE Nº 2-C1"
      .Range(.Cells(4, 1), .Cells(4, 6)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(6, 1), .Cells(6, 6)).Merge
      .Range(.Cells(7, 1), .Cells(7, 6)).Merge
      .Range(.Cells(8, 1), .Cells(8, 6)).Merge
      .Cells(6, 1) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO OPERACIONAL"
      .Cells(7, 1) = "METODO DEL INDICADOR BASICO"
      .Cells(8, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_str_PerAno, "0,000")
      .Range(.Cells(8, 1), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
      .Range(.Cells(8, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(10, 1) = "ENTIDAD: EDPYME MICASITA"
      .Range(.Cells(10, 1), .Cells(10, 6)).Merge
               
      .Cells(12, 1) = "PERIODOS DE 12 MESES"
      .Cells(12, 2) = "INGRESOS FINANCIEROS 1"
      .Cells(12, 3) = "INGRESOS POR SERVICIOS 1"
      .Cells(12, 4) = "GASTOS FINANCIEROS 1"
      .Cells(12, 5) = "GASTOS POR SERVICIOS 1"
      .Cells(12, 6) = "MARGEN OPERACIONAL BRUTO"
      
      .Cells(12, 1).RowHeight = 45
      .Cells(13, 1).RowHeight = 0
   
      .Range(.Cells(12, 1), .Cells(12, 6)).Font.Bold = True
      '.Range(.Cells(12, 1), .Cells(12, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(12, 1), .Cells(29, 6)).VerticalAlignment = xlVAlignCenter
      
      .Range(.Cells(17, 1), .Cells(17, 1)).Font.Bold = True
            
      .Cells(17, 1).RowHeight = 30
      .Cells(18, 1).RowHeight = 0
      
      .Range(.Cells(26, 1), .Cells(26, 10)).Merge
            
      .Range(.Cells(27, 1), .Cells(27, 10)).Merge
            
      .Cells(26, 1).RowHeight = 30
      .Cells(27, 1).RowHeight = 30
      
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(14, 1), .Cells(16, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(19, 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(17, 6), .Cells(19, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(21, 1), .Cells(22, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
      
      .Range(.Cells(12, 1), .Cells(12, 6)).WrapText = True
      .Range(.Cells(17, 1), .Cells(17, 1)).WrapText = True
      .Range(.Cells(26, 1), .Cells(26, 10)).WrapText = True
      .Range(.Cells(27, 1), .Cells(27, 10)).WrapText = True
              
      .Columns("A").ColumnWidth = 40
      .Columns("A").NumberFormat = "@"
      
      .Columns("B").ColumnWidth = 15
      .Columns("B").NumberFormat = "###,###,##0.00"
            
      .Columns("C").ColumnWidth = 15
      .Columns("C").NumberFormat = "###,###,##0.00"
            
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,##0.00"
            
      .Columns("E").ColumnWidth = 15
      .Columns("E").NumberFormat = "###,###,##0.00"
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,##0.00"
                
   End With
   
   r_obj_Excel.ActiveSheet.Cells(14, 1) = r_str_Descri(0)
   r_obj_Excel.ActiveSheet.Cells(15, 1) = r_str_Descri(1)
   r_obj_Excel.ActiveSheet.Cells(16, 1) = r_str_Descri(2)
   r_obj_Excel.ActiveSheet.Cells(17, 1) = r_str_Descri(3) & " 2"
   r_obj_Excel.ActiveSheet.Cells(19, 1) = r_str_Descri(4) & " 3"
   r_obj_Excel.ActiveSheet.Cells(21, 1) = r_str_Descri(5)
   r_obj_Excel.ActiveSheet.Cells(22, 1) = r_str_Descri(6)
   
   r_int_ConAux = -1
   
   For r_int_Contad = 14 To 18 Step 1
   
      r_dbl_MaOpBr = 0
      r_str_Cadena = ""
      
      If r_int_Contad < 17 Then
                  
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 2) = r_dbl_Period(r_int_ConAux + 1)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 3) = r_dbl_Period(r_int_ConAux + 2)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 4) = r_dbl_Period(r_int_ConAux + 3)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 5) = r_dbl_Period(r_int_ConAux + 4)
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 6) = r_dbl_Period(r_int_ConAux + 5)
      
      ElseIf r_int_Contad = 17 Then
        
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 2) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 5) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 6) = r_dbl_Period(19)
         
      ElseIf r_int_Contad = 18 Then
         
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 2) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 5) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 1, 6) = r_dbl_Period(24)
            
      End If
      
      r_dbl_RePaEf = r_dbl_RePaEf + r_dbl_MaOpBr
      
      r_int_ConAux = r_int_ConAux + 5
   
   Next
   
   r_obj_Excel.ActiveSheet.Cells(21, 2) = r_dbl_InLiGl
   r_obj_Excel.ActiveSheet.Cells(22, 2) = r_dbl_FacAju
   
   r_obj_Excel.ActiveSheet.Cells(24, 1) = "NOTAS"
   
   r_obj_Excel.ActiveSheet.Cells(25, 1) = "1. La información correspondiente a las cuentas contables establecidas en el articulo 5º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   r_obj_Excel.ActiveSheet.Cells(26, 1) = "2. El requerimiento de patrimonio efectivo por riesgo operacional será el promedio de los valores positivos del margen operacional bruto multiplicado por 15%, según lo indicado en el artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   r_obj_Excel.ActiveSheet.Cells(27, 1) = "3. El APR por riesgo operacional se halla multiplicado el requerimiento de patrimonio efectivo por riesgo operacional por la inversa del limite global que establece la Ley General en el artículo 199º y la Vigésima Cuarta Disposición Transitoria y por los factores de ajsute que se consignan al final del artículo 3º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo Operacional."
   
   Call fs_GeneDB
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
      
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

End Sub


Private Sub fs_GenExc_Detall()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
      
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 3
   r_obj_Excel.Workbooks.Add
   
   For r_int_Contad = 1 To 3 Step 1
   
    g_str_Parame = "SELECT * FROM CNTBL_ASIENTO_DET "
    g_str_Parame = g_str_Parame & "WHERE TO_NUMBER(CONCAT(ANO,DECODE(LENGTH(MES),1,CONCAT('0',MES),MES))) <= " & (r_str_PerAno - r_int_Contad) & Format(r_str_PerMes, "00") & " AND "
    g_str_Parame = g_str_Parame & "TO_NUMBER(CONCAT(ANO,DECODE(LENGTH(MES),1,CONCAT('0',MES),MES))) > " & (r_str_PerAno - (r_int_Contad + 1)) & Format(r_str_PerMes, "00") & " "
    'g_str_Parame = g_str_Parame & "ORDER BY ANO DESC, MES DESC, CNTA_CTBL DESC"
    g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL DESC"
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
       Exit Sub
    End If
    
    If r_int_Contad = 1 Then
       r_obj_Excel.Sheets(r_int_Contad).Name = r_str_PerAno & "-" & Format(r_str_PerMes, "00") & " AL " & r_str_PerAno - 1 & "-" & Format(r_str_PerMes + 1, "00")
    ElseIf r_int_Contad = 2 Then
       r_obj_Excel.Sheets(r_int_Contad).Name = r_str_PerAno - 1 & "-" & Format(r_str_PerMes, "00") & " AL " & r_str_PerAno - 2 & "-" & Format(r_str_PerMes + 1, "00")
    Else
       r_obj_Excel.Sheets(r_int_Contad).Name = r_str_PerAno - 2 & "-" & Format(r_str_PerMes, "00") & " AL " & r_str_PerAno - 3 & "-" & Format(r_str_PerMes + 1, "00")
    End If
    
    With r_obj_Excel.Sheets(r_int_Contad)
    
       .Cells(1, 1) = "ITEM"
       .Cells(1, 2) = "AÑO"
       .Cells(1, 3) = "MES"
       .Cells(1, 4) = "CUENTA"
       .Cells(1, 5) = "GLOSA"
       .Cells(1, 6) = "DEBE-HABER"
       .Cells(1, 7) = "IMPORTE"
    
       .Range(.Cells(1, 1), .Cells(1, 7)).Font.Bold = True
       .Range(.Cells(1, 1), .Cells(1, 7)).HorizontalAlignment = xlHAlignCenter
        
       .Columns("A").ColumnWidth = 8
       
       .Columns("B").ColumnWidth = 5
       .Columns("B").HorizontalAlignment = xlHAlignCenter
       
       .Columns("C").ColumnWidth = 5
       .Columns("C").HorizontalAlignment = xlHAlignCenter
       
       .Columns("D").ColumnWidth = 15
       .Columns("D").HorizontalAlignment = xlHAlignCenter
       .Columns("D").NumberFormat = "@"
       
       .Columns("E").ColumnWidth = 67
       .Columns("E").NumberFormat = "@"
       
       .Columns("F").ColumnWidth = 12
       .Columns("F").HorizontalAlignment = xlHAlignCenter
       
       .Columns("G").ColumnWidth = 12
       .Columns("G").NumberFormat = "###,###,##0.00"
          
       g_rst_Princi.MoveFirst
         
       r_int_ConVer = 2
       
       Do While Not g_rst_Princi.EOF
          'Buscando datos de la Garantía en Registro de Hipotecas
          If ((Left(g_rst_Princi!CNTA_CTBL, 2) = "41" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "42" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "49") And Trim(g_rst_Princi!FLAG_DEBHAB) = "D") Or ((Left(g_rst_Princi!CNTA_CTBL, 2) = "51" Or Left(g_rst_Princi!CNTA_CTBL, 2) = "52") And Trim(g_rst_Princi!FLAG_DEBHAB) = "H") Then
          
            .Cells(r_int_ConVer, 1) = r_int_ConVer - 1
            .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!ANO)
            .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!MES)
            .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!CNTA_CTBL)
            .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DET_GLOSA)
            .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!FLAG_DEBHAB)
            .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVSOL)
                                    
            r_int_ConVer = r_int_ConVer + 1
          End If
          g_rst_Princi.MoveNext
          DoEvents
       Loop
    
    End With
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
   
   Next
     
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub


Private Sub fs_GenDat_DB()

   Erase r_dbl_Period
   Erase r_str_Descri

   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_dbl_InLiGl = 10.5
   r_dbl_FacAju = 0.4
   
   r_str_Descri(0) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 1
   r_str_Descri(1) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno - 1 & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 2
   r_str_Descri(2) = "DEL " & Format(r_str_PerMes, "00") & "/" & r_str_PerAno - 2 & " AL " & Format(r_str_PerMes + 1, "00") & "/" & r_str_PerAno - 3
   r_str_Descri(3) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO OPERACIONAL"
   r_str_Descri(4) = "APR RIESGO OPERACIONAL"
   r_str_Descri(5) = "INVERSA DEL LIMITE GLOBAL"
   r_str_Descri(6) = "FACTOR DE AJUSTE"

   r_int_ConAux = -1
      
   g_str_Parame = "SELECT * FROM HIS_RPEROP WHERE "
   g_str_Parame = g_str_Parame & "RPEROP_PERMES = " & r_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "RPEROP_PERANO = " & r_str_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY RPEROP_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
          
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
                      
      Do While Not g_rst_Princi.EOF

         r_dbl_Period(r_int_ConAux + 1) = r_dbl_Period(r_int_ConAux + 1) + g_rst_Princi!RPEROP_INGFIN
         r_dbl_Period(r_int_ConAux + 2) = r_dbl_Period(r_int_ConAux + 2) + g_rst_Princi!RPEROP_INGSER
         r_dbl_Period(r_int_ConAux + 3) = r_dbl_Period(r_int_ConAux + 3) + g_rst_Princi!RPEROP_GASFIN
         r_dbl_Period(r_int_ConAux + 4) = r_dbl_Period(r_int_ConAux + 4) + g_rst_Princi!RPEROP_GASSER
         r_dbl_Period(r_int_ConAux + 5) = r_dbl_Period(r_int_ConAux + 5) + g_rst_Princi!RPEROP_MAOPBR
               
         r_int_ConAux = r_int_ConAux + 5
               
         g_rst_Princi.MoveNext
         DoEvents
         
      Loop
    
   End If
               
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
      
   'r_dbl_Period(19) = r_dbl_RePaEf * 0.15 / 3
   'r_dbl_Period(24) = r_dbl_RePaEf * 0.15 / 3 * r_dbl_InLiGl * r_dbl_FacAju
   'r_dbl_Period(25) = r_dbl_InLiGl
   'r_dbl_Period(30) = r_dbl_FacAju
   
End Sub

