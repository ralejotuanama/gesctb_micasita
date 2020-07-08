VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_21 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   7290
   ClientTop       =   4710
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   6720
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   4207
      _StockProps     =   15
      BackColor       =   14215660
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
         Width           =   6645
         _Version        =   65536
         _ExtentX        =   11721
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
            Caption         =   "Anexo N° 16-B"
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
            Width           =   5925
            _Version        =   65536
            _ExtentX        =   10451
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0116-04 Simulación de Escenario de Estres y Plan de Contingencia"
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
            Picture         =   "GesCtb_frm_834.frx":0000
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   750
         Width           =   6645
         _Version        =   65536
         _ExtentX        =   11721
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
            Picture         =   "GesCtb_frm_834.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6030
            Picture         =   "GesCtb_frm_834.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_834.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_834.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_834.frx":11A2
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
         Width           =   6645
         _Version        =   65536
         _ExtentX        =   11721
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
Attribute VB_Name = "frm_RepSbs_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim r_dbl_Evalua(600)   As Double
   Dim r_str_Denomi(100)   As String
   Dim r_str_Fechas(26)    As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_SIESCO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_SIESCO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

Private Sub fs_Genera_ArcPla()

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
      
   r_str_NomRes = "C:\04" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".116"
   
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
   
   Print #r_int_NumRes, Format(116, "0000") & Format(4, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_int_ConAux = 0
   
   
   For r_int_Contad = 100 To 3400 Step 100
      r_str_Cadena = ""
      
      If r_int_Contad <> 100 And r_int_Contad <> 1000 And r_int_Contad <> 3000 Then
         For r_int_ConTem = 0 To 11 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                         
      Else
         For r_int_ConTem = 0 To 11 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 15)
            r_int_ConAux = r_int_ConAux + 1
         Next
                     
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "0000") & r_str_Cadena
      
      
      If r_int_Contad = 200 Or r_int_Contad = 800 Then
         r_int_Contad = r_int_Contad - 50
      ElseIf r_int_Contad = 600 Then
         r_int_Contad = r_int_Contad + 100
      ElseIf r_int_Contad = 1000 Then
         r_int_Contad = r_int_Contad + 100
      ElseIf r_int_Contad = 1500 Then
         r_int_Contad = r_int_Contad - 80
      ElseIf r_int_Contad = 1520 Then
         r_int_Contad = r_int_Contad - 70
      ElseIf r_int_Contad = 1550 Then
         r_int_Contad = r_int_Contad + 50
      ElseIf r_int_Contad = 2000 Then
         r_int_Contad = r_int_Contad + 100
      ElseIf r_int_Contad = 2200 Then
         r_int_Contad = r_int_Contad + 50
      ElseIf r_int_Contad = 2350 Then
         r_int_Contad = r_int_Contad + 50
      ElseIf r_int_Contad = 2500 Then
         r_int_Contad = r_int_Contad + 100
      End If
      
      If r_int_Contad = 250 Or r_int_Contad = 850 Then
         r_int_Contad = r_int_Contad - 50
      End If
      
   Next
          
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
         
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt

  
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
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(3, 12), .Cells(5, 12)).HorizontalAlignment = xlHAlignRight
      .Cells(3, 12) = "Anexo Nº 16B"
      .Cells(5, 12) = "CODIGO S.B.S.: 240"
      .Range(.Cells(3, 12), .Cells(3, 13)).Merge
      .Range(.Cells(5, 12), .Cells(5, 13)).Merge
            
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
      .Cells(6, 1) = "SIMULACRO DE ESCENARIO DE ESTRES Y PLAN DE CONTINGENCIA 1/"
      .Cells(7, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Cells(8, 1) = "(En Nuevos Soles y de Dólares Americanos)"
                  
      .Cells(15, 1) = "CUENTAS /2"
      .Cells(15, 2) = "DENOMINACION"
      .Cells(13, 3) = "TOTAL"
      .Cells(15, 3) = "M.N."
      .Cells(15, 4) = "M.E."
      .Cells(13, 5) = "HASTA 1 MES"
      .Cells(15, 5) = "M.N."
      .Cells(15, 6) = "M.E."
      .Cells(13, 7) = "MAS DE 1 MES HASTA 2 MESES"
      .Cells(15, 7) = "M.N."
      .Cells(15, 8) = "M.E."
      .Cells(13, 9) = "MAS DE 2 MESES HASTA 3 MESES"
      .Cells(15, 9) = "M.N."
      .Cells(15, 10) = "M.E."
      .Cells(13, 11) = "MAS DE 3 MESES HASTA 6 MESES"
      .Cells(15, 11) = "M.N."
      .Cells(15, 12) = "M.E."
      .Cells(13, 13) = "MAS DE 6 MESES"
      .Cells(15, 13) = "M.N."
      .Cells(15, 14) = "M.E."
            
      .Cells(16, 2) = "Activos /3"
      .Cells(17, 1) = "1100"
      .Cells(17, 2) = "Disponible 4/"
      .Cells(18, 1) = "1200"
      .Cells(18, 2) = "Fondos Interbancarios"
      .Cells(19, 1) = "1300"
      .Cells(19, 2) = "Inversiones Negociables y a Vencimiento 5/"
      .Cells(20, 1) = "1401+1403+1404+1407+1408"
      .Cells(20, 2) = "Creditos"
      .Cells(21, 1) = "1500"
      .Cells(21, 2) = "Cuentas por Cobrar"
      .Cells(22, 1) = "1601+1602"
      .Cells(22, 2) = "Bienes Realizables, Recibidos en Pagos y Adjudicados"
      .Cells(23, 2) = "Posiciones Activas en Inst.Financieros Derivados Delivery 6/"
      .Cells(24, 2) = "Contingentes 7/"
      .Cells(25, 2) = "Total (I)"
      .Cells(26, 2) = "Pasivos"
      .Cells(27, 1) = "2101"
      .Cells(27, 2) = "Obligaciones a la Vista"
      .Cells(28, 1) = "2102"
      .Cells(28, 2) = "Obligaciones por Cuentas de Ahorro"
      .Cells(29, 1) = "2103"
      .Cells(29, 2) = "Obligaciones por Cuentas a Plazos"
      .Cells(30, 1) = "2105-2105.02-2105.03-2105.04"
      .Cells(30, 2) = "Obligaciones Relacionadas con Inversiones Negociables y a Vencimiento"
      .Cells(31, 1) = "2104+2106+2107+2108-2108.05"
      .Cells(31, 2) = "Otras Obligaciones con el Público"
      '.Cells(32, 1) = "2108.05"
      '.Cells(32, 2) = ""
      .Cells(32, 1) = "2200"
      .Cells(32, 2) = "Fondos Interbancarios"
      .Cells(33, 1) = "2300"
      .Cells(33, 2) = "Depósitos de Empresas del Sistema Financiero y OFI"
      .Cells(34, 1) = "2105-2105.02-2105.03-2105.04-2408.01(p)"
      .Cells(34, 2) = "Adeudados y Otras Obligaciones Financieras 8/"
      '.Cells(36, 1) = "2408.01 (p)"
      '.Cells(36, 2) = ""
      .Cells(35, 1) = "2500"
      .Cells(35, 2) = "Cuentas por Pagar"
      .Cells(36, 1) = "2508"
      .Cells(36, 2) = "Valores, Titulos y Obligaciones en Circulación 9/"
      .Cells(37, 1) = ""
      .Cells(37, 2) = "Posiciones Pasivas en Inst. Finac. Derivados - Delivery"
      .Cells(38, 1) = ""
      .Cells(38, 2) = "Contingentes /10"
      .Cells(39, 1) = ""
      .Cells(39, 2) = "Total (II)"
      .Cells(40, 1) = ""
      .Cells(40, 2) = "Brecha (I)-(II)"
      .Cells(41, 1) = ""
      .Cells(41, 2) = "Brecha Acumulada (III)"
      .Cells(42, 1) = ""
      .Cells(42, 2) = "Brecha Acumulada (III) Y Patrimonio Efectivo 11/"
      
      .Cells(44, 2) = "Plan de Contingencia 18/"
      .Cells(45, 2) = "Total (IV)"
      .Cells(47, 2) = "Total (I)-(II)+(IV)"
      .Cells(48, 2) = "Total Acumulado"
      .Cells(49, 2) = "Total Acumulado/Patrimonio Efectivo 17/"
            
      .Columns("A").ColumnWidth = 24
      .Columns("A").NumberFormat = "@"
            
      .Columns("B").ColumnWidth = 44
      '.Columns("B").NumberFormat = "@"
      '.Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Range("C:P").NumberFormat = "###,###,##0.00"
      .Range("C:P").ColumnWidth = 10
            
      .Range(.Cells(13, 1), .Cells(15, 1)).Merge
      .Range(.Cells(13, 1), .Cells(15, 1)).WrapText = True
      .Range(.Cells(13, 1), .Cells(15, 1)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 1), .Cells(15, 1)).HorizontalAlignment = xlHAlignCenter
      
            
      .Range(.Cells(13, 2), .Cells(15, 2)).Merge
      .Range(.Cells(13, 2), .Cells(15, 2)).WrapText = True
      .Range(.Cells(13, 2), .Cells(15, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 2), .Cells(15, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(13, 3), .Cells(14, 4)).Merge
      .Range(.Cells(13, 3), .Cells(14, 4)).WrapText = True
      .Range(.Cells(13, 3), .Cells(14, 4)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 3), .Cells(14, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 3), .Cells(14, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 5), .Cells(14, 6)).Merge
      .Range(.Cells(13, 5), .Cells(14, 6)).WrapText = True
      .Range(.Cells(13, 5), .Cells(14, 6)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 5), .Cells(14, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 5), .Cells(14, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 7), .Cells(14, 8)).Merge
      .Range(.Cells(13, 7), .Cells(14, 8)).WrapText = True
      .Range(.Cells(13, 7), .Cells(14, 8)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 7), .Cells(14, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 7), .Cells(14, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 9), .Cells(14, 10)).Merge
      .Range(.Cells(13, 9), .Cells(14, 10)).WrapText = True
      .Range(.Cells(13, 9), .Cells(14, 10)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 9), .Cells(14, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 9), .Cells(14, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      
      .Range(.Cells(13, 11), .Cells(14, 12)).Merge
      .Range(.Cells(13, 11), .Cells(14, 12)).WrapText = True
      .Range(.Cells(13, 11), .Cells(14, 12)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 11), .Cells(14, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 11), .Cells(14, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 13), .Cells(14, 14)).Merge
      .Range(.Cells(13, 13), .Cells(14, 14)).WrapText = True
      .Range(.Cells(13, 13), .Cells(14, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 13), .Cells(14, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 13), .Cells(14, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                  
      '.Range(.Cells(13, 15), .Cells(14, 16)).Merge
      '.Range(.Cells(13, 15), .Cells(14, 16)).VerticalAlignment = xlCenter
      '.Range(.Cells(13, 15), .Cells(14, 16)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(13, 15), .Cells(14, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
           
      .Range(.Cells(16, 2), .Cells(16, 2)).Font.Bold = True
      .Range(.Cells(16, 2), .Cells(16, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(16, 2), .Cells(16, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(16, 2), .Cells(16, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(25, 1), .Cells(25, 14)).Font.Bold = True
      .Range(.Cells(25, 1), .Cells(25, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(25, 1), .Cells(25, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(25, 1), .Cells(25, 14)).VerticalAlignment = xlCenter
            
      .Range(.Cells(26, 1), .Cells(26, 14)).Font.Bold = True
      .Range(.Cells(26, 1), .Cells(26, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(26, 1), .Cells(26, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(26, 1), .Cells(26, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(26, 1), .Cells(26, 14)).HorizontalAlignment = xlHAlignCenter
            
      Dim r_int_Contad As Integer
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 39 To 42 Step 1
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      .Range(.Cells(13, 1), .Cells(15, 14)).Font.Bold = True
      .Range(.Cells(13, 1), .Cells(15, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 1), .Cells(15, 14)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(13, 1), .Cells(15, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 44 To 45 Step 1
         '.Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 47 To 49 Step 1
         '.Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 44 To 52 Step 1
         If r_int_Contad <> 49 Then
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 2)).Merge
         End If
      Next
      
      .Range(.Cells(44, 2), .Cells(44, 2)).Font.Bold = True
      .Range(.Cells(44, 2), .Cells(44, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(44, 2), .Cells(44, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 7
                        
   End With
   
   r_int_ConTem = 0
   
   For r_int_ConAux = 17 To 49 Step 1
      If r_int_ConAux <> 26 And r_int_ConAux <> 43 And r_int_ConAux <> 44 And r_int_ConAux <> 46 Then
         For r_int_Contad = 3 To 14 Step 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad) = r_dbl_Evalua(r_int_ConTem)
            r_int_ConTem = r_int_ConTem + 1
         Next
      End If
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

   g_str_Parame = "DELETE FROM HIS_SIESCO WHERE "
   g_str_Parame = g_str_Parame & "SIESCO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "SIESCO_PERANO = " & r_int_PerAno & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   r_int_ConTem = 0
      
   For r_int_Contad = 0 To 28 Step 1

      r_str_Cadena = "USP_HIS_SIESCO ("
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

      For r_int_ConAux = 0 To 11 Step 1
         r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & ", "
         r_int_ConTem = r_int_ConTem + 1
      Next

      r_str_Cadena = Left(r_str_Cadena, Len(r_str_Cadena) - 2) & ") "

      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_SIESCO.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If

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
   r_str_Denomi(6) = "1401+1403+1404+1407+1408"
   r_str_Denomi(7) = "Creditos"
   r_str_Denomi(8) = "1500"
   r_str_Denomi(9) = "Cuentas por Cobrar"
   r_str_Denomi(10) = "1601+1602"
   r_str_Denomi(11) = "Bienes Realizables, Recibidos en Pagos y Adjudicados"
   r_str_Denomi(13) = "Posiciones Activas en Inst.Financieros Derivados Delivery "
   r_str_Denomi(15) = "Contingentes "
   r_str_Denomi(17) = "Total (I)"
   r_str_Denomi(18) = "2101"
   r_str_Denomi(19) = "Obligaciones a la Vista"
   r_str_Denomi(20) = "2102"
   r_str_Denomi(21) = "Obligaciones por Cuentas de Ahorro"
   r_str_Denomi(22) = "2103"
   r_str_Denomi(23) = "Obligaciones por Cuentas a Plazos"
   r_str_Denomi(24) = "2105-2105.02-2105.03-2105.04"
   r_str_Denomi(25) = "Obligaciones Relacionadas con Inversiones Negociables y a Vencimiento"
   r_str_Denomi(26) = "2104+2106+2107+2108-2108.05"
   r_str_Denomi(27) = "Otras Obligaciones con el Público"
   r_str_Denomi(28) = "2200"
   r_str_Denomi(29) = "Fondos Interbancarios"
   r_str_Denomi(30) = "2300"
   r_str_Denomi(31) = "Depósitos de Empresas del Sistema Financiero y OFI"
   r_str_Denomi(32) = "2105-2105.02-2105.03-2105.04-2408.01(p)"
   r_str_Denomi(33) = "Adeudados y Otras Obligaciones Financieras "
   r_str_Denomi(34) = "2500"
   r_str_Denomi(35) = "Cuentas por Pagar"
   r_str_Denomi(36) = "2508"
   r_str_Denomi(37) = "Valores, Titulos y Obligaciones en Circulación"
   r_str_Denomi(39) = "Posiciones Pasivas en Inst. Finac. Derivados - Delivery"
   r_str_Denomi(41) = "Contingentes "
   r_str_Denomi(43) = "Total (II)"
   r_str_Denomi(45) = "Brecha (I)-(II)"
   r_str_Denomi(47) = "Brecha Acumulada (III)"
   r_str_Denomi(49) = "Brecha Acumulada (III) Y Patrimonio Efectivo "
   r_str_Denomi(51) = "Total (IV)"
   r_str_Denomi(53) = "Total (I)-(II)+(IV)"
   r_str_Denomi(55) = "Total Acumulado"
   r_str_Denomi(57) = "Total Acumulado/Patrimonio Efectivo "
   
   Erase r_dbl_Evalua

   g_str_Parame = "SELECT * FROM HIS_SIESCO WHERE "
   g_str_Parame = g_str_Parame & "SIESCO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "SIESCO_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY SIESCO_NUMITE ASC "

   
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
   
   'Call fs_GenDat_Exc
   Call fs_GeneDB

   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      '.Pictures.Insert ("C:\miCasita\Desarrollo\Graficos\Logo.jpg")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      .Range(.Cells(3, 12), .Cells(5, 12)).HorizontalAlignment = xlHAlignRight
      .Cells(3, 12) = "Anexo Nº 16B"
      .Cells(5, 12) = "CODIGO S.B.S.: 240"
      .Range(.Cells(3, 12), .Cells(3, 13)).Merge
      .Range(.Cells(5, 12), .Cells(5, 13)).Merge
            
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
      .Cells(6, 1) = "SIMULACRO DE ESCENARIO DE ESTRES Y PLAN DE CONTINGENCIA 1/"
      .Cells(7, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " del " & Format(r_int_PerAno, "0000")
      .Cells(8, 1) = "(En Nuevos Soles y de Dólares Americanos)"
                  
      .Cells(15, 1) = "CUENTAS /2"
      .Cells(15, 2) = "DENOMINACION"
      .Cells(13, 3) = "TOTAL"
      .Cells(15, 3) = "M.N."
      .Cells(15, 4) = "M.E."
      .Cells(13, 5) = "HASTA 1 MES"
      .Cells(15, 5) = "M.N."
      .Cells(15, 6) = "M.E."
      .Cells(13, 7) = "MAS DE 1 MES HASTA 2 MESES"
      .Cells(15, 7) = "M.N."
      .Cells(15, 8) = "M.E."
      .Cells(13, 9) = "MAS DE 2 MESES HASTA 3 MESES"
      .Cells(15, 9) = "M.N."
      .Cells(15, 10) = "M.E."
      .Cells(13, 11) = "MAS DE 3 MESES HASTA 6 MESES"
      .Cells(15, 11) = "M.N."
      .Cells(15, 12) = "M.E."
      .Cells(13, 13) = "MAS DE 6 MESES"
      .Cells(15, 13) = "M.N."
      .Cells(15, 14) = "M.E."
            
      .Cells(16, 2) = "Activos /3"
      .Cells(17, 1) = "1100"
      .Cells(17, 2) = "Disponible 4/"
      .Cells(18, 1) = "1200"
      .Cells(18, 2) = "Fondos Interbancarios"
      .Cells(19, 1) = "1300"
      .Cells(19, 2) = "Inversiones Negociables y a Vencimiento 5/"
      .Cells(20, 1) = "1401+1403+1404+1407+1408"
      .Cells(20, 2) = "Creditos"
      .Cells(21, 1) = "1500"
      .Cells(21, 2) = "Cuentas por Cobrar"
      .Cells(22, 1) = "1601+1602"
      .Cells(22, 2) = "Bienes Realizables, Recibidos en Pagos y Adjudicados"
      .Cells(23, 2) = "Posiciones Activas en Inst.Financieros Derivados Delivery 6/"
      .Cells(24, 2) = "Contingentes 7/"
      .Cells(25, 2) = "Total (I)"
      .Cells(26, 2) = "Pasivos"
      .Cells(27, 1) = "2101"
      .Cells(27, 2) = "Obligaciones a la Vista"
      .Cells(28, 1) = "2102"
      .Cells(28, 2) = "Obligaciones por Cuentas de Ahorro"
      .Cells(29, 1) = "2103"
      .Cells(29, 2) = "Obligaciones por Cuentas a Plazos"
      .Cells(30, 1) = "2105-2105.02-2105.03-2105.04"
      .Cells(30, 2) = "Obligaciones Relacionadas con Inversiones Negociables y a Vencimiento"
      .Cells(31, 1) = "2104+2106+2107+2108-2108.05"
      .Cells(31, 2) = "Otras Obligaciones con el Público"
      '.Cells(32, 1) = "2108.05"
      '.Cells(32, 2) = ""
      .Cells(32, 1) = "2200"
      .Cells(32, 2) = "Fondos Interbancarios"
      .Cells(33, 1) = "2300"
      .Cells(33, 2) = "Depósitos de Empresas del Sistema Financiero y OFI"
      .Cells(34, 1) = "2105-2105.02-2105.03-2105.04-2408.01(p)"
      .Cells(34, 2) = "Adeudados y Otras Obligaciones Financieras 8/"
      '.Cells(36, 1) = "2408.01 (p)"
      '.Cells(36, 2) = ""
      .Cells(35, 1) = "2500"
      .Cells(35, 2) = "Cuentas por Pagar"
      .Cells(36, 1) = "2508"
      .Cells(36, 2) = "Valores, Titulos y Obligaciones en Circulación 9/"
      .Cells(37, 1) = ""
      .Cells(37, 2) = "Posiciones Pasivas en Inst. Finac. Derivados - Delivery"
      .Cells(38, 1) = ""
      .Cells(38, 2) = "Contingentes /10"
      .Cells(39, 1) = ""
      .Cells(39, 2) = "Total (II)"
      .Cells(40, 1) = ""
      .Cells(40, 2) = "Brecha (I)-(II)"
      .Cells(41, 1) = ""
      .Cells(41, 2) = "Brecha Acumulada (III)"
      .Cells(42, 1) = ""
      .Cells(42, 2) = "Brecha Acumulada (III) Y Patrimonio Efectivo 11/"
      
      .Cells(44, 2) = "Plan de Contingencia 18/"
      .Cells(45, 2) = "Total (IV)"
      .Cells(47, 2) = "Total (I)-(II)+(IV)"
      .Cells(48, 2) = "Total Acumulado"
      .Cells(49, 2) = "Total Acumulado/Patrimonio Efectivo 17/"
            
      .Columns("A").ColumnWidth = 24
      .Columns("A").NumberFormat = "@"
            
      .Columns("B").ColumnWidth = 44
      '.Columns("B").NumberFormat = "@"
      '.Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Range("C:P").NumberFormat = "###,###,##0.00"
      .Range("C:P").ColumnWidth = 10
            
      .Range(.Cells(13, 1), .Cells(15, 1)).Merge
      .Range(.Cells(13, 1), .Cells(15, 1)).WrapText = True
      .Range(.Cells(13, 1), .Cells(15, 1)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 1), .Cells(15, 1)).HorizontalAlignment = xlHAlignCenter
      
            
      .Range(.Cells(13, 2), .Cells(15, 2)).Merge
      .Range(.Cells(13, 2), .Cells(15, 2)).WrapText = True
      .Range(.Cells(13, 2), .Cells(15, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 2), .Cells(15, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(13, 3), .Cells(14, 4)).Merge
      .Range(.Cells(13, 3), .Cells(14, 4)).WrapText = True
      .Range(.Cells(13, 3), .Cells(14, 4)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 3), .Cells(14, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 3), .Cells(14, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 5), .Cells(14, 6)).Merge
      .Range(.Cells(13, 5), .Cells(14, 6)).WrapText = True
      .Range(.Cells(13, 5), .Cells(14, 6)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 5), .Cells(14, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 5), .Cells(14, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 7), .Cells(14, 8)).Merge
      .Range(.Cells(13, 7), .Cells(14, 8)).WrapText = True
      .Range(.Cells(13, 7), .Cells(14, 8)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 7), .Cells(14, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 7), .Cells(14, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 9), .Cells(14, 10)).Merge
      .Range(.Cells(13, 9), .Cells(14, 10)).WrapText = True
      .Range(.Cells(13, 9), .Cells(14, 10)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 9), .Cells(14, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 9), .Cells(14, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      
      .Range(.Cells(13, 11), .Cells(14, 12)).Merge
      .Range(.Cells(13, 11), .Cells(14, 12)).WrapText = True
      .Range(.Cells(13, 11), .Cells(14, 12)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 11), .Cells(14, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 11), .Cells(14, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(13, 13), .Cells(14, 14)).Merge
      .Range(.Cells(13, 13), .Cells(14, 14)).WrapText = True
      .Range(.Cells(13, 13), .Cells(14, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 13), .Cells(14, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 13), .Cells(14, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                  
      '.Range(.Cells(13, 15), .Cells(14, 16)).Merge
      '.Range(.Cells(13, 15), .Cells(14, 16)).VerticalAlignment = xlCenter
      '.Range(.Cells(13, 15), .Cells(14, 16)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(13, 15), .Cells(14, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
           
      .Range(.Cells(16, 2), .Cells(16, 2)).Font.Bold = True
      .Range(.Cells(16, 2), .Cells(16, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(16, 2), .Cells(16, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(16, 2), .Cells(16, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(25, 1), .Cells(25, 14)).Font.Bold = True
      .Range(.Cells(25, 1), .Cells(25, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(25, 1), .Cells(25, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(25, 1), .Cells(25, 14)).VerticalAlignment = xlCenter
            
      .Range(.Cells(26, 1), .Cells(26, 14)).Font.Bold = True
      .Range(.Cells(26, 1), .Cells(26, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(26, 1), .Cells(26, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(26, 1), .Cells(26, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(26, 1), .Cells(26, 14)).HorizontalAlignment = xlHAlignCenter
            
      Dim r_int_Contad As Integer
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(16, 1), .Cells(42, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 39 To 42 Step 1
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      .Range(.Cells(13, 1), .Cells(15, 14)).Font.Bold = True
      .Range(.Cells(13, 1), .Cells(15, 14)).VerticalAlignment = xlCenter
      .Range(.Cells(13, 1), .Cells(15, 14)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(13, 1), .Cells(15, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(15, 14)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(44, 1), .Cells(45, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 44 To 45 Step 1
         '.Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 1 To 14 Step 1
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(47, 1), .Cells(49, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 47 To 49 Step 1
         '.Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 5)).Font.Bold = True
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
      
      For r_int_Contad = 44 To 52 Step 1
         If r_int_Contad <> 49 Then
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 2)).Merge
         End If
      Next
      
      .Range(.Cells(44, 2), .Cells(44, 2)).Font.Bold = True
      .Range(.Cells(44, 2), .Cells(44, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(44, 2), .Cells(44, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 7
                        
   End With
   
   r_int_ConTem = 0
   
   For r_int_ConAux = 17 To 49 Step 1
      If r_int_ConAux <> 26 And r_int_ConAux <> 43 And r_int_ConAux <> 44 And r_int_ConAux <> 46 Then
         For r_int_Contad = 3 To 14 Step 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad) = r_dbl_Evalua(r_int_ConTem)
            r_int_ConTem = r_int_ConTem + 1
         Next
      End If
   Next
   
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

   r_str_Denomi(0) = "1100"
   r_str_Denomi(1) = "Disponible "
   r_str_Denomi(2) = "1200"
   r_str_Denomi(3) = "Fondos Interbancarios"
   r_str_Denomi(4) = "1300"
   r_str_Denomi(5) = "Inversiones Negociables y a Vencimiento "
   r_str_Denomi(6) = "1401+1403+1404+1407+1408"
   r_str_Denomi(7) = "Creditos"
   r_str_Denomi(8) = "1500"
   r_str_Denomi(9) = "Cuentas por Cobrar"
   r_str_Denomi(10) = "1601+1602"
   r_str_Denomi(11) = "Bienes Realizables, Recibidos en Pagos y Adjudicados"
   r_str_Denomi(13) = "Posiciones Activas en Inst.Financieros Derivados Delivery "
   r_str_Denomi(15) = "Contingentes "
   r_str_Denomi(17) = "Total (I)"
   r_str_Denomi(18) = "2101"
   r_str_Denomi(19) = "Obligaciones a la Vista"
   r_str_Denomi(20) = "2102"
   r_str_Denomi(21) = "Obligaciones por Cuentas de Ahorro"
   r_str_Denomi(22) = "2103"
   r_str_Denomi(23) = "Obligaciones por Cuentas a Plazos"
   r_str_Denomi(24) = "2105-2105.02-2105.03-2105.04"
   r_str_Denomi(25) = "Obligaciones Relacionadas con Inversiones Negociables y a Vencimiento"
   r_str_Denomi(26) = "2104+2106+2107+2108-2108.05"
   r_str_Denomi(27) = "Otras Obligaciones con el Público"
   r_str_Denomi(28) = "2200"
   r_str_Denomi(29) = "Fondos Interbancarios"
   r_str_Denomi(30) = "2300"
   r_str_Denomi(31) = "Depósitos de Empresas del Sistema Financiero y OFI"
   r_str_Denomi(32) = "2105-2105.02-2105.03-2105.04-2408.01(p)"
   r_str_Denomi(33) = "Adeudados y Otras Obligaciones Financieras "
   r_str_Denomi(34) = "2500"
   r_str_Denomi(35) = "Cuentas por Pagar"
   r_str_Denomi(36) = "2508"
   r_str_Denomi(37) = "Valores, Titulos y Obligaciones en Circulación"
   r_str_Denomi(39) = "Posiciones Pasivas en Inst. Finac. Derivados - Delivery"
   r_str_Denomi(41) = "Contingentes "
   r_str_Denomi(43) = "Total (II)"
   r_str_Denomi(45) = "Brecha (I)-(II)"
   r_str_Denomi(47) = "Brecha Acumulada (III)"
   r_str_Denomi(49) = "Brecha Acumulada (III) Y Patrimonio Efectivo "
   r_str_Denomi(51) = "Total (IV)"
   r_str_Denomi(53) = "Total (I)-(II)+(IV)"
   r_str_Denomi(55) = "Total Acumulado"
   r_str_Denomi(57) = "Total Acumulado/Patrimonio Efectivo "
   
   Erase r_dbl_Evalua

   g_str_Parame = "SELECT * FROM HIS_SIESCO WHERE "
   g_str_Parame = g_str_Parame & "SIESCO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "SIESCO_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY SIESCO_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Erase r_dbl_Evalua
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
      
      r_int_ConAux = -1
            
      Do While Not g_rst_Princi.EOF
      
         r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!SIESCO_TOTAMN
         r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!SIESCO_TOTAME
         r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!SIESCO_1MESMN
         r_dbl_Evalua(r_int_ConAux + 4) = g_rst_Princi!SIESCO_1MESME
         r_dbl_Evalua(r_int_ConAux + 5) = g_rst_Princi!SIESCO_2MESMN
         r_dbl_Evalua(r_int_ConAux + 6) = g_rst_Princi!SIESCO_2MESME
         r_dbl_Evalua(r_int_ConAux + 7) = g_rst_Princi!SIESCO_3MESMN
         r_dbl_Evalua(r_int_ConAux + 8) = g_rst_Princi!SIESCO_3MESME
         r_dbl_Evalua(r_int_ConAux + 9) = g_rst_Princi!SIESCO_6MESMN
         r_dbl_Evalua(r_int_ConAux + 10) = g_rst_Princi!SIESCO_6MESME
         r_dbl_Evalua(r_int_ConAux + 11) = g_rst_Princi!SIESCO_MAMEMN
         r_dbl_Evalua(r_int_ConAux + 12) = g_rst_Princi!SIESCO_MAMEME
         
         r_int_ConAux = r_int_ConAux + 12

         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub




