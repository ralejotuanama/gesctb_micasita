VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form9"
   ClientHeight    =   2325
   ClientLeft      =   7605
   ClientTop       =   5730
   ClientWidth     =   5085
   Icon            =   "GesCtb_frm_700.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
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
         TabIndex        =   7
         Top             =   30
         Width           =   5025
         _Version        =   65536
         _ExtentX        =   8864
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
            TabIndex        =   8
            Top             =   30
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   476
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   9
            Top             =   270
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0214-01 Créditos Según dias de Incumplimiento"
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
            Picture         =   "GesCtb_frm_700.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   5025
         _Version        =   65536
         _ExtentX        =   8864
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
            Picture         =   "GesCtb_frm_700.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_700.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_700.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4410
            Picture         =   "GesCtb_frm_700.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_700.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   4
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
         TabIndex        =   11
         Top             =   1440
         Width           =   5025
         _Version        =   65536
         _ExtentX        =   8864
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim r_str_FecIni        As String
   Dim r_str_FecFin        As String
   Dim r_str_Evalua(1000)  As Double
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_dbl_MulUso        As Double
   Dim r_int_PerMes        As Integer
   Dim r_int_PerAno        As Integer
   Dim r_str_Cadena        As String
   Dim r_int_Cantid       As Integer
   Dim r_int_FlgRpr       As Integer

Private Sub cmd_ExpArc_Click()

   Dim r_int_MsgBox As Integer
   
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
      
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   r_int_Cantid = modsec_gf_CanReg("HIS_DIAINC", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

   Dim r_int_MsgBox As Integer
      
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
      
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   r_int_Cantid = modsec_gf_CanReg("HIS_DIAINC", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

Private Sub fs_GenExc()

   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer

   Dim r_str_TipMon     As String
   
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
      .Range(.Cells(1, 1), .Cells(12, 21)).Font.Bold = True
      .Cells(6, 2).HorizontalAlignment = xlHAlignLeft
            
      .Cells(5, 2) = "CREDITOS SEGUN DIAS DE INCUMPLIMIENTO"
      .Cells(6, 2) = "CODIGO: 240"
      .Cells(7, 2) = "Al " & Right(r_str_FecIni, 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Left(r_str_FecFin, 4)
      .Cells(8, 2) = "(En Nuevos Soles)"
      
      .Range("C10") = "CREDITOS SIN ATRASO 2/"
      .Range("D10") = "CREDITOS CON ALGUN DIA DE INCUMPLIMIENTO DE ACUERDO AL CRONOGRAMA DE PAGOS 1/"
      .Range("T10") = "SALDO TOTAL DE CREDITOS 3/"
      
      .Cells(13, 1) = "CORPORATIVOS"
      .Cells(15, 1) = "TRATADOS COMO CORPORATIVOS"
      .Cells(17, 1) = "GRANDES EMPRESAS"
      .Cells(19, 1) = "MEDIANAS EMPRESAS"
      .Cells(21, 1) = "PEQUEÑAS EMPRESAS"
      .Cells(23, 1) = "MICRO EMPRESAS"
      .Cells(25, 1) = "CONSUMO"
      .Cells(27, 1) = "HIPOTECAS PARA VIVIENDAS 18/"
      .Cells(28, 1) = "TOTAL"
      
      For r_int_ConVar = 13 To 23 Step 2
         .Cells(r_int_ConVar, 2) = "ARREND. FINANCIERO + CAPIT. INMOBILIARIA " & r_int_ConVar - 9 & "/"

      Next
      
      .Range(.Cells(10, 1), .Cells(12, 2)).Merge
      
      For r_int_ConVar = 13 To 26 Step 2
         .Range(.Cells(r_int_ConVar, 1), .Cells(r_int_ConVar + 1, 1)).Merge

      Next
      
      .Cells(14, 2) = "OTROS CREDITOS CORPORATIVOS 5/"
      .Cells(16, 2) = "OTROS CREDITOS TRATADOS COMO CORPORATIVOS 7/"
      .Cells(18, 2) = "OTROS CREDITOS A GRANDES EMPRESAS /9"
      .Cells(20, 2) = "OTROS CREDITOS A MEDIANAS EMPRESAS 11/"
      .Cells(22, 2) = "OTROS CREDITOS A PEQUEÑAS EMPRESAS 13/"
      .Cells(24, 2) = "OTROS CREDITOS A MICROEMPRESAS 15/ "
      .Cells(25, 2) = "TARJETA DE CREDITO 16/"
      .Cells(26, 2) = "OTROS CREDITOS DE COSUMO 17/"
            
      '.Range("B22:T22").Font.Bold = True
      
      '.Range("B30") = "TOTAL (I + II)"
      .Range("B28:T28").Font.Bold = True
      
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
      
      '.Range("E39:G39").Borders(xlEdgeTop).LineStyle = xlContinuous
      '.Range("L39:N39").Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(39, 3), .Cells(40, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(39, 3), .Cells(40, 12)).VerticalAlignment = xlVAlignCenter
      
      .Range("D11") = "DE 1 A 15 DIAS"
      .Range("D11:E11").Merge
            
      .Range("F11") = "ENTRE 16 Y 30 DIAS"
      .Range("F11:G11").Merge
            
      .Range("H11") = "ENTRE 31 Y 60 DIAS"
      .Range("H11:I11").Merge
            
      .Range("J11") = "ENTRE 61 Y 90 DIAS"
      .Range("J11:K11").Merge
            
      .Range("L11") = "ENTRE 91 Y 120 DIAS"
      .Range("L11:M11").Merge
            
      .Range("N11") = "ENTRE 121 Y 180 DIAS"
      .Range("N11:O11").Merge
            
      .Range("P11") = "ENTRE 181 Y 365 DIAS"
      .Range("P11:Q11").Merge
            
      .Range("R11") = "MAYOR A 365 DIAS"
      .Range("R11:S11").Merge
            
      .Range(.Cells(10, 1), .Cells(12, 20)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 20)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 20)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 20)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(12, 20)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range("D11:S11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D11:S11").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range("A13:T13").Borders(xlEdgeTop).LineStyle = xlContinuous
      '.Range("A13:T13").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      '.Range("A22:T22").Borders(xlEdgeTop).LineStyle = xlContinuous
      '.Range("A22:T22").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range("A28:T28").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("A28:T28").Borders(xlEdgeBottom).LineStyle = xlContinuous
                
      For r_int_ConVar = 1 To 20 Step 1
         .Range(.Cells(13, 1), .Cells(28, r_int_ConVar)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(28, r_int_ConVar)).Borders(xlEdgeRight).LineStyle = xlContinuous
      Next
      
      For r_int_ConVar = 4 To 18 Step 2
         .Cells(12, r_int_ConVar) = "PORCION NO AMORTIZADA"
         .Cells(12, r_int_ConVar + 1) = "SALDO"
      Next
      
      .Range(.Cells(13, 1), .Cells(28, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(28, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
      .Range(.Cells(10, 3), .Cells(12, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 21)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 21)).VerticalAlignment = xlHAlignFill
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 28
      .Columns("B").ColumnWidth = 47
      .Columns("C").ColumnWidth = 14
      .Columns("T").ColumnWidth = 14
      .Columns("U").ColumnWidth = 2
      
      .Columns("C:T").NumberFormat = "###,###,##0.00"
                   
   End With
   
   
   r_int_ConAux = 13
   r_int_ConTem = 0
      
   For r_int_Contad = 0 To 287 Step 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_ConTem + 3) = r_str_Evalua(r_int_Contad)
      r_int_ConTem = r_int_ConTem + 1
      
      If (r_int_Contad + 1) Mod 18 = 0 Then
         r_int_ConAux = r_int_ConAux + 1
         r_int_ConTem = 0
      End If
      
      'If r_int_Contad = 712 Then
      '   r_int_ConAux = r_int_ConAux + 1
      'End If
      
   Next
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenArc()
   
   Dim r_int_NumRes     As Integer
   Dim r_int_PerMes     As Integer
   Dim r_int_PerAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_ConGen     As Integer
   Dim r_int_CodEmp     As Integer
   
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\01" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".214"
 
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
   
   Print #r_int_NumRes, Format(214, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
      
   r_int_ConTem = 0
   r_int_ConAux = 0
      
   For r_int_ConGen = 100 To 2300 Step 100
      r_str_Cadena = ""
                        
      If r_int_ConGen = 100 Or r_int_ConGen = 400 Or r_int_ConGen = 700 Or r_int_ConGen = 1000 Or r_int_ConGen = 1300 Or r_int_ConGen = 1600 Or r_int_ConGen = 1900 Then
      
         r_int_ConAux = r_int_ConTem
         For r_int_Contad = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_str_Evalua(r_int_ConTem) + r_str_Evalua(r_int_ConTem + 18), "########0.00"), 1, "0", 18)
            r_int_ConTem = r_int_ConTem + 1
         Next
         
         r_int_ConTem = r_int_ConAux
         
      Else
         For r_int_Contad = 0 To 17 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_str_Evalua(r_int_ConTem), "########0.00"), 1, "0", 18)
            r_int_ConTem = r_int_ConTem + 1
         Next
         
      End If
          
      Print #r_int_NumRes, Format(r_int_ConGen, "000000") & r_str_Cadena
      
   Next
         
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt
   
   
End Sub


Private Sub fs_GenDat()

   Erase r_str_Evalua()

   'Creditos Comerciales
   g_str_Parame = "SELECT COMCIE_NUMOPE, COMCIE_CAPVIG, COMCIE_CAPVEN, COMCIE_TIPCAM, COMCIE_TIPMON, COMCIE_DIAMOR, COMCIE_NUECRE FROM CRE_COMCIE WHERE "
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
           
   Do While Not g_rst_Princi.EOF
   
      If g_rst_Princi!COMCIE_NUECRE = 6 Then
           
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 18
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 18
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
         
      ElseIf g_rst_Princi!COMCIE_NUECRE = 1 Or g_rst_Princi!COMCIE_NUECRE = 2 Or g_rst_Princi!COMCIE_NUECRE = 3 Or g_rst_Princi!COMCIE_NUECRE = 4 Or g_rst_Princi!COMCIE_NUECRE = 5 Then

         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 54
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 54
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
         
      ElseIf g_rst_Princi!COMCIE_NUECRE = 7 Then
      
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 90
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 90
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
            
      ElseIf g_rst_Princi!COMCIE_NUECRE = 8 Then
      
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 126
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 126
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
      
      ElseIf g_rst_Princi!COMCIE_NUECRE = 9 Then
      
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 162
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 162
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
      
      ElseIf g_rst_Princi!COMCIE_NUECRE = 10 Then
      
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 198
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 198
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
      
      ElseIf g_rst_Princi!COMCIE_NUECRE = 11 Or g_rst_Princi!COMCIE_NUECRE = 12 Then
   
         If g_rst_Princi!COMCIE_TIPMON = 1 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 234
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!COMCIE_CAPVIG
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!COMCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!COMCIE_CAPVIG
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!COMCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!COMCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!COMCIE_TIPMON = 2 Then
            If g_rst_Princi!COMCIE_CAPVEN = 0 Then
               r_int_Contad = 234
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!COMCIE_DIAMOR >= 1 And g_rst_Princi!COMCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 16 And g_rst_Princi!COMCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 31 And g_rst_Princi!COMCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 61 And g_rst_Princi!COMCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 91 And g_rst_Princi!COMCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 121 And g_rst_Princi!COMCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 181 And g_rst_Princi!COMCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!COMCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!COMCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!COMCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If

      End If
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
       
         
   'Credito Hipotecario
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_CAPVIG, HIPCIE_CAPVEN, HIPCIE_TIPCAM, HIPCIE_TIPMON, HIPCIE_DIAMOR, HIPCIE_NUECRE FROM CRE_HIPCIE WHERE "
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
     
   Do While Not g_rst_Princi.EOF
   
         If g_rst_Princi!HIPCIE_NUECRE = 13 Then
           
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
               r_int_Contad = 252
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + g_rst_Princi!HIPCIE_CAPVIG
            Else
               If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = r_str_Evalua(r_int_Contad + 1) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 2) = r_str_Evalua(r_int_Contad + 2) + g_rst_Princi!HIPCIE_CAPVIG
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + g_rst_Princi!HIPCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + g_rst_Princi!HIPCIE_CAPVIG
                  
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + g_rst_Princi!HIPCIE_CAPVIG
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + g_rst_Princi!HIPCIE_CAPVIG
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + g_rst_Princi!HIPCIE_CAPVIG
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + g_rst_Princi!HIPCIE_CAPVIG
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + g_rst_Princi!HIPCIE_CAPVEN
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + g_rst_Princi!HIPCIE_CAPVIG
   
               End If
            End If
            
         ElseIf g_rst_Princi!HIPCIE_TIPMON = 2 Then
            If g_rst_Princi!HIPCIE_CAPVEN = 0 Then
               r_int_Contad = 252
               r_str_Evalua(r_int_Contad) = r_str_Evalua(r_int_Contad) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               
            Else
               If (g_rst_Princi!HIPCIE_DIAMOR >= 1 And g_rst_Princi!HIPCIE_DIAMOR <= 15) Then
                  r_str_Evalua(r_int_Contad + 1) = Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 2) = Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 16 And g_rst_Princi!HIPCIE_DIAMOR <= 30) Then
                  r_str_Evalua(r_int_Contad + 3) = r_str_Evalua(r_int_Contad + 3) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 4) = r_str_Evalua(r_int_Contad + 4) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 31 And g_rst_Princi!HIPCIE_DIAMOR <= 60) Then
                  r_str_Evalua(r_int_Contad + 5) = r_str_Evalua(r_int_Contad + 5) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 6) = r_str_Evalua(r_int_Contad + 6) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 61 And g_rst_Princi!HIPCIE_DIAMOR <= 90) Then
                  r_str_Evalua(r_int_Contad + 7) = r_str_Evalua(r_int_Contad + 7) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 8) = r_str_Evalua(r_int_Contad + 8) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 91 And g_rst_Princi!HIPCIE_DIAMOR <= 120) Then
                  r_str_Evalua(r_int_Contad + 9) = r_str_Evalua(r_int_Contad + 9) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 10) = r_str_Evalua(r_int_Contad + 10) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 121 And g_rst_Princi!HIPCIE_DIAMOR <= 180) Then
                  r_str_Evalua(r_int_Contad + 11) = r_str_Evalua(r_int_Contad + 11) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 12) = r_str_Evalua(r_int_Contad + 12) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 181 And g_rst_Princi!HIPCIE_DIAMOR <= 365) Then
                  r_str_Evalua(r_int_Contad + 13) = r_str_Evalua(r_int_Contad + 13) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 14) = r_str_Evalua(r_int_Contad + 14) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   
               ElseIf (g_rst_Princi!HIPCIE_DIAMOR >= 365) Then
                  r_str_Evalua(r_int_Contad + 15) = r_str_Evalua(r_int_Contad + 15) + Format(g_rst_Princi!HIPCIE_CAPVEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_str_Evalua(r_int_Contad + 16) = r_str_Evalua(r_int_Contad + 16) + Format(g_rst_Princi!HIPCIE_CAPVIG * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   
               End If
            End If
         End If
         
      End If

            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
      
         
   For r_int_ConAux = 0 To 268 Step 18
      r_dbl_MulUso = 0
      
      For r_int_Contad = 0 To 16 Step 1
         r_dbl_MulUso = r_dbl_MulUso + r_str_Evalua(r_int_Contad + r_int_ConAux)
         
         If r_int_Contad = 16 Then
            r_str_Evalua(r_int_Contad + r_int_ConAux + 1) = r_dbl_MulUso
         End If
         
      Next
      
   Next
      
   For r_int_ConAux = 0 To 17 Step 1
      r_dbl_MulUso = 0
      
      For r_int_Contad = 0 To 252 Step 18
         r_dbl_MulUso = r_dbl_MulUso + r_str_Evalua(r_int_Contad + r_int_ConAux)
         
      Next
      
      r_str_Evalua(r_int_ConAux + 270) = r_dbl_MulUso
      
   Next

End Sub


Private Sub fs_GeneDB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM HIS_DIAINC WHERE "
   g_str_Parame = g_str_Parame & "DIAINC_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "DIAINC_PERANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "DIAINC_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = 0
         
   For r_int_Contad = 0 To 14 Step 1
  
      r_str_Cadena = "USP_HIS_DIAINC ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_09', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      
      For r_int_ConTem = 0 To 17 Step 1
         r_str_Cadena = r_str_Cadena & Format(r_str_Evalua(r_int_ConAux), "########0.00") & ", "
         r_int_ConAux = r_int_ConAux + 1
      Next
     
      r_str_Cadena = Left(r_str_Cadena, Len(r_str_Cadena) - 2) & ")"
          
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_DIAINC.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If

   Next

End Sub

Private Sub fs_GenRpt()

   Call fs_GenDat
   Call fs_GeneDB
           
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
    
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".HIS_DIAINC"
    
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_20.RPT"
           
   crp_Imprim.SelectionFormula = "{HIS_DIAINC.DIAINC_PERMES} = " & r_int_PerMes & " AND {HIS_DIAINC.DIAINC_PERANO} = " & r_int_PerAno & " AND {HIS_DIAINC.DIAINC_USUCRE} = '" & modgen_g_str_CodUsu & "' "
         
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   Screen.MousePointer = 0
  
End Sub


Private Sub fs_GenDat_DB()

   Erase r_str_Evalua()

   g_str_Parame = "SELECT * FROM HIS_DIAINC WHERE "
   g_str_Parame = g_str_Parame & "DIAINC_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "DIAINC_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY DIAINC_NUMITE ASC "
  
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
   
   r_int_ConAux = -1
           
   Do While Not g_rst_Princi.EOF
   
      r_str_Evalua(r_int_ConAux + 1) = g_rst_Princi!DIAINC_CRESAT
      r_str_Evalua(r_int_ConAux + 2) = g_rst_Princi!DIAINC_POAM01
      r_str_Evalua(r_int_ConAux + 3) = g_rst_Princi!DIAINC_SALD01
      r_str_Evalua(r_int_ConAux + 4) = g_rst_Princi!DIAINC_POAM16
      r_str_Evalua(r_int_ConAux + 5) = g_rst_Princi!DIAINC_SALD16
      r_str_Evalua(r_int_ConAux + 6) = g_rst_Princi!DIAINC_POAM31
      r_str_Evalua(r_int_ConAux + 7) = g_rst_Princi!DIAINC_SALD31
      r_str_Evalua(r_int_ConAux + 8) = g_rst_Princi!DIAINC_POAM61
      r_str_Evalua(r_int_ConAux + 9) = g_rst_Princi!DIAINC_SALD61
      r_str_Evalua(r_int_ConAux + 10) = g_rst_Princi!DIAINC_POAM91
      r_str_Evalua(r_int_ConAux + 11) = g_rst_Princi!DIAINC_SALD91
      r_str_Evalua(r_int_ConAux + 12) = g_rst_Princi!DIAINC_POA121
      r_str_Evalua(r_int_ConAux + 13) = g_rst_Princi!DIAINC_SAL121
      r_str_Evalua(r_int_ConAux + 14) = g_rst_Princi!DIAINC_POA181
      r_str_Evalua(r_int_ConAux + 15) = g_rst_Princi!DIAINC_SAL181
      r_str_Evalua(r_int_ConAux + 16) = g_rst_Princi!DIAINC_POA365
      r_str_Evalua(r_int_ConAux + 17) = g_rst_Princi!DIAINC_SAL365
      r_str_Evalua(r_int_ConAux + 18) = g_rst_Princi!DIAINC_SALTOT
      
      r_int_ConAux = r_int_ConAux + 18
            
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
         
         
   For r_int_ConAux = 0 To 268 Step 18
      r_dbl_MulUso = 0
      
      For r_int_Contad = 0 To 16 Step 1
         r_dbl_MulUso = r_dbl_MulUso + r_str_Evalua(r_int_Contad + r_int_ConAux)
         
         If r_int_Contad = 16 Then
            r_str_Evalua(r_int_Contad + r_int_ConAux + 1) = r_dbl_MulUso
         End If
         
      Next
   Next
      
   For r_int_ConAux = 0 To 17 Step 1
      r_dbl_MulUso = 0
      
      For r_int_Contad = 0 To 252 Step 18
         r_dbl_MulUso = r_dbl_MulUso + r_str_Evalua(r_int_Contad + r_int_ConAux)
         
      Next
      r_str_Evalua(r_int_ConAux + 270) = r_dbl_MulUso
      
   Next

End Sub


