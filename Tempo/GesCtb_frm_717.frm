VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   6855
   ClientTop       =   4350
   ClientWidth     =   8625
   Icon            =   "GesCtb_frm_717.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
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
         Width           =   8565
         _Version        =   65536
         _ExtentX        =   15108
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
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte 2-B1 Anexo N° 3"
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
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0202-22 Metodo Estandar Requerimientos de Patrimonio Efectivo por Riesgo Cambiario II"
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
            Picture         =   "GesCtb_frm_717.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   8565
         _Version        =   65536
         _ExtentX        =   15108
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
            Picture         =   "GesCtb_frm_717.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7950
            Picture         =   "GesCtb_frm_717.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_717.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_717.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_717.frx":11AE
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
         Width           =   8565
         _Version        =   65536
         _ExtentX        =   15108
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
Attribute VB_Name = "frm_RepSbs_17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim r_dbl_Evalua(600)   As Double
   Dim r_str_Descri(30)    As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String
   Dim r_dbl_MulUso        As Double
   Dim r_str_Cadena        As String
   Dim r_str_NomRes        As String
   Dim r_str_ParAux        As String
   Dim r_dbl_LimGlo        As Double
   Dim r_dbl_FacAju        As Double
   Dim r_dbl_TipCam        As Double
   Dim r_int_Cantid        As Integer
   Dim r_int_FlgRpr        As Integer
   Dim r_int_TemAux        As Integer
   
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
   
   r_dbl_LimGlo = 10.5
   r_dbl_FacAju = 0.96
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_REPAEF", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("¿Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_GenAr1
         Call fs_GenAr2
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
         
   Call fs_GenAr1
   Call fs_GenAr2
      
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_REPAEF", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

   Erase r_str_Descri

   r_str_Descri(0) = "Dólar Americano"
   r_str_Descri(1) = "Euro"
   r_str_Descri(2) = "Yen Japonés"
   r_str_Descri(3) = "Libra Esterlina"
   r_str_Descri(4) = "Dólar Canadienes"
   r_str_Descri(5) = "Franco Suizo"
   r_str_Descri(6) = "Corona Danesa"
   r_str_Descri(7) = "Corona Sueca"
   r_str_Descri(8) = "Corona Sueca"
   r_str_Descri(9) = "Real"
   r_str_Descri(10) = "Boliviano"
   r_str_Descri(11) = "Otras Divisas"
   r_str_Descri(12) = "Oro"
   r_str_Descri(13) = "TOTAL ME"
   r_str_Descri(14) = "Mayor Posición Global Agregada en Divisas de Oro"
   r_str_Descri(15) = "Posición Global de Oro"
   r_str_Descri(16) = "Impacto Gamma de las Divisas"
   r_str_Descri(17) = "Impacto Vega de las Divisas"
   r_str_Descri(18) = "Total"
   r_str_Descri(19) = "APR por Riesgo de Mercado - Riesgo Cambiario"
   r_str_Descri(20) = "Inversa del Límite Global"
   r_str_Descri(21) = "Factor de Ajuste"
      

   Erase r_dbl_Evalua()
   
   For r_int_TemAux = 1 To 2 Step 1
   
      If r_int_TemAux = 1 Then
         r_int_ConAux = 0
      Else
         r_int_ConAux = 300
      End If
            
      For r_int_Contad = 0 To 11 Step 1
         r_dbl_Evalua(r_int_ConAux + 2) = r_dbl_Evalua(r_int_ConAux + 0) - r_dbl_Evalua(r_int_ConAux + 1)
         r_dbl_Evalua(r_int_ConAux + 7) = r_dbl_Evalua(r_int_ConAux + 6) + r_dbl_Evalua(r_int_ConAux + 5) + r_dbl_Evalua(r_int_ConAux + 4) + r_dbl_Evalua(r_int_ConAux + 3)
         r_dbl_Evalua(r_int_ConAux + 9) = r_dbl_Evalua(r_int_ConAux + 8) + r_dbl_Evalua(r_int_ConAux + 7) + r_dbl_Evalua(r_int_ConAux + 2)
         r_int_ConAux = r_int_ConAux + 12
      Next
      
      For r_int_Contad = r_int_ConAux To 11 Step 1
         r_dbl_MulUso = 0
         For r_int_ConTem = (0 + r_int_Contad) To (144 + r_int_Contad) Step 12
            r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_ConTem)
         Next
         
         r_dbl_Evalua(156 + r_int_Contad) = r_dbl_MulUso
         
      Next
      
      If r_int_TemAux = 1 Then
         r_int_ConAux = 0
      Else
         r_int_ConAux = 300
      End If
      
      r_dbl_Evalua(r_int_ConAux + 168) = r_dbl_Evalua(r_int_ConAux + 9)
      r_dbl_Evalua(r_int_ConAux + 169) = 9.5
      r_dbl_Evalua(r_int_ConAux + 170) = Format(r_dbl_Evalua(r_int_ConAux + 168) * 0.095, "###########0.00")
      r_dbl_Evalua(r_int_ConAux + 181) = 9.5
      r_dbl_Evalua(r_int_ConAux + 193) = 100
      r_dbl_Evalua(r_int_ConAux + 205) = 100
      
      r_dbl_MulUso = 0
      
      For r_int_Contad = 170 + r_int_ConAux To 206 + r_int_ConAux Step 12
         r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_Contad)
         
      Next
      
      r_dbl_Evalua(r_int_ConAux + 218) = r_dbl_MulUso
      r_dbl_Evalua(r_int_ConAux + 230) = r_dbl_Evalua(r_int_ConAux + 218) * r_dbl_LimGlo * r_dbl_FacAju
      r_dbl_Evalua(r_int_ConAux + 240) = r_dbl_LimGlo
      r_dbl_Evalua(r_int_ConAux + 252) = r_dbl_FacAju
   
   Next
         
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
                     
         .Range(.Cells(2, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(2, 1), .Cells(6, 1)).Font.Bold = True
               
         .Cells(2, 1) = "ANEXO 3"
         .Cells(4, 1) = "REPORTE Nº 2-B1 ANEXO 3"
         .Cells(5, 1) = "MÉTODO ESTÁNDAR"
         .Cells(6, 1) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO CAMBIARIO"
         .Cells(7, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
         
         If r_int_TemAux = 1 Then
            .Cells(8, 1) = "(Expresado En Nuevos Soles)"
         Else
            .Cells(8, 1) = "(Expresado En Dolares Americanos)"
         End If
         
         .Cells(10, 1) = "EMPRESA: Edpyme MiCasita S.A."
         .Range(.Cells(10, 1), .Cells(10, 2)).Merge
         
         .Cells(12, 1) = "I. EXPOSICIÓN EN MONEDA EXTRANJERA"
         .Range(.Cells(12, 1), .Cells(12, 2)).Merge
         .Range(.Cells(12, 1), .Cells(12, 1)).Font.Bold = True
         
         .Range(.Cells(10, 1), .Cells(12, 2)).HorizontalAlignment = xlHAlignLeft
         
         .Cells(13, 1) = "Divisas"
         .Cells(13, 2) = "Activos (a)"
         .Cells(13, 3) = "Pasivos (b)"
         .Cells(13, 4) = "Posición de Cambio del Balance en ME (a)-(b) (1)"
         .Cells(13, 5) = "Posición en Forward de ME"
         .Cells(14, 5) = "Largas (2)"
         .Cells(14, 6) = "Cortas (3)"
         .Cells(13, 7) = "Posición en Otros Productos Finacieros Derivados de ME"
         .Cells(14, 7) = "Largas (4)"
         .Cells(14, 8) = "Cortas (5)"
         .Cells(13, 9) = "Posición Neta en Derivados de ME (6)=(2+3+4+5)"
         .Cells(13, 10) = "Valor Delta Neto en ME (7)"
         .Cells(13, 11) = "Posición Global en ME (8)=(1+6+7)"
         .Cells(13, 12) = "Sensibilidad de Opciones"
         .Cells(14, 12) = "Gamma (9)"
         .Cells(14, 13) = "Vega (10)"
         
         .Cells(15, 1) = r_str_Descri(0)
         .Cells(16, 1) = r_str_Descri(1)
         .Cells(17, 1) = r_str_Descri(2)
         .Cells(18, 1) = r_str_Descri(3)
         .Cells(19, 1) = r_str_Descri(4)
         .Cells(20, 1) = r_str_Descri(5)
         .Cells(21, 1) = r_str_Descri(6)
         .Cells(22, 1) = r_str_Descri(7)
         .Cells(23, 1) = r_str_Descri(8)
         .Cells(24, 1) = r_str_Descri(9)
         .Cells(25, 1) = r_str_Descri(10)
         .Cells(26, 1) = r_str_Descri(11)
         .Cells(27, 1) = r_str_Descri(12)
         .Cells(28, 1) = r_str_Descri(13)
        
         .Range(.Cells(2, 1), .Cells(2, 13)).Merge
         .Range(.Cells(4, 1), .Cells(4, 13)).Merge
         .Range(.Cells(5, 1), .Cells(5, 13)).Merge
         .Range(.Cells(6, 1), .Cells(6, 13)).Merge
         .Range(.Cells(7, 1), .Cells(7, 13)).Merge
         .Range(.Cells(8, 1), .Cells(8, 13)).Merge
         
         .Range(.Cells(13, 1), .Cells(14, 1)).Merge
         .Range(.Cells(13, 2), .Cells(14, 2)).Merge
         .Range(.Cells(13, 3), .Cells(14, 3)).Merge
         .Range(.Cells(13, 4), .Cells(14, 4)).Merge
         .Range(.Cells(13, 5), .Cells(13, 6)).Merge
         .Range(.Cells(13, 7), .Cells(13, 8)).Merge
         .Range(.Cells(13, 9), .Cells(14, 9)).Merge
         .Range(.Cells(13, 10), .Cells(14, 10)).Merge
         .Range(.Cells(13, 11), .Cells(14, 11)).Merge
         .Range(.Cells(13, 12), .Cells(13, 13)).Merge
               
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
         .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         
         .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         .Range(.Cells(27, 1), .Cells(27, 13)).Borders(xlEdgeTop).LineStyle = xlDouble
         .Range(.Cells(27, 1), .Cells(27, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         .Range(.Cells(28, 1), .Cells(28, 13)).Font.Bold = True
               
         .Range(.Cells(13, 1), .Cells(14, 13)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(13, 1), .Cells(14, 13)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(13, 1), .Cells(14, 13)).WrapText = True
         .Range(.Cells(13, 1), .Cells(14, 13)).Font.Bold = True
         
         .Cells(30, 1) = "II. REQUERIMIENTO DE PATRIMONIO EFECTIVO"
         .Range(.Cells(30, 1), .Cells(30, 2)).Merge
         .Range(.Cells(30, 1), .Cells(30, 1)).Font.Bold = True
         
         .Cells(31, 1).RowHeight = 40
         .Cells(31, 1) = "Concepto"
         .Cells(31, 3) = "Importe Base"
         .Cells(31, 4) = "Factor"
         .Cells(31, 5) = "Requerimiento de Patrimonio Efectivo"
         
         .Range(.Cells(31, 1), .Cells(31, 2)).Merge
         .Range(.Cells(36, 1), .Cells(36, 2)).Merge
         .Range(.Cells(37, 1), .Cells(37, 2)).Merge
         
         .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
         .Range(.Cells(31, 1), .Cells(31, 5)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(31, 1), .Cells(31, 5)).VerticalAlignment = xlVAlignCenter
         .Range(.Cells(31, 1), .Cells(31, 5)).WrapText = True
         .Range(.Cells(31, 1), .Cells(31, 5)).Font.Bold = True
         
         For r_int_Contad = 1 To 5 Step 1
            If r_int_Contad = 1 Then
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               '.Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               '.Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
               
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).HorizontalAlignment = xlHAlignCenter
               .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Font.Bold = True
               
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
               
            ElseIf r_int_Contad = 2 Then
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).HorizontalAlignment = xlHAlignCenter
               
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
   
            ElseIf r_int_Contad = 4 Then
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
               
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
               '.Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).HorizontalAlignment = xlHAlignCenter
               .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Font.Bold = True
            ElseIf r_int_Contad <> 2 Then
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
               .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
         Next
               
         .Cells(32, 1) = r_str_Descri(14)
         .Cells(32, 2) = "A"
         .Cells(33, 1) = r_str_Descri(15)
         .Cells(33, 2) = "B"
         .Cells(34, 1) = r_str_Descri(16)
         .Cells(34, 2) = "C"
         .Cells(35, 1) = r_str_Descri(17)
         .Cells(35, 2) = "D"
         .Cells(36, 1) = r_str_Descri(18)
         .Cells(37, 1) = r_str_Descri(19) & " /1"
         
         .Range(.Cells(32, 4), .Cells(35, 4)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(32, 4), .Cells(35, 4)).NumberFormat = "@"
         
         .Cells(32, 4) = "9.5%"
         .Cells(33, 4) = "k%"
         .Cells(34, 4) = "100.0%"
         .Cells(35, 4) = "100.0%"
         
         
         .Cells(39, 1) = r_str_Descri(20)
         .Cells(40, 1) = r_str_Descri(21)
         
         .Cells(39, 2) = r_dbl_LimGlo
         .Cells(40, 2) = r_dbl_FacAju
         
         .Range(.Cells(43, 1), .Cells(43, 13)).Merge
         .Cells(43, 1).RowHeight = 30
         .Cells(43, 1).WrapText = True
         
         .Cells(43, 1) = "1. APR por riesgo de mercado se halla multiplicando el requerimiento de patrimonio efectivo por riesgo de mercado por la inversa del límite global que establece la Ley General en el artículo " & _
                           "199º y la Vigésima Cuarta Disposición Transitoria y por el factor de ajuste que se consigna al final del artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo de Mercado."
                                                   
         .Range(.Cells(48, 4), .Cells(48, 5)).Merge
         .Range(.Cells(48, 8), .Cells(48, 9)).Merge
         
         .Range(.Cells(49, 4), .Cells(49, 5)).Merge
         .Range(.Cells(49, 8), .Cells(49, 9)).Merge
         
         .Range(.Cells(49, 4), .Cells(49, 5)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(49, 8), .Cells(49, 9)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(49, 4), .Cells(49, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(49, 8), .Cells(49, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
         
         .Cells(49, 4) = "Gerente General"
         .Cells(49, 8) = "Gerente Unidad de Riesgos"
              
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
               
         .Columns("A").ColumnWidth = 42
         .Columns("B").ColumnWidth = 15
         .Columns("C").ColumnWidth = 15
         .Columns("D").ColumnWidth = 20
         .Columns("E").ColumnWidth = 15
         .Columns("F").ColumnWidth = 15
         .Columns("G").ColumnWidth = 15
         .Columns("H").ColumnWidth = 15
         .Columns("I").ColumnWidth = 25
         .Columns("J").ColumnWidth = 15
         .Columns("K").ColumnWidth = 20
         .Columns("L").ColumnWidth = 15
         .Columns("M").ColumnWidth = 15
               
         .Columns("B:M").NumberFormat = "###,###,##0.00"
      
         r_int_ConAux = 15
         
         r_int_ConTem = 0
            
         If r_int_TemAux = 1 Then
            
            For r_int_Contad = 0 To 167 Step 1
               .Cells(r_int_ConAux, r_int_ConTem + 2) = r_dbl_Evalua(r_int_Contad)
               r_int_ConTem = r_int_ConTem + 1
               
               If (r_int_Contad + 1) Mod 12 = 0 Then
                  r_int_ConAux = r_int_ConAux + 1
                  r_int_ConTem = 0
               End If
                     
            Next
         
         Else
         
            For r_int_Contad = 300 To 467 Step 1
               .Cells(r_int_ConAux, r_int_ConTem + 2) = r_dbl_Evalua(r_int_Contad)
               r_int_ConTem = r_int_ConTem + 1
               
               If (r_int_Contad + 1) Mod 12 = 0 Then
                  r_int_ConAux = r_int_ConAux + 1
                  r_int_ConTem = 0
               End If
         
            Next
         
         End If
         
         If r_int_TemAux = 1 Then
            r_int_ConTem = 168
         Else
            r_int_ConTem = 468
         End If
         
         For r_int_Contad = 32 To 35 Step 1
            .Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConTem)
            .Cells(r_int_Contad, 5) = r_dbl_Evalua(r_int_ConTem + 2)
            r_int_ConTem = r_int_ConTem + 12
                  
         Next
         
         If r_int_TemAux = 1 Then
            .Cells(36, 5) = r_dbl_Evalua(218)
            .Cells(37, 5) = r_dbl_Evalua(230)
         Else
            .Cells(36, 5) = r_dbl_Evalua(518)
            .Cells(37, 5) = r_dbl_Evalua(530)
         End If
         
         
         
         'Bloquear el archivo
         'r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=123
         
      End With
            
   Next
            
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenAr1()
  
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
   
   Screen.MousePointer = 11
      
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\21" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".202"
   
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
   
   Print #r_int_NumRes, Format(202, "0000") & Format(21, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
        
   r_str_Cadena = ""
   
   For r_int_Contad = 0 To 11 Step 1
      r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_Contad), "########0.00"), 1, "0", 18)
   Next
   
   Print #r_int_NumRes, Format(1, "000000") & "USD" & r_str_Cadena
   
   r_int_ConAux = 144
   
   For r_int_Contad = 10000 To 10100 Step 100
      r_str_Cadena = ""
      
      For r_int_ConTem = 0 To 11 Step 1
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConTem + r_int_ConAux), "########0.00"), 1, "0", 18)
      Next
      
      r_int_ConAux = r_int_ConAux + 12
      
      Print #r_int_NumRes, Format(r_int_Contad, "000000") & "   " & r_str_Cadena
      
   Next
         
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
         
   MsgBox "Archivo 1 creado.", vbInformation, modgen_g_str_NomPlt

End Sub

Private Sub fs_GenAr2()
  
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
   
   Screen.MousePointer = 11
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\22" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".202"
   
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
   
   Print #r_int_NumRes, Format(202, "0000") & Format(22, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_int_ConAux = 168
   
   For r_int_Contad = 100 To 600 Step 100
      r_str_Cadena = ""
      
      'If r_int_Contad < 500 Then
         'r_str_Cadena = gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18) & gs_modsec_Genera(0, 1, "0", 9) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 2), "########0.00"), 1, "0", 18)
         r_str_Cadena = gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 1), "000.000000"), 1, "0", 9) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux + 2), "########0.00"), 1, "0", 18)
         r_int_ConAux = r_int_ConAux + 12
      'Else
      '   r_str_Cadena = gs_modsec_Genera(0, 1, "0", 18) & gs_modsec_Genera(0, 1, "0", 9) & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
      '   r_int_ConAux = r_int_ConAux + 1
         
      'End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "000000") & r_str_Cadena
      
   Next
         
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
         
   Screen.MousePointer = 0
         
   MsgBox "Archivo 2 creado.", vbInformation, modgen_g_str_NomPlt

End Sub


Private Sub fs_GeneDB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   For r_int_TemAux = 1 To 2 Step 1
   
      g_str_Parame = "DELETE FROM HIS_REPAEF WHERE "
      g_str_Parame = g_str_Parame & "REPAEF_PERMES = " & r_int_PerMes & " AND "
      g_str_Parame = g_str_Parame & "REPAEF_PERANO = " & r_int_PerAno & " AND "
      g_str_Parame = g_str_Parame & "REPAEF_MONEDA = " & r_int_TemAux & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
                 
               
      If r_int_TemAux = 1 Then
         r_int_ConTem = 0
      Else
         r_int_ConTem = 300
      End If
      
      For r_int_Contad = 0 To 21 Step 1
     
         r_str_Cadena = "USP_HIS_REPAEF ("
         r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
         r_str_Cadena = r_str_Cadena & 0 & ", "
         r_str_Cadena = r_str_Cadena & 0 & ", "
         r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
         r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
         r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
         r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
         r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
         r_str_Cadena = r_str_Cadena & "'" & r_str_Descri(r_int_Contad) & "', "
                                       
         For r_int_ConAux = 0 To 11 Step 1
            'If r_int_ConAux = 11 Then
            '   r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & " "
            'Else
               r_str_Cadena = r_str_Cadena & r_dbl_Evalua(r_int_ConTem) & ", "
            'End If
            
            r_int_ConTem = r_int_ConTem + 1
         Next
         
         r_str_Cadena = r_str_Cadena & CInt(r_int_TemAux) & ")"
             
         If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
            MsgBox "Error al ejecutar el Procedimiento USP_HIS_REPAEF.", vbCritical, modgen_g_str_NomPlt
            Exit Sub
         End If
   
      Next
   
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
                  
      .Range(.Cells(2, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 1), .Cells(6, 1)).Font.Bold = True
            
      .Cells(2, 1) = "ANEXO 3"
      .Cells(4, 1) = "REPORTE Nº 2-B1 ANEXO 3"
      .Cells(5, 1) = "MÉTODO ESTÁNDAR"
      .Cells(6, 1) = "REQUERIMIENTO DE PATRIMONIO EFECTIVO POR RIESGO CAMBIARIO"
      .Cells(7, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      .Cells(8, 1) = "(Expresado En Nuevos Soles)"
      
      .Cells(10, 1) = "EMPRESA: Edpyme MiCasita S.A."
      .Range(.Cells(10, 1), .Cells(10, 2)).Merge
      
      .Cells(12, 1) = "I. EXPOSICIÓN EN MONEDA EXTRANJERA"
      .Range(.Cells(12, 1), .Cells(12, 2)).Merge
      .Range(.Cells(12, 1), .Cells(12, 1)).Font.Bold = True
      
      .Range(.Cells(10, 1), .Cells(12, 2)).HorizontalAlignment = xlHAlignLeft
      
      .Cells(13, 1) = "Divisas"
      .Cells(13, 2) = "Activos (a)"
      .Cells(13, 3) = "Pasivos (b)"
      .Cells(13, 4) = "Posición de Cambio del Balance en ME (a)-(b) (1)"
      .Cells(13, 5) = "Posición en Forward de ME"
      .Cells(14, 5) = "Largas (2)"
      .Cells(14, 6) = "Cortas (3)"
      .Cells(13, 7) = "Posición en Otros Productos Finacieros Derivados de ME"
      .Cells(14, 7) = "Largas (4)"
      .Cells(14, 8) = "Cortas (5)"
      .Cells(13, 9) = "Posición Neta en Derivados de ME (6)=(2+3+4+5)"
      .Cells(13, 10) = "Valor Delta Neto en ME (7)"
      .Cells(13, 11) = "Posición Global en ME (8)=(1+6+7)"
      .Cells(13, 12) = "Sensibilidad de Opciones"
      .Cells(14, 12) = "Gamma (9)"
      .Cells(14, 13) = "Vega (10)"
      
      .Cells(15, 1) = r_str_Descri(0)
      .Cells(16, 1) = r_str_Descri(1)
      .Cells(17, 1) = r_str_Descri(2)
      .Cells(18, 1) = r_str_Descri(3)
      .Cells(19, 1) = r_str_Descri(4)
      .Cells(20, 1) = r_str_Descri(5)
      .Cells(21, 1) = r_str_Descri(6)
      .Cells(22, 1) = r_str_Descri(7)
      .Cells(23, 1) = r_str_Descri(8)
      .Cells(24, 1) = r_str_Descri(9)
      .Cells(25, 1) = r_str_Descri(10)
      .Cells(26, 1) = r_str_Descri(11)
      .Cells(27, 1) = r_str_Descri(12)
      .Cells(28, 1) = r_str_Descri(13)
     
      .Range(.Cells(2, 1), .Cells(2, 13)).Merge
      .Range(.Cells(4, 1), .Cells(4, 13)).Merge
      .Range(.Cells(5, 1), .Cells(5, 13)).Merge
      .Range(.Cells(6, 1), .Cells(6, 13)).Merge
      .Range(.Cells(7, 1), .Cells(7, 13)).Merge
      .Range(.Cells(8, 1), .Cells(8, 13)).Merge
      
      .Range(.Cells(13, 1), .Cells(14, 1)).Merge
      .Range(.Cells(13, 2), .Cells(14, 2)).Merge
      .Range(.Cells(13, 3), .Cells(14, 3)).Merge
      .Range(.Cells(13, 4), .Cells(14, 4)).Merge
      .Range(.Cells(13, 5), .Cells(13, 6)).Merge
      .Range(.Cells(13, 7), .Cells(13, 8)).Merge
      .Range(.Cells(13, 9), .Cells(14, 9)).Merge
      .Range(.Cells(13, 10), .Cells(14, 10)).Merge
      .Range(.Cells(13, 11), .Cells(14, 11)).Merge
      .Range(.Cells(13, 12), .Cells(13, 13)).Merge
            
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(13, 1), .Cells(14, 13)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
      .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(15, 1), .Cells(28, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(27, 1), .Cells(27, 13)).Borders(xlEdgeTop).LineStyle = xlDouble
      .Range(.Cells(27, 1), .Cells(27, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(28, 1), .Cells(28, 13)).Font.Bold = True
            
      .Range(.Cells(13, 1), .Cells(14, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(13, 1), .Cells(14, 13)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(13, 1), .Cells(14, 13)).WrapText = True
      .Range(.Cells(13, 1), .Cells(14, 13)).Font.Bold = True
      
      .Cells(30, 1) = "II. REQUERIMIENTO DE PATRIMONIO EFECTIVO"
      .Range(.Cells(30, 1), .Cells(30, 2)).Merge
      .Range(.Cells(30, 1), .Cells(30, 1)).Font.Bold = True
      
      .Cells(31, 1).RowHeight = 40
      .Cells(31, 1) = "Concepto"
      .Cells(31, 3) = "Importe Base"
      .Cells(31, 4) = "Factor"
      .Cells(31, 5) = "Requerimiento de Patrimonio Efectivo"
      
      .Range(.Cells(31, 1), .Cells(31, 2)).Merge
      .Range(.Cells(36, 1), .Cells(36, 2)).Merge
      .Range(.Cells(37, 1), .Cells(37, 2)).Merge
      
      .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(31, 1), .Cells(31, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(31, 1), .Cells(31, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(31, 1), .Cells(31, 5)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(31, 1), .Cells(31, 5)).WrapText = True
      .Range(.Cells(31, 1), .Cells(31, 5)).Font.Bold = True
      
      For r_int_Contad = 1 To 5 Step 1
         If r_int_Contad = 1 Then
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            '.Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            '.Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(36, r_int_Contad), .Cells(37, r_int_Contad + 1)).Font.Bold = True
            
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            
         ElseIf r_int_Contad = 2 Then
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).HorizontalAlignment = xlHAlignCenter
            
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(39, r_int_Contad), .Cells(40, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous

         ElseIf r_int_Contad = 4 Then
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            '.Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(36, r_int_Contad + 1), .Cells(37, r_int_Contad + 1)).Font.Bold = True
         ElseIf r_int_Contad <> 2 Then
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(32, r_int_Contad), .Cells(35, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         End If
      Next
            
      .Cells(32, 1) = r_str_Descri(14)
      .Cells(32, 2) = "A"
      .Cells(33, 1) = r_str_Descri(15)
      .Cells(33, 2) = "B"
      .Cells(34, 1) = r_str_Descri(16)
      .Cells(34, 2) = "C"
      .Cells(35, 1) = r_str_Descri(17)
      .Cells(35, 2) = "D"
      .Cells(36, 1) = r_str_Descri(18)
      .Cells(37, 1) = r_str_Descri(19) & " /1"
      
      .Range(.Cells(32, 4), .Cells(35, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(32, 4), .Cells(35, 4)).NumberFormat = "@"
      
      .Cells(32, 4) = "9.5%"
      .Cells(33, 4) = "k%"
      .Cells(34, 4) = "100.0%"
      .Cells(35, 4) = "100.0%"
      
      
      .Cells(39, 1) = r_str_Descri(20)
      .Cells(40, 1) = r_str_Descri(21)
      
      .Cells(39, 2) = r_dbl_LimGlo
      .Cells(40, 2) = r_dbl_FacAju
      
      .Range(.Cells(43, 1), .Cells(43, 13)).Merge
      .Cells(43, 1).RowHeight = 30
      .Cells(43, 1).WrapText = True
      
      .Cells(43, 1) = "1. APR por riesgo de mercado se halla multiplicando el requerimiento de patrimonio efectivo por riesgo de mercado por la inversa del límite global que establece la Ley General en el artículo " & _
                        "199º y la Vigésima Cuarta Disposición Transitoria y por el factor de ajuste que se consigna al final del artículo 6º del Reglamento para el Requerimiento de Patrimonio Efectivo por Riesgo de Mercado."
                                                
      .Range(.Cells(48, 4), .Cells(48, 5)).Merge
      .Range(.Cells(48, 8), .Cells(48, 9)).Merge
      
      .Range(.Cells(49, 4), .Cells(49, 5)).Merge
      .Range(.Cells(49, 8), .Cells(49, 9)).Merge
      
      .Range(.Cells(49, 4), .Cells(49, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(49, 8), .Cells(49, 9)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(49, 4), .Cells(49, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(49, 8), .Cells(49, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(49, 4) = "Gerente General"
      .Cells(49, 8) = "Gerente Unidad de Riesgos"
           
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 42
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 20
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 25
      .Columns("J").ColumnWidth = 15
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 15
            
      .Columns("B:M").NumberFormat = "###,###,##0.00"
                   
   End With
   
   
   r_int_ConAux = 15
   r_int_ConTem = 0
      
   For r_int_Contad = 0 To 167 Step 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_ConTem + 2) = r_dbl_Evalua(r_int_Contad)
      r_int_ConTem = r_int_ConTem + 1
      
      If (r_int_Contad + 1) Mod 12 = 0 Then
         r_int_ConAux = r_int_ConAux + 1
         r_int_ConTem = 0
      End If
            
   Next
   
   
   r_int_ConTem = 168
      
   For r_int_Contad = 32 To 35 Step 1
      r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConTem)
      r_obj_Excel.ActiveSheet.Cells(r_int_Contad, 5) = r_dbl_Evalua(r_int_ConTem + 2)
      r_int_ConTem = r_int_ConTem + 12
            
   Next
   
   r_obj_Excel.ActiveSheet.Cells(36, 5) = r_dbl_Evalua(218)
   r_obj_Excel.ActiveSheet.Cells(37, 5) = r_dbl_Evalua(230)
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
            
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenRpt1()

   Call fs_GenDat
   Call fs_GeneDB
           
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
    
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".HIS_PATEFE"
    
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_19.RPT"
           
   crp_Imprim.SelectionFormula = "{HIS_PATEFE.PATEFE_PERMES} = " & r_int_PerMes & " AND {HIS_PATEFE.PATEFE_PERANO} = " & r_int_PerAno & " AND {HIS_PATEFE.PATEFE_USUCRE} = '" & modgen_g_str_CodUsu & "' "
         
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   Screen.MousePointer = 0
  
End Sub


Private Sub fs_GenDat_DB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Descri

   r_str_Descri(0) = "Dólar Americano"
   r_str_Descri(1) = "Euro"
   r_str_Descri(2) = "Yen Japonés"
   r_str_Descri(3) = "Libra Esterlina"
   r_str_Descri(4) = "Dólar Canadienes"
   r_str_Descri(5) = "Franco Suizo"
   r_str_Descri(6) = "Corona Danesa"
   r_str_Descri(7) = "Corona Sueca"
   r_str_Descri(8) = "Corona Sueca"
   r_str_Descri(9) = "Real"
   r_str_Descri(10) = "Boliviano"
   r_str_Descri(11) = "Otras Divisas"
   r_str_Descri(12) = "Oro"
   r_str_Descri(13) = "TOTAL ME"
   r_str_Descri(14) = "Mayor Posición Global Agregada en Divisas de Oro"
   r_str_Descri(15) = "Posición Global de Oro"
   r_str_Descri(16) = "Impacto Gamma de las Divisas"
   r_str_Descri(17) = "Impacto Vega de las Divisas"
   r_str_Descri(18) = "Total"
   r_str_Descri(19) = "APR por Riesgo de Mercado - Riesgo Cambiario"
   r_str_Descri(20) = "Inversa del Límite Global"
   r_str_Descri(21) = "Factor de Ajuste"
      
   Erase r_dbl_Evalua()
   
   For r_int_Contad = 1 To 2 Step 1
   
      g_str_Parame = "SELECT * FROM HIS_REPAEF WHERE "
      g_str_Parame = g_str_Parame & "REPAEF_PERMES = " & r_int_PerMes & " AND "
      g_str_Parame = g_str_Parame & "REPAEF_PERANO = " & r_int_PerAno & " AND "
      g_str_Parame = g_str_Parame & "REPAEF_MONEDA = " & r_int_Contad & " "
      g_str_Parame = g_str_Parame & "ORDER BY REPAEF_NUMITE ASC "
        
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
       End If
                   
       If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       
          g_rst_Princi.MoveFirst
          
          If r_int_Contad = 1 Then
            r_int_ConAux = -1
          Else
            r_int_ConAux = 299
          End If
                
          Do While Not g_rst_Princi.EOF
                         
             r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!REPAEF_ACTIVO
             r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!REPAEF_PASIVO
             r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!REPAEF_POCABA
             r_dbl_Evalua(r_int_ConAux + 4) = g_rst_Princi!REPAEF_PFMELA
             r_dbl_Evalua(r_int_ConAux + 5) = g_rst_Princi!REPAEF_PFMECO
             r_dbl_Evalua(r_int_ConAux + 6) = g_rst_Princi!REPAEF_PDMELA
             r_dbl_Evalua(r_int_ConAux + 7) = g_rst_Princi!REPAEF_PDMECO
             r_dbl_Evalua(r_int_ConAux + 8) = g_rst_Princi!REPAEF_PNDEME
             r_dbl_Evalua(r_int_ConAux + 9) = g_rst_Princi!REPAEF_VDNEME
             r_dbl_Evalua(r_int_ConAux + 10) = g_rst_Princi!REPAEF_POGLME
             r_dbl_Evalua(r_int_ConAux + 11) = g_rst_Princi!REPAEF_SEOPGA
             r_dbl_Evalua(r_int_ConAux + 12) = g_rst_Princi!REPAEF_SEOPVE
             
             r_int_ConAux = r_int_ConAux + 12
             
             g_rst_Princi.MoveNext
             DoEvents
          Loop
       
       End If
      
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
    
    Next
     
         
   'r_dbl_Evalua(2) = r_dbl_Evalua(0) - r_dbl_Evalua(1)
   'r_dbl_Evalua(7) = r_dbl_Evalua(6) + r_dbl_Evalua(5) + r_dbl_Evalua(4) + r_dbl_Evalua(3)
   'r_dbl_Evalua(9) = r_dbl_Evalua(8) + r_dbl_Evalua(7) + r_dbl_Evalua(2)
   
   'For r_int_Contad = 0 To 11 Step 1
   '   r_dbl_MulUso = 0
   '   For r_int_ConTem = (0 + r_int_Contad) To (144 + r_int_Contad) Step 12
   '      r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_ConTem)
   '   Next
      
   '   r_dbl_Evalua(156 + r_int_Contad) = r_dbl_MulUso
      
   'Next
   
   'r_dbl_Evalua(168) = r_dbl_Evalua(9)
   'r_dbl_Evalua(169) = 9.5
   'r_dbl_Evalua(170) = Format(r_dbl_Evalua(168) * 0.095, "###########0.00")
   'r_dbl_Evalua(181) = 9.5
   'r_dbl_Evalua(193) = 100
   'r_dbl_Evalua(205) = 100
   
   'r_dbl_MulUso = 0
   
   'For r_int_Contad = 170 To 206 Step 12
   '   r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_Contad)
      
   'Next
   
   'r_dbl_Evalua(218) = r_dbl_MulUso
   'r_dbl_Evalua(230) = r_dbl_Evalua(218) * r_dbl_LimGlo * r_dbl_FacAju
   'r_dbl_Evalua(240) = r_dbl_LimGlo
   'r_dbl_Evalua(252) = r_dbl_FacAju
         
End Sub






