VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   2325
   ClientLeft      =   7845
   ClientTop       =   5400
   ClientWidth     =   6195
   Icon            =   "GesCtb_frm_703.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6225
      _Version        =   65536
      _ExtentX        =   10980
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
         TabIndex        =   7
         Top             =   30
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
            Caption         =   "Anexo Nº 5-A"
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
            Width           =   5325
            _Version        =   65536
            _ExtentX        =   9393
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0102-03 Informe de Clasificación de Deudores y Provisiones"
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
            Picture         =   "GesCtb_frm_703.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
            Picture         =   "GesCtb_frm_703.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5520
            Picture         =   "GesCtb_frm_703.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_703.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_703.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_703.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3090
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
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim r_dbl_EvaPar(100)    As Double
   Dim r_dbl_Evalua(1500)  As Double
   Dim r_str_Denomi(500)   As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_TemAux        As Integer
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String
   Dim r_dbl_MonNac        As Double
   Dim r_dbl_MonExt        As Double
   Dim r_str_FecRpt        As String
   Dim r_dbl_MulUso        As Double
   Dim r_str_Cadena        As String
   Dim r_str_NomRes        As String
   Dim r_int_ConGen        As Integer
   Dim r_str_ParAux        As String
   Dim r_int_Cantid       As Integer
   Dim r_int_FlgRpr       As Integer
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_PROPRO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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
   
   r_int_Cantid = modsec_gf_CanReg("HIS_PROPRO", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
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

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer        As Integer
   Dim r_str_Cadena        As String
   Dim r_str_FecRpt        As String
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
         
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
     
      .Range(.Cells(5, 3), .Cells(5, 3)).HorizontalAlignment = xlHAlignLeft
      .Cells(3, 1) = "Anexo Nº 5-A"
      .Cells(5, 3) = "Empresa: MiCasita S.A."
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Range(.Cells(5, 1), .Cells(5, 4)).Merge
      
      .Range(.Cells(3, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Underline = xlUnderlineStyleSingle
       
      .Range(.Cells(6, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(6, 1), .Cells(8, 1)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
      
      .Range(.Cells(6, 1), .Cells(6, 4)).Merge
      .Range(.Cells(7, 1), .Cells(7, 4)).Merge
      .Range(.Cells(8, 1), .Cells(8, 4)).Merge
      
      .Cells(8, 1) = "Resumen de Provisiones Procíclicas"
      .Cells(6, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Cells(7, 1) = "(En Miles de Nuevos Soles)"
      
      .Cells(12, 1) = ""
      .Cells(12, 2) = "Endeudamiento en Categoría Normal 1/"
      .Cells(12, 3) = "Provisión Procíclica Constituida"
      .Cells(12, 4) = "Provisión Procíclica Requerida 2/"
            
      .Range(.Cells(12, 1), .Cells(12, 4)).Font.Bold = True
      .Range(.Cells(12, 1), .Cells(12, 4)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(12, 1), .Cells(12, 4)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(12, 1), .Cells(12, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).WrapText = True
      
      For r_int_Contad = 1 To 4 Step 1
         .Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         '.Range(.Cells(13, r_int_Contad), .Cells(25, r_int_Contad)).RowHeight = 30
      Next
      
      r_int_ConAux = -1
      
      For r_int_Contad = 13 To 25 Step 1
         .Cells(r_int_Contad, 1) = r_str_Denomi(r_int_Contad - 13)
         .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux + 1)
         .Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConAux + 2)
         .Cells(r_int_Contad, 4) = r_dbl_Evalua(r_int_ConAux + 3)
         r_int_ConAux = r_int_ConAux + 3
      Next
      
      'For r_int_ConTem = 0 To 2 Step 1
      '   r_dbl_MulUso = 0
      '   For r_int_TemAux = r_int_ConTem To 54 + r_int_ConTem Step 3
      '      r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_TemAux)
      '   Next
      '   .Cells(32, r_int_ConTem + 2) = r_dbl_MulUso
         
      'Next
      
      'For r_int_Contad = 2 To 4 Step 1
      '   .Cells(32, r_int_Contad) = r_dbl_Evalua(0)
      'Next
      
      .Cells(25, 1) = "Total"
      
      .Cells(27, 1) = "Indicar si se encuentra en periodo de adecuación"
      .Cells(28, 1) = "Si esta en periodo de adecuación, indicar mes"
      
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(27, 2), .Cells(28, 2)).NumberFormat = "@"
      .Range(.Cells(27, 2), .Cells(28, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(27, 2) = r_dbl_Evalua(r_int_ConAux + 1)
      r_int_ConAux = r_int_ConAux + 3
      .Cells(28, 2) = r_dbl_Evalua(r_int_ConAux + 1)
      
      .Cells(27, 3) = "1 = si, 0 = no"
      .Cells(28, 3) = "0 = primer mes, 1 = segundo y tercer mes, 2 = cuarto y quinto mes"
      .Range(.Cells(28, 3), .Cells(28, 4)).Merge
      
      .Range(.Cells(12, 1), .Cells(12, 1)).HorizontalAlignment = xlHAlignCenter
   
      .Range(.Cells(1, 1), .Cells(35, 5)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(35, 5)).Font.Size = 8
                  
      .Range(.Cells(13, 2), .Cells(25, 4)).NumberFormat = "###,###,##0.00"
      
      .Columns("A").ColumnWidth = 60
      '.Columns("A").HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").ColumnWidth = 23
      '.Columns("B").NumberFormat = "###,###,##0.00"
            
      .Columns("C").ColumnWidth = 23
      '.Columns("C").NumberFormat = "###,###,##0.00"
            
      .Columns("D").ColumnWidth = 23
      '.Columns("D").NumberFormat = "###,###,##0.00"
      
      .Cells(32, 1) = "_________________________"
      .Cells(32, 2) = "___________________________"
      .Cells(32, 4) = "_________________________"
      
      .Cells(33, 1) = "Sr. Roberto Baba Yamamoto"
      .Cells(33, 2) = "Srta. Rossana Mesa Bustamente"
      .Cells(33, 4) = "Sr. Javier Delgado Blanco"
           
      .Cells(34, 1) = "Gerente General"
      .Cells(34, 2) = "Contador General"
      .Cells(34, 4) = "Unidad de Riesgos"
            
      .Range(.Cells(32, 1), .Cells(35, 4)).HorizontalAlignment = xlHAlignCenter
   
   End With
   

   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
      
End Sub

Private Sub fs_GenDat()

   Erase r_str_Denomi
   Erase r_dbl_Evalua
   Erase r_dbl_EvaPar

   r_str_Denomi(0) = "Corporativo"
   r_str_Denomi(1) = "Corporativo con Garantía Autoliquidable"
   r_str_Denomi(2) = "Grandes Empresas"
   r_str_Denomi(3) = "Grandes Empresas con Garantías Autoliquidables"
   r_str_Denomi(4) = "Medianas Empresas"
   r_str_Denomi(5) = "Pequeñas Empresas"
   r_str_Denomi(6) = "Microempresas"
   r_str_Denomi(7) = "Consumo Revolvente"
   r_str_Denomi(8) = "Consumo no Revolvente"
   r_str_Denomi(9) = "Consumo no Revolvente bajo convenios elegibles"
   r_str_Denomi(10) = "Hipotecario para Vivienda"
   r_str_Denomi(11) = "Hipotecario para Vivienda con Garantía Autoliquidable"
   r_str_Denomi(12) = "Total"
   r_str_Denomi(13) = "Indicador de periodo de adecuación"
   r_str_Denomi(14) = "Mes de periodo de adecuación"
   
   
   g_str_Parame = "SELECT * FROM CTB_MNTIND WHERE MNTIND_CODIGO = '001'"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
         
   Do While Not g_rst_Princi.EOF
      
      r_dbl_Evalua(39) = Trim(g_rst_Princi!MNTIND_CODIND)
      r_dbl_Evalua(42) = IIf(IsNull(Trim(g_rst_Princi!MNTIND_CODPER)) = True, 0, Trim(g_rst_Princi!MNTIND_CODPER))
         
         
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
 
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
      
   g_str_Parame = "SELECT * FROM HIS_CLADEU WHERE "
   g_str_Parame = g_str_Parame & "CLADEU_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "CLADEU_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "(CLADEU_SUBCAB <> 8 AND CLADEU_SUBCAB <> 17 AND CLADEU_SUBCAB <> 26 AND CLADEU_SUBCAB <> 35 AND  "
   g_str_Parame = g_str_Parame & "CLADEU_SUBCAB <> 44 AND CLADEU_SUBCAB <> 53 AND CLADEU_SUBCAB <> 61 AND CLADEU_SUBCAB <> 70 AND "
   g_str_Parame = g_str_Parame & "CLADEU_SUBCAB <> 80 AND CLADEU_SUBCAB <> 89 AND CLADEU_SUBCAB <> 98 AND CLADEU_SUBCAB <> 107) "
   g_str_Parame = g_str_Parame & "ORDER BY CLADEU_SUBCAB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_int_Contad = 0
   
   Do While Not g_rst_Princi.EOF
   
      r_dbl_EvaPar(r_int_Contad) = Trim(g_rst_Princi!CLADEU_MTONOR)
      r_int_Contad = r_int_Contad + 1
   
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
 
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   
   
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_int_Contad = 0
   
   Do While Not g_rst_Princi.EOF
      
      If g_rst_Princi!HIPCIE_NUECRE = 13 Then
         r_dbl_Evalua(31) = r_dbl_Evalua(31) + IIf(IsNull(Trim(g_rst_Princi!HIPCIE_PRVCIC)), 0, IIf(g_rst_Princi!HIPCIE_TIPMON = 1, g_rst_Princi!HIPCIE_PRVCIC, g_rst_Princi!HIPCIE_PRVCIC * g_rst_Princi!HIPCIE_TIPCAM))
         
      End If
      
      If g_rst_Princi!HIPCIE_FECDES >= 20100701 Then
      
         If g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "010" Then
            If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            
               If Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 0 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * 0, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 1 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * 0.0015, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 2 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * 0.003, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 0 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * 0.004, "###,###,##0.00")
                  
               End If
            
            Else
                        
               If Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 0 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM * 0, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 1 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM * 0.0015, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 2 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM * 0.003, "###,###,##0.00")
                  
               ElseIf Int(r_dbl_Evalua(39)) = 0 Then
                  r_dbl_Evalua(1) = r_dbl_Evalua(1) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM * 0.004, "###,###,##0.00")
                  
               End If
            
            
            End If
         End If
      
      End If
         
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
 
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   g_str_Parame = "SELECT * FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_int_Contad = 0
   
   Do While Not g_rst_Princi.EOF
      
      If g_rst_Princi!COMCIE_NUECRE = 6 Then
         r_dbl_Evalua(1) = r_dbl_Evalua(1) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 7 Then
         r_dbl_Evalua(7) = r_dbl_Evalua(7) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 8 Then
         r_dbl_Evalua(13) = r_dbl_Evalua(13) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 9 Then
         r_dbl_Evalua(16) = r_dbl_Evalua(16) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 10 Then
         r_dbl_Evalua(19) = r_dbl_Evalua(19) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 11 Then
         r_dbl_Evalua(22) = r_dbl_Evalua(22) + Trim(g_rst_Princi!COMCIE_PRVCIC)
      ElseIf g_rst_Princi!COMCIE_NUECRE = 12 Then
         r_dbl_Evalua(25) = r_dbl_Evalua(25) + Trim(g_rst_Princi!COMCIE_PRVCIC)
                  
      End If
         
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
 
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   r_dbl_Evalua(0) = r_dbl_EvaPar(8) - r_dbl_EvaPar(40) - r_dbl_EvaPar(24) + r_dbl_EvaPar(32)
   r_dbl_Evalua(3) = r_dbl_EvaPar(40)
   r_dbl_Evalua(6) = r_dbl_EvaPar(9) - r_dbl_EvaPar(41) - r_dbl_EvaPar(25) + r_dbl_EvaPar(33)
   r_dbl_Evalua(9) = r_dbl_EvaPar(41)
   r_dbl_Evalua(12) = r_dbl_EvaPar(10) - r_dbl_EvaPar(42) - r_dbl_EvaPar(26)
   r_dbl_Evalua(15) = r_dbl_EvaPar(11) - r_dbl_EvaPar(43) - r_dbl_EvaPar(27)
   r_dbl_Evalua(18) = r_dbl_EvaPar(12) - r_dbl_EvaPar(44) - r_dbl_EvaPar(28)
   r_dbl_Evalua(21) = r_dbl_EvaPar(13) - r_dbl_EvaPar(45) - r_dbl_EvaPar(29)
   r_dbl_Evalua(24) = r_dbl_EvaPar(14) - r_dbl_EvaPar(46) - r_dbl_EvaPar(48) - r_dbl_EvaPar(30)
   r_dbl_Evalua(27) = r_dbl_EvaPar(48)
   r_dbl_Evalua(30) = r_dbl_EvaPar(15) - r_dbl_EvaPar(47) - r_dbl_EvaPar(63) - r_dbl_EvaPar(31)
   r_dbl_Evalua(33) = r_dbl_EvaPar(47)
     
   If Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 0 Then
      r_dbl_Evalua(2) = Format(r_dbl_Evalua(0) * 0, "###,###,##0.00")
      r_dbl_Evalua(5) = Format(r_dbl_Evalua(3) * 0, "###,###,##0.00")
      r_dbl_Evalua(8) = Format(r_dbl_Evalua(6) * 0, "###,###,##0.00")
      r_dbl_Evalua(11) = Format(r_dbl_Evalua(9) * 0, "###,###,##0.00")
      r_dbl_Evalua(14) = Format(r_dbl_Evalua(12) * 0, "###,###,##0.00")
      r_dbl_Evalua(17) = Format(r_dbl_Evalua(15) * 0, "###,###,##0.00")
      r_dbl_Evalua(20) = Format(r_dbl_Evalua(18) * 0, "###,###,##0.00")
      r_dbl_Evalua(23) = Format(r_dbl_Evalua(21) * 0, "###,###,##0.00")
      r_dbl_Evalua(26) = Format(r_dbl_Evalua(24) * 0, "###,###,##0.00")
      r_dbl_Evalua(29) = Format(r_dbl_Evalua(27) * 0, "###,###,##0.00")
      r_dbl_Evalua(32) = Format(r_dbl_Evalua(30) * 0, "###,###,##0.00")
      r_dbl_Evalua(35) = Format(r_dbl_Evalua(33) * 0, "###,###,##0.00")
      
   ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 1 Then
      r_dbl_Evalua(2) = Format(r_dbl_Evalua(0) * 0.0015, "###,###,##0.00")
      r_dbl_Evalua(5) = Format(r_dbl_Evalua(3) * 0.0015, "###,###,##0.00")
      r_dbl_Evalua(8) = Format(r_dbl_Evalua(6) * 0.0015, "###,###,##0.00")
      r_dbl_Evalua(11) = Format(r_dbl_Evalua(9) * 0.0015, "###,###,##0.00")
      r_dbl_Evalua(14) = Format(r_dbl_Evalua(12) * 0.001, "###,###,##0.00")
      r_dbl_Evalua(17) = Format(r_dbl_Evalua(15) * 0.002, "###,###,##0.00")
      r_dbl_Evalua(20) = Format(r_dbl_Evalua(18) * 0.002, "###,###,##0.00")
      r_dbl_Evalua(23) = Format(r_dbl_Evalua(21) * 0.005, "###,###,##0.00")
      r_dbl_Evalua(26) = Format(r_dbl_Evalua(24) * 0.004, "###,###,##0.00")
      r_dbl_Evalua(29) = Format(r_dbl_Evalua(27) * 0.0025, "###,###,##0.00")
      r_dbl_Evalua(32) = Format(r_dbl_Evalua(30) * 0.0015, "###,###,##0.00")
      r_dbl_Evalua(35) = Format(r_dbl_Evalua(33) * 0.0015, "###,###,##0.00")
      
   ElseIf Int(r_dbl_Evalua(39)) = 1 And Int(r_dbl_Evalua(42)) = 2 Then
      r_dbl_Evalua(2) = Format(r_dbl_Evalua(0) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(5) = Format(r_dbl_Evalua(3) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(8) = Format(r_dbl_Evalua(6) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(11) = Format(r_dbl_Evalua(9) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(14) = Format(r_dbl_Evalua(12) * 0.002, "###,###,##0.00")
      r_dbl_Evalua(17) = Format(r_dbl_Evalua(15) * 0.004, "###,###,##0.00")
      r_dbl_Evalua(20) = Format(r_dbl_Evalua(18) * 0.004, "###,###,##0.00")
      r_dbl_Evalua(23) = Format(r_dbl_Evalua(21) * 0.01, "###,###,##0.00")
      r_dbl_Evalua(26) = Format(r_dbl_Evalua(24) * 0.007, "###,###,##0.00")
      r_dbl_Evalua(29) = Format(r_dbl_Evalua(27) * 0.025, "###,###,##0.00")
      r_dbl_Evalua(32) = Format(r_dbl_Evalua(30) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(35) = Format(r_dbl_Evalua(33) * 0.003, "###,###,##0.00")
      
   ElseIf Int(r_dbl_Evalua(39)) = 0 Then
      r_dbl_Evalua(2) = Format(r_dbl_Evalua(0) * 0.004, "###,###,##0.00")
      r_dbl_Evalua(5) = Format(r_dbl_Evalua(3) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(8) = Format(r_dbl_Evalua(6) * 0.0045, "###,###,##0.00")
      r_dbl_Evalua(11) = Format(r_dbl_Evalua(9) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(14) = Format(r_dbl_Evalua(12) * 0.003, "###,###,##0.00")
      r_dbl_Evalua(17) = Format(r_dbl_Evalua(15) * 0.005, "###,###,##0.00")
      r_dbl_Evalua(20) = Format(r_dbl_Evalua(18) * 0.005, "###,###,##0.00")
      r_dbl_Evalua(23) = Format(r_dbl_Evalua(21) * 0.015, "###,###,##0.00")
      r_dbl_Evalua(26) = Format(r_dbl_Evalua(24) * 0.01, "###,###,##0.00")
      r_dbl_Evalua(29) = Format(r_dbl_Evalua(27) * 0.025, "###,###,##0.00")
      r_dbl_Evalua(32) = Format(r_dbl_Evalua(30) * 0.004, "###,###,##0.00")
      r_dbl_Evalua(35) = Format(r_dbl_Evalua(33) * 0.003, "###,###,##0.00")
      
   End If
      
   
   For r_int_ConTem = 0 To 2 Step 1
      r_dbl_MulUso = 0
      For r_int_TemAux = r_int_ConTem To 33 + r_int_ConTem Step 3
         r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_TemAux)
      Next
      r_dbl_Evalua(r_int_ConTem + 36) = r_dbl_MulUso
      
   Next
      
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
      
   r_str_NomRes = "C:\03" & Right(ipp_PerAno.Text, 2) & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".105"
   
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
   
   Print #r_int_NumRes, Format(105, "0000") & Format(3, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
      
   r_int_ConAux = 0
   
   For r_int_Contad = 100 To 1500 Step 100
      r_str_Cadena = ""
      If r_int_Contad < 1300 Then
         For r_int_ConTem = 0 To 2 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
            r_int_ConAux = r_int_ConAux + 1
         Next
         
      ElseIf r_int_Contad = 1300 Then
         r_str_Cadena = ""
         r_int_ConAux = r_int_ConAux + 21
         For r_int_ConTem = 0 To 2 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
            r_int_ConAux = r_int_ConAux + 1
         Next
         
         'For r_int_ConTem = 0 To 2 Step 1
         '   r_dbl_MulUso = 0
         '   For r_int_TemAux = r_int_ConTem To 33 + r_int_ConTem Step 3
         '      r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_TemAux)
         '   Next
         '   r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_MulUso, "########0.00"), 1, "0", 18)
            
         'Next
         
         'r_int_ConAux = r_int_ConAux + 24
      
      ElseIf r_int_Contad > 1300 Then
         r_str_Cadena = ""
         For r_int_ConTem = 0 To 2 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
            r_int_ConAux = r_int_ConAux + 1
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

   r_int_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   r_int_PerAno = ipp_PerAno.Text

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardarán los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM HIS_PROPRO WHERE "
   g_str_Parame = g_str_Parame & "PROPRO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "PROPRO_PERANO = " & r_int_PerAno & "  "
   'g_str_Parame = g_str_Parame & "PROPRO_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   r_int_ConAux = 0
   
   For r_int_Contad = 0 To 14 Step 1
  
      r_str_Cadena = "USP_HIS_PROPRO ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_??', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      r_str_Cadena = r_str_Cadena & "'" & r_str_Denomi(r_int_Contad) & "', "
      
      For r_int_ConTem = 0 To 2 Step 1
         If r_int_ConTem = 2 Then
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux)
         Else
            r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux) & ", "
         End If
         r_int_ConAux = r_int_ConAux + 1
      Next
      
      r_str_Cadena = r_str_Cadena & ")"
          
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_PROPRO.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If

   Next
   
End Sub

Private Sub fs_GenRpt()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer        As Integer
   Dim r_str_Cadena        As String
   Dim r_str_FecRpt        As String
   
   Call fs_GenDat
         
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
     
      .Range(.Cells(5, 3), .Cells(5, 3)).HorizontalAlignment = xlHAlignLeft
      .Cells(3, 1) = "Anexo Nº 5-A"
      .Cells(5, 3) = "Empresa: MiCasita S.A."
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Range(.Cells(5, 1), .Cells(5, 4)).Merge
      
      .Range(.Cells(3, 1), .Cells(3, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 1)).Font.Underline = xlUnderlineStyleSingle
       
      .Range(.Cells(6, 1), .Cells(8, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(6, 1), .Cells(8, 1)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(8, 1)).Font.Underline = xlUnderlineStyleSingle
      
      .Range(.Cells(6, 1), .Cells(6, 4)).Merge
      .Range(.Cells(7, 1), .Cells(7, 4)).Merge
      .Range(.Cells(8, 1), .Cells(8, 4)).Merge
      
      .Cells(8, 1) = "Resumen de Provisiones Procíclicas"
      .Cells(6, 1) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Cells(7, 1) = "(En Miles de Nuevos Soles)"
      
      .Cells(12, 1) = ""
      .Cells(12, 2) = "Endeudamiento en Categoría Normal 1/"
      .Cells(12, 3) = "Provisión Procíclica Constituida"
      .Cells(12, 4) = "Provisión Procíclica Requerida 2/"
            
      .Range(.Cells(12, 1), .Cells(12, 4)).Font.Bold = True
      .Range(.Cells(12, 1), .Cells(12, 4)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(12, 1), .Cells(12, 4)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(12, 1), .Cells(12, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(12, 1), .Cells(12, 4)).WrapText = True
      
      For r_int_Contad = 1 To 4 Step 1
         .Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         .Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         '.Range(.Cells(13, r_int_Contad), .Cells(32, r_int_Contad)).RowHeight = 30
      Next
      
      r_int_ConAux = 0
      
      For r_int_Contad = 13 To 31 Step 1
         .Cells(r_int_Contad, 1) = r_str_Denomi(r_int_Contad - 13)
         .Cells(r_int_Contad, 2) = r_dbl_Evalua(r_int_ConAux)
         .Cells(r_int_Contad, 3) = r_dbl_Evalua(r_int_ConAux)
         .Cells(r_int_Contad, 4) = r_dbl_Evalua(r_int_ConAux)
         r_int_ConAux = r_int_ConAux + 3
      Next
      
      For r_int_Contad = 2 To 4 Step 1
         .Cells(32, r_int_Contad) = r_dbl_Evalua(0)
      Next
      
      .Cells(32, 1) = "Total"
      
      .Cells(34, 1) = "Indicar si se encuentra en periodo de adecuación"
      .Cells(35, 1) = "Si esta en periodo de adecuación, indicar mes"
      
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(35, 2)).NumberFormat = "@"
      .Range(.Cells(34, 2), .Cells(35, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(34, 2) = 0
      .Cells(35, 2) = 0
      
      .Cells(34, 3) = "1 = si, 0 = no"
      .Cells(35, 3) = "0 = primer mes, 1 = segundo y tercer mes, 2 = cuarto y quinto mes"
      .Range(.Cells(35, 3), .Cells(35, 4)).Merge
      
      .Range(.Cells(12, 1), .Cells(12, 1)).HorizontalAlignment = xlHAlignCenter
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(56, 99)).Font.Size = 10
                  
      .Range(.Cells(13, 2), .Cells(32, 4)).NumberFormat = "###,###,##0.00"
      
      .Columns("A").ColumnWidth = 50
      '.Columns("A").HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").ColumnWidth = 30
      '.Columns("B").NumberFormat = "###,###,##0.00"
            
      .Columns("C").ColumnWidth = 30
      '.Columns("C").NumberFormat = "###,###,##0.00"
            
      .Columns("D").ColumnWidth = 30
      '.Columns("D").NumberFormat = "###,###,##0.00"
   
   End With
   
   Call fs_GeneDB
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing


End Sub


Private Sub fs_GenDat_DB()

   Erase r_str_Denomi
   Erase r_dbl_Evalua

   r_str_Denomi(0) = "Corporativo"
   r_str_Denomi(1) = "Corporativo con Garantía Autoliquidable"
   r_str_Denomi(2) = "Grandes Empresas"
   r_str_Denomi(3) = "Grandes Empresas con Garantías Autoliquidables"
   r_str_Denomi(4) = "Medianas Empresas"
   r_str_Denomi(5) = "Pequeñas Empresas"
   r_str_Denomi(6) = "Microempresas"
   r_str_Denomi(7) = "Consumo Revolvente"
   r_str_Denomi(8) = "Consumo no Revolvente"
   r_str_Denomi(9) = "Consumo no Revolvente bajo convenios elegibles"
   r_str_Denomi(10) = "Hipotecario para Vivienda"
   r_str_Denomi(11) = "Hipotecario para Vivienda con Garantía Autoliquidable"
   r_str_Denomi(12) = "Grandes Empresas - Sustitución de Contraparte"
   r_str_Denomi(13) = "Medianas Empresas - Sustitución de Contraparte"
   r_str_Denomi(14) = "Pequeñas Empresas - Sustitución de Contraparte"
   r_str_Denomi(15) = "Microempresas - Sustitución de Contraparte"
   r_str_Denomi(16) = "Consumo Revolvente - Sustitución de Contraparte"
   r_str_Denomi(17) = "Consumo no Revolvente - Sustitución de Contraparte"
   r_str_Denomi(18) = "Hipotecario para Vivienda - Sustitución de Contraparte"
   r_str_Denomi(19) = "Total"
   r_str_Denomi(20) = "Indicador de periodo de adecuación"
   r_str_Denomi(21) = "Mes de periodo de adecuación"
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   g_str_Parame = "SELECT * FROM HIS_PROPRO WHERE "
   g_str_Parame = g_str_Parame & "PROPRO_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "PROPRO_PERANO = " & r_int_PerAno & ""
   g_str_Parame = g_str_Parame & "ORDER BY PROPRO_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_ConTem = -1
                                       
      Do While Not g_rst_Princi.EOF
      
         r_dbl_Evalua(r_int_ConTem + 1) = g_rst_Princi!PROPRO_ENCANO
         r_dbl_Evalua(r_int_ConTem + 2) = g_rst_Princi!PROPRO_PRPRCO
         r_dbl_Evalua(r_int_ConTem + 3) = g_rst_Princi!PROPRO_PRPRRE
         
         r_int_ConTem = r_int_ConTem + 3
                     
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
   End If
         
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   For r_int_ConTem = 0 To 2 Step 1
      r_dbl_MulUso = 0
      For r_int_TemAux = r_int_ConTem To 33 + r_int_ConTem Step 3
         r_dbl_MulUso = r_dbl_MulUso + r_dbl_Evalua(r_int_TemAux)
      Next
      r_dbl_Evalua(r_int_ConTem + 36) = r_dbl_MulUso
      
   Next
   
      
End Sub



