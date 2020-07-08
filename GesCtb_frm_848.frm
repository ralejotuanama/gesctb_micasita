VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_22 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   1425
   ClientTop       =   2175
   ClientWidth     =   14025
   Icon            =   "GesCtb_frm_848.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel5 
      Height          =   10095
      Left            =   0
      TabIndex        =   7
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
         TabIndex        =   8
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
            TabIndex        =   6
            Top             =   60
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   13256
            _Version        =   393216
            Rows            =   50
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            Redraw          =   -1  'True
            FocusRect       =   0
            FillStyle       =   1
            SelectionMode   =   1
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
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   180
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Balance General"
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
            Picture         =   "GesCtb_frm_848.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
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
         Begin VB.CommandButton cmd_ExpExcDet 
            Height          =   585
            Left            =   1245
            Picture         =   "GesCtb_frm_848.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel - Detallado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_848.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   645
            Picture         =   "GesCtb_frm_848.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel - Resumido"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13320
            Picture         =   "GesCtb_frm_848.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   12
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
            TabIndex        =   14
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   450
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer

Private Sub cmd_ExpExcDet_Click()
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
   Call fs_GenExcDet
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

Private Sub cmd_Proces_Click()
Dim r_str_PerMes    As String
Dim r_str_PerAno    As String
Dim r_bol_FlagTA    As Boolean
Dim r_bol_FlagTP    As Boolean
Dim p               As Integer
Dim q               As Integer
Dim r               As Integer
Dim k               As Integer
Dim i               As Integer
Dim anvigente       As Integer
Dim mesvigente      As Integer
   
   r_bol_FlagTA = True
   r_bol_FlagTP = True
    
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
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   grd_LisEEBG.Redraw = False
   Call gs_LimpiaGrid(grd_LisEEBG)
   Call fs_Recorset_nc
    
   'Consulta para obtener el mes y año vigente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT max(PERMES_CODMES) AS MES, max(PERMES_CODANO) AS ANO "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "Error al ejecutar la consulta para obtener año vigente.", vbCritical, modgen_g_str_NomPlt
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      anvigente = g_rst_Genera!ANO
      mesvigente = g_rst_Genera!Mes
   End If
  
   'Prepara SP
   g_str_Parame = "USP_CUR_GEN_EEBG ("
   g_str_Parame = g_str_Parame & 12 & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) - 1 & ", 1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "
   
   'Ejecuta consulta
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEBG.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'ALMACENA RESULTADOS DEL AÑO ANTERIOR, EN UN RECORSET NO CONECTADO
   Dim g_rst_Auxiliar As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TT_EEBG "
   g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "'"
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "'"
   'g_str_Parame = g_str_Parame & " ORDER BY grupo, subgrp, item, indtipo "
   g_str_Parame = g_str_Parame & " ORDER BY orden, subgrp, item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Auxiliar, 3) Then
      MsgBox "Error al ejecutar la consulta para recorset no conectado.", vbCritical, modgen_g_str_NomPlt 'Exit Sub
   End If
   
   If Not (g_rst_Auxiliar.BOF And g_rst_Auxiliar.EOF) Then
      g_rst_Auxiliar.MoveFirst
      Do While Not g_rst_Auxiliar.EOF
         g_rst_GenAux.AddNew
         g_rst_GenAux.Fields(0).Value = g_rst_Auxiliar!GRUPO
         g_rst_GenAux.Fields(1).Value = g_rst_Auxiliar!NOMGRUPO
         g_rst_GenAux.Fields(2).Value = g_rst_Auxiliar!SUBGRP
         g_rst_GenAux.Fields(3).Value = g_rst_Auxiliar!NOMSUBGRP
         g_rst_GenAux.Fields(4).Value = g_rst_Auxiliar!CNTACTBLE
         g_rst_GenAux.Fields(5).Value = g_rst_Auxiliar!NOMCTA
         g_rst_GenAux.Fields(6).Value = g_rst_Auxiliar!MES01
         g_rst_GenAux.Fields(7).Value = g_rst_Auxiliar!MES02
         g_rst_GenAux.Fields(8).Value = g_rst_Auxiliar!MES03
         g_rst_GenAux.Fields(9).Value = g_rst_Auxiliar!MES04
         g_rst_GenAux.Fields(10).Value = g_rst_Auxiliar!MES05
         g_rst_GenAux.Fields(11).Value = g_rst_Auxiliar!MES06
         g_rst_GenAux.Fields(12).Value = g_rst_Auxiliar!MES07
         g_rst_GenAux.Fields(13).Value = g_rst_Auxiliar!MES08
         g_rst_GenAux.Fields(14).Value = g_rst_Auxiliar!MES09
         g_rst_GenAux.Fields(15).Value = g_rst_Auxiliar!MES10
         g_rst_GenAux.Fields(16).Value = g_rst_Auxiliar!MES11
         g_rst_GenAux.Fields(17).Value = g_rst_Auxiliar!MES12
         g_rst_GenAux.Fields(18).Value = g_rst_Auxiliar!ACUMU
         g_rst_GenAux.Fields(19).Value = g_rst_Auxiliar!INDTIPO
         g_rst_GenAux.Fields(20).Value = g_rst_Auxiliar!Item
         g_rst_GenAux.Fields(21).Value = g_rst_Auxiliar!CTATIPO
         
         g_rst_GenAux.Update
         g_rst_Auxiliar.MoveNext
      Loop
   End If
   
   'muestra la información
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.FixedRows = 1

   grd_LisEEBG.Row = 0
   grd_LisEEBG.Col = 3
   grd_LisEEBG.FixedCols = 4
   grd_LisEEBG.BackColorFixed = &H80000005
   grd_LisEEBG.Text = "EJERCICIOS"
   
   If CInt(r_str_PerMes) = 12 Then
      GoTo SALTO1
    
S1:
      grd_LisEEBG.Col = p:        grd_LisEEBG.Text = "FEB-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S2:
      grd_LisEEBG.Col = p + 1:    grd_LisEEBG.Text = "MAR-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S3:
      grd_LisEEBG.Col = p + 2:    grd_LisEEBG.Text = "ABR-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S4:
      grd_LisEEBG.Col = p + 3:    grd_LisEEBG.Text = "MAY-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S5:
      grd_LisEEBG.Col = p + 4:    grd_LisEEBG.Text = "JUN-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S6:
      grd_LisEEBG.Col = p + 5:    grd_LisEEBG.Text = "JUL-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S7:
      grd_LisEEBG.Col = p + 6:    grd_LisEEBG.Text = "AGO-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S8:
      grd_LisEEBG.Col = p + 7:    grd_LisEEBG.Text = "SET-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S9:
      grd_LisEEBG.Col = p + 8:    grd_LisEEBG.Text = "OCT-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S10:
      grd_LisEEBG.Col = p + 9:    grd_LisEEBG.Text = "NOV-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
S11:
      grd_LisEEBG.Col = p + 10:   grd_LisEEBG.Text = "DIC-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
  
   ElseIf CInt(r_str_PerMes) = 11 Then
      p = -6
      GoTo S11
   ElseIf CInt(r_str_PerMes) = 10 Then
      p = -5
      GoTo S10
   ElseIf CInt(r_str_PerMes) = 9 Then
      p = -4
      GoTo S9
   ElseIf CInt(r_str_PerMes) = 8 Then
      p = -3
      GoTo S8
   ElseIf CInt(r_str_PerMes) = 7 Then
      p = -2
      GoTo S7
   ElseIf CInt(r_str_PerMes) = 6 Then
      p = -1
      GoTo S6
   ElseIf CInt(r_str_PerMes) = 5 Then
      p = 0
      GoTo S5
   ElseIf CInt(r_str_PerMes) = 4 Then
      p = 1
      GoTo S4
   ElseIf CInt(r_str_PerMes) = 3 Then
      p = 2
      GoTo S3
   ElseIf CInt(r_str_PerMes) = 2 Then
      p = 3
      GoTo S2
   ElseIf CInt(r_str_PerMes) = 1 Then
      p = 4
      GoTo S1
   End If

SALTO1:
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.Col = 2
   grd_LisEEBG.Text = "T"

   'Titulo
   grd_LisEEBG.Col = 3
   grd_LisEEBG.CellFontName = "Arial"
   grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
   grd_LisEEBG.CellFontBold = True
   grd_LisEEBG.CellFontSize = 8
   grd_LisEEBG.Text = "ACTIVO"

   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   
   If CInt(r_str_PerAno) - 1 < anvigente Then

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
               
               grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
               grd_LisEEBG.CellFontBold = True
               grd_LisEEBG.CellFontName = "Arial"
               grd_LisEEBG.CellFontSize = 8
               
               If CInt(r_str_PerMes) = 12 Then
                  q = 5
                  grd_LisEEBG.Col = q - 1
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES01, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S12:
                  grd_LisEEBG.Col = q
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES02, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S13:
                  grd_LisEEBG.Col = q + 1
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES03, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S14:
                  grd_LisEEBG.Col = q + 2
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES04, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S15:
                  grd_LisEEBG.Col = q + 3
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES05, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S16:
                  grd_LisEEBG.Col = q + 4
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES06, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S17:
                  grd_LisEEBG.Col = q + 5
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES07, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S18:
                  grd_LisEEBG.Col = q + 6
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES08, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S19:
                  grd_LisEEBG.Col = q + 7
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES09, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S20:
                  grd_LisEEBG.Col = q + 8
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES10, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S21:
                  grd_LisEEBG.Col = q + 9
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES11, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S22:
                  grd_LisEEBG.Col = q + 10
                  grd_LisEEBG.Text = Format(g_rst_Princi!MES12, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
                  
               ElseIf CInt(r_str_PerMes) = 11 Then
                  q = -6
                  GoTo S22
               ElseIf CInt(r_str_PerMes) = 10 Then
                   q = -5
                   GoTo S21
               ElseIf CInt(r_str_PerMes) = 9 Then
                   q = -4
                   GoTo S20
               ElseIf CInt(r_str_PerMes) = 8 Then
                   q = -3
                   GoTo S19
               ElseIf CInt(r_str_PerMes) = 7 Then
                   q = -2
                   GoTo S18
               ElseIf CInt(r_str_PerMes) = 6 Then
                   q = -1
                   GoTo S17
               ElseIf CInt(r_str_PerMes) = 5 Then
                   q = 0
                   GoTo S16
               ElseIf CInt(r_str_PerMes) = 4 Then
                   q = 1
                   GoTo S15
               ElseIf CInt(r_str_PerMes) = 3 Then
                   q = 2
                   GoTo S14
               ElseIf CInt(r_str_PerMes) = 2 Then
                   q = 3
                   GoTo S13
               ElseIf CInt(r_str_PerMes) = 1 Then
                   q = 4
                   GoTo S12
               End If
               
               'grd_LisEEBG.Col = 16
               'grd_LisEEBG.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0.00")
               'grd_LisEEBG.CellAlignment = flexAlignRightCenter
               'grd_LisEEBG.CellFontName = "Arial"
               'grd_LisEEBG.CellFontSize = 8
   
               grd_LisEEBG.Col = 17
               grd_LisEEBG.Text = Trim(g_rst_Princi!NOMGRUPO)
               grd_LisEEBG.CellFontName = "Arial"
               grd_LisEEBG.CellFontSize = 8
   
               grd_LisEEBG.Col = 18
               grd_LisEEBG.Text = Trim(g_rst_Princi!NOMSUBGRP & "")
               grd_LisEEBG.CellFontName = "Arial"
               grd_LisEEBG.CellFontSize = 8
   
               If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
                  grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
                  grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
                  
                  grd_LisEEBG.Col = 2
                  grd_LisEEBG.Text = "T"

                  grd_LisEEBG.Col = 3
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
                  grd_LisEEBG.CellFontBold = True
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

   ElseIf CInt(r_str_PerAno) - 1 >= anvigente Then
 
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         Do While Not g_rst_GenAux.EOF
            If mesvigente = 12 Then
                  GoTo Seguir
            ElseIf mesvigente = 11 Then
                  For i = 17 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 10 Then
                  For i = 16 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 9 Then
                  For i = 15 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 8 Then
                  For i = 14 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 7 Then
                  For i = 13 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 6 Then
                  For i = 12 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 5 Then
                  For i = 11 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 4 Then
                  For i = 10 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 3 Then
                  For i = 9 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 2 Then
                  For i = 8 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            ElseIf mesvigente = 1 Then
                  For i = 7 To 17
                     g_rst_GenAux.Fields(i).Value = 0
                  Next i
            End If
            
            g_rst_GenAux.Update
            g_rst_GenAux.MoveNext
         Loop
      End If
    
Seguir:
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         Do While Not g_rst_GenAux.EOF
      
            If Trim(g_rst_GenAux!INDTIPO) <> "D" Then
               If Trim(g_rst_GenAux!INDTIPO) <> "L" Then
                  If Trim(g_rst_GenAux!INDTIPO) = "B" Then GoTo SALTO2

                  grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
                  grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
                  
                  grd_LisEEBG.Col = 0
                  grd_LisEEBG.Text = Trim(g_rst_GenAux!GRUPO)
                  grd_LisEEBG.Col = 1
                  grd_LisEEBG.Text = Trim(g_rst_GenAux!SUBGRP)
                  
                  grd_LisEEBG.Col = 2
                  grd_LisEEBG.Text = Trim(g_rst_GenAux!INDTIPO)
                  
                  If Trim(g_rst_GenAux!INDTIPO) = "S" Or Trim(g_rst_GenAux!INDTIPO) = "N" Or Trim(g_rst_GenAux!INDTIPO) = "R" Then
                     grd_LisEEBG.Col = 3
                     grd_LisEEBG.Text = Space(5) & Trim(g_rst_GenAux!NOMSUBGRP)
                  ElseIf Trim(g_rst_GenAux!INDTIPO) = "L" Then
                     grd_LisEEBG.Col = 3
                     grd_LisEEBG.Text = ""
                  ElseIf Trim(g_rst_GenAux!INDTIPO) = "D" Then
                     grd_LisEEBG.Col = 3
                     grd_LisEEBG.Text = Trim(g_rst_GenAux!NOMGRUPO)
                  ElseIf Trim(g_rst_GenAux!INDTIPO) = "G" Or Trim(g_rst_GenAux!INDTIPO) = "T" Or Trim(g_rst_GenAux!INDTIPO) = "A" Then
                     grd_LisEEBG.Col = 3
                     grd_LisEEBG.Text = Trim(g_rst_GenAux!NOMGRUPO)
                  End If
                  grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
                  grd_LisEEBG.CellFontBold = True
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
               
                  If CInt(r_str_PerMes) = 12 Then
                      GoTo rellenar
                
'''                r = 15
'''                grd_LisEEBG.Col = r - 11
'''                grd_LisEEBG.Text = Format(g_rst_GenAux!MES01, "###,###,###,##0.00")
'''                grd_LisEEBG.CellAlignment = flexAlignRightCenter
'''                grd_LisEEBG.CellFontName = "Arial"
'''                grd_LisEEBG.CellFontSize = 8

S23:
                  grd_LisEEBG.Col = r - 10
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES02, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S24:
                  grd_LisEEBG.Col = r - 9
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES03, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S25:
                  grd_LisEEBG.Col = r - 8
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES04, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S26:
                  grd_LisEEBG.Col = r - 7
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES05, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S27:
                  grd_LisEEBG.Col = r - 6
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES06, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S28:
                  grd_LisEEBG.Col = r - 5
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES07, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S29:
                  grd_LisEEBG.Col = r - 4
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES08, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S30:
                  grd_LisEEBG.Col = r - 3
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES09, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S31:
                  grd_LisEEBG.Col = r - 2
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES10, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S32:
                  grd_LisEEBG.Col = r - 1
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES11, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S33:
                  grd_LisEEBG.Col = r
                  grd_LisEEBG.Text = Format(g_rst_GenAux!MES12, "###,###,###,##0.00")
                  grd_LisEEBG.CellAlignment = flexAlignRightCenter
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
                
                  ElseIf CInt(r_str_PerMes) = 11 Then
                          r = 4
                          GoTo S33
                  ElseIf CInt(r_str_PerMes) = 10 Then
                          r = 5
                          GoTo S32
                  ElseIf CInt(r_str_PerMes) = 9 Then
                          r = 6
                          GoTo S31
                  ElseIf CInt(r_str_PerMes) = 8 Then
                          r = 7
                          GoTo S30
                  ElseIf CInt(r_str_PerMes) = 7 Then
                          r = 8
                          GoTo S29
                  ElseIf CInt(r_str_PerMes) = 6 Then
                          r = 9
                          GoTo S28
                  ElseIf CInt(r_str_PerMes) = 5 Then
                          r = 10
                          GoTo S27
                  ElseIf CInt(r_str_PerMes) = 4 Then
                          r = 11
                          GoTo S26
                  ElseIf CInt(r_str_PerMes) = 3 Then
                          r = 12
                          GoTo S25
                  ElseIf CInt(r_str_PerMes) = 2 Then
                          r = 13
                          GoTo S24
                  ElseIf CInt(r_str_PerMes) = 1 Then
                          r = 14
                          GoTo S23
                  End If

                  If CInt(r_str_PerMes) = 12 Then
rellenar:
                     For k = 4 To 15
                        grd_LisEEBG.Col = k
                        grd_LisEEBG.Text = Format(0, "###,###,###,##0.00")
                        grd_LisEEBG.CellFontName = "Arial"
                        grd_LisEEBG.CellFontSize = 8
                     Next
                  End If

                  grd_LisEEBG.Col = 17
                  grd_LisEEBG.Text = Trim(g_rst_GenAux!NOMGRUPO)
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
                  
                  grd_LisEEBG.Col = 18
                  grd_LisEEBG.Text = Trim(g_rst_GenAux!NOMSUBGRP & "")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8

                  If Trim(g_rst_GenAux!NOMGRUPO) = "TOTAL ACTIVO" Then
                     grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
                     grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
                     
                     grd_LisEEBG.Col = 2
                     grd_LisEEBG.Text = "T"

                     grd_LisEEBG.Col = 3
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
                     grd_LisEEBG.CellFontBold = True
                     grd_LisEEBG.CellFontSize = 8
                     grd_LisEEBG.Text = "PASIVO"
                  End If

               Else
                  grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
                  grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
               End If
            End If
            
SALTO2:
            g_rst_GenAux.MoveNext
         Loop
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Prepara SP
   g_str_Parame = "USP_CUR_GEN_EEBG ("
   g_str_Parame = g_str_Parame & CInt(r_str_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) & ", 1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "
   
   'Ejecuta consulta
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEBG.", vbCritical, modgen_g_str_NomPlt
       Screen.MousePointer = 0
       Exit Sub
   End If
   
   Call fs_llenar(CInt(r_str_PerMes), CInt(r_str_PerAno), anvigente, mesvigente)
   grd_LisEEBG.Redraw = True
   
   If grd_LisEEBG.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEBG)
      Call fs_Activa(True)
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_llenar(ByVal Columna As Integer, ByVal anno As Integer, ByVal anvigente, ByVal mesvigente)
Dim Fila As Integer
Dim j As Integer
Dim k As Integer
Dim L As Integer
        
   Fila = 2
   grd_LisEEBG.Row = 0
    
   If Columna = 12 Then
      j = -19
      grd_LisEEBG.Col = 4 + Columna + j + 18:  grd_LisEEBG.Text = "DIC-" & Right(anno, 2)
S1:
      grd_LisEEBG.Col = 4 + Columna + j + 17:  grd_LisEEBG.Text = "NOV-" & Right(anno, 2)
S2:
      grd_LisEEBG.Col = 4 + Columna + j + 16:  grd_LisEEBG.Text = "OCT-" & Right(anno, 2)
S3:
      grd_LisEEBG.Col = 4 + Columna + j + 15:  grd_LisEEBG.Text = "SET-" & Right(anno, 2)
S4:
      grd_LisEEBG.Col = 4 + Columna + j + 14:  grd_LisEEBG.Text = "AGO-" & Right(anno, 2)
S5:
      grd_LisEEBG.Col = 4 + Columna + j + 13:  grd_LisEEBG.Text = "JUL-" & Right(anno, 2)
S6:
      grd_LisEEBG.Col = 4 + Columna + j + 12:  grd_LisEEBG.Text = "JUN-" & Right(anno, 2)
S7:
      grd_LisEEBG.Col = 4 + Columna + j + 11:  grd_LisEEBG.Text = "MAY-" & Right(anno, 2)
S8:
      grd_LisEEBG.Col = 4 + Columna + j + 10:  grd_LisEEBG.Text = "ABR-" & Right(anno, 2)
S9:
      grd_LisEEBG.Col = 4 + Columna + j + 9:   grd_LisEEBG.Text = "MAR-" & Right(anno, 2)
S10:
      grd_LisEEBG.Col = 4 + Columna + j + 8:   grd_LisEEBG.Text = "FEB-" & Right(anno, 2)
S11:
      grd_LisEEBG.Col = 4 + Columna + j + 7:   grd_LisEEBG.Text = "ENE-" & Right(anno, 2)
            
   ElseIf Columna = 11 Then
      j = -17
      GoTo S1
   ElseIf Columna = 10 Then
       j = -15
      GoTo S2
   ElseIf Columna = 9 Then
       j = -13
       GoTo S3
   ElseIf Columna = 8 Then
       j = -11
       GoTo S4
   ElseIf Columna = 7 Then
       j = -9
       GoTo S5
   ElseIf Columna = 6 Then
       j = -7
       GoTo S6
   ElseIf Columna = 5 Then
       j = -5
       GoTo S7
   ElseIf Columna = 4 Then
       j = -3
       GoTo S8
   ElseIf Columna = 3 Then
       j = -1
       GoTo S9
   ElseIf Columna = 2 Then
       j = 1
       GoTo S10
   ElseIf Columna = 1 Then
       j = 3
       GoTo S11
   End If
             
   If Me.ipp_PerAno.Text < anvigente Then
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            For Fila = 4 To grd_LisEEBG.Rows - 1
               If Trim(g_rst_Princi!INDTIPO) <> "L" Then
                  If Trim(g_rst_Princi!INDTIPO) = "B" Then GoTo SALTO
                  
                  If Columna = 12 Then        'DICIEMBRE
                     k = 13
                  
                     grd_LisEEBG.TextMatrix(Fila, k + 2) = Format(g_rst_Princi!MES12, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S12:
                     grd_LisEEBG.TextMatrix(Fila, k + 1) = Format(g_rst_Princi!MES11, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S13:
                     grd_LisEEBG.TextMatrix(Fila, k) = Format(g_rst_Princi!MES10, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S14:
                     grd_LisEEBG.TextMatrix(Fila, k - 1) = Format(g_rst_Princi!MES09, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S15:
                     grd_LisEEBG.TextMatrix(Fila, k - 2) = Format(g_rst_Princi!MES08, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S16:
                     grd_LisEEBG.TextMatrix(Fila, k - 3) = Format(g_rst_Princi!MES07, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S17:
                     grd_LisEEBG.TextMatrix(Fila, k - 4) = Format(g_rst_Princi!MES06, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S18:
                     grd_LisEEBG.TextMatrix(Fila, k - 5) = Format(g_rst_Princi!MES05, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S19:
                     grd_LisEEBG.TextMatrix(Fila, k - 6) = Format(g_rst_Princi!MES04, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S20:
                     grd_LisEEBG.TextMatrix(Fila, k - 7) = Format(g_rst_Princi!MES03, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S21:
                     grd_LisEEBG.TextMatrix(Fila, k - 8) = Format(g_rst_Princi!MES02, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
S22:
                     grd_LisEEBG.TextMatrix(Fila, k - 9) = Format(g_rst_Princi!MES01, "###,###,###,##0.00")
                     grd_LisEEBG.CellAlignment = flexAlignRightCenter
                     grd_LisEEBG.CellFontName = "Arial"
                     grd_LisEEBG.CellFontSize = 8
                    
                  ElseIf Columna = 11 Then    'NOVIEMBRE
                     k = 14
                     GoTo S12
                  ElseIf Columna = 10 Then    'OCTUBRE
                      k = 15
                      GoTo S13
                  ElseIf Columna = 9 Then     'SETIEMBRE
                      k = 16
                      GoTo S14
                  ElseIf Columna = 8 Then     'AGOSTO
                      k = 17
                      GoTo S15
                  ElseIf Columna = 7 Then     'JULIO
                      k = 18
                      GoTo S16
                  ElseIf Columna = 6 Then     'JUNIO
                      k = 19
                      GoTo S17
                  ElseIf Columna = 5 Then     'MAYO
                      k = 20
                      GoTo S18
                  ElseIf Columna = 4 Then     'ABRIL
                      k = 21
                      GoTo S19
                  ElseIf Columna = 3 Then     'MARZO
                      k = 22
                       GoTo S20
                  ElseIf Columna = 2 Then     'FEBRERO
                      k = 23
                      GoTo S21
                  ElseIf Columna = 1 Then     'ENERO
                      k = 24
                      GoTo S22
                  End If
                  
                  If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
                     Fila = Fila + 2
                  End If
               
               Else
                  GoTo SALTO 'NO ES NECESARIO
               End If
               
SALTO:
               g_rst_Princi.MoveNext
            Next
         Loop
      End If
      
   ElseIf Me.ipp_PerAno.Text >= anvigente Then
               
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            For Fila = 4 To grd_LisEEBG.Rows - 1
               If Trim(g_rst_Princi!INDTIPO) <> "L" Then
                  If Trim(g_rst_Princi!INDTIPO) = "B" Then GoTo SALTO1
                  
                  If cmb_PerMes.ListIndex = 0 Then        'ENERO
                      j = 16
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S33
                      
                  ElseIf cmb_PerMes.ListIndex = 1 Then    'FEBRERO
                      j = 15
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S32
                      
                  ElseIf cmb_PerMes.ListIndex = 2 Then    'MARZO
                      j = 14
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S31
                      
                  ElseIf cmb_PerMes.ListIndex = 3 Then    'ABRIL
                      j = 13
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S30
                      
                  ElseIf cmb_PerMes.ListIndex = 4 Then    'MAYO
                      j = 12
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S29
                      
                  ElseIf cmb_PerMes.ListIndex = 5 Then    'JUNIO
                      j = 11
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S28
                      
                  ElseIf cmb_PerMes.ListIndex = 6 Then    'JULIO
                      j = 10
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S27
                            
                  ElseIf cmb_PerMes.ListIndex = 7 Then    'AGOSTO
                      j = 9
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S26
                      
                  ElseIf cmb_PerMes.ListIndex = 8 Then    'SETIEMBRE
                      j = 8
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S25
                      
                  ElseIf cmb_PerMes.ListIndex = 9 Then    'OCTUBRE
                      j = 7
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S24
                      
                  ElseIf cmb_PerMes.ListIndex = 10 Then   'NOVIEMBRE
                      j = 6
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S23
                      
                  ElseIf cmb_PerMes.ListIndex = 11 Then   'DICIEMBRE
                      j = 5
                      If cmb_PerMes.ListIndex < mesvigente Then GoTo S34
                            
                  End If
                        
                  If mesvigente = 12 Then
                     L = 10
S34:
                     grd_LisEEBG.TextMatrix(Fila, j + 10) = Format(g_rst_Princi!MES12, "###,###,###,##0.00")
S23:
                     grd_LisEEBG.TextMatrix(Fila, j + 9) = Format(g_rst_Princi!MES11, "###,###,###,##0.00")
S24:
                     grd_LisEEBG.TextMatrix(Fila, j + 8) = Format(g_rst_Princi!MES10, "###,###,###,##0.00")
S25:
                     grd_LisEEBG.TextMatrix(Fila, j + 7) = Format(g_rst_Princi!MES09, "###,###,###,##0.00")
S26:
                     grd_LisEEBG.TextMatrix(Fila, j + 6) = Format(g_rst_Princi!MES08, "###,###,###,##0.00")
S27:
                     grd_LisEEBG.TextMatrix(Fila, j + 5) = Format(g_rst_Princi!MES07, "###,###,###,##0.00")
S28:
                     grd_LisEEBG.TextMatrix(Fila, j + 4) = Format(g_rst_Princi!MES06, "###,###,###,##0.00")
S29:
                     grd_LisEEBG.TextMatrix(Fila, j + 3) = Format(g_rst_Princi!MES05, "###,###,###,##0.00")
S30:
                     grd_LisEEBG.TextMatrix(Fila, j + 2) = Format(g_rst_Princi!MES04, "###,###,###,##0.00")
S31:
                     grd_LisEEBG.TextMatrix(Fila, j + 1) = Format(g_rst_Princi!MES03, "###,###,###,##0.00")
S32:
                     grd_LisEEBG.TextMatrix(Fila, j) = Format(g_rst_Princi!MES02, "###,###,###,##0.00")
S33:
                     grd_LisEEBG.TextMatrix(Fila, j - 1) = Format(g_rst_Princi!MES01, "###,###,###,##0.00")
                        
                  ElseIf mesvigente = 11 Then
                     L = 9
                     GoTo S23
                  ElseIf mesvigente = 10 Then
                     L = 8
                     GoTo S24
                  ElseIf mesvigente = 9 Then
                     L = 7
                     GoTo S25
                  ElseIf mesvigente = 8 Then
                     L = 6
                     GoTo S26
                  ElseIf mesvigente = 7 Then
                     L = 5
                     GoTo S27
                  ElseIf mesvigente = 6 Then
                      L = 4
                      GoTo S28
                  ElseIf mesvigente = 5 Then
                      L = 3
                      GoTo S29
                  ElseIf mesvigente = 4 Then
                      L = 2
                      GoTo S30
                  ElseIf mesvigente = 3 Then
                      L = 1
                      GoTo S31
                  ElseIf mesvigente = 2 Then
                      L = 0
                      GoTo S32
                  ElseIf mesvigente = 1 Then
                      L = -1
                      GoTo S33
                  End If
                        
                  For k = 1 To CInt(cmb_PerMes.ListIndex + 1 - mesvigente)
                      grd_LisEEBG.TextMatrix(Fila, k + j + L) = Format(0, "###,###,###,##0.00")
                  Next
                        
                  If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
                      Fila = Fila + 2
                  End If
               End If
SALTO1:
               g_rst_Princi.MoveNext
            Next
         Loop
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Centra fila (meses) del encabezado
   For j = 4 To 15
        grd_LisEEBG.Row = 0
        grd_LisEEBG.Col = j
        grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   Next j

End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   Call fs_Activa(False)
   Call fs_Recorset_nc
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Recorset_nc()
    Set g_rst_GenAux = New ADODB.Recordset
    
    g_rst_GenAux.Fields.Append "GRUPO", adBigInt, 2, adFldFixed
    g_rst_GenAux.Fields.Append "NOMGRUPO", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "SUBGRP", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "NOMSUBGRP", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "CNTACTBLE", adChar, 30, adFldIsNullable
    g_rst_GenAux.Fields.Append "NOMCTA", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "MES01", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES02", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES03", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES04", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES05", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES06", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES07", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES08", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES09", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES10", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES11", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES12", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "ACUMU", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "INDTIPO", adChar, 5, adFldFixed
    g_rst_GenAux.Fields.Append "ITEM", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "CTATIPO", adChar, 3, adFldFixed
    
    g_rst_GenAux.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   ipp_PerAno.Text = Year(date)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
   'LISTADO INGRESOS MANUALES
   grd_LisEEBG.ColWidth(0) = 0       ' GRUPO
   grd_LisEEBG.ColWidth(1) = 0       ' COD SUBGRUPO
   grd_LisEEBG.ColWidth(2) = 0       ' INDICA TIPO
   grd_LisEEBG.ColWidth(3) = 4200    ' DESCRIPCION
   grd_LisEEBG.ColWidth(4) = 1300    ' MES 1
   grd_LisEEBG.ColWidth(5) = 1300    ' MES 2
   grd_LisEEBG.ColWidth(6) = 1300    ' MES 3
   grd_LisEEBG.ColWidth(7) = 1300    ' MES 4
   grd_LisEEBG.ColWidth(8) = 1300    ' MES 5
   grd_LisEEBG.ColWidth(9) = 1300    ' MES 6
   grd_LisEEBG.ColWidth(10) = 1300   ' MES 7
   grd_LisEEBG.ColWidth(11) = 1300   ' MES 8
   grd_LisEEBG.ColWidth(12) = 1300   ' MES 9
   grd_LisEEBG.ColWidth(13) = 1300   ' MES 10
   grd_LisEEBG.ColWidth(14) = 1300   ' MES 11
   grd_LisEEBG.ColWidth(15) = 1300   ' MES 12
   grd_LisEEBG.ColWidth(16) = 0      ' ACUMULADO 930
   grd_LisEEBG.ColWidth(17) = 0      ' NOMBRE GRUPO
   grd_LisEEBG.ColWidth(18) = 0      ' NOMBRE SUBGRUPO
   grd_LisEEBG.ColAlignment(3) = flexAlignLeftCenter
   grd_LisEEBG.ColAlignment(4) = flexAlignRightCenter
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
      
   Call gs_LimpiaGrid(grd_LisEEBG)
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExcRes.Enabled = estado
    cmd_ExpExcDet.Enabled = estado
End Sub

Private Sub grd_LisEEBG_DblClick()
Dim r_str_FecRpt        As String

   If grd_LisEEBG.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   moddat_g_str_FecIng = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
   
   grd_LisEEBG.Col = 0
   moddat_g_str_CodPrd = Trim(grd_LisEEBG & "")
   
   grd_LisEEBG.Col = 1
   moddat_g_str_CodSub = Trim(grd_LisEEBG)
   
   grd_LisEEBG.Col = 2
   moddat_g_str_TipCre = Trim(grd_LisEEBG)
         
   grd_LisEEBG.Col = 17
   moddat_g_str_NomPrd = UCase(Trim(grd_LisEEBG))
   
   grd_LisEEBG.Col = 18
   moddat_g_str_NomPrd = moddat_g_str_NomPrd & " " & IIf(Len(Trim(grd_LisEEBG)) > 0, " - " & UCase(Trim(grd_LisEEBG)), "")
   
   Call gs_RefrescaGrid(grd_LisEEBG)
   
   If moddat_g_str_TipCre <> "A" And moddat_g_str_TipCre <> "L" And moddat_g_str_TipCre <> "" And moddat_g_str_TipCre <> "N" And moddat_g_str_TipCre <> "X" And moddat_g_str_TipCre <> "R" Then
      frm_RptCtb_23.Show 1
   End If
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_nrofil = 5
    r_int_NoFlLi = 2
    r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
    r_int_PerAno = CInt(ipp_PerAno.Text)
    r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
        .Cells(1, 2) = "REPORTE DE BALANCE GENERAL"
        .Range(.Cells(1, 2), .Cells(1, 3)).Merge
        .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
        .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
        .Range(.Cells(2, 2), .Cells(2, 3)).Merge
        .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
        .Cells(3, 2) = "( En Soles )"
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 4), "-") = 0 Then
            .Cells(r_int_nrofil, 4) = "'" & "ENE " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 4) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 4) & ""

        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 5), "-") = 0 Then
            .Cells(r_int_nrofil, 5) = "'" & "FEB " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 5) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 5) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 6), "-") = 0 Then
           .Cells(r_int_nrofil, 6) = "'" & "MAR " & Right(r_int_PerAno, 2)
        Else
           .Cells(r_int_nrofil, 6) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 6) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 7), "-") = 0 Then
           .Cells(r_int_nrofil, 7) = "'" & "ABR " & Right(r_int_PerAno, 2)
        Else
           .Cells(r_int_nrofil, 7) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 7) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 8), "-") = 0 Then
            .Cells(r_int_nrofil, 8) = "'" & "MAY " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 8) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 8) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 9), "-") = 0 Then
            .Cells(r_int_nrofil, 9) = "'" & "JUN " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 9) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 9) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 10), "-") = 0 Then
            .Cells(r_int_nrofil, 10) = "'" & "JUL " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 10) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 10) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 11), "-") = 0 Then
            .Cells(r_int_nrofil, 11) = "'" & "AGO " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 11) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 11) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 12), "-") = 0 Then
            .Cells(r_int_nrofil, 12) = "'" & "SET " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 12) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 12) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 13), "-") = 0 Then
            .Cells(r_int_nrofil, 13) = "'" & "OCT " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 13) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 13) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 14), "-") = 0 Then
            .Cells(r_int_nrofil, 14) = "'" & "NOV " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 14) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 14) & ""
        End If
        
        If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 15), "-") = 0 Then
            .Cells(r_int_nrofil, 15) = "'" & "DIC " & Right(r_int_PerAno, 2)
        Else
            .Cells(r_int_nrofil, 15) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 15) & ""
        End If
                
        
        .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Font.Bold = True
                
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 37
        .Columns("D").ColumnWidth = 13.5
        .Columns("D").HorizontalAlignment = xlHAlignRight
        .Columns("D").NumberFormat = "###,###,###,##0.00"
        .Columns("E").ColumnWidth = 13.5
        .Columns("E").NumberFormat = "###,###,###,##0.00"
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("F").ColumnWidth = 13.5
        .Columns("F").NumberFormat = "###,###,###,##0.00"
        .Columns("F").HorizontalAlignment = xlHAlignRight
        .Columns("G").ColumnWidth = 13.5
        .Columns("G").NumberFormat = "###,###,###,##0.00"
        .Columns("G").HorizontalAlignment = xlHAlignRight
        .Columns("H").ColumnWidth = 13.5
        .Columns("H").NumberFormat = "###,###,###,##0.00"
        .Columns("H").HorizontalAlignment = xlHAlignRight
        .Columns("I").ColumnWidth = 13.5
        .Columns("I").NumberFormat = "###,###,###,##0.00"
        .Columns("I").HorizontalAlignment = xlHAlignRight
        .Columns("J").ColumnWidth = 13.5
        .Columns("J").NumberFormat = "###,###,###,##0.00"
        .Columns("J").HorizontalAlignment = xlHAlignRight
        .Columns("K").ColumnWidth = 13.5
        .Columns("K").NumberFormat = "###,###,###,##0.00"
        .Columns("K").HorizontalAlignment = xlHAlignRight
        .Columns("L").ColumnWidth = 13.5
        .Columns("L").NumberFormat = "###,###,###,##0.00"
        .Columns("L").HorizontalAlignment = xlHAlignRight
        .Columns("M").ColumnWidth = 13.5
        .Columns("M").NumberFormat = "###,###,###,##0.00"
        .Columns("M").HorizontalAlignment = xlHAlignRight
        .Columns("N").ColumnWidth = 13.5
        .Columns("N").NumberFormat = "###,###,###,##0.00"
        .Columns("N").HorizontalAlignment = xlHAlignRight
        .Columns("O").ColumnWidth = 13.5
        .Columns("O").NumberFormat = "###,###,###,##0.00"
        .Columns("O").HorizontalAlignment = xlHAlignRight
        
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
         
        r_int_nrofil = r_int_nrofil + 2
        .Range(.Cells(4, 4), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
         
        For r_int_NoFlLi = 2 To grd_LisEEBG.Rows - 1
            If Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "G" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "F" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "A" _
                 Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "T" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)) = "X" Then
                'TITULO
                .Cells(r_int_nrofil, 2) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 3)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Interior.Color = RGB(146, 208, 80)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Font.Bold = True
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
            
            r_int_nrofil = r_int_nrofil + 1
        Next r_int_NoFlLi
   End With
   
   r_obj_Excel.Visible = True
End Sub

Private Sub fs_GenExcDet()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_str_PerMes        As String
Dim r_str_PerAno        As String
Dim L                   As Integer
Dim r                   As Integer
Dim k                   As Integer
Dim anovigente          As Integer
Dim mesvigente          As Integer
    
   r_int_Contad = 5
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE"
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE BALANCE GENERAL"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(3, 2) = "( En Soles )"
     
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 4), "-") = 0 Then
          .Cells(r_int_Contad, 4) = "'" & "ENE " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 4) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 4)
      End If
        
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 5), "-") = 0 Then
          .Cells(r_int_Contad, 5) = "'" & "FEB " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 5) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 5)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 6), "-") = 0 Then
         .Cells(r_int_Contad, 6) = "'" & "MAR " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 6) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 6)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 7), "-") = 0 Then
         .Cells(r_int_Contad, 7) = "'" & "ABR " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 7) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 7)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 8), "-") = 0 Then
         .Cells(r_int_Contad, 8) = "'" & "MAY " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 8) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 8)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 9), "-") = 0 Then
         .Cells(r_int_Contad, 9) = "'" & "JUN " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 9) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 9)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 10), "-") = 0 Then
         .Cells(r_int_Contad, 10) = "'" & "JUL " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 10) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 10)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 11), "-") = 0 Then
         .Cells(r_int_Contad, 11) = "'" & "AGO " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 11) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 11)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 12), "-") = 0 Then
         .Cells(r_int_Contad, 12) = "'" & "SET " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 12) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 12)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 13), "-") = 0 Then
        .Cells(r_int_Contad, 13) = "'" & "OCT " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 13) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 13)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 14), "-") = 0 Then
         .Cells(r_int_Contad, 14) = "'" & "NOV " & Right(r_int_PerAno, 2)
      Else
        .Cells(r_int_Contad, 14) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 14)
      End If
      
      If InStr(frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 15), "-") = 0 Then
         .Cells(r_int_Contad, 15) = "'" & "DIC " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_Contad, 15) = "'" & frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 15)
      End If
      
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
      .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 37
      .Columns("D").ColumnWidth = 13.5
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0.00"
      .Columns("E").ColumnWidth = 13.5
      .Columns("E").NumberFormat = "###,###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 13.5
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 13.5
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 13.5
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 13.5
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 13.5
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 13.5
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 13.5
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11

      'Consulta para obtener el mes y año vigente
      g_str_Parame = "SELECT max(PERMES_CODMES) AS MES, max(PERMES_CODANO) AS ANO FROM CTB_PERMES "
      g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         MsgBox "Error al ejecutar la consulta para obtener año vigente.", vbCritical, modgen_g_str_NomPlt
      End If
          
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         anovigente = g_rst_Genera!ANO
         mesvigente = g_rst_Genera!Mes
      End If
      
      'PARA EL AÑO ACTUAL
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TT_EEBG  "
      g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
      'g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "
      g_str_Parame = g_str_Parame & " ORDER BY ORDEN, SUBGRP, ITEM, INDTIPO "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
        
      r_int_Contad = r_int_Contad + 2
      .Cells(r_int_Contad, 2) = "ACTIVO"
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
      r_int_Contad = r_int_Contad + 1

      If g_rst_GenAux.RecordCount > 0 Then g_rst_GenAux.MoveFirst
      
      Do While Not g_rst_Princi.EOF
            r_int_Contad = r_int_Contad + 1
            If Trim(g_rst_Princi!INDTIPO) = "L" Then
              If Trim(g_rst_Princi!NOMGRUPO) = "TOTAL ACTIVO" Then
                  r_int_Contad = r_int_Contad + 1
                  .Cells(r_int_Contad, 2) = "PASIVO"
                  .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
                  .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
                  r_int_Contad = r_int_Contad + 1
              End If
              If g_rst_GenAux.RecordCount > 0 Then g_rst_GenAux.MoveNext
              g_rst_Princi.MoveNext
              r_int_Contad = r_int_Contad + 1
            End If
            If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Or Trim(g_rst_Princi!INDTIPO) = "A" _
                  Or Trim(g_rst_Princi!INDTIPO) = "T" Or Trim(g_rst_Princi!INDTIPO) = "X" Then
              .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
            End If
            If Trim(g_rst_Princi!INDTIPO) = "S" Or Trim(g_rst_Princi!INDTIPO) = "N" Or Trim(g_rst_Princi!INDTIPO) = "R" Then
              .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
            End If
            If Trim(g_rst_Princi!INDTIPO) = "D" Then
              .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
              .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
            End If
            If Trim(g_rst_Princi!INDTIPO) = "B" Then
              .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
              .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP & "")
            End If

        
            If Me.cmb_PerMes.ListIndex = 11 Then                'DICIEMBRE
            
               r = 7
               
              .Cells(r_int_Contad, r + 8) = g_rst_Princi!MES12
              .Cells(r_int_Contad, r + 7) = g_rst_Princi!MES11
              .Cells(r_int_Contad, r + 6) = g_rst_Princi!MES10
              .Cells(r_int_Contad, r + 5) = g_rst_Princi!MES09
              .Cells(r_int_Contad, r + 4) = g_rst_Princi!MES08
              .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
              .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
              .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
              .Cells(r_int_Contad, r) = g_rst_Princi!MES04
              .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
              .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
              .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
              
              GoTo SALTO2
              
              .Cells(r_int_Contad, r - 15) = g_rst_GenAux!MES01
S1:
              .Cells(r_int_Contad, r - 14) = g_rst_GenAux!MES02
S2:
              .Cells(r_int_Contad, r - 13) = g_rst_GenAux!MES03
S3:
              .Cells(r_int_Contad, r - 12) = g_rst_GenAux!MES04
S4:
              .Cells(r_int_Contad, r - 11) = g_rst_GenAux!MES05
S5:
              .Cells(r_int_Contad, r - 10) = g_rst_GenAux!MES06
S6:
              .Cells(r_int_Contad, r - 9) = g_rst_GenAux!MES07
S7:
              .Cells(r_int_Contad, r - 8) = g_rst_GenAux!MES08
S8:
              .Cells(r_int_Contad, r - 7) = g_rst_GenAux!MES09
S9:
              .Cells(r_int_Contad, r - 6) = g_rst_GenAux!MES10
S10:
              .Cells(r_int_Contad, r - 5) = g_rst_GenAux!MES11
S11:
              .Cells(r_int_Contad, r - 4) = g_rst_GenAux!MES12
            
            ElseIf Me.cmb_PerMes.ListIndex = 10 Then                'NOVIEMBRE
                r = 8
                .Cells(r_int_Contad, r + 7) = g_rst_Princi!MES11
                .Cells(r_int_Contad, r + 6) = g_rst_Princi!MES10
                .Cells(r_int_Contad, r + 5) = g_rst_Princi!MES09
                .Cells(r_int_Contad, r + 4) = g_rst_Princi!MES08
                .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
              
                GoTo S11
            ElseIf Me.cmb_PerMes.ListIndex = 9 Then                 'OCTUBRE
                r = 9
                .Cells(r_int_Contad, r + 6) = g_rst_Princi!MES10
                .Cells(r_int_Contad, r + 5) = g_rst_Princi!MES09
                .Cells(r_int_Contad, r + 4) = g_rst_Princi!MES08
                .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
              
                GoTo S10
            ElseIf Me.cmb_PerMes.ListIndex = 8 Then                 'SETIEMBRE
                r = 10
                .Cells(r_int_Contad, r + 5) = g_rst_Princi!MES09
                .Cells(r_int_Contad, r + 4) = g_rst_Princi!MES08
                .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
              GoTo S9
            ElseIf Me.cmb_PerMes.ListIndex = 7 Then                 'AGOSTO
                r = 11
                .Cells(r_int_Contad, r + 4) = g_rst_Princi!MES08
                .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                GoTo S8
            ElseIf Me.cmb_PerMes.ListIndex = 6 Then                 'JULIO
                r = 12
                .Cells(r_int_Contad, r + 3) = g_rst_Princi!MES07
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                GoTo S7
            ElseIf Me.cmb_PerMes.ListIndex = 5 Then                 'JUNIO
                r = 13
                .Cells(r_int_Contad, r + 2) = g_rst_Princi!MES06
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                GoTo S6
            ElseIf Me.cmb_PerMes.ListIndex = 4 Then                 'MAYO
                r = 14
                .Cells(r_int_Contad, r + 1) = g_rst_Princi!MES05
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                
                GoTo S5
            ElseIf Me.cmb_PerMes.ListIndex = 3 Then                 'ABRIL
                r = 15
                .Cells(r_int_Contad, r) = g_rst_Princi!MES04
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                
                GoTo S4
            ElseIf Me.cmb_PerMes.ListIndex = 2 Then                 'MARZO
                r = 16
                .Cells(r_int_Contad, r - 1) = g_rst_Princi!MES03
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01

                GoTo S3
            ElseIf Me.cmb_PerMes.ListIndex = 1 Then                 'FEBRERO
                r = 17
                .Cells(r_int_Contad, r - 2) = g_rst_Princi!MES02
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                GoTo S2
            ElseIf Me.cmb_PerMes.ListIndex = 0 Then                 'ENERO
                r = 18
                .Cells(r_int_Contad, r - 3) = g_rst_Princi!MES01
                GoTo S1
            End If
            
SALTO2:
            If anovigente = CInt(Me.ipp_PerAno.Text) Then
               For k = 1 To CInt(cmb_PerMes.ListIndex + 1 - mesvigente)
                   .Cells(r_int_Contad, k + r - 2) = Format(0, "###,###,###,##0.00")
               Next
            End If

          g_rst_Princi.MoveNext
          If g_rst_GenAux.RecordCount > 0 Then g_rst_GenAux.MoveNext
          DoEvents
       Loop
       
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
   End With
   
   '---------------------------hoja 2------------------------------------------------
   r_int_Contad = 6
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_obj_Excel.Sheets(2).Name = "RESUMEN"
   With r_obj_Excel.Sheets(2)
      .Cells(3, 2) = "RESUMEN PARA LA SUNAT"
      .Range(.Cells(3, 2), .Cells(3, 3)).Merge
      .Range(.Cells(3, 2), .Cells(3, 3)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 3)).HorizontalAlignment = xlHAlignCenter
      .Cells(5, 2) = "CODIGO"
      .Cells(5, 3) = "MONTO"
      
      .Range(.Cells(5, 2), .Cells(5, 3)).Font.Bold = True
      .Range(.Cells(5, 2), .Cells(5, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 2), .Cells(5, 3)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 3)).Borders(xlInsideVertical).LineStyle = xlContinuous

      .Columns("B").ColumnWidth = 8
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 18
      .Columns("C").NumberFormat = "###,###,##0.00"
      
      'llama al SP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "USP_RPT_BG_RESUMSUNAT("
      g_str_Parame = g_str_Parame & CInt(r_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(r_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'BALANCE GENERAL', "
      g_str_Parame = g_str_Parame & "0)"
         
      'EJECUTA CONSULTA
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + " SELECT RPT_CODIGO, RPT_VALNUM" & Format(r_str_PerMes, "00") & " "
      g_str_Parame = g_str_Parame + "   FROM RPT_TABLA_TEMP WHERE RPT_PERMES = '" & CInt(r_str_PerMes) & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_PERANO = '" & CInt(r_str_PerAno) & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'BALANCE GENERAL' "
      g_str_Parame = g_str_Parame + "    AND RPT_MONEDA = 0 "
      g_str_Parame = g_str_Parame + " ORDER BY RPT_CODIGO"
      
      'EJECUTA CONSULTA
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!RPT_CODIGO)
            .Cells(r_int_Contad, 3) = g_rst_Princi.Fields(1)
            
            g_rst_Princi.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcDet2()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
   
   r_int_Contad = 5
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE BALANCE GENERAL"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(3, 2) = "( En Soles )"
      
      .Cells(r_int_Contad, 4) = "'" & "ENE " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 5) = "'" & "FEB " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 6) = "'" & "MAR " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 7) = "'" & "ABR " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 8) = "'" & "MAY " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 9) = "'" & "JUN " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 10) = "'" & "JUL " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 11) = "'" & "AGO " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 12) = "'" & "SET " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 13) = "'" & "OCT " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 14) = "'" & "NOV " & Right(r_int_PerAno, 2)
      .Cells(r_int_Contad, 15) = "'" & "DIC " & Right(r_int_PerAno, 2)
      
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
      .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 37
      .Columns("D").ColumnWidth = 13.5
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0.00"
      .Columns("E").ColumnWidth = 13.5
      .Columns("E").NumberFormat = "###,###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 13.5
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 13.5
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 13.5
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 13.5
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 13.5
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 13.5
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 13.5
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TT_EEBG  "
      g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
      'g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "
      g_str_Parame = g_str_Parame & " ORDER BY ORDEN, SUBGRP, ITEM, INDTIPO "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
        
      r_int_Contad = r_int_Contad + 2
      .Cells(r_int_Contad, 2) = "ACTIVO"
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
      r_int_Contad = r_int_Contad + 1

      Do While Not g_rst_Princi.EOF
             
          r_int_Contad = r_int_Contad + 1
          If Trim(g_rst_Princi!INDTIPO) = "L" Then
            g_rst_Princi.MoveNext
            r_int_Contad = r_int_Contad + 1
          End If
          If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Or Trim(g_rst_Princi!INDTIPO) = "A" _
                Or Trim(g_rst_Princi!INDTIPO) = "T" Or Trim(g_rst_Princi!INDTIPO) = "X" Then
            .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
          End If
          If Trim(g_rst_Princi!INDTIPO) = "S" Or Trim(g_rst_Princi!INDTIPO) = "N" Or Trim(g_rst_Princi!INDTIPO) = "R" Then
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 15)).Font.Bold = True
          End If
          If Trim(g_rst_Princi!INDTIPO) = "D" Then
            .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
          End If
          If Trim(g_rst_Princi!INDTIPO) = "B" Then
            .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP & "")
          End If
          
          .Cells(r_int_Contad, 4) = g_rst_Princi!MES01
          .Cells(r_int_Contad, 5) = g_rst_Princi!MES02
          .Cells(r_int_Contad, 6) = g_rst_Princi!MES03
          .Cells(r_int_Contad, 7) = g_rst_Princi!MES04
          .Cells(r_int_Contad, 8) = g_rst_Princi!MES05
          .Cells(r_int_Contad, 9) = g_rst_Princi!MES06
          .Cells(r_int_Contad, 10) = g_rst_Princi!MES07
          .Cells(r_int_Contad, 11) = g_rst_Princi!MES08
          .Cells(r_int_Contad, 12) = g_rst_Princi!MES09
          .Cells(r_int_Contad, 13) = g_rst_Princi!MES10
          .Cells(r_int_Contad, 14) = g_rst_Princi!MES11
          .Cells(r_int_Contad, 15) = g_rst_Princi!MES12
          g_rst_Princi.MoveNext
          DoEvents
       Loop
        
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
   End With
         
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If cmb_PerMes.ListIndex > -1 Then
      If KeyAscii = 13 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub
