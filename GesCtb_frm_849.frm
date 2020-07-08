VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_23 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   3735
   ClientTop       =   3900
   ClientWidth     =   14130
   Icon            =   "GesCtb_frm_849.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel108 
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14130
      _Version        =   65536
      _ExtentX        =   24924
      _ExtentY        =   13891
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   5775
         Left            =   45
         TabIndex        =   1
         Top             =   2025
         Width           =   14055
         _Version        =   65536
         _ExtentX        =   24791
         _ExtentY        =   10186
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
            Height          =   5655
            Left            =   45
            TabIndex        =   2
            Top             =   60
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   9975
            _Version        =   393216
            Rows            =   21
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
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
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
            TabIndex        =   4
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
            Picture         =   "GesCtb_frm_849.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   5
         Top             =   780
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
         Font3D          =   2
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   45
            Picture         =   "GesCtb_frm_849.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13305
            Picture         =   "GesCtb_frm_849.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   525
         Left            =   60
         TabIndex        =   8
         Top             =   1470
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   926
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
         Begin VB.Label lblconcepto 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1245
            TabIndex        =   10
            Top             =   120
            Width           =   12705
         End
         Begin VB.Label Label5 
            Caption         =   "Concepto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_LisEEBG)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmd_ExpExc.Enabled = False
   lblConcepto.Caption = moddat_g_str_NomPrd
   
   'LISTADO INGRESOS MANUALES
   grd_LisEEBG.ColWidth(0) = 1200    ' CUENTA
   grd_LisEEBG.ColWidth(1) = 3600    ' DESCRIPCION
   grd_LisEEBG.ColWidth(2) = 1200    ' MES 1
   grd_LisEEBG.ColWidth(3) = 1200    ' MES 2
   grd_LisEEBG.ColWidth(4) = 1200    ' MES 3
   grd_LisEEBG.ColWidth(5) = 1200    ' MES 4
   grd_LisEEBG.ColWidth(6) = 1200    ' MES 5
   grd_LisEEBG.ColWidth(7) = 1200    ' MES 6
   grd_LisEEBG.ColWidth(8) = 1200    ' MES 7
   grd_LisEEBG.ColWidth(9) = 1200    ' MES 8
   grd_LisEEBG.ColWidth(10) = 1200   ' MES 9
   grd_LisEEBG.ColWidth(11) = 1200   ' MES 10
   grd_LisEEBG.ColWidth(12) = 1200   ' MES 11
   grd_LisEEBG.ColWidth(13) = 1200   ' MES 12
   grd_LisEEBG.ColWidth(14) = 0      ' ACUMULADO 900
   
   grd_LisEEBG.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEEBG.ColAlignment(1) = flexAlignLeftCenter
   grd_LisEEBG.ColAlignment(2) = flexAlignRightCenter
   grd_LisEEBG.ColAlignment(3) = flexAlignRightCenter
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
'   grd_LisEEBG.ColAlignment(14) = flexAlignRightCenter
   
   Call gs_LimpiaGrid(grd_LisEEBG)
End Sub

Private Sub fs_Buscar()
   Dim p                As Integer
   Dim k                As Integer
   Dim anovigente       As Integer
   Dim mesvigente       As Integer
   
   'Obteniendo Información
   If IsNull(moddat_g_str_CodPrd) Or moddat_g_str_CodPrd = "" And IsNull(moddat_g_str_CodSub) Or moddat_g_str_CodSub = "" Then
       Exit Sub
   End If
   
   'Consulta para obtener el mes y año vigente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT max(PERMES_CODMES) AS MES, max(PERMES_CODANO) AS ANO "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "Error al ejecutar la consulta para obtener año vigente.", vbCritical, modgen_g_str_NomPlt
   End If
     
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      anovigente = g_rst_Genera!ANO
      mesvigente = g_rst_Genera!Mes
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TT_EEBG "
   g_str_Parame = g_str_Parame & " WHERE NOT INDTIPO= 'L' AND NOT INDTIPO= 'A' "
   g_str_Parame = g_str_Parame & "   AND GRUPO = " & CInt(moddat_g_str_CodPrd) & " "
   If moddat_g_str_TipCre = "S" Then
      g_str_Parame = g_str_Parame & "   AND subgrp = " & CInt(moddat_g_str_CodSub) & "  "
   End If
   g_str_Parame = g_str_Parame & "   AND USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & " ORDER BY grupo, subgrp, item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_LisEEBG.Rows = grd_LisEEBG.Rows + 2
   grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
   grd_LisEEBG.FixedRows = 1
    
   grd_LisEEBG.Row = 0
   grd_LisEEBG.Col = 0
   grd_LisEEBG.Text = "CUENTA"
   grd_LisEEBG.Col = 1
   grd_LisEEBG.Text = "DESCRIPCION"
   grd_LisEEBG.Col = 2
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 4): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 3
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 5): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 4
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 6): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 5
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 7): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 6
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 8): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 7
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 9): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 8
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 10): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 9
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 11): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 10
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 12): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 11
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 13): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 12
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 14): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
   grd_LisEEBG.Col = 13
   grd_LisEEBG.Text = frm_RptCtb_22.grd_LisEEBG.TextMatrix(0, 15): grd_LisEEBG.CellAlignment = flexAlignCenterCenter
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If (Trim(g_rst_Princi!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_Princi!GRUPO)) Then
             GoTo SALTO
         End If
         
         grd_LisEEBG.Rows = grd_LisEEBG.Rows + 1
         grd_LisEEBG.Row = grd_LisEEBG.Rows - 1
         
         grd_LisEEBG.Col = 0
         grd_LisEEBG.Text = Trim(g_rst_Princi!CNTACTBLE & "")
         grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
         grd_LisEEBG.CellFontBold = False
         grd_LisEEBG.CellFontName = "Arial"
         grd_LisEEBG.CellFontSize = 8
         
         grd_LisEEBG.Col = 1
         If Trim(g_rst_Princi!INDTIPO) = "S" Or Trim(g_rst_Princi!INDTIPO) = "N" Then
            grd_LisEEBG.Text = Space(5) & UCase(Trim(g_rst_Princi!NOMSUBGRP))
         ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Or Trim(g_rst_Princi!INDTIPO) = "T" Then
            grd_LisEEBG.Text = Trim(g_rst_Princi!NOMGRUPO)
         ElseIf Trim(g_rst_Princi!INDTIPO) = "B" Then
            grd_LisEEBG.Text = IIf(Len(Trim(g_rst_Princi!CNTACTBLE & "")) > 0, " - " & Trim(g_rst_Princi!NOMSUBGRP & ""), "")
         Else
            grd_LisEEBG.Text = IIf(Len(Trim(g_rst_Princi!NOMCTA & "")) > 0, " - " & Trim(g_rst_Princi!NOMCTA & ""), "")
         End If
         
         grd_LisEEBG.CellForeColor = modgen_g_con_ColVer
         grd_LisEEBG.CellFontBold = False
         grd_LisEEBG.CellFontName = "Arial"
         grd_LisEEBG.CellFontSize = 8
         
         'PARA EL AÑO ACTUAL
         If frm_RptCtb_22.cmb_PerMes.ListIndex = 11 Then  'DICIEMBRE
            p = 3
            grd_LisEEBG.Col = p + 10
            grd_LisEEBG.Text = Format(g_rst_Princi!MES12, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S1:
            grd_LisEEBG.Col = p + 9
            grd_LisEEBG.Text = Format(g_rst_Princi!MES11, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S2:
            grd_LisEEBG.Col = p + 8
            grd_LisEEBG.Text = Format(g_rst_Princi!MES10, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S3:
            grd_LisEEBG.Col = p + 7
            grd_LisEEBG.Text = Format(g_rst_Princi!MES09, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S4:
            grd_LisEEBG.Col = p + 6
            grd_LisEEBG.Text = Format(g_rst_Princi!MES08, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S5:
            grd_LisEEBG.Col = p + 5
            grd_LisEEBG.Text = Format(g_rst_Princi!MES07, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S6:
            grd_LisEEBG.Col = p + 4
            grd_LisEEBG.Text = Format(g_rst_Princi!MES06, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S7:
            grd_LisEEBG.Col = p + 3
            grd_LisEEBG.Text = Format(g_rst_Princi!MES05, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S8:
            grd_LisEEBG.Col = p + 2
            grd_LisEEBG.Text = Format(g_rst_Princi!MES04, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S9:
            grd_LisEEBG.Col = p + 1
            grd_LisEEBG.Text = Format(g_rst_Princi!MES03, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S10:
            grd_LisEEBG.Col = p
            grd_LisEEBG.Text = Format(g_rst_Princi!MES02, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
S11:
            grd_LisEEBG.Col = p - 1
            grd_LisEEBG.Text = Format(g_rst_Princi!MES01, "###,###,###,##0.00")
            grd_LisEEBG.CellFontName = "Arial"
            grd_LisEEBG.CellFontSize = 8
                    
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 10 Then 'NOVIEMBRE
             p = 4
             GoTo S1
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 9 Then  'OCTUBRE
             p = 5
             GoTo S2
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 8 Then  'SETIEMBRE
             p = 6
             GoTo S3
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 7 Then  'AGOSTO
             p = 7
             GoTo S4
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 6 Then  'JULIO
             p = 8
             GoTo S5
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 5 Then  'JUNIO
             p = 9
             GoTo S6
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 4 Then  'MAYO
             p = 10
             GoTo S7
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 3 Then  'ABRIL
             p = 11
             GoTo S8
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 2 Then  'MARZO
             p = 12
             GoTo S9
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 1 Then  'FEBRERO
             p = 13
             GoTo S10
         ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 0 Then  'ENERO
             p = 14
             GoTo S11
         End If
         
         grd_LisEEBG.Col = 14
         grd_LisEEBG.Text = Trim(g_rst_Princi!INDTIPO)
            
'            grd_LisEEBG.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0.00")
'            grd_LisEEBG.CellFontName = "Arial"
'            grd_LisEEBG.CellFontSize = 8
            
SALTO:
         If frm_RptCtb_22.ipp_PerAno.Text = anovigente And p <> 0 Then
         
            For k = 1 To CInt(frm_RptCtb_22.cmb_PerMes.ListIndex + 1 - mesvigente)
               grd_LisEEBG.Col = p + k
               grd_LisEEBG.Text = Format(0, "###,###,###,##0.00")
               grd_LisEEBG.CellFontName = "Arial"
               grd_LisEEBG.CellFontSize = 8
            Next
         End If
                    
         g_rst_Princi.MoveNext
      Loop
   End If
   
   'AÑO ANTERIOR
   grd_LisEEBG.Row = 1
   
   'Informaciòn del recordset no conectado
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
      
        If g_rst_GenAux!INDTIPO <> "L" And CInt(g_rst_GenAux!GRUPO) = CInt(moddat_g_str_CodPrd) Then
            If moddat_g_str_TipCre = "S" Then
               If g_rst_GenAux!SUBGRP = CInt(moddat_g_str_CodSub) Then
                  GoTo Ingresar
               End If
            Else
Ingresar:
               If (Trim(g_rst_GenAux!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_GenAux!GRUPO)) Then
                  GoTo SALTO1
               End If
                    
               If grd_LisEEBG.Row < grd_LisEEBG.Rows - 1 Then
                  grd_LisEEBG.Row = grd_LisEEBG.Row + 1
               End If
                    
               If frm_RptCtb_22.cmb_PerMes.ListIndex = 11 Then 'DICIMEBRE
                  GoTo SALTO2
S12:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p) = Format(g_rst_GenAux!MES02, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S13:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 1) = Format(g_rst_GenAux!MES03, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S14:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 2) = Format(g_rst_GenAux!MES04, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S15:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 3) = Format(g_rst_GenAux!MES05, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S16:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 4) = Format(g_rst_GenAux!MES06, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S17:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 5) = Format(g_rst_GenAux!MES07, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S18:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 6) = Format(g_rst_GenAux!MES08, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S19:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 7) = Format(g_rst_GenAux!MES09, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S20:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 8) = Format(g_rst_GenAux!MES10, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S21:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 9) = Format(g_rst_GenAux!MES11, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
S22:
                  grd_LisEEBG.TextMatrix(grd_LisEEBG.Row, p + 10) = Format(g_rst_GenAux!MES12, "###,###,###,##0.00")
                  grd_LisEEBG.CellFontName = "Arial"
                  grd_LisEEBG.CellFontSize = 8
                            
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 10 Then 'NOVIEMBRE
                   p = -8
                   GoTo S22
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 9 Then  'OCTUBRE
                   p = -7
                   GoTo S21
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 8 Then  'SEPTIEMBRE
                   p = -6
                   GoTo S20
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 7 Then  'AGOSTO
                   p = -5
                   GoTo S19
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 6 Then  'JULIO
                   p = -4
                   GoTo S18
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 5 Then  'JUNIO
                   p = -3
                   GoTo S17
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 4 Then  'MAYO
                   p = -2
                   GoTo S16
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 3 Then  'ABRIL
                   p = -1
                   GoTo S15
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 2 Then  'MARZO
                   p = 0
                   GoTo S14
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 1 Then  'FEBRERO
                   p = 1
                   GoTo S13
               ElseIf frm_RptCtb_22.cmb_PerMes.ListIndex = 0 Then  'ENERO
                   p = 2
                   GoTo S12
               End If
            End If
         End If
            
SALTO1:
         g_rst_GenAux.MoveNext
      Loop
      
SALTO2:
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   grd_LisEEBG.Redraw = True
   If grd_LisEEBG.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEBG)
      cmd_ExpExc.Enabled = True
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
   
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer
            
   r_int_nrofil = 5
    
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE BALANCE GERENAL"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = moddat_g_str_FecIng
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(3, 2) = "( En Soles )"
  
      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 2), "-") = 0 Then
         .Cells(r_int_nrofil, 4) = "'" & "ENE " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(r_int_nrofil, 4) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 2) & ""
      End If
      
      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 3), "-") = 0 Then
         .Cells(r_int_nrofil, 5) = "'" & "FEB " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(r_int_nrofil, 5) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 3) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 4), "-") = 0 Then
         .Cells(r_int_nrofil, 6) = "'" & "MAR " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(r_int_nrofil, 6) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 4) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 5), "-") = 0 Then
         .Cells(r_int_nrofil, 7) = "'" & "ABR " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(r_int_nrofil, 7) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 5) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 6), "-") = 0 Then
          .Cells(r_int_nrofil, 8) = "'" & "MAY " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 8) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 6) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 7), "-") = 0 Then
          .Cells(r_int_nrofil, 9) = "'" & "JUN " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 9) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 7) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 8), "-") = 0 Then
          .Cells(r_int_nrofil, 10) = "'" & "JUL " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 10) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 8) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 9), "-") = 0 Then
          .Cells(r_int_nrofil, 11) = "'" & "AGO " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 11) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 9) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 10), "-") = 0 Then
          .Cells(r_int_nrofil, 12) = "'" & "SET " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 12) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 10) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 11), "-") = 0 Then
          .Cells(r_int_nrofil, 13) = "'" & "OCT " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 13) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 11) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 12), "-") = 0 Then
          .Cells(r_int_nrofil, 14) = "'" & "NOV " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 14) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 12) & ""
      End If

      If InStr(frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 13), "-") = 0 Then
          .Cells(r_int_nrofil, 15) = "'" & "DIC " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(r_int_nrofil, 15) = "'" & frm_RptCtb_23.grd_LisEEBG.TextMatrix(0, 13) & ""
      End If

      '.Cells(4, 16) = "'" & "ACUM " & Right(moddat_g_str_FecIng, 2)
        
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Font.Bold = True
      
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
      
      .Range(.Cells(r_int_nrofil, 3), .Cells(r_int_nrofil, 15)).HorizontalAlignment = xlHAlignCenter

      r_int_nrofil = r_int_nrofil + 2
      For r_int_NoFlLi = 2 To grd_LisEEBG.Rows - 1
         If Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 14)) = "G" Then
            'TITULO
            .Cells(r_int_nrofil, 2) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 1)
            .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Font.Bold = True
            
         ElseIf Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 14)) = "S" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 14)) = "T" Or Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 14)) = "N" Then
             .Cells(r_int_nrofil, 3) = Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 1))
             .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 15)).Font.Bold = True
         Else
             .Cells(r_int_nrofil, 2) = "'" & Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 0))
             .Cells(r_int_nrofil, 3) = Trim(grd_LisEEBG.TextMatrix(r_int_NoFlLi, 1))
         End If

         .Cells(r_int_nrofil, 4) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 2)
         .Cells(r_int_nrofil, 5) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 3)
         .Cells(r_int_nrofil, 6) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 4)
         .Cells(r_int_nrofil, 7) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 5)
         .Cells(r_int_nrofil, 8) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 6)
         .Cells(r_int_nrofil, 9) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 7)
         .Cells(r_int_nrofil, 10) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 8)
         .Cells(r_int_nrofil, 11) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 9)
         .Cells(r_int_nrofil, 12) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 10)
         .Cells(r_int_nrofil, 13) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 11)
         .Cells(r_int_nrofil, 14) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 12)
         .Cells(r_int_nrofil, 15) = grd_LisEEBG.TextMatrix(r_int_NoFlLi, 13)

         r_int_nrofil = r_int_nrofil + 1
      Next r_int_NoFlLi
   End With
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
End Sub
