VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_20 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   4290
   ClientTop       =   4665
   ClientWidth     =   15150
   Icon            =   "GesCtb_frm_818.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel3 
      Height          =   5655
      Left            =   0
      TabIndex        =   3
      Top             =   1965
      Width           =   15105
      _Version        =   65536
      _ExtentX        =   26644
      _ExtentY        =   9975
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
      Begin MSFlexGridLib.MSFlexGrid grd_LisEEFF 
         Height          =   5565
         Left            =   45
         TabIndex        =   2
         Top             =   15
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   9816
         _Version        =   393216
         Rows            =   20
         Cols            =   16
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15120
      _Version        =   65536
      _ExtentX        =   26670
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
         TabIndex        =   5
         Top             =   180
         Width           =   4575
         _Version        =   65536
         _ExtentX        =   8070
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Reporte de Estados Financieros"
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
         Picture         =   "GesCtb_frm_818.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   15120
      _Version        =   65536
      _ExtentX        =   26670
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
         Picture         =   "GesCtb_frm_818.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   14475
         Picture         =   "GesCtb_frm_818.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   615
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   1410
      Width           =   15105
      _Version        =   65536
      _ExtentX        =   26644
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
         TabIndex        =   9
         Top             =   120
         Width           =   13755
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
         TabIndex        =   8
         Top             =   150
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frm_RptCtb_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   'Call fs_GenExc_AntVer
   Call fs_GenExc_NueVer
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   'Call fs_Inicia_AntVer
   'Call fs_Buscar_AntVer
   
   Call fs_Inicia_NueVer
   Call fs_Buscar_NueVer
   
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_LisEEFF)
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia_NueVer()
   cmd_ExpExc.Enabled = False
   lblConcepto.Caption = moddat_g_str_NomPrd
   
   'LISTADO INGRESOS MANUALES
   grd_LisEEFF.ColWidth(0) = 1020   ' CUENTA
   grd_LisEEFF.ColWidth(1) = 3030   ' DESCRIPCION

   Call gs_LimpiaGrid(grd_LisEEFF)
End Sub
Private Sub fs_Inicia_AntVer()
   cmd_ExpExc.Enabled = False
   lblConcepto.Caption = moddat_g_str_NomPrd
   
   'LISTADO INGRESOS MANUALES
   grd_LisEEFF.ColWidth(0) = 1020   ' CUENTA
   grd_LisEEFF.ColWidth(1) = 3030   ' DESCRIPCION
   grd_LisEEFF.ColWidth(2) = 810    ' MES 1
   grd_LisEEFF.ColWidth(3) = 810    ' MES 2
   grd_LisEEFF.ColWidth(4) = 810    ' MES 3
   grd_LisEEFF.ColWidth(5) = 810    ' MES 4
   grd_LisEEFF.ColWidth(6) = 810    ' MES 5
   grd_LisEEFF.ColWidth(7) = 810    ' MES 6
   grd_LisEEFF.ColWidth(8) = 810    ' MES 7
   grd_LisEEFF.ColWidth(9) = 810    ' MES 8
   grd_LisEEFF.ColWidth(10) = 810   ' MES 9
   grd_LisEEFF.ColWidth(11) = 810   ' MES 10
   grd_LisEEFF.ColWidth(12) = 810   ' MES 11
   grd_LisEEFF.ColWidth(13) = 810   ' MES 12
   grd_LisEEFF.ColWidth(14) = 900   ' ACUMULADO AÑO ACTUAL
   grd_LisEEFF.ColWidth(15) = 0     ' INDTIPO
   
   grd_LisEEFF.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEEFF.ColAlignment(1) = flexAlignLeftCenter
   grd_LisEEFF.ColAlignment(2) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(3) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(4) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(5) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(6) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(7) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(8) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(9) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(10) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(11) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(12) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(13) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(14) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(15) = flexAlignRightCenter
    
   Call gs_LimpiaGrid(grd_LisEEFF)
End Sub
Private Sub fs_Buscar_NueVer()

   Dim r_int_VarAux1    As Integer
   Dim r_int_ConAux     As Integer
   
   Dim r_int_PerAnoi    As Integer
   Dim r_int_PerAnof    As Integer
   Dim r_int_PerMesi    As Integer
   
   Dim r_int_ConAnn     As Integer
   Dim r_int_ConMes     As Integer
   
   'Obteniendo Información
   If IsNull(moddat_g_str_CodPrd) Or moddat_g_str_CodPrd = "" And IsNull(moddat_g_str_CodSub) Or moddat_g_str_CodSub = "" Then
       Exit Sub
   End If
   
   r_int_PerAnoi = frm_RptCtb_19.ipp_PerAnoi.Text
   r_int_PerAnof = frm_RptCtb_19.ipp_PerAnof.Text
   
  
   'ÚLTIMO AÑO CONSULTADO
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TT_EEFF WHERE "
   g_str_Parame = g_str_Parame & "  GRUPO = " & CInt(moddat_g_str_CodPrd) & " "
   If moddat_g_str_TipCre = "S" Then
      g_str_Parame = g_str_Parame & "   AND subgrp = " & CInt(moddat_g_str_CodSub) & "  "
   End If
   g_str_Parame = g_str_Parame & "   AND INDTIPO <> 'L' "
   g_str_Parame = g_str_Parame & "   AND USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & " ORDER BY grupo, subgrp, item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_LisEEFF.Cols = frm_RptCtb_19.grd_LisEEFF.Cols - 3
   
   grd_LisEEFF.Rows = grd_LisEEFF.Rows + 2
   grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
   grd_LisEEFF.FixedRows = 1
    
   grd_LisEEFF.Row = 0
   grd_LisEEFF.Col = 0
   grd_LisEEFF.Text = "CUENTA"
   grd_LisEEFF.Col = 1
   grd_LisEEFF.Text = "DESCRIPCION"
   grd_LisEEFF.FixedCols = 2
   
   r_int_VarAux1 = 4
   For r_int_ConAux = 2 To grd_LisEEFF.Cols - 1
      grd_LisEEFF.Col = r_int_ConAux
      grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, r_int_VarAux1): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
      r_int_VarAux1 = r_int_VarAux1 + 1
   Next r_int_ConAux
 
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If (Trim(g_rst_Princi!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_Princi!GRUPO)) Then
             GoTo SALTO
         End If
         
         grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
         grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
         
         grd_LisEEFF.Col = 0
         grd_LisEEFF.Text = Trim(g_rst_Princi!CNTACTBLE & "")
         grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
         grd_LisEEFF.CellFontBold = False
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 7
         
         grd_LisEEFF.Col = 1
         If Trim(g_rst_Princi!INDTIPO) = "S" Then
            grd_LisEEFF.Text = Space(5) & UCase(Trim(g_rst_Princi!NOMSUBGRP))
         ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
            grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
         Else
            grd_LisEEFF.Text = IIf(Len(Trim(g_rst_Princi!NOMCTA & "")) > 0, " - " & Trim(g_rst_Princi!NOMCTA & ""), "")
         End If
         
         grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
         grd_LisEEFF.CellFontBold = False
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 8
         
         r_int_VarAux1 = grd_LisEEFF.Cols - 1
         
            If frm_RptCtb_19.cmb_PerMesf.ListIndex = 11 Then                  'DICIEMBRE
               r_int_VarAux1 = r_int_VarAux1 - 6
               grd_LisEEFF.Col = r_int_VarAux1 + 4
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES12, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

S1:
               grd_LisEEFF.Col = r_int_VarAux1 + 3
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES11, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250
   
S2:
               grd_LisEEFF.Col = r_int_VarAux1 + 2
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES10, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

S3:
               grd_LisEEFF.Col = r_int_VarAux1 + 1
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES09, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

S4:
               grd_LisEEFF.Col = r_int_VarAux1
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES08, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

S5:
               grd_LisEEFF.Col = r_int_VarAux1 - 1
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES07, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 1) = 1250

S6:
               grd_LisEEFF.Col = r_int_VarAux1 - 2
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES06, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 2) = 1250

S7:
               grd_LisEEFF.Col = r_int_VarAux1 - 3
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES05, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 3) = 1250

S8:
               grd_LisEEFF.Col = r_int_VarAux1 - 4
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES04, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 4) = 1250

S9:
               grd_LisEEFF.Col = r_int_VarAux1 - 5
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES03, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 5) = 1250

S10:
               grd_LisEEFF.Col = r_int_VarAux1 - 6
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES02, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 6) = 1250

S11:
               grd_LisEEFF.Col = r_int_VarAux1 - 7
               grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES01, 2)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(r_int_VarAux1 - 7) = 1250
                    
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 10 Then           'NOVIEMBRE
             r_int_VarAux1 = r_int_VarAux1 - 5
             GoTo S1
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 9 Then            'OCTUBRE
             r_int_VarAux1 = r_int_VarAux1 - 4
             GoTo S2
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 8 Then            'SETIEMBRE
             r_int_VarAux1 = r_int_VarAux1 - 3
             GoTo S3
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 7 Then            'AGOSTO
             r_int_VarAux1 = r_int_VarAux1 - 2
             GoTo S4
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 6 Then            'JULIO
             r_int_VarAux1 = r_int_VarAux1 - 1
             GoTo S5
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 5 Then            'JUNIO
             r_int_VarAux1 = r_int_VarAux1
             GoTo S6
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 4 Then            'MAYO
             r_int_VarAux1 = r_int_VarAux1 + 1
             GoTo S7
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 3 Then            'ABRIL
             r_int_VarAux1 = r_int_VarAux1 + 2
             GoTo S8
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 2 Then            'MARZO
             r_int_VarAux1 = r_int_VarAux1 + 3
             GoTo S9
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 1 Then            'FEBRERO
             r_int_VarAux1 = r_int_VarAux1 + 4
             GoTo S10
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 0 Then            'ENERO
             r_int_VarAux1 = r_int_VarAux1 + 5
             GoTo S11
         End If

Salir:

         grd_LisEEFF.Col = grd_LisEEFF.Cols - 2
         If CInt(frm_RptCtb_19.cmb_PerMesi.ListIndex) > 0 Then
            grd_LisEEFF.Text = FormatNumber(Sumar(grd_LisEEFF, grd_LisEEFF.Row), 2)
         Else
            grd_LisEEFF.Text = FormatNumber(g_rst_Princi!ACUMU, 2)
         End If
         grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 2) = 1250
         
         grd_LisEEFF.Col = grd_LisEEFF.Cols - 1
         grd_LisEEFF.Text = Trim(g_rst_Princi!INDTIPO)
         grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 1) = 0
                       
SALTO:
         g_rst_Princi.MoveNext
      Loop
     
   End If
   
   'AÑO ANTERIOR AL ÚLTIMO AÑO CONSULTADO
   
   For r_int_ConAnn = r_int_PerAnoi To r_int_PerAnof
      grd_LisEEFF.Row = 1
      
      If r_int_ConAnn > r_int_PerAnoi And r_int_ConAnn <> r_int_PerAnof Then
         g_rst_GenAux.MoveFirst
         g_rst_GenAux.Find "anno = '" & r_int_ConAnn & "'"
         r_int_ConMes = 0
         
         If r_int_VarAux1 = 2 Then
            r_int_VarAux1 = (12 - (frm_RptCtb_19.cmb_PerMesi.ListIndex + 1) + 1) + 2
         Else
            r_int_VarAux1 = r_int_VarAux1 + 12
         End If

         GoTo Ingresar1
         
      ElseIf r_int_ConAnn = r_int_PerAnof Then
         Exit For
         
      Else
         g_rst_GenAux.MoveFirst
         g_rst_GenAux.Find "anno = '" & r_int_ConAnn & "'"
         r_int_ConMes = frm_RptCtb_19.cmb_PerMesi.ListIndex
         r_int_VarAux1 = 2

         'Informaciòn del recordset no conectado
         If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
            'g_rst_GenAux.MoveFirst
            Do While Not g_rst_GenAux.EOF
Ingresar1:
              If CInt(g_rst_GenAux!anno) <> r_int_ConAnn Then GoTo SALTO1

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
                          
                     If grd_LisEEFF.Row < grd_LisEEFF.Rows - 1 Then
                        grd_LisEEFF.Row = grd_LisEEFF.Row + 1
                     End If
                          
                     If r_int_ConMes = 11 Then           'DICIMEBRE

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250
                                  
                     ElseIf r_int_ConMes = 10 Then       'NOVIEMBRE
                        
                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250
                        
                     ElseIf r_int_ConMes = 9 Then        'OCTUBRE
                        
                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250
                        
                     ElseIf r_int_ConMes = 8 Then        'SEPTIEMBRE
                        
                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250
                        
                     ElseIf r_int_ConMes = 7 Then        'AGOSTO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250
                        
                     ElseIf r_int_ConMes = 6 Then        'JULIO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250
                        
                     ElseIf r_int_ConMes = 5 Then        'JUNIO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250
                        
                     ElseIf r_int_ConMes = 4 Then        'MAYO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES05, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1250
                        
                     ElseIf r_int_ConMes = 3 Then        'ABRIL

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES04, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES05, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 8) = 1250
                        
                     ElseIf r_int_ConMes = 2 Then        'MARZO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES03, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES04, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES05, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 8) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 9) = 1250
                        
                     ElseIf r_int_ConMes = 1 Then        'FEBRERO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES02, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES03, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES04, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES05, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 8) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 9) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 10) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 10) = 1250
                        
                     ElseIf r_int_ConMes = 0 Then        'ENERO

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES01, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1) = 1250
                        
                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES02, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 1) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES03, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 2) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES04, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES05, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES06, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES07, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES08, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES09, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 8) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES10, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 9) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 10) = FormatNumber(g_rst_GenAux!MES11, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 10) = 1250

                        grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, r_int_VarAux1 + 11) = FormatNumber(g_rst_GenAux!MES12, 2)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(r_int_VarAux1 + 11) = 1250
                         
                     End If
                  End If
               End If
SALTO1:
               g_rst_GenAux.MoveNext
            Loop

         End If
      End If
   Next r_int_ConAnn
     
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   grd_LisEEFF.Redraw = True
   If grd_LisEEFF.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEFF)
      cmd_ExpExc.Enabled = True
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
   
End Sub
Private Sub fs_Buscar_AntVer()
Dim p            As Integer
Dim q            As Integer
Dim k            As Integer
Dim anovigente   As Integer
Dim mesvigente   As Integer
   
   If IsNull(moddat_g_str_CodPrd) Or moddat_g_str_CodPrd = "" And IsNull(moddat_g_str_CodSub) Or moddat_g_str_CodSub = "" Then
       Exit Sub
   End If
   
   'Consulta para obtener el mes y año vigente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT MAX(PERMES_CODMES) AS MES, MAX(PERMES_CODANO) AS ANO "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "Error al ejecutar la consulta para obtener año vigente.", vbCritical, modgen_g_str_NomPlt
   End If
     
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      anovigente = g_rst_Genera!ANO
      mesvigente = g_rst_Genera!Mes
   End If
        
   'Obteniendo Información
   g_str_Parame = "SELECT * FROM TT_EEFF WHERE "
   g_str_Parame = g_str_Parame & " grupo = " & CInt(moddat_g_str_CodPrd) & " "
   If moddat_g_str_TipCre = "S" Then
      g_str_Parame = g_str_Parame & " AND subgrp = " & CInt(moddat_g_str_CodSub) & "  "
   End If
   g_str_Parame = g_str_Parame & "   AND USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "   ORDER BY grupo, subgrp, item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_LisEEFF.Rows = grd_LisEEFF.Rows + 2
   grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
   grd_LisEEFF.FixedRows = 1
   
   grd_LisEEFF.Row = 0
   grd_LisEEFF.Col = 0
   grd_LisEEFF.Text = "CUENTA"
   grd_LisEEFF.Col = 1
   grd_LisEEFF.Text = "DESCRIPCION"
   grd_LisEEFF.Col = 2
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"ENERO"
   grd_LisEEFF.Col = 3
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"FEBRERO"
   grd_LisEEFF.Col = 4
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"MARZO"
   grd_LisEEFF.Col = 5
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"ABRIL"
   grd_LisEEFF.Col = 6
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"MAYO"
   grd_LisEEFF.Col = 7
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9): grd_LisEEFF.CellAlignment = flexAlignCenterCenter     '"JUNIO"
   grd_LisEEFF.Col = 8
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"JULIO"
   grd_LisEEFF.Col = 9
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"AGOSTO"
   grd_LisEEFF.Col = 10
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"SETIEMBRE"
   grd_LisEEFF.Col = 11
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"OCTUBRE"
   grd_LisEEFF.Col = 12
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"NOVIEMBRE"
   grd_LisEEFF.Col = 13
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"DICIEMBRE"
   grd_LisEEFF.Col = 14
   grd_LisEEFF.Text = frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 16): grd_LisEEFF.CellAlignment = flexAlignCenterCenter    '"ACUMULADO-AÑO-ACTUAL"
    
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If Trim(g_rst_Princi!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_Princi!GRUPO) Then
             GoTo SALTO
         End If
         
         grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
         grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
         
         grd_LisEEFF.Col = 0
         grd_LisEEFF.Text = Trim(g_rst_Princi!CNTACTBLE & "")
         grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
         grd_LisEEFF.CellFontBold = True
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 7
         
         If Trim(g_rst_Princi!INDTIPO) = "S" Then
            grd_LisEEFF.Col = 1
            grd_LisEEFF.Text = Space(5) & UCase(Trim(g_rst_Princi!NOMSUBGRP))
         ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
            grd_LisEEFF.Col = 1
            grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
         Else
            grd_LisEEFF.Col = 1
            grd_LisEEFF.Text = IIf(Len(Trim(g_rst_Princi!NOMCTA & "")) > 0, " - " & Trim(g_rst_Princi!NOMCTA & ""), "")
         End If
         grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
         grd_LisEEFF.CellFontBold = True
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 7
         
         'PARA EL AÑO ACTUAL
         If frm_RptCtb_19.cmb_PerMesf.ListIndex = 11 Then     'DICIEMBRE
            p = 3
            grd_LisEEFF.Col = p + 10
            grd_LisEEFF.Text = Format(g_rst_Princi!MES12, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S1:
            grd_LisEEFF.Col = p + 9
            grd_LisEEFF.Text = Format(g_rst_Princi!MES11, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S2:
            grd_LisEEFF.Col = p + 8
            grd_LisEEFF.Text = Format(g_rst_Princi!MES10, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S3:
            grd_LisEEFF.Col = p + 7
            grd_LisEEFF.Text = Format(g_rst_Princi!MES09, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S4:
            grd_LisEEFF.Col = p + 6
            grd_LisEEFF.Text = Format(g_rst_Princi!MES08, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S5:
            grd_LisEEFF.Col = p + 5
            grd_LisEEFF.Text = Format(g_rst_Princi!MES07, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S6:
            grd_LisEEFF.Col = p + 4
            grd_LisEEFF.Text = Format(g_rst_Princi!MES06, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S7:
            grd_LisEEFF.Col = p + 3
            grd_LisEEFF.Text = Format(g_rst_Princi!MES05, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S8:
            grd_LisEEFF.Col = p + 2
            grd_LisEEFF.Text = Format(g_rst_Princi!MES04, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S9:
            grd_LisEEFF.Col = p + 1
            grd_LisEEFF.Text = Format(g_rst_Princi!MES03, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S10:
            grd_LisEEFF.Col = p
            grd_LisEEFF.Text = Format(g_rst_Princi!MES02, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
S11:
            grd_LisEEFF.Col = p - 1
            grd_LisEEFF.Text = Format(g_rst_Princi!MES01, "###,###,###,##0")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
                
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 10 Then 'NOVIEMBRE
             p = 4
             GoTo S1
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 9 Then  'OCTUBRE
             p = 5
             GoTo S2
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 8 Then  'SETIEMBRE
             p = 6
             GoTo S3
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 7 Then  'AGOSTO
             p = 7
             GoTo S4
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 6 Then  'JULIO
             p = 8
             GoTo S5
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 5 Then  'JUNIO
             p = 9
             GoTo S6
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 4 Then  'MAYO
             p = 10
             GoTo S7
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 3 Then  'ABRIL
             p = 11
             GoTo S8
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 2 Then  'MARZO
             p = 12
             GoTo S9
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 1 Then  'FEBRERO
             p = 13
             GoTo S10
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 0 Then  'ENERO
             p = 14
             GoTo S11
         End If
            
         grd_LisEEFF.Col = 14
         If g_rst_Princi!ACUMU = 0 Then
            grd_LisEEFF.Text = Format(Sumar(grd_LisEEFF, grd_LisEEFF.Row), "###,###,###,##0")
         Else
            grd_LisEEFF.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0")
         End If
         
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 8
         
         grd_LisEEFF.Col = 15
         grd_LisEEFF.Text = Trim(g_rst_Princi!INDTIPO)
         grd_LisEEFF.CellFontName = "Arial"
         grd_LisEEFF.CellFontSize = 8
         
SALTO:
         If frm_RptCtb_19.ipp_PerAnof.Text = anovigente And p <> 0 Then
            For k = 1 To CInt(frm_RptCtb_19.cmb_PerMesf.ListIndex + 1 - mesvigente)
               grd_LisEEFF.Col = k + p
               grd_LisEEFF.Text = 0
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
            Next
         End If

         g_rst_Princi.MoveNext
      Loop
   End If
   
   'AÑO ANTERIOR
   grd_LisEEFF.Row = 1
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
      
         If CInt(g_rst_GenAux!GRUPO) = CInt(moddat_g_str_CodPrd) Then
            If moddat_g_str_TipCre = "S" Then
               If g_rst_GenAux!SUBGRP = CInt(moddat_g_str_CodSub) Then
                   GoTo Ingresar
               End If
            Else
Ingresar:
               If (Trim(g_rst_GenAux!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_GenAux!GRUPO)) Then
                   GoTo SALTO1
               End If
               
               If grd_LisEEFF.Row < grd_LisEEFF.Rows - 1 Then
                   grd_LisEEFF.Row = grd_LisEEFF.Row + 1
               End If
            
               If frm_RptCtb_19.cmb_PerMesi.ListIndex = 0 Then      'ENERO
                  q = 2

                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q) = Format(g_rst_GenAux!MES02, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S12:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 1) = Format(g_rst_GenAux!MES03, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S13:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 2) = Format(g_rst_GenAux!MES04, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S14:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 3) = Format(g_rst_GenAux!MES05, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S15:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 4) = Format(g_rst_GenAux!MES06, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S16:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 5) = Format(g_rst_GenAux!MES07, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S17:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 6) = Format(g_rst_GenAux!MES08, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S18:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 7) = Format(g_rst_GenAux!MES09, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S19:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 8) = Format(g_rst_GenAux!MES10, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S20:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 9) = Format(g_rst_GenAux!MES11, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S21:
                  grd_LisEEFF.TextMatrix(grd_LisEEFF.Row, q + 10) = Format(g_rst_GenAux!MES12, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                        
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 1 Then  'FEBRERO
                   q = 1
                   GoTo S12
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 2 Then  'MARZO
                   q = 0
                   GoTo S13
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 3 Then  'ABRIL
                   q = -1
                   GoTo S14
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 4 Then  'MAYO
                   q = -2
                   GoTo S15
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 5 Then  'JUNIO
                   q = -3
                   GoTo S16
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 6 Then  'JULIO
                   q = -4
                   GoTo S17
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 7 Then  'AGOSTO
                   q = -5
                   GoTo S18
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 8 Then  'SEPTIEMBRE
                   q = -6
                   GoTo S19
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 9 Then  'OCTUBRE
                   q = -7
                   GoTo S20
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 10 Then  'NOVIEMBRE
                   q = -8
                   GoTo S21
               End If
            End If
         End If
      
SALTO1:
         g_rst_GenAux.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_LisEEFF.Redraw = True
   If grd_LisEEFF.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEFF)
      cmd_ExpExc.Enabled = True
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Function Sumar(MSHFlexGrid As Object, Fila As Integer) As Currency
   On Error GoTo error_function
  
   With MSHFlexGrid
        Dim total As Currency
        Dim i As Long
        For i = 1 To .Cols - 1 '.Rows - 1
            If IsNumeric(.TextMatrix(Fila, i)) Then
                total = total + .TextMatrix(Fila, i)
            End If
        Next
        Sumar = total
    End With
       
   Exit Function
   
error_function:
   MsgBox Err.Description, vbCritical, "error al sumar"
End Function
Private Sub fs_GenExc_NueVer()

Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer

Dim r_int_ConAux        As Integer
            
   r_int_nrofil = 5
    
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
       
       r_str_FecRpt = Format(ff_Ultimo_Dia_Mes(frm_RptCtb_19.cmb_PerMesf.ListIndex + 1, frm_RptCtb_19.ipp_PerAnof.Text), "00") & "/" & Format(frm_RptCtb_19.cmb_PerMesf.ListIndex + 1, "00") & "/" & frm_RptCtb_19.ipp_PerAnof.Text

      .Cells(2, 2) = "Del " & "01 de " & Left(frm_RptCtb_19.cmb_PerMesi.Text, 1) & LCase(Mid(frm_RptCtb_19.cmb_PerMesi.Text, 2, Len(frm_RptCtb_19.cmb_PerMesi.Text))) & " del " & frm_RptCtb_19.ipp_PerAnoi.Text & " Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(frm_RptCtb_19.cmb_PerMesf.Text, 1) & LCase(Mid(frm_RptCtb_19.cmb_PerMesf.Text, 2, Len(frm_RptCtb_19.cmb_PerMesf.Text))) & " del " & Format(frm_RptCtb_19.ipp_PerAnof.Text, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(3, 2) = "( En Soles )"
      .Cells(5, 2) = "EJERCICIOS"
      
      For r_int_ConAux = 4 To grd_LisEEFF.Cols
         .Cells(r_int_nrofil, r_int_ConAux) = "'" & frm_RptCtb_20.grd_LisEEFF.TextMatrix(0, r_int_ConAux - 2) & ""
      Next r_int_ConAux
   
        
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_ConAux - 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_ConAux - 1)).Font.Bold = True
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 37
      
       For r_int_ConAux = 4 To grd_LisEEFF.Cols
           .Columns(r_int_ConAux).HorizontalAlignment = xlHAlignRight
           .Columns(r_int_ConAux).NumberFormat = "###,###,###,##0.00"
           .Columns(r_int_ConAux).ColumnWidth = 14.5
       Next r_int_ConAux

     
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      .Range(.Cells(r_int_nrofil, 3), .Cells(r_int_nrofil, r_int_ConAux)).HorizontalAlignment = xlHAlignCenter

      r_int_nrofil = r_int_nrofil + 2
      For r_int_NoFlLi = 2 To grd_LisEEFF.Rows - 1
         If Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, r_int_ConAux - 2)) = "G" Then
            'TITULO
            .Cells(r_int_nrofil, 2) = grd_LisEEFF.TextMatrix(r_int_NoFlLi, 1)
            .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_ConAux - 1)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_ConAux - 1)).Font.Bold = True
            
         ElseIf Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, grd_LisEEFF.Cols - 1)) = "S" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, grd_LisEEFF.Cols - 1)) = "T" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, grd_LisEEFF.Cols - 1)) = "N" Then
             .Cells(r_int_nrofil, 3) = Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 1))
             .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_ConAux - 1)).Font.Bold = True
         Else
             .Cells(r_int_nrofil, 2) = "'" & Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 0))
             .Cells(r_int_nrofil, 3) = Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 1))
         End If

         For r_int_Contad = 4 To grd_LisEEFF.Cols
               .Cells(r_int_nrofil, r_int_Contad) = grd_LisEEFF.TextMatrix(r_int_NoFlLi, r_int_Contad - 2)
         Next r_int_Contad

         r_int_nrofil = r_int_nrofil + 1
      Next r_int_NoFlLi
      
      .Columns("C:C").EntireColumn.AutoFit
   End With
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
End Sub
Private Sub fs_GenExc_AntVer()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim p                   As Integer
Dim q                   As Integer
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = moddat_g_str_FecIng
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      
      .Cells(4, 2) = "EJERCICIOS"
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4), "-") = 0 Then
         .Cells(4, 4) = "'" & "ENE " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(4, 4) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5), "-") = 0 Then
          .Cells(4, 5) = "'" & "FEB " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 5) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6), "-") = 0 Then
         .Cells(4, 6) = "'" & "MAR " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(4, 6) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7), "-") = 0 Then
         .Cells(4, 7) = "'" & "ABR " & Right(moddat_g_str_FecIng, 2)
      Else
         .Cells(4, 7) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8), "-") = 0 Then
          .Cells(4, 8) = "'" & "MAY " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 8) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9), "-") = 0 Then
          .Cells(4, 9) = "'" & "JUN " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 9) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10), "-") = 0 Then
          .Cells(4, 10) = "'" & "JUL " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 10) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11), "-") = 0 Then
          .Cells(4, 11) = "'" & "AGO " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 11) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12), "-") = 0 Then
          .Cells(4, 12) = "'" & "SET " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 12) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13), "-") = 0 Then
          .Cells(4, 13) = "'" & "OCT " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 13) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14), "-") = 0 Then
          .Cells(4, 14) = "'" & "NOV " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 14) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14) & ""
      End If

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15), "-") = 0 Then
          .Cells(4, 15) = "'" & "DIC " & Right(moddat_g_str_FecIng, 2)
      Else
          .Cells(4, 15) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15) & ""
      End If

      .Cells(4, 16) = "'" & "ACUM " & Right(moddat_g_str_FecIng, 2)
      .Range(.Cells(4, 2), .Cells(4, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 16)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      '.Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 37
      '.Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 11
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0"
      
      .Columns("E").ColumnWidth = 11
      .Columns("E").NumberFormat = "###,###,###,##0"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      
      .Columns("F").ColumnWidth = 11
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      
      .Columns("G").ColumnWidth = 11
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("H").HorizontalAlignment = xlHAlignRight
            
      .Columns("I").ColumnWidth = 11
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      
      .Columns("J").ColumnWidth = 11
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("J").HorizontalAlignment = xlHAlignRight
            
      .Columns("K").ColumnWidth = 11
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("K").HorizontalAlignment = xlHAlignRight
            
      .Columns("L").ColumnWidth = 11
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 11
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 11
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 11
      .Columns("O").NumberFormat = "###,###,###,##0"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 12
      .Columns("P").NumberFormat = "###,###,###,##0"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      'Obteniendo Información
      g_str_Parame = "SELECT * FROM TT_EEFF WHERE "
      g_str_Parame = g_str_Parame & " GRUPO = " & CInt(moddat_g_str_CodPrd) & " "
      If moddat_g_str_TipCre = "S" Then
         g_str_Parame = g_str_Parame & " AND SUBGRP = " & CInt(moddat_g_str_CodSub) & "  "
      End If
      g_str_Parame = g_str_Parame & "   AND USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "
             
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_Contad = 4
      Do While Not g_rst_Princi.EOF
         If Trim(g_rst_Princi!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_Princi!GRUPO) Then
             GoTo SALTO
         End If
             
         r_int_Contad = r_int_Contad + 1
         If Trim(g_rst_Princi!INDTIPO) = "L" Then
            g_rst_Princi.MoveNext
            r_int_Contad = r_int_Contad + 1
            If g_rst_Princi.EOF Then
               GoTo FIN
            End If
         End If
         If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
            .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Font.Bold = True
         End If
         If Trim(g_rst_Princi!INDTIPO) = "S" Then
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Font.Bold = True
         End If
         If Trim(g_rst_Princi!INDTIPO) = "D" Then
            .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
         End If
          
         If frm_RptCtb_19.cmb_PerMesf.ListIndex = 11 Then     'DICIEMBRE
            p = 4
            .Cells(r_int_Contad, p + 11) = g_rst_Princi!MES12
S1:
            .Cells(r_int_Contad, p + 10) = g_rst_Princi!MES11
S2:
            .Cells(r_int_Contad, p + 9) = g_rst_Princi!MES10
S3:
            .Cells(r_int_Contad, p + 8) = g_rst_Princi!MES09
S4:
            .Cells(r_int_Contad, p + 7) = g_rst_Princi!MES08
S5:
            .Cells(r_int_Contad, p + 6) = g_rst_Princi!MES07
S6:
            .Cells(r_int_Contad, p + 5) = g_rst_Princi!MES06
S7:
            .Cells(r_int_Contad, p + 4) = g_rst_Princi!MES05
S8:
            .Cells(r_int_Contad, p + 3) = g_rst_Princi!MES04
S9:
            .Cells(r_int_Contad, p + 2) = g_rst_Princi!MES03
S10:
            .Cells(r_int_Contad, p + 1) = g_rst_Princi!MES02
S11:
            .Cells(r_int_Contad, p) = g_rst_Princi!MES01
                
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 10 Then     'NOVIEMBRE
            p = 5
            GoTo S1
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 9 Then      'OCTUBRE
            p = 6
            GoTo S2
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 8 Then      'SETIEMBRE
            p = 7
            GoTo S3
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 7 Then      'AGOSTO
            p = 8
            GoTo S4
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 6 Then      'JULIO
            p = 9
            GoTo S5
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 5 Then      'JUNIO
            p = 10
            GoTo S6
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 4 Then      'MAYO
            p = 11
            GoTo S7
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 3 Then      'ABRIL
            p = 12
            GoTo S8
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 2 Then      'MARZO
            p = 13
            GoTo S9
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 1 Then      'FEBRERO
            p = 14
            GoTo S10
         ElseIf frm_RptCtb_19.cmb_PerMesf.ListIndex = 0 Then      'ENERO
            p = 15
            GoTo S11
         End If
       
         Dim total As Currency
       
         If g_rst_Princi!ACUMU = 0 Then
            total = g_rst_Princi!MES01 + g_rst_Princi!MES02 + g_rst_Princi!MES03 + g_rst_Princi!MES04 + g_rst_Princi!MES05 + g_rst_Princi!MES06 + g_rst_Princi!MES07 + g_rst_Princi!MES08 + _
                     g_rst_Princi!MES09 + g_rst_Princi!MES10 + g_rst_Princi!MES11 + g_rst_Princi!MES12
            .Cells(r_int_Contad, 16) = total
         Else
            .Cells(r_int_Contad, 16) = g_rst_Princi!ACUMU
         End If
         
SALTO:
         g_rst_Princi.MoveNext
         
FIN:
         DoEvents
      Loop
        
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'AÑO ANTERIOR
      g_rst_GenAux.MoveFirst
      r_int_Contad = 4
     
      Do While Not g_rst_GenAux.EOF
 
         If g_rst_GenAux!GRUPO = CInt(moddat_g_str_CodPrd) Then
            If moddat_g_str_TipCre = "S" Then
               If g_rst_GenAux!SUBGRP = CInt(moddat_g_str_CodSub) Then
                  GoTo Ingresar
               End If
            Else
Ingresar:
               If Trim(g_rst_GenAux!INDTIPO) = "G" And moddat_g_str_TipCre = "S" And CInt(moddat_g_str_CodPrd) = Trim(g_rst_GenAux!GRUPO) Then
                  GoTo SALTO1
               End If
               
               r_int_Contad = r_int_Contad + 1
               If Trim(g_rst_GenAux!INDTIPO) = "L" Then
                  g_rst_GenAux.MoveNext
                  r_int_Contad = r_int_Contad + 1
                  If g_rst_GenAux!GRUPO <> CInt(moddat_g_str_CodPrd) Then
                     GoTo FIN1
                  End If
               End If
               
               If frm_RptCtb_19.cmb_PerMesi.ListIndex = 11 Then
                  'q = 4
                  GoTo SALTO1
              
                  .Cells(r_int_Contad, q - 12) = g_rst_GenAux!MES01
S12:
                  .Cells(r_int_Contad, q - 11) = g_rst_GenAux!MES02
S13:
                  .Cells(r_int_Contad, q - 10) = g_rst_GenAux!MES03
S14:
                  .Cells(r_int_Contad, q - 9) = g_rst_GenAux!MES04
S15:
                  .Cells(r_int_Contad, q - 8) = g_rst_GenAux!MES05
S16:
                  .Cells(r_int_Contad, q - 7) = g_rst_GenAux!MES06
S17:
                  .Cells(r_int_Contad, q - 6) = g_rst_GenAux!MES07
S18:
                  .Cells(r_int_Contad, q - 5) = g_rst_GenAux!MES08
S19:
                  .Cells(r_int_Contad, q - 4) = g_rst_GenAux!MES09
S20:
                  .Cells(r_int_Contad, q - 3) = g_rst_GenAux!MES10
S21:
                  .Cells(r_int_Contad, q - 2) = g_rst_GenAux!MES11
S22:
                  .Cells(r_int_Contad, q - 1) = g_rst_GenAux!MES12
                    
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 10 Then
                  q = 5
                  GoTo S22
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 9 Then
                  q = 6
                  GoTo S21
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 8 Then
                  q = 7
                  GoTo S20
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 7 Then
                  q = 8
                  GoTo S19
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 6 Then
                  q = 9
                  GoTo S18
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 5 Then
                  q = 10
                  GoTo S17
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 4 Then
                  q = 11
                  GoTo S16
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 3 Then
                  q = 12
                  GoTo S15
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 2 Then
                  q = 13
                  GoTo S14
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 1 Then
                  q = 14
                  GoTo S13
               ElseIf frm_RptCtb_19.cmb_PerMesi.ListIndex = 0 Then
                  q = 15
                  GoTo S12
               End If
            End If
         End If
    
SALTO1:
         g_rst_GenAux.MoveNext
FIN1:
         DoEvents
      Loop
   End With
           
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
