VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_27 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   6720
   ClientTop       =   6675
   ClientWidth     =   14175
   Icon            =   "GesCtb_frm_853.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7335
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14175
      _Version        =   65536
      _ExtentX        =   25003
      _ExtentY        =   12938
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
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   14055
         _Version        =   65536
         _ExtentX        =   24791
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   7
            Top             =   150
            Width           =   5205
            _Version        =   65536
            _ExtentX        =   9181
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Resumen de Provisiones"
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
            Picture         =   "GesCtb_frm_853.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   780
         Width           =   14055
         _Version        =   65536
         _ExtentX        =   24791
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_853.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_853.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13440
            Picture         =   "GesCtb_frm_853.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   885
         Left            =   60
         TabIndex        =   9
         Top             =   1470
         Width           =   14055
         _Version        =   65536
         _ExtentX        =   24791
         _ExtentY        =   1561
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
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   135
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1170
            TabIndex        =   1
            Top             =   450
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
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   165
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   495
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4860
         Left            =   60
         TabIndex        =   12
         Top             =   2400
         Width           =   14055
         _Version        =   65536
         _ExtentX        =   24791
         _ExtentY        =   8572
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
         Begin TabDlg.SSTab tab_Clasif 
            Height          =   4695
            Left            =   90
            TabIndex        =   13
            Top             =   90
            Width           =   13875
            _ExtentX        =   24474
            _ExtentY        =   8281
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "General"
            TabPicture(0)   =   "GesCtb_frm_853.frx":0D6C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_LisCla"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Por Tipo"
            TabPicture(1)   =   "GesCtb_frm_853.frx":0D88
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisTip"
            Tab(1).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_LisCla 
               Height          =   4245
               Left            =   90
               TabIndex        =   14
               Top             =   390
               Width           =   13710
               _ExtentX        =   24183
               _ExtentY        =   7488
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               Redraw          =   -1  'True
               MergeCells      =   1
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisTip 
               Height          =   4245
               Left            =   -74910
               TabIndex        =   15
               Top             =   390
               Width           =   13710
               _ExtentX        =   24183
               _ExtentY        =   7488
               _Version        =   393216
               Rows            =   7
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               Redraw          =   -1  'True
               MergeCells      =   1
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
      End
   End
End
Attribute VB_Name = "frm_RptCtb_27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_Mes01      As Double
Dim l_dbl_Mes02      As Double
Dim l_dbl_Mes03      As Double
Dim l_dbl_Mes04      As Double
Dim l_dbl_Mes05      As Double
Dim l_dbl_Mes06      As Double
Dim l_dbl_Mes07      As Double
Dim l_dbl_Mes08      As Double
Dim l_dbl_Mes09      As Double
Dim l_dbl_Mes10      As Double
Dim l_dbl_Mes11      As Double
Dim l_dbl_Mes12      As Double

Private Sub cmd_Proces_Click()
Dim r_int_MesAct     As String
Dim r_int_AnoAct     As String
Dim r_int_MesSel     As String
Dim r_int_AnoSel     As String
Dim r_int_MesAnt     As String
Dim r_int_AnoAnt     As String
Dim r_int_NumFil     As Integer
Dim r_int_Count      As Integer
Dim r_dbl_Monto1     As Double
Dim r_dbl_Monto2     As Double
Dim r_dbl_Monto3     As Double
Dim r_dbl_Monto4     As Double
Dim r_dbl_Monto5     As Double
Dim r_dbl_Monto6     As Double
Dim r_dbl_Monto7     As Double
Dim r_dbl_Monto8     As Double
Dim r_dbl_Monto9     As Double
Dim r_dbl_Monto10    As Double
Dim r_dbl_Total      As Double

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
   r_int_MesAct = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   r_int_AnoAct = ipp_PerAno.Text
   l_dbl_Mes01 = 0
   l_dbl_Mes02 = 0
   l_dbl_Mes03 = 0
   l_dbl_Mes04 = 0
   l_dbl_Mes05 = 0
   l_dbl_Mes06 = 0
   l_dbl_Mes07 = 0
   l_dbl_Mes08 = 0
   l_dbl_Mes09 = 0
   l_dbl_Mes10 = 0
   l_dbl_Mes11 = 0
   l_dbl_Mes12 = 0
   
   Me.Enabled = False
   
   'Muestra Provisiones General
   Call fs_Setea_ColumnasGeneral
   grd_LisCla.Redraw = False
   grd_LisTip.Redraw = False
   
   For r_int_Count = 1 To CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
      If r_int_Count > r_int_MesAct Then
         Exit For
      End If
      
      r_dbl_Total = 0
      r_dbl_Monto1 = 0
      r_dbl_Monto2 = 0
      r_dbl_Monto3 = 0
      r_dbl_Monto4 = 0
      r_dbl_Monto5 = 0
      r_dbl_Monto6 = 0
      r_dbl_Monto7 = 0
      r_dbl_Monto8 = 0
      r_dbl_Monto9 = 0
      r_dbl_Monto10 = 0
      r_int_MesSel = r_int_Count
      r_int_AnoSel = ipp_PerAno.Text
      r_int_MesAnt = IIf(r_int_MesSel = 1, 12, r_int_MesSel - 1)
      r_int_AnoAnt = IIf(r_int_MesSel = 1, CLng(ipp_PerAno.Text) - 1, ipp_PerAno.Text)
      
      Call fs_GenExc(r_int_MesSel, r_int_AnoSel, r_int_MesAnt, r_int_AnoAnt, r_dbl_Monto1, r_dbl_Monto2, r_dbl_Monto3, r_dbl_Monto4, r_dbl_Monto5, r_dbl_Monto6, r_dbl_Monto7, r_dbl_Monto8, r_dbl_Monto9, r_dbl_Monto10)
      
      grd_LisCla.Col = r_int_Count + 1
      grd_LisCla.Row = 2
      grd_LisCla.Text = Format(r_dbl_Monto1, "###,##0.00")
      grd_LisCla.Row = 3
      grd_LisCla.Text = Format(r_dbl_Monto2, "###,##0.00")
      grd_LisCla.Row = 6
      grd_LisCla.Text = Format(r_dbl_Monto3, "###,##0.00")
      grd_LisCla.Row = 7
      grd_LisCla.Text = Format(r_dbl_Monto4, "###,##0.00")
      grd_LisCla.Row = 8
      grd_LisCla.Text = Format(r_dbl_Monto5, "###,##0.00")
      grd_LisCla.Row = 9
      grd_LisCla.Text = Format(r_dbl_Monto6, "###,##0.00")
      grd_LisCla.Row = 10
      grd_LisCla.Text = Format(r_dbl_Monto7, "###,##0.00")
      grd_LisCla.Row = 11
      grd_LisCla.Text = Format(r_dbl_Monto8, "###,##0.00")
      grd_LisCla.Row = 12
      grd_LisCla.Text = Format(r_dbl_Monto9, "###,##0.00")
      grd_LisCla.Row = 13
      grd_LisCla.Text = Format(r_dbl_Monto10, "###,##0.00")
      grd_LisCla.Row = 14
      grd_LisCla.Text = Format(r_dbl_Monto1 + r_dbl_Monto2 + r_dbl_Monto3 + r_dbl_Monto4 + r_dbl_Monto5 + r_dbl_Monto6 + r_dbl_Monto7 + r_dbl_Monto8 + r_dbl_Monto9 + r_dbl_Monto10, "###,##0.00")
   Next r_int_Count
   
   'Muestra Provisiones Por Tipo
   Call fs_Setea_ColumnasTipo
   
   'Procesa infomacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_RPT_TIPOPROV("
   g_str_Parame = g_str_Parame & CInt(r_int_MesAct) & ", "
   g_str_Parame = g_str_Parame & CInt(r_int_AnoAct) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'TIPOPROVISION', "
   g_str_Parame = g_str_Parame & "0)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Consulta informacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * "
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_PERMES = '" & CInt(r_int_MesAct) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_PERANO = '" & CInt(r_int_AnoAct) & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'TIPOPROVISION' "
   g_str_Parame = g_str_Parame & "    AND RPT_MONEDA = 0 "
   g_str_Parame = g_str_Parame & "  ORDER BY RPT_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_NumFil = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If r_int_NumFil < 5 Then
            Select Case r_int_MesAct
               Case 1
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
               Case 2
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
               Case 3
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
               Case 4
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
               Case 5
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
               Case 6
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
               Case 7
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
               Case 8
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0.00")
               Case 9
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0.00")
               Case 10
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(g_rst_Princi!RPT_VALNUM10, "###,###,##0.00")
               Case 11
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(g_rst_Princi!RPT_VALNUM10, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(g_rst_Princi!RPT_VALNUM11, "###,###,##0.00")
               Case 12
                  grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(g_rst_Princi!RPT_VALNUM10, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(g_rst_Princi!RPT_VALNUM11, "###,###,##0.00")
                  grd_LisTip.TextMatrix(r_int_NumFil, 13) = Format(g_rst_Princi!RPT_VALNUM12, "###,###,##0.00")
            End Select
         End If
         
         r_int_NumFil = r_int_NumFil + 1
         g_rst_Princi.MoveNext
      Loop
      
      Select Case r_int_MesAct
         Case 1
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            
         Case 2
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            
         Case 3
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            
         Case 4
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            
         Case 5
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            
         Case 6
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            
         Case 7
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            
         Case 8
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(l_dbl_Mes08, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(grd_LisTip.TextMatrix(4, 9) - l_dbl_Mes08, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(CDbl(grd_LisTip.TextMatrix(1, 9)) + CDbl(grd_LisTip.TextMatrix(2, 9)) + CDbl(grd_LisTip.TextMatrix(3, 9)) + CDbl(grd_LisTip.TextMatrix(4, 9)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = ""
            
         Case 9
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(l_dbl_Mes09, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(grd_LisTip.TextMatrix(4, 9) - l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(grd_LisTip.TextMatrix(4, 10) - l_dbl_Mes09, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(CDbl(grd_LisTip.TextMatrix(1, 9)) + CDbl(grd_LisTip.TextMatrix(2, 9)) + CDbl(grd_LisTip.TextMatrix(3, 9)) + CDbl(grd_LisTip.TextMatrix(4, 9)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(CDbl(grd_LisTip.TextMatrix(1, 10)) + CDbl(grd_LisTip.TextMatrix(2, 10)) + CDbl(grd_LisTip.TextMatrix(3, 10)) + CDbl(grd_LisTip.TextMatrix(4, 10)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = ""
            
         Case 10
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(l_dbl_Mes10, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(grd_LisTip.TextMatrix(4, 9) - l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(grd_LisTip.TextMatrix(4, 10) - l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(grd_LisTip.TextMatrix(4, 11) - l_dbl_Mes10, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(CDbl(grd_LisTip.TextMatrix(1, 9)) + CDbl(grd_LisTip.TextMatrix(2, 9)) + CDbl(grd_LisTip.TextMatrix(3, 9)) + CDbl(grd_LisTip.TextMatrix(4, 9)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(CDbl(grd_LisTip.TextMatrix(1, 10)) + CDbl(grd_LisTip.TextMatrix(2, 10)) + CDbl(grd_LisTip.TextMatrix(3, 10)) + CDbl(grd_LisTip.TextMatrix(4, 10)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(CDbl(grd_LisTip.TextMatrix(1, 11)) + CDbl(grd_LisTip.TextMatrix(2, 11)) + CDbl(grd_LisTip.TextMatrix(3, 11)) + CDbl(grd_LisTip.TextMatrix(4, 11)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = ""
            
         Case 11
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(l_dbl_Mes10, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(l_dbl_Mes11, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(grd_LisTip.TextMatrix(4, 9) - l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(grd_LisTip.TextMatrix(4, 10) - l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(grd_LisTip.TextMatrix(4, 11) - l_dbl_Mes10, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(grd_LisTip.TextMatrix(4, 12) - l_dbl_Mes11, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(CDbl(grd_LisTip.TextMatrix(1, 9)) + CDbl(grd_LisTip.TextMatrix(2, 9)) + CDbl(grd_LisTip.TextMatrix(3, 9)) + CDbl(grd_LisTip.TextMatrix(4, 9)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(CDbl(grd_LisTip.TextMatrix(1, 10)) + CDbl(grd_LisTip.TextMatrix(2, 10)) + CDbl(grd_LisTip.TextMatrix(3, 10)) + CDbl(grd_LisTip.TextMatrix(4, 10)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(CDbl(grd_LisTip.TextMatrix(1, 11)) + CDbl(grd_LisTip.TextMatrix(2, 11)) + CDbl(grd_LisTip.TextMatrix(3, 11)) + CDbl(grd_LisTip.TextMatrix(4, 11)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(CDbl(grd_LisTip.TextMatrix(1, 12)) + CDbl(grd_LisTip.TextMatrix(2, 12)) + CDbl(grd_LisTip.TextMatrix(3, 12)) + CDbl(grd_LisTip.TextMatrix(4, 12)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = ""
            
         Case 12
            r_int_NumFil = 6
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(l_dbl_Mes10, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(l_dbl_Mes11, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 13) = Format(l_dbl_Mes12, "###,###,##0.00")
            r_int_NumFil = 5
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(grd_LisTip.TextMatrix(4, 2) - l_dbl_Mes01, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(grd_LisTip.TextMatrix(4, 3) - l_dbl_Mes02, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(grd_LisTip.TextMatrix(4, 4) - l_dbl_Mes03, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(grd_LisTip.TextMatrix(4, 5) - l_dbl_Mes04, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(grd_LisTip.TextMatrix(4, 6) - l_dbl_Mes05, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(grd_LisTip.TextMatrix(4, 7) - l_dbl_Mes06, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(grd_LisTip.TextMatrix(4, 8) - l_dbl_Mes07, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(grd_LisTip.TextMatrix(4, 9) - l_dbl_Mes08, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(grd_LisTip.TextMatrix(4, 10) - l_dbl_Mes09, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(grd_LisTip.TextMatrix(4, 11) - l_dbl_Mes10, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(grd_LisTip.TextMatrix(4, 12) - l_dbl_Mes11, "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 13) = Format(grd_LisTip.TextMatrix(4, 13) - l_dbl_Mes12, "###,###,##0.00")
            r_int_NumFil = 7
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = Format(CDbl(grd_LisTip.TextMatrix(1, 2)) + CDbl(grd_LisTip.TextMatrix(2, 2)) + CDbl(grd_LisTip.TextMatrix(3, 2)) + CDbl(grd_LisTip.TextMatrix(4, 2)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = Format(CDbl(grd_LisTip.TextMatrix(1, 3)) + CDbl(grd_LisTip.TextMatrix(2, 3)) + CDbl(grd_LisTip.TextMatrix(3, 3)) + CDbl(grd_LisTip.TextMatrix(4, 3)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = Format(CDbl(grd_LisTip.TextMatrix(1, 4)) + CDbl(grd_LisTip.TextMatrix(2, 4)) + CDbl(grd_LisTip.TextMatrix(3, 4)) + CDbl(grd_LisTip.TextMatrix(4, 4)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = Format(CDbl(grd_LisTip.TextMatrix(1, 5)) + CDbl(grd_LisTip.TextMatrix(2, 5)) + CDbl(grd_LisTip.TextMatrix(3, 5)) + CDbl(grd_LisTip.TextMatrix(4, 5)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = Format(CDbl(grd_LisTip.TextMatrix(1, 6)) + CDbl(grd_LisTip.TextMatrix(2, 6)) + CDbl(grd_LisTip.TextMatrix(3, 6)) + CDbl(grd_LisTip.TextMatrix(4, 6)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = Format(CDbl(grd_LisTip.TextMatrix(1, 7)) + CDbl(grd_LisTip.TextMatrix(2, 7)) + CDbl(grd_LisTip.TextMatrix(3, 7)) + CDbl(grd_LisTip.TextMatrix(4, 7)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = Format(CDbl(grd_LisTip.TextMatrix(1, 8)) + CDbl(grd_LisTip.TextMatrix(2, 8)) + CDbl(grd_LisTip.TextMatrix(3, 8)) + CDbl(grd_LisTip.TextMatrix(4, 8)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = Format(CDbl(grd_LisTip.TextMatrix(1, 9)) + CDbl(grd_LisTip.TextMatrix(2, 9)) + CDbl(grd_LisTip.TextMatrix(3, 9)) + CDbl(grd_LisTip.TextMatrix(4, 9)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = Format(CDbl(grd_LisTip.TextMatrix(1, 10)) + CDbl(grd_LisTip.TextMatrix(2, 10)) + CDbl(grd_LisTip.TextMatrix(3, 10)) + CDbl(grd_LisTip.TextMatrix(4, 10)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = Format(CDbl(grd_LisTip.TextMatrix(1, 11)) + CDbl(grd_LisTip.TextMatrix(2, 11)) + CDbl(grd_LisTip.TextMatrix(3, 11)) + CDbl(grd_LisTip.TextMatrix(4, 11)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = Format(CDbl(grd_LisTip.TextMatrix(1, 12)) + CDbl(grd_LisTip.TextMatrix(2, 12)) + CDbl(grd_LisTip.TextMatrix(3, 12)) + CDbl(grd_LisTip.TextMatrix(4, 12)), "###,###,##0.00")
            grd_LisTip.TextMatrix(r_int_NumFil, 13) = Format(CDbl(grd_LisTip.TextMatrix(1, 13)) + CDbl(grd_LisTip.TextMatrix(2, 13)) + CDbl(grd_LisTip.TextMatrix(3, 13)) + CDbl(grd_LisTip.TextMatrix(4, 13)), "###,###,##0.00")
            r_int_NumFil = 4
            grd_LisTip.TextMatrix(r_int_NumFil, 2) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 3) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 4) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 5) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 6) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 7) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 8) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 9) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 10) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 11) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 12) = ""
            grd_LisTip.TextMatrix(r_int_NumFil, 13) = ""
      End Select
   End If
   
   grd_LisTip.Redraw = True
   grd_LisCla.Redraw = True
   cmd_ExpExc.Enabled = True
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
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
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   cmd_ExpExc.Enabled = False
End Sub

Private Sub fs_Setea_ColumnasGeneral()
   grd_LisCla.Redraw = False
   Call gs_LimpiaGrid(grd_LisCla)
      
   'Ancho de columnas
   grd_LisCla.Cols = 14
   grd_LisCla.ColWidth(0) = 0
   grd_LisCla.ColWidth(1) = 4200
   grd_LisCla.ColWidth(2) = 1200
   grd_LisCla.ColWidth(3) = 1200
   grd_LisCla.ColWidth(4) = 1200
   grd_LisCla.ColWidth(5) = 1200
   grd_LisCla.ColWidth(6) = 1200
   grd_LisCla.ColWidth(7) = 1200
   grd_LisCla.ColWidth(8) = 1200
   grd_LisCla.ColWidth(9) = 1200
   grd_LisCla.ColWidth(10) = 1200
   grd_LisCla.ColWidth(11) = 1200
   grd_LisCla.ColWidth(12) = 1200
   grd_LisCla.ColWidth(13) = 1200
   grd_LisCla.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCla.ColAlignment(2) = flexAlignRightCenter
   grd_LisCla.ColAlignment(3) = flexAlignRightCenter
   grd_LisCla.ColAlignment(4) = flexAlignRightCenter
   grd_LisCla.ColAlignment(5) = flexAlignRightCenter
   grd_LisCla.ColAlignment(6) = flexAlignRightCenter
   grd_LisCla.ColAlignment(7) = flexAlignRightCenter
   grd_LisCla.ColAlignment(8) = flexAlignRightCenter
   grd_LisCla.ColAlignment(9) = flexAlignRightCenter
   grd_LisCla.ColAlignment(10) = flexAlignRightCenter
   grd_LisCla.ColAlignment(11) = flexAlignRightCenter
   grd_LisCla.ColAlignment(12) = flexAlignRightCenter
   grd_LisCla.ColAlignment(13) = flexAlignRightCenter
   
   'Cabecera
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Row = 0: grd_LisCla.Text = ""
   grd_LisCla.Col = 1: grd_LisCla.Text = "RESUMEN"
   grd_LisCla.Col = 2: grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 3: grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 4: grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 5: grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 6: grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 7: grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 8: grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 9: grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 10: grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 11: grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 12: grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 13: grd_LisCla.Text = "DICIEMBRE"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Créditos Comerciales"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "1"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Nuevos Desembolsos"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "2"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Por Amortización Comerciales"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "3"
   grd_LisCla.Col = 1:   grd_LisCla.Text = ""
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "4"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Créditos Hipotecarios"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "5"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Nuevos Desembolsos"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "6"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Amortización, Cancelación, Regresan, Salen"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "7"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Clientes que han ingresado como Morosos"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "8"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Clientes Morosos que revierten Provisión"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "9"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Clientes Morosos que han incrementado su provisión"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "10"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Clientes Morosos que ya están al día"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "11"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Provisión Neta por Clientes Alineados"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "12"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "   Provisión Voluntaria"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "13"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "TOTALES"
   
   With grd_LisCla
      .MergeCells = flexMergeFree
      .FixedCols = 2
      .FixedRows = 1
   End With
   grd_LisCla.Redraw = True
End Sub

Private Sub fs_Setea_ColumnasTipo()
   grd_LisTip.Redraw = False
   Call gs_LimpiaGrid(grd_LisTip)
      
   'Ancho de columnas
   grd_LisTip.Cols = 14
   grd_LisTip.ColWidth(0) = 0
   grd_LisTip.ColWidth(1) = 2400
   grd_LisTip.ColWidth(2) = 1200
   grd_LisTip.ColWidth(3) = 1200
   grd_LisTip.ColWidth(4) = 1200
   grd_LisTip.ColWidth(5) = 1200
   grd_LisTip.ColWidth(6) = 1200
   grd_LisTip.ColWidth(7) = 1200
   grd_LisTip.ColWidth(8) = 1200
   grd_LisTip.ColWidth(9) = 1200
   grd_LisTip.ColWidth(10) = 1200
   grd_LisTip.ColWidth(11) = 1200
   grd_LisTip.ColWidth(12) = 1200
   grd_LisTip.ColWidth(13) = 1200
   grd_LisTip.ColAlignment(1) = flexAlignLeftCenter
   grd_LisTip.ColAlignment(2) = flexAlignRightCenter
   grd_LisTip.ColAlignment(3) = flexAlignRightCenter
   grd_LisTip.ColAlignment(4) = flexAlignRightCenter
   grd_LisTip.ColAlignment(5) = flexAlignRightCenter
   grd_LisTip.ColAlignment(6) = flexAlignRightCenter
   grd_LisTip.ColAlignment(7) = flexAlignRightCenter
   grd_LisTip.ColAlignment(8) = flexAlignRightCenter
   grd_LisTip.ColAlignment(9) = flexAlignRightCenter
   grd_LisTip.ColAlignment(10) = flexAlignRightCenter
   grd_LisTip.ColAlignment(11) = flexAlignRightCenter
   grd_LisTip.ColAlignment(12) = flexAlignRightCenter
   grd_LisTip.ColAlignment(13) = flexAlignRightCenter
   
   'Cabecera
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Row = 0: grd_LisTip.Text = ""
   grd_LisTip.Col = 1: grd_LisTip.Text = "DETALLE"
   grd_LisTip.Col = 2: grd_LisTip.Text = "ENERO"
   grd_LisTip.Col = 3: grd_LisTip.Text = "FEBRERO"
   grd_LisTip.Col = 4: grd_LisTip.Text = "MARZO"
   grd_LisTip.Col = 5: grd_LisTip.Text = "ABRIL"
   grd_LisTip.Col = 6: grd_LisTip.Text = "MAYO"
   grd_LisTip.Col = 7: grd_LisTip.Text = "JUNIO"
   grd_LisTip.Col = 8: grd_LisTip.Text = "JULIO"
   grd_LisTip.Col = 9: grd_LisTip.Text = "AGOSTO"
   grd_LisTip.Col = 10: grd_LisTip.Text = "SETIEMBRE"
   grd_LisTip.Col = 11: grd_LisTip.Text = "OCTUBRE"
   grd_LisTip.Col = 12: grd_LisTip.Text = "NOVIEMBRE"
   grd_LisTip.Col = 13: grd_LisTip.Text = "DICIEMBRE"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "0"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "Provisión Genérica"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "1"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "Provisión Prociclica"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "2"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "Provisión Voluntaria"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "3"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "Provisión Específica"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "4"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "    Calidad de Cartera"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "5"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "    Alineados"
   
   grd_LisTip.Rows = grd_LisTip.Rows + 1
   grd_LisTip.Row = grd_LisTip.Rows - 1
   grd_LisTip.Col = 0:   grd_LisTip.Text = "6"
   grd_LisTip.Col = 1:   grd_LisTip.Text = "Total Provisión"
   
   With grd_LisTip
      .MergeCells = flexMergeFree
      .FixedCols = 2
      .FixedRows = 1
   End With
   grd_LisTip.Redraw = True
End Sub

Private Sub fs_GenExc(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer, ByRef p_Monto1 As Double, ByRef p_Monto2 As Double, ByRef p_Monto3 As Double, ByRef p_Monto4 As Double, ByRef p_Monto5 As Double, ByRef p_Monto6 As Double, ByRef p_Monto7 As Double, ByRef p_Monto8 As Double, ByRef p_Monto9 As Double, ByRef p_Monto10 As Double)
Dim r_int_Nindex     As Integer
Dim r_int_nroaux     As Integer
Dim r_bol_FlagOp     As Boolean
Dim r_dbl_NvDHip     As Double
Dim r_dbl_NvDCom     As Double
Dim r_dbl_PorHip     As Double
Dim r_dbl_xAmCom     As Double
Dim r_dbl_CMoros     As Double
Dim r_dbl_HipPrv     As Double
Dim r_dbl_CliMoI     As Double
Dim r_dbl_CMoDia     As Double
Dim r_dbl_PrvNet     As Double
Dim r_dbl_SumNor     As Double

   'inicializar valores
   r_dbl_NvDHip = 0
   r_dbl_NvDCom = 0
   r_dbl_PorHip = 0
   r_dbl_xAmCom = 0
   r_dbl_CMoros = 0
   r_dbl_HipPrv = 0
   r_dbl_CliMoI = 0
   r_dbl_CMoDia = 0
   r_dbl_PrvNet = 0
   r_dbl_SumNor = 0
   r_bol_FlagOp = False
   g_str_Parame = gf_Query(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       'MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
       Exit Sub
   End If

   Do While r_int_Nindex <= 4
      r_int_nroaux = 1
   
      Do While r_int_nroaux < 3
         g_rst_Princi.Filter = adFilterNone
         g_rst_Princi.MoveFirst
            
         If r_int_nroaux = 1 Then
             g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR > 2"
         Else
             g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR < 3"
         End If
         
         If g_rst_Princi.EOF Then
             r_bol_FlagOp = True
         End If
         
         Do While Not g_rst_Princi.EOF
             'Clientes que han ingresado como Morosos
             If CDbl(Format(g_rst_Princi!PROVISION2, "###,###,##0.00")) = 0 And (fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "NUEVO") Then
                 r_dbl_CMoros = r_dbl_CMoros + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
             End If
             
             'Por Hipotecario: Amortización, Cancelación, Regresan, Salen de Provisionados
             If CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00")) < 1 And (fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "MEJOR" Or fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "IGUAL") Then
                 r_dbl_HipPrv = r_dbl_HipPrv + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
             End If
             
             'Por Clientes Morosos Hipotecarios
             If CInt(g_rst_Princi!HIPCIE_CLACLI) = CInt(g_rst_Princi!HIPCIE_CLAALI) And CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00")) > 0 And CDbl(g_rst_Princi!PROVISION2) > 0 Then
                 r_dbl_CliMoI = r_dbl_CliMoI + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
             End If
             
             'Por Alineamiento a favor
             If CInt(g_rst_Princi!HIPCIE_CLACLI) <> CInt(g_rst_Princi!HIPCIE_CLAALI) And CInt(g_rst_Princi!HIPCIE_CLAPRV) > 1 Then
                 r_dbl_PrvNet = r_dbl_PrvNet + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
             End If
             
            'Clientes Normales (Suma de Provision Voluntaria)
            If g_rst_Princi!CLASIFICACION = "NOR" Then
               If g_rst_Princi!PROV_VOLUNT1 <> g_rst_Princi!PROV_VOLUNT2 Then
                  r_dbl_SumNor = r_dbl_SumNor + g_rst_Princi!PROV_VOLUNT1
               End If
            End If
             
             Select Case p_MesAct
                  Case 1: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes01 = l_dbl_Mes01 + g_rst_Princi!PROVISION
                  Case 2: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes02 = l_dbl_Mes02 + g_rst_Princi!PROVISION
                  Case 3: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes03 = l_dbl_Mes03 + g_rst_Princi!PROVISION
                  Case 4: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes04 = l_dbl_Mes04 + g_rst_Princi!PROVISION
                  Case 5: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes05 = l_dbl_Mes05 + g_rst_Princi!PROVISION
                  Case 6: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes06 = l_dbl_Mes06 + g_rst_Princi!PROVISION
                  Case 7: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes07 = l_dbl_Mes07 + g_rst_Princi!PROVISION
                  Case 8: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes08 = l_dbl_Mes08 + g_rst_Princi!PROVISION
                  Case 9: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes09 = l_dbl_Mes09 + g_rst_Princi!PROVISION
                  Case 10: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes10 = l_dbl_Mes10 + g_rst_Princi!PROVISION
                  Case 11: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes11 = l_dbl_Mes11 + g_rst_Princi!PROVISION
                  Case 12: If g_rst_Princi!CLAINT1 = "NOR" Then l_dbl_Mes12 = l_dbl_Mes12 + g_rst_Princi!PROVISION
             End Select
             
             g_rst_Princi.MoveNext
             DoEvents
         Loop
            
          r_int_nroaux = r_int_nroaux + 1
      Loop
        
      r_bol_FlagOp = False
      r_int_Nindex = r_int_Nindex + 1
   Loop
    
   p_Monto1 = gf_NvoCreCom(p_MesAct, p_AnoAct)
   p_Monto2 = gf_AmoCreCom(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
   p_Monto3 = gf_QNvoDH(p_MesAct, p_AnoAct)
   p_Monto4 = gf_PorHip(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
   p_Monto5 = r_dbl_CMoros
   p_Monto6 = r_dbl_HipPrv
   p_Monto7 = r_dbl_CliMoI
   p_Monto8 = gf_SalFav(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
   p_Monto9 = r_dbl_PrvNet
   p_Monto10 = r_dbl_SumNor
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function fs_MuestraTipo(ByVal p_PrvAct As String, ByVal p_PrvAnt As Variant) As String
   fs_MuestraTipo = ""
   
   If p_PrvAct = "NOR" Then
      fs_MuestraTipo = "NORMAL"
      Exit Function
   End If
   If IsNull(p_PrvAnt) Or p_PrvAnt = "" Then
      fs_MuestraTipo = "NUEVO"
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      End If
      Exit Function
   End If
   If p_PrvAct = p_PrvAnt Then
      fs_MuestraTipo = "IGUAL"
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      End If
   Else
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      Else
         If p_PrvAct = "CPP" Then
            fs_MuestraTipo = "MEJOR"
         End If
         If p_PrvAct = "DEF" Then
            If p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            Else
               fs_MuestraTipo = "MEJOR"
            End If
         End If
         If p_PrvAct = "DUD" Then
            If p_PrvAnt = "DEF" Or p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            Else
               fs_MuestraTipo = "MEJOR"
            End If
         End If
         If p_PrvAct = "PER" Then
            If p_PrvAnt = "DEF" Or p_PrvAnt = "DUD" Or p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            End If
         End If
         If p_PrvAnt = "PER" Then
            fs_MuestraTipo = "MEJOR"
         End If
      End If
   End If
End Function

Private Function gf_NvoCreCom(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer) As Double
Dim r_dbl_MtoAct     As Double

   gf_NvoCreCom = 0
   r_dbl_MtoAct = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(DECODE(COMCIE_TIPMON, 1, (COMCIE_PRVGEN+COMCIE_PRVCIC), (COMCIE_PRVGEN+COMCIE_PRVCIC)*COMCIE_TIPCAM)) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE  "
   g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_FECDES >= " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "01 "
   g_str_Parame = g_str_Parame & "   AND COMCIE_CLAPRV = 0  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAct = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
   
   gf_NvoCreCom = r_dbl_MtoAct
End Function

Private Function gf_AmoCreCom(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_MtoAct     As Double
Dim r_dbl_MtoAnt     As Double
Dim r_dbl_MtoDif     As Double
Dim r_dbl_TipCam     As Double
Dim r_rst_PerAnt     As ADODB.Recordset
Dim r_rst_PerAct     As ADODB.Recordset

   gf_AmoCreCom = 0
   r_dbl_MtoAct = 0
   r_dbl_MtoAnt = 0
   r_dbl_MtoDif = 0
   r_dbl_TipCam = 0
   
   'Operaciones del periodo anterior
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COMCIE_NUMOPE, COMCIE_TOTPRE "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE "
   g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAnt & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_TIPGAR = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerAnt, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_PerAnt.BOF And r_rst_PerAnt.EOF) Then
      r_rst_PerAnt.MoveFirst
      Do While Not r_rst_PerAnt.EOF
         r_dbl_MtoAnt = r_rst_PerAnt!comcie_totpre
         
         'Busca operaciones del periodo anterior en el periodo actual
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT COMCIE_NUMOPE, COMCIE_TOTPRE "
         g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE "
         g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAct & " "
         g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAct & " "
         g_str_Parame = g_str_Parame & "   AND COMCIE_NUMOPE = " & Trim(r_rst_PerAnt!COMCIE_NUMOPE) & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerAct, 3) Then
            Exit Function
         End If
         
         If Not (r_rst_PerAct.BOF And r_rst_PerAct.EOF) Then
            r_rst_PerAct.MoveFirst
            r_dbl_MtoAct = r_rst_PerAct!comcie_totpre
            r_dbl_MtoDif = r_dbl_MtoDif + (r_dbl_MtoAct - r_dbl_MtoAnt)
         Else
            r_dbl_MtoDif = r_dbl_MtoDif - r_dbl_MtoAnt
         End If
         
         r_rst_PerAnt.MoveNext
      Loop
      
      gf_AmoCreCom = r_dbl_MtoDif * (1.3 / 100)
   End If
   
End Function

Private Function gf_QNvoDH(ByVal p_Mes As Integer, ByVal p_Ano As Integer) As Double
   gf_QNvoDH = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT (SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_MTOPRE, HIPCIE_MTOPRE*HIPCIE_TIPCAM))*0.7)/100 AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_Ano & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_Mes & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES >='" & p_Ano & Format(p_Mes, "00") & "01' "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES <='" & p_Ano & Format(p_Mes, "00") & "31' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!total) Then
         gf_QNvoDH = CDbl(g_rst_GenAux!total)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Function gf_PorHip(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_MtoAct     As Double
Dim r_dbl_MtoAnt     As Double
Dim r_dbl_TipCam     As Double

   gf_PorHip = 0
   r_dbl_MtoAct = 0
   r_dbl_MtoAnt = 0
   r_dbl_TipCam = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC)*HIPCIE_TIPCAM)) AS MONTO_TOTAL, "
   g_str_Parame = g_str_Parame & "       MAX(HIPCIE_TIPCAM) AS TIPO_CAMBIO"
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES < " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "01 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAct = CDbl(g_rst_GenAux!MONTO_TOTAL)
         r_dbl_TipCam = CDbl(g_rst_GenAux!TIPO_CAMBIO)
      End If
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC)*" & CStr(r_dbl_TipCam) & ")) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAnt & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
    
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAnt = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
    
   gf_PorHip = r_dbl_MtoAct - r_dbl_MtoAnt
End Function

Private Function gf_SalFav(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_TipCam     As Double
   
   gf_SalFav = 0
   
   'Obtiene el tipo de cambio del mes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT HIPCIE_TIPCAM "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_MesAct & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      r_dbl_TipCam = CDbl(g_rst_GenAux!HIPCIE_TIPCAM)
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND( SUM( DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*" & CStr(r_dbl_TipCam) & ") )* -1 ,2) AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_AnoAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_CLAPRV <> 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_NUMOPE NOT IN ( "
   g_str_Parame = g_str_Parame & "          SELECT HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "            FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "           WHERE HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "             AND HIPCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "             AND HIPCIE_CLAPRV <> 0 )"
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
    
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!total) Then
         gf_SalFav = CDbl(g_rst_GenAux!total)
      End If
   End If
    
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

''QUERY 1
Private Function gf_Query(ByVal p_MesAct1 As Integer, ByVal p_AnoAct1 As Integer, ByVal p_MesAnt1 As Integer, ByVal p_AnoAnt1 As Integer) As String
    gf_Query = ""
    gf_Query = gf_Query & " SELECT * FROM ( "
    gf_Query = gf_Query & " WITH QUERY1 AS ( "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
    gf_Query = gf_Query & "                HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='003' THEN 'CME' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_Query = gf_Query & "                TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_Query = gf_Query & "                HIPCIE_DIAMOR AS DIASATRASO, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=2 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' END AS TIPOGARAN, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR IN (1,2) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "                     ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
    gf_Query = gf_Query & "                NVL((SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                       FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                      WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                        AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2,1)), 0) AS TASA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                       ELSE "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                     END "
    gf_Query = gf_Query & "                END AS CLASIFICACION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT1 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "            AND HIPCIE_CLAPRV <> 0 "
    gf_Query = gf_Query & "         UNION "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
    gf_Query = gf_Query & "                HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='003' THEN 'CME' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_Query = gf_Query & "                TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_Query = gf_Query & "                HIPCIE_DIAMOR AS DIASATRASO, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=2 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' END AS TIPOGARAN, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR IN (1,2) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "                     ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
    gf_Query = gf_Query & "                NVL((SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                       FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                      WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                        AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2,1)), 0) AS TASA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                       ELSE "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                     END "
    gf_Query = gf_Query & "                END AS CLASIFICACION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT1 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "            AND HIPCIE_PRVVOL > 0 "
    gf_Query = gf_Query & "            AND HIPCIE_CLAALI = 0 ), "
    
    gf_Query = gf_Query & " QUERY2 AS ("
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROVISION2, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROV_VOLUNT2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                       ELSE"
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                     END"
    gf_Query = gf_Query & "                END AS CLASIFICACION2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT2 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAnt1
    gf_Query = gf_Query & "            AND HIPCIE_CLAPRV <> 0 "
    gf_Query = gf_Query & "         UNION "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROVISION2, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROV_VOLUNT2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                       ELSE"
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                     END"
    gf_Query = gf_Query & "                END AS CLASIFICACION2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT2 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PRVVOL > 0 "
    gf_Query = gf_Query & "            AND HIPCIE_CLAALI = 0 )"
    
    gf_Query = gf_Query & "  SELECT PRODUCTO, NOMBRECLIENTE, DIASATRASO, TIPOGARAN "
    gf_Query = gf_Query & "         ,ROUND(VALORGARANTIA,2) AS VALORGARANTIA, TASA "
    gf_Query = gf_Query & "         ,ROUND(CAPITAL,2) AS CAPITAL, ROUND(CAPITAL_BAL,2) AS CAPITAL_BAL "
    gf_Query = gf_Query & "         ,ROUND(PROVISION,2) AS PROVISION, CLASIFICACION, CLAINT1 "
    gf_Query = gf_Query & "         ,ROUND(NVL(PROVISION2,0),2) AS PROVISION2, CLASIFICACION2, CLAINT2 "
    gf_Query = gf_Query & "         ,ROUND((NVL(PROVISION,0) + NVL(PROV_VOLUNT1,0) - NVL(PROVISION2,0) - NVL(PROV_VOLUNT2,0)),2) AS AJUSTE "
    gf_Query = gf_Query & "         ,HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_TDOCLI, HIPCIE_NDOCLI "
    gf_Query = gf_Query & "         ,HIPCIE_CLACLI, HIPCIE_CLAALI, PROV_VOLUNT1, PROV_VOLUNT2 "
    gf_Query = gf_Query & "         ,HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, OPERACION "
    gf_Query = gf_Query & "   FROM QUERY1 "
    gf_Query = gf_Query & " LEFT JOIN QUERY2 ON (HIPCIE_NUMOPE = NUMOPE) "
    gf_Query = gf_Query & "  ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO "
    gf_Query = gf_Query & "  )"
     
End Function

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "GENERAL"
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(1, 1) = "RESUMEN DE PROVISIONES A " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text) & " EN SOLES "
      .Range(.Cells(1, 1), .Cells(1, 13)).Merge
      .Range("A1:M1").HorizontalAlignment = xlHAlignCenter
      
      'Primera Linea
      r_int_NroFil = 3
      .Cells(r_int_NroFil, 1) = "RESUMEN"
      .Cells(r_int_NroFil, 2) = "ENERO"
      .Cells(r_int_NroFil, 3) = "FEBRERO"
      .Cells(r_int_NroFil, 4) = "MARZO"
      .Cells(r_int_NroFil, 5) = "ABRIL"
      .Cells(r_int_NroFil, 6) = "MAYO"
      .Cells(r_int_NroFil, 7) = "JUNIO"
      .Cells(r_int_NroFil, 8) = "JULIO"
      .Cells(r_int_NroFil, 9) = "AGOSTO"
      .Cells(r_int_NroFil, 10) = "SETIEMBRE"
      .Cells(r_int_NroFil, 11) = "OCTUBRE"
      .Cells(r_int_NroFil, 12) = "NOVIEMBRE"
      .Cells(r_int_NroFil, 13) = "DICIEMBRE"
      
      'Segunda Linea
      .Columns("A").ColumnWidth = 50:     .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Columns("B").ColumnWidth = 12:     .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12:     .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 12:     .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 12:     .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 12:     .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 12:     .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 12:     .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 12:     .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 12:     .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 12:     .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 12:     .Cells(r_int_NroFil, 12).HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12:     .Cells(r_int_NroFil, 13).HorizontalAlignment = xlHAlignCenter
            
      'Formatea titulo
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Size = 11
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      'Exporta filas
      For r_int_Contad = 4 To 17
         .Cells(r_int_Contad, 1) = grd_LisCla.TextMatrix(r_int_Contad - 3, 1)
         .Cells(r_int_Contad, 2) = grd_LisCla.TextMatrix(r_int_Contad - 3, 2)
         .Cells(r_int_Contad, 3) = grd_LisCla.TextMatrix(r_int_Contad - 3, 3)
         .Cells(r_int_Contad, 4) = grd_LisCla.TextMatrix(r_int_Contad - 3, 4)
         .Cells(r_int_Contad, 5) = grd_LisCla.TextMatrix(r_int_Contad - 3, 5)
         .Cells(r_int_Contad, 6) = grd_LisCla.TextMatrix(r_int_Contad - 3, 6)
         .Cells(r_int_Contad, 7) = grd_LisCla.TextMatrix(r_int_Contad - 3, 7)
         .Cells(r_int_Contad, 8) = grd_LisCla.TextMatrix(r_int_Contad - 3, 8)
         .Cells(r_int_Contad, 9) = grd_LisCla.TextMatrix(r_int_Contad - 3, 9)
         .Cells(r_int_Contad, 10) = grd_LisCla.TextMatrix(r_int_Contad - 3, 10)
         .Cells(r_int_Contad, 11) = grd_LisCla.TextMatrix(r_int_Contad - 3, 11)
         .Cells(r_int_Contad, 12) = grd_LisCla.TextMatrix(r_int_Contad - 3, 12)
         .Cells(r_int_Contad, 13) = grd_LisCla.TextMatrix(r_int_Contad - 3, 13)
      Next
   End With
   
   
   r_obj_Excel.Sheets(2).Name = "POR TIPO"
   With r_obj_Excel.Sheets(2)
      'Titulo
      .Cells(1, 1) = "RESUMEN DE PROVISIONES A " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text) & " EN SOLES "
      .Range(.Cells(1, 1), .Cells(1, 13)).Merge
      .Range("A1:M1").HorizontalAlignment = xlHAlignCenter

      'Primera Linea
      r_int_NroFil = 3
      .Cells(r_int_NroFil, 1) = "DETALLE"
      .Cells(r_int_NroFil, 2) = "ENERO"
      .Cells(r_int_NroFil, 3) = "FEBRERO"
      .Cells(r_int_NroFil, 4) = "MARZO"
      .Cells(r_int_NroFil, 5) = "ABRIL"
      .Cells(r_int_NroFil, 6) = "MAYO"
      .Cells(r_int_NroFil, 7) = "JUNIO"
      .Cells(r_int_NroFil, 8) = "JULIO"
      .Cells(r_int_NroFil, 9) = "AGOSTO"
      .Cells(r_int_NroFil, 10) = "SETIEMBRE"
      .Cells(r_int_NroFil, 11) = "OCTUBRE"
      .Cells(r_int_NroFil, 12) = "NOVIEMBRE"
      .Cells(r_int_NroFil, 13) = "DICIEMBRE"
         
      'Segunda Linea
      .Columns("A").ColumnWidth = 30:     .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Columns("B").ColumnWidth = 12:     .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12:     .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 12:     .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 12:     .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 12:     .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 12:     .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 12:     .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 12:     .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 12:     .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 12:     .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 12:     .Cells(r_int_NroFil, 12).HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12:     .Cells(r_int_NroFil, 13).HorizontalAlignment = xlHAlignCenter
            
      'Formatea titulo
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Size = 11
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      'Exporta filas
      For r_int_Contad = 4 To 10
         .Cells(r_int_Contad, 1) = grd_LisTip.TextMatrix(r_int_Contad - 3, 1)
         .Cells(r_int_Contad, 2) = grd_LisTip.TextMatrix(r_int_Contad - 3, 2)
         .Cells(r_int_Contad, 3) = grd_LisTip.TextMatrix(r_int_Contad - 3, 3)
         .Cells(r_int_Contad, 4) = grd_LisTip.TextMatrix(r_int_Contad - 3, 4)
         .Cells(r_int_Contad, 5) = grd_LisTip.TextMatrix(r_int_Contad - 3, 5)
         .Cells(r_int_Contad, 6) = grd_LisTip.TextMatrix(r_int_Contad - 3, 6)
         .Cells(r_int_Contad, 7) = grd_LisTip.TextMatrix(r_int_Contad - 3, 7)
         .Cells(r_int_Contad, 8) = grd_LisTip.TextMatrix(r_int_Contad - 3, 8)
         .Cells(r_int_Contad, 9) = grd_LisTip.TextMatrix(r_int_Contad - 3, 9)
         .Cells(r_int_Contad, 10) = grd_LisTip.TextMatrix(r_int_Contad - 3, 10)
         .Cells(r_int_Contad, 11) = grd_LisTip.TextMatrix(r_int_Contad - 3, 11)
         .Cells(r_int_Contad, 12) = grd_LisTip.TextMatrix(r_int_Contad - 3, 12)
         .Cells(r_int_Contad, 13) = grd_LisTip.TextMatrix(r_int_Contad - 3, 13)
      Next
   End With

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

