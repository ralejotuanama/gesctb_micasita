VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_CtbIntPBP_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   6855
   ClientTop       =   1965
   ClientWidth     =   14130
   Icon            =   "GesCtb_frm_847.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14160
      _Version        =   65536
      _ExtentX        =   24977
      _ExtentY        =   12303
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
         TabIndex        =   11
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   570
            TabIndex        =   17
            Top             =   60
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Proceso"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   570
            TabIndex        =   18
            Top             =   330
            Width           =   5235
            _Version        =   65536
            _ExtentX        =   9234
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Contabilización de Provisión de intereses PBP (CRC-CME)"
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
            Picture         =   "GesCtb_frm_847.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   12
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_847.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar pantalla"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_847.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar información"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13395
            Picture         =   "GesCtb_frm_847.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   1845
            Picture         =   "GesCtb_frm_847.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generar asientos automaticos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_847.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   885
         Left            =   60
         TabIndex        =   13
         Top             =   1470
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
            Top             =   105
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
            TabIndex        =   15
            Top             =   130
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   495
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4530
         Left            =   60
         TabIndex        =   16
         Top             =   2400
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   7990
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   4320
            Left            =   105
            TabIndex        =   8
            Top             =   105
            Width           =   13815
            _ExtentX        =   24368
            _ExtentY        =   7620
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Nuevos Soles"
            TabPicture(0)   =   "GesCtb_frm_847.frx":1380
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_LisSoles"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Dolares Americanos"
            TabPicture(1)   =   "GesCtb_frm_847.frx":139C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisDolar"
            Tab(1).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_LisSoles 
               Height          =   3870
               Left            =   60
               TabIndex        =   7
               Top             =   375
               Width           =   13680
               _ExtentX        =   24130
               _ExtentY        =   6826
               _Version        =   393216
               Rows            =   10
               Cols            =   13
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisDolar 
               Height          =   3870
               Left            =   -74940
               TabIndex        =   9
               Top             =   375
               Width           =   13680
               _ExtentX        =   24130
               _ExtentY        =   6826
               _Version        =   393216
               Rows            =   10
               Cols            =   13
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
Attribute VB_Name = "frm_Pro_CtbIntPBP_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_PerMes        As String
Dim l_str_PerAno        As String
Dim l_str_FinPer        As String
Dim l_dbl_TipCam        As Double

Private Sub cmd_Buscar_Click()
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
   l_str_PerMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   l_str_PerAno = ipp_PerAno.Text
   l_str_FinPer = l_str_PerAno & l_str_PerMes & ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno))
   Call fs_Procesa
   Call fs_Habilita(True)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Habilita(False)
   Call gs_SetFocus(cmb_PerMes)
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
   If grd_LisSoles.Rows <= 2 Then
      MsgBox "Debe buscar la información para un período.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Buscar)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   l_str_PerMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   l_str_PerAno = ipp_PerAno.Text
   l_str_FinPer = l_str_PerAno & l_str_PerMes & ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno))
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Proces_Click()
Dim r_str_Cadena     As String
Dim r_int_NumVec     As Integer
Dim r_rst_Record     As ADODB.Recordset

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
   If grd_LisSoles.Rows <= 2 Then
      MsgBox "Debe buscar la información para un período.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Buscar)
      Exit Sub
   End If
   
   l_str_PerMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   l_str_PerAno = ipp_PerAno.Text
   l_str_FinPer = l_str_PerAno & l_str_PerMes & ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno))
   
   'Valida año ejecucion
   If l_str_PerAno < 2015 Then
      MsgBox "Los asientos contables se generaran a partir del año 2015.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Valida ejecucion
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS NUM_EJEC "
   r_str_Cadena = r_str_Cadena & "  FROM CTB_PERPRO "
   r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(l_str_PerAno) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(l_str_PerMes) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = 1 "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 3) Then
      Exit Sub
   End If
   
   r_rst_Record.MoveFirst
   r_int_NumVec = r_rst_Record!NUM_EJEC
   
   r_rst_Record.Close
   Set r_rst_Record = Nothing
   
   'Confirmacion
   If r_int_NumVec > 0 Then
      MsgBox "Los asientos contables ya han sido generados para el periodo seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   Else
      If MsgBox("¿Está seguro de generar los asientos automaticos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   Call fs_GeneraAsientos
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Habilita(False)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub fs_Limpia()
   SSTab1.Tab = 0
   grd_LisDolar.Rows = 2
   grd_LisSoles.Rows = 2
   cmb_PerMes.ListIndex = -1
   Call fs_Formacion_Grid(grd_LisDolar)
   Call fs_Formacion_Grid(grd_LisSoles)
End Sub

Private Sub fs_Habilita(ByVal p_Habilitado As Boolean)
   cmd_Buscar.Enabled = Not p_Habilitado
   cmd_ExpExc.Enabled = p_Habilitado
   cmd_Proces.Enabled = p_Habilitado
   cmb_PerMes.Enabled = Not p_Habilitado
   ipp_PerAno.Enabled = Not p_Habilitado
End Sub

Private Sub fs_Formacion_Grid(NombreGrid As MSFlexGrid)
Dim r_int_nrofil  As Integer
Dim r_int_Nindex  As Integer
Dim i             As Integer

   With NombreGrid
      .Col = 0: .CellAlignment = flexAlignCenterCenter
      .Col = 1: .CellAlignment = flexAlignCenterCenter
      .Col = 2: .CellAlignment = flexAlignCenterCenter
      .Col = 3: .CellAlignment = flexAlignCenterCenter
      .Col = 4: .CellAlignment = flexAlignCenterCenter
      .Col = 5: .CellAlignment = flexAlignCenterCenter
      .Col = 6: .CellAlignment = flexAlignCenterCenter
      .Col = 7: .CellAlignment = flexAlignCenterCenter
      .Col = 8: .CellAlignment = flexAlignCenterCenter
      .Col = 9: .CellAlignment = flexAlignCenterCenter
      .Col = 10: .CellAlignment = flexAlignCenterCenter
      .Col = 11: .CellAlignment = flexAlignCenterCenter
      .Col = 12: .CellAlignment = flexAlignCenterCenter
      .ColWidth(0) = 1000
      .ColWidth(1) = 3500
      .ColWidth(2) = 900
      .ColWidth(3) = 900
      .ColWidth(4) = 900
      .ColWidth(5) = 900
      .ColWidth(6) = 900
      .ColWidth(7) = 900
      .ColWidth(8) = 600
      .ColWidth(9) = 900
      .ColWidth(10) = 900
      .ColWidth(11) = 900
      .ColWidth(12) = 0
      .WordWrap = True
      
      r_int_nrofil = 0
      .TextMatrix(r_int_nrofil, 0) = "OPERACION"
      .TextMatrix(r_int_nrofil, 1) = "NOMBRE DE CLIENTE"
      For r_int_Nindex = 2 To 6
         .TextMatrix(r_int_nrofil, r_int_Nindex) = "AJUSTE"
      Next r_int_Nindex
      .TextMatrix(r_int_nrofil, 7) = "TOTAL"
      .TextMatrix(r_int_nrofil, 8) = "GANA PBP"
      .TextMatrix(r_int_nrofil, 9) = "SALDO ANT."
      .TextMatrix(r_int_nrofil, 10) = "PROVIS. MES"
      .TextMatrix(r_int_nrofil, 11) = "SALDO SIG."
      
      r_int_nrofil = 1
      .TextMatrix(r_int_nrofil, 0) = "OPERACION"
      .TextMatrix(r_int_nrofil, 1) = "NOMBRE DE CLIENTE"
      For r_int_Nindex = 2 To 6
         .TextMatrix(r_int_nrofil, r_int_Nindex) = "AJUSTE"
      Next r_int_Nindex
      .TextMatrix(r_int_nrofil, 7) = "TOTAL"
      .TextMatrix(r_int_nrofil, 8) = "GANA PBP"
      .TextMatrix(r_int_nrofil, 9) = "SALDO ANT."
      .TextMatrix(r_int_nrofil, 10) = "PROVIS. MES"
      .TextMatrix(r_int_nrofil, 11) = "SALDO SIG."
      
      .MergeCells = flexMergeRestrictRows
      .MergeCol(0) = True:      .MergeRow(0) = True
      .MergeCol(1) = True:      .MergeRow(0) = True
      .MergeCol(2) = True:      .MergeRow(0) = True
      .MergeCol(3) = True:      .MergeRow(0) = True
      .MergeCol(4) = True:      .MergeRow(0) = True
      .MergeCol(5) = True:      .MergeRow(0) = True
      .MergeCol(6) = True:      .MergeRow(0) = True
      .MergeCol(7) = True:      .MergeRow(0) = True
      .MergeCol(8) = True:      .MergeRow(0) = True
      .MergeCol(9) = True:      .MergeRow(0) = True
      .MergeCol(10) = True:     .MergeRow(0) = True
      .MergeCol(11) = True:     .MergeRow(0) = True
      
      For i = 0 To 1
         .Row = i
         .Col = 0: .CellAlignment = flexAlignCenterCenter
         .Col = 1: .CellAlignment = flexAlignCenterCenter
         .Col = 2: .CellAlignment = flexAlignCenterCenter
         .Col = 3: .CellAlignment = flexAlignCenterCenter
         .Col = 4: .CellAlignment = flexAlignCenterCenter
         .Col = 5: .CellAlignment = flexAlignCenterCenter
         .Col = 6: .CellAlignment = flexAlignCenterCenter
         .Col = 7: .CellAlignment = flexAlignCenterCenter
         .Col = 8: .CellAlignment = flexAlignCenterCenter
         .Col = 9: .CellAlignment = flexAlignCenterCenter
         .Col = 10: .CellAlignment = flexAlignCenterCenter
         .Col = 11: .CellAlignment = flexAlignCenterCenter
      Next
      .Rows = .Rows + 1
      .FixedRows = 2
      If .FixedRows = 2 Then .Rows = .Rows - 1
   End With
End Sub

Private Sub fs_Procesa()
Dim r_str_Parame     As String
Dim r_int_NumDia     As Integer
Dim r_int_nrofil     As Integer
Dim r_int_Nindex     As Integer
Dim r_int_NroCol     As Integer
Dim r_dbl_MtoCan     As Double
Dim r_dbl_SumTot1    As Double
Dim r_dbl_SumTot2    As Double
Dim r_dbl_SumTot3    As Double
Dim r_dbl_SumTot4    As Double
Dim r_dbl_SumTot5    As Double
Dim r_dbl_SumTot     As Double
Dim r_dbl_SumProv    As Double
Dim r_dbl_SumIni     As Double
Dim r_dbl_SumFin     As Double

   g_str_Parame = gf_Query_InteresPBP(l_str_PerMes, l_str_PerAno)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      grd_LisDolar.Rows = 2
      grd_LisSoles.Rows = 2
      Exit Sub
   End If
   
   '*****************
   'IMPRIME SOLES
   r_dbl_SumTot1 = 0
   r_dbl_SumTot2 = 0
   r_dbl_SumTot3 = 0
   r_dbl_SumTot4 = 0
   r_dbl_SumTot5 = 0
   r_dbl_SumProv = 0
   r_dbl_SumTot = 0
   r_dbl_SumIni = 0
   r_dbl_SumFin = 0
   
   With grd_LisSoles
      '********
      'Cabecera
      .Rows = 3
      r_int_NroCol = 6
      
      r_int_nrofil = 0
      .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      r_int_nrofil = 1
      .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      .WordWrap = True
      
      For r_int_Nindex = 1 To 4
         r_int_NroCol = r_int_NroCol - 1
         r_int_nrofil = 0
         .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
         r_int_nrofil = 1
         .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
      Next r_int_Nindex
      
      .MergeCells = flexMergeRestrictRows
      .MergeCol(0) = True:  .MergeRow(0) = True
      .MergeCol(1) = True:  .MergeRow(0) = True
      .MergeCol(2) = True:  .MergeRow(0) = True
      .MergeCol(3) = True:  .MergeRow(0) = True
      .MergeCol(4) = True:  .MergeRow(0) = True
      .MergeCol(5) = True:  .MergeRow(0) = True
      .MergeCol(6) = True:  .MergeRow(0) = True
      .MergeCol(7) = True:  .MergeRow(0) = True
      .MergeCol(8) = True:  .MergeRow(0) = True
      .MergeCol(9) = True:  .MergeRow(0) = True
      .MergeCol(10) = True: .MergeRow(0) = True
      .MergeCol(11) = True: .MergeRow(0) = True
      
      .Row = 1
      .Col = 0:  .CellAlignment = flexAlignCenterCenter
      .Col = 1:  .CellAlignment = flexAlignCenterCenter
      .Col = 2:  .CellAlignment = flexAlignCenterCenter
      .Col = 3:  .CellAlignment = flexAlignCenterCenter
      .Col = 4:  .CellAlignment = flexAlignCenterCenter
      .Col = 5:  .CellAlignment = flexAlignCenterCenter
      .Col = 6:  .CellAlignment = flexAlignCenterCenter
      .Col = 7:  .CellAlignment = flexAlignCenterCenter
      .Col = 8:  .CellAlignment = flexAlignCenterCenter
      .Col = 9:  .CellAlignment = flexAlignCenterCenter
      .Col = 10: .CellAlignment = flexAlignCenterCenter
      .Col = 11: .CellAlignment = flexAlignCenterCenter
      
      '***********
      'Informacion
      r_int_nrofil = 2
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            .TextMatrix(r_int_nrofil, 0) = Trim(g_rst_Princi!HIPCIE_NUMOPE)
            .TextMatrix(r_int_nrofil, 1) = Trim(g_rst_Princi!CLIENTE)
            .TextMatrix(r_int_nrofil, 7) = Format(CDbl(g_rst_Princi!HIPCIE_DEVPBP) * CInt(g_rst_Princi!NROMESES), "###,##0.00")
            
            .Row = r_int_nrofil
            .Col = 8: .CellAlignment = flexAlignCenterCenter
           
            Select Case CInt(g_rst_Princi!FLAGPBP)
               Case 0
                  r_str_Parame = ""
                  r_str_Parame = r_str_Parame & "SELECT CUOCIE_SITUAC, CUOCIE_FECVCT, CUOCIE_FECPAG "
                  r_str_Parame = r_str_Parame & "  FROM CRE_CUOCIE "
                  r_str_Parame = r_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_NUMOPE = '" & Trim(g_rst_Princi!HIPCIE_NUMOPE) & "' "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_TIPCRO = 1 "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_FECVCT <= " & l_str_FinPer & " "
                  r_str_Parame = r_str_Parame & " ORDER BY CUOCIE_NUMCUO DESC"
                  
                  If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  
                  r_int_Nindex = 0
                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                     g_rst_Genera.MoveFirst
                     Do While Not g_rst_Genera.EOF
                        If r_int_Nindex > CInt(g_rst_Princi!NROMESES) Then
                           Exit Do
                        End If
                        
                        r_int_NumDia = 0
                        If g_rst_Genera!CUOCIE_SITUAC = 2 Then
                           r_int_NumDia = DateDiff("d", Format(Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 7, 2) & "/" & Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 5, 2) & "/" & Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 1, 4), "DD/MM/YYYY"), Format(Mid(Trim(l_str_FinPer), 7, 2) & "/" & Mid(Trim(l_str_FinPer), 5, 2) & "/" & Mid(Trim(l_str_FinPer), 1, 4), "DD/MM/YYYY"))
                        Else
                           r_int_NumDia = DateDiff("d", Format(Mid(g_rst_Genera!CUOCIE_FECVCT, 7, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECVCT, 5, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECVCT, 1, 4), "yyyy/mm/dd"), Format(Mid(g_rst_Genera!CUOCIE_FECPAG, 7, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECPAG, 5, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECPAG, 1, 4), "yyyy/mm/dd"))
                        End If
                        If Mid(Trim(g_rst_Princi!HIPCIE_NUMOPE), 1, 3) = "001" Or Mid(Trim(g_rst_Princi!HIPCIE_NUMOPE), 1, 3) = "003" Then
                           If r_int_NumDia > 15 Then
                              .TextMatrix(r_int_nrofil, 8) = "NO"
                           End If
                        Else
                           If r_int_NumDia > 30 Then
                              .TextMatrix(r_int_nrofil, 8) = "NO"
                           End If
                        End If
                        
                        r_int_Nindex = r_int_Nindex + 1
                        g_rst_Genera.MoveNext
                     Loop
                  End If
                  .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
                  
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
                  
               Case 1
                  .TextMatrix(r_int_nrofil, 6) = Format(CDbl(g_rst_Princi!INTADE), "###,##0.00")
                  .TextMatrix(r_int_nrofil, 7) = Format(0, "###,##0.00")
                  .TextMatrix(r_int_nrofil, 8) = "SI"
                  .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
                  
               Case 2
                  .TextMatrix(r_int_nrofil, 8) = "NO"
                  .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
            End Select
            
            If Not CInt(g_rst_Princi!NROMESES) = 0 Then
               r_int_NroCol = 6
               For r_int_Nindex = CInt(g_rst_Princi!NROMESES) To 1 Step -1
                  .TextMatrix(r_int_nrofil, r_int_NroCol) = Format(CDbl(g_rst_Princi!HIPCIE_DEVPBP), "###,##0.00")
                  r_int_NroCol = r_int_NroCol - 1
               Next r_int_Nindex
            End If
            
            'Calcula Monto a Provisionar
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "" Then
                If .TextMatrix(r_int_nrofil, 6) = "" Then
                    .TextMatrix(r_int_nrofil, 10) = Format(CDbl(0), "###,##0.00")
               Else
                    .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)), "###,##0.00")
               End If
            End If
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "SI" Then
               .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)) - CDbl(.TextMatrix(r_int_nrofil, 9)), "###,##0.00")
            End If
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "NO" Then
               .TextMatrix(r_int_nrofil, 10) = Format(0, "###,##0.00")
            End If
            
            'Calcula saldo final
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "NO" Or Trim(.TextMatrix(r_int_nrofil, 8)) = "SI" Then
               .TextMatrix(r_int_nrofil, 11) = Format(0, "###,##0.00")
            Else
               .TextMatrix(r_int_nrofil, 11) = Format(CDbl(.TextMatrix(r_int_nrofil, 9)) + CDbl(.TextMatrix(r_int_nrofil, 10)), "###,##0.00")
            End If
            
            If .TextMatrix(r_int_nrofil, 2) <> "" Then r_dbl_SumTot1 = r_dbl_SumTot1 + (CDbl(.TextMatrix(r_int_nrofil, 2)))
            If .TextMatrix(r_int_nrofil, 3) <> "" Then r_dbl_SumTot2 = r_dbl_SumTot2 + (CDbl(.TextMatrix(r_int_nrofil, 3)))
            If .TextMatrix(r_int_nrofil, 4) <> "" Then r_dbl_SumTot3 = r_dbl_SumTot3 + (CDbl(.TextMatrix(r_int_nrofil, 4)))
            If .TextMatrix(r_int_nrofil, 5) <> "" Then r_dbl_SumTot4 = r_dbl_SumTot4 + (CDbl(.TextMatrix(r_int_nrofil, 5)))
            If .TextMatrix(r_int_nrofil, 6) <> "" Then r_dbl_SumTot5 = r_dbl_SumTot5 + (CDbl(.TextMatrix(r_int_nrofil, 6)))
            If .TextMatrix(r_int_nrofil, 7) <> "" Then r_dbl_SumTot = r_dbl_SumTot + (CDbl(.TextMatrix(r_int_nrofil, 7)))
            r_dbl_SumIni = r_dbl_SumIni + CDbl(.TextMatrix(r_int_nrofil, 9))
            r_dbl_SumProv = r_dbl_SumProv + CDbl(.TextMatrix(r_int_nrofil, 10))
            r_dbl_SumFin = r_dbl_SumFin + CDbl(.TextMatrix(r_int_nrofil, 11))
            
            r_int_nrofil = r_int_nrofil + 1
            .Rows = .Rows + 1
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
            
      'Creditos cancelados
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT PP.PPGCAB_NUMOPE, PP.PPGCAB_FECPRO, PP.PPGCAB_FECPPG, PP.PPGCAB_MTODEP, PP.PPGCAB_MTOTOT, "
      r_str_Parame = r_str_Parame & "       TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE "
      r_str_Parame = r_str_Parame & "  FROM CRE_PPGCAB PP "
      r_str_Parame = r_str_Parame & " INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
      r_str_Parame = r_str_Parame & " INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI "
      r_str_Parame = r_str_Parame & " WHERE SUBSTR(PP.PPGCAB_NUMOPE,1,3) IN ('003') "
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_FECPPG >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01"
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_FECPPG <= " & l_str_PerAno & Format(l_str_PerMes, "00") & "31"
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_TIPPPG = 2 "
      
      If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            r_dbl_MtoCan = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Genera!PPGCAB_NUMOPE), "###,##0.00")
            If r_dbl_MtoCan > 0 Then
               .TextMatrix(r_int_nrofil, 0) = Trim(g_rst_Genera!PPGCAB_NUMOPE)
               .TextMatrix(r_int_nrofil, 1) = Trim(g_rst_Genera!CLIENTE)
               .TextMatrix(r_int_nrofil, 6) = Format(0, "###,##0.00")
               .TextMatrix(r_int_nrofil, 7) = Format(0, "###,##0.00")
               .TextMatrix(r_int_nrofil, 8) = ""
               .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Genera!PPGCAB_NUMOPE), "###,##0.00")
               .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)) - CDbl(.TextMatrix(r_int_nrofil, 9)), "###,##0.00")
               .TextMatrix(r_int_nrofil, 11) = Format(CDbl(.TextMatrix(r_int_nrofil, 9)) + CDbl(.TextMatrix(r_int_nrofil, 10)), "###,##0.00")
               
               r_dbl_SumIni = r_dbl_SumIni + CDbl(.TextMatrix(r_int_nrofil, 9))
               r_dbl_SumProv = r_dbl_SumProv + CDbl(.TextMatrix(r_int_nrofil, 10))
               r_dbl_SumFin = r_dbl_SumFin + CDbl(.TextMatrix(r_int_nrofil, 11))
               r_int_nrofil = r_int_nrofil + 1
               .Rows = .Rows + 1
            End If
            
            g_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      .TextMatrix(r_int_nrofil, 2) = Format(r_dbl_SumTot1, "###,##0.00")
      .TextMatrix(r_int_nrofil, 3) = Format(r_dbl_SumTot2, "###,##0.00")
      .TextMatrix(r_int_nrofil, 4) = Format(r_dbl_SumTot3, "###,##0.00")
      .TextMatrix(r_int_nrofil, 5) = Format(r_dbl_SumTot4, "###,##0.00")
      .TextMatrix(r_int_nrofil, 6) = Format(r_dbl_SumTot5, "###,##0.00")
      .TextMatrix(r_int_nrofil, 7) = Format(r_dbl_SumTot, "###,##0.00")
      .TextMatrix(r_int_nrofil, 9) = Format(r_dbl_SumIni, "###,##0.00")
      .TextMatrix(r_int_nrofil, 10) = Format(r_dbl_SumProv, "###,##0.00")
      .TextMatrix(r_int_nrofil, 11) = Format(r_dbl_SumFin, "###,##0.00")
   End With
   
   '******************
   'IMPRIME DOLARES
   r_dbl_SumTot1 = 0
   r_dbl_SumTot2 = 0
   r_dbl_SumTot3 = 0
   r_dbl_SumTot4 = 0
   r_dbl_SumTot5 = 0
   r_dbl_SumProv = 0
   r_dbl_SumTot = 0
   r_dbl_SumIni = 0
   r_dbl_SumFin = 0
   
   g_rst_Princi.MoveFirst
   With grd_LisDolar
      '********
      'Cabecera
      .Rows = 3
      r_int_NroCol = 6
      
      r_int_nrofil = 0
      .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      r_int_nrofil = 1
      .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      
      .WordWrap = True
      For r_int_Nindex = 1 To 4
         r_int_NroCol = r_int_NroCol - 1
         r_int_nrofil = 0
         .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
         
         r_int_nrofil = 1
         .TextMatrix(r_int_nrofil, r_int_NroCol) = "AJUSTE " & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
      Next r_int_Nindex
      
      .MergeCells = flexMergeRestrictRows
      .MergeCol(0) = True:      .MergeRow(0) = True
      .MergeCol(1) = True:      .MergeRow(0) = True
      .MergeCol(2) = True:      .MergeRow(0) = True
      .MergeCol(3) = True:      .MergeRow(0) = True
      .MergeCol(4) = True:      .MergeRow(0) = True
      .MergeCol(5) = True:      .MergeRow(0) = True
      .MergeCol(6) = True:      .MergeRow(0) = True
      .MergeCol(7) = True:      .MergeRow(0) = True
      .MergeCol(8) = True:      .MergeRow(0) = True
      .MergeCol(9) = True:      .MergeRow(0) = True
      .MergeCol(10) = True:     .MergeRow(0) = True
      .MergeCol(11) = True:     .MergeRow(0) = True
      
      .Row = 1
      .Col = 0:  .CellAlignment = flexAlignCenterCenter
      .Col = 1:  .CellAlignment = flexAlignCenterCenter
      .Col = 2:  .CellAlignment = flexAlignCenterCenter
      .Col = 3:  .CellAlignment = flexAlignCenterCenter
      .Col = 4:  .CellAlignment = flexAlignCenterCenter
      .Col = 5:  .CellAlignment = flexAlignCenterCenter
      .Col = 6:  .CellAlignment = flexAlignCenterCenter
      .Col = 7:  .CellAlignment = flexAlignCenterCenter
      .Col = 8:  .CellAlignment = flexAlignCenterCenter
      .Col = 9:  .CellAlignment = flexAlignCenterCenter
      .Col = 10: .CellAlignment = flexAlignCenterCenter
      .Col = 11: .CellAlignment = flexAlignCenterCenter
      
      '***********
      'Informacion
      r_int_nrofil = 2
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_TIPMON = 2 Then
            .TextMatrix(r_int_nrofil, 0) = Trim(g_rst_Princi!HIPCIE_NUMOPE)
            .TextMatrix(r_int_nrofil, 1) = Trim(g_rst_Princi!CLIENTE)
            .TextMatrix(r_int_nrofil, 7) = Format(CDbl(g_rst_Princi!HIPCIE_DEVPBP) * CInt(g_rst_Princi!NROMESES), "###,##0.00")
            
            .Row = r_int_nrofil
            .Col = 8: .CellAlignment = flexAlignCenterCenter
            
            Select Case CInt(g_rst_Princi!FLAGPBP)
               Case 0
                  r_str_Parame = ""
                  r_str_Parame = r_str_Parame & "SELECT CUOCIE_SITUAC, CUOCIE_FECVCT, CUOCIE_FECPAG "
                  r_str_Parame = r_str_Parame & "  FROM CRE_CUOCIE "
                  r_str_Parame = r_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_NUMOPE = '" & Trim(g_rst_Princi!HIPCIE_NUMOPE) & "' "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_TIPCRO = 1 "
                  r_str_Parame = r_str_Parame & "   AND CUOCIE_FECVCT <= " & l_str_FinPer & " "
                  r_str_Parame = r_str_Parame & " ORDER BY CUOCIE_NUMCUO DESC"
                  
                  If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  
                  r_int_Nindex = 0
                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                     g_rst_Genera.MoveFirst
                     Do While Not g_rst_Genera.EOF
                        If r_int_Nindex > CInt(g_rst_Princi!NROMESES) Then
                           Exit Do
                        End If
                        
                        r_int_NumDia = 0
                        If g_rst_Genera!CUOCIE_SITUAC = 2 Then
                           r_int_NumDia = DateDiff("d", Format(Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 7, 2) & "/" & Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 5, 2) & "/" & Mid(Trim(g_rst_Genera!CUOCIE_FECVCT), 1, 4), "DD/MM/YYYY"), Format(Mid(Trim(l_str_FinPer), 7, 2) & "/" & Mid(Trim(l_str_FinPer), 5, 2) & "/" & Mid(Trim(l_str_FinPer), 1, 4), "DD/MM/YYYY"))
                        Else
                           r_int_NumDia = DateDiff("d", Format(Mid(g_rst_Genera!CUOCIE_FECVCT, 7, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECVCT, 5, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECVCT, 1, 4), "yyyy/mm/dd"), Format(Mid(g_rst_Genera!CUOCIE_FECPAG, 7, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECPAG, 5, 2) & "/" & Mid(g_rst_Genera!CUOCIE_FECPAG, 1, 4), "yyyy/mm/dd"))
                        End If
                        If Mid(Trim(g_rst_Princi!HIPCIE_NUMOPE), 1, 3) = "001" Or Mid(Trim(g_rst_Princi!HIPCIE_NUMOPE), 1, 3) = "003" Then
                           If r_int_NumDia > 15 Then
                              .TextMatrix(r_int_nrofil, 8) = "NO"
                           End If
                        Else
                           If r_int_NumDia > 30 Then
                              .TextMatrix(r_int_nrofil, 8) = "NO"
                           End If
                        End If
                        .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
                        
                        r_int_Nindex = r_int_Nindex + 1
                        g_rst_Genera.MoveNext
                     Loop
                  End If
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
                  
               Case 1
                  .TextMatrix(r_int_nrofil, 6) = Format(CDbl(g_rst_Princi!INTADE), "###,##0.00")
                  .TextMatrix(r_int_nrofil, 7) = Format(0, "###,##0.00")
                  .TextMatrix(r_int_nrofil, 8) = "SI"
                  .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
                  
               Case 2
                  .TextMatrix(r_int_nrofil, 8) = "NO"
                  .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Princi!HIPCIE_NUMOPE), "###,##0.00")
                  
            End Select
            
            If Not CInt(g_rst_Princi!NROMESES) = 0 Then
               r_int_NroCol = 6
               For r_int_Nindex = CInt(g_rst_Princi!NROMESES) To 1 Step -1
                  .TextMatrix(r_int_nrofil, r_int_NroCol) = Format(CDbl(g_rst_Princi!HIPCIE_DEVPBP), "###,##0.00")
                  r_int_NroCol = r_int_NroCol - 1
               Next r_int_Nindex
            End If
             
            'Calcula Monto a Provisionar
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "" Then
                If .TextMatrix(r_int_nrofil, 6) = "" Then
                    .TextMatrix(r_int_nrofil, 10) = Format(CDbl(0), "###,##0.00")
                Else
                    .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)), "###,##0.00")
                End If
            End If
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "SI" Then
               .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)) - CDbl(.TextMatrix(r_int_nrofil, 9)), "###,##0.00")
            End If
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "NO" Then
               .TextMatrix(r_int_nrofil, 10) = Format(0, "###,##0.00")
            End If
            
            'Calcula el saldo final
            If Trim(.TextMatrix(r_int_nrofil, 8)) = "NO" Or Trim(.TextMatrix(r_int_nrofil, 8)) = "SI" Then
               .TextMatrix(r_int_nrofil, 11) = Format(0, "###,##0.00")
            Else
               .TextMatrix(r_int_nrofil, 11) = Format(CDbl(.TextMatrix(r_int_nrofil, 9)) + CDbl(.TextMatrix(r_int_nrofil, 10)), "###,##0.00")
            End If
            
            If .TextMatrix(r_int_nrofil, 2) <> "" Then r_dbl_SumTot1 = r_dbl_SumTot1 + (CDbl(.TextMatrix(r_int_nrofil, 2)))
            If .TextMatrix(r_int_nrofil, 3) <> "" Then r_dbl_SumTot2 = r_dbl_SumTot2 + (CDbl(.TextMatrix(r_int_nrofil, 3)))
            If .TextMatrix(r_int_nrofil, 4) <> "" Then r_dbl_SumTot3 = r_dbl_SumTot3 + (CDbl(.TextMatrix(r_int_nrofil, 4)))
            If .TextMatrix(r_int_nrofil, 5) <> "" Then r_dbl_SumTot4 = r_dbl_SumTot4 + (CDbl(.TextMatrix(r_int_nrofil, 5)))
            If .TextMatrix(r_int_nrofil, 6) <> "" Then r_dbl_SumTot5 = r_dbl_SumTot5 + (CDbl(.TextMatrix(r_int_nrofil, 6)))
            If .TextMatrix(r_int_nrofil, 7) <> "" Then r_dbl_SumTot = r_dbl_SumTot + (CDbl(.TextMatrix(r_int_nrofil, 7)))
            r_dbl_SumIni = r_dbl_SumIni + CDbl(.TextMatrix(r_int_nrofil, 9))
            r_dbl_SumProv = r_dbl_SumProv + CDbl(.TextMatrix(r_int_nrofil, 10))
            r_dbl_SumFin = r_dbl_SumFin + CDbl(.TextMatrix(r_int_nrofil, 11))
            
            r_int_nrofil = r_int_nrofil + 1
            .Rows = .Rows + 1
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
            
      'Creditos cancelados
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT PP.PPGCAB_NUMOPE, PP.PPGCAB_FECPRO, PP.PPGCAB_FECPPG, PP.PPGCAB_MTODEP, PP.PPGCAB_MTOTOT, "
      r_str_Parame = r_str_Parame & "       TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE "
      r_str_Parame = r_str_Parame & "  FROM CRE_PPGCAB PP "
      r_str_Parame = r_str_Parame & " INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
      r_str_Parame = r_str_Parame & " INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI "
      r_str_Parame = r_str_Parame & " WHERE SUBSTR(PP.PPGCAB_NUMOPE,1,3) IN ('001') "
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_FECPPG >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01"
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_FECPPG <= " & l_str_PerAno & Format(l_str_PerMes, "00") & "31"
      r_str_Parame = r_str_Parame & "   AND PP.PPGCAB_TIPPPG = 2 "
      
      If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            r_dbl_MtoCan = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Genera!PPGCAB_NUMOPE), "###,##0.00")
            If r_dbl_MtoCan > 0 Then
               .TextMatrix(r_int_nrofil, 0) = Trim(g_rst_Genera!PPGCAB_NUMOPE)
               .TextMatrix(r_int_nrofil, 1) = Trim(g_rst_Genera!CLIENTE)
               .TextMatrix(r_int_nrofil, 6) = Format(0, "###,##0.00")
               .TextMatrix(r_int_nrofil, 7) = Format(0, "###,##0.00")
               .TextMatrix(r_int_nrofil, 8) = ""
               .TextMatrix(r_int_nrofil, 9) = Format(fs_Obtiene_SaldoPerAnterior(g_rst_Genera!PPGCAB_NUMOPE), "###,##0.00")
               .TextMatrix(r_int_nrofil, 10) = Format(CDbl(.TextMatrix(r_int_nrofil, 6)) - CDbl(.TextMatrix(r_int_nrofil, 9)), "###,##0.00")
               .TextMatrix(r_int_nrofil, 11) = Format(CDbl(.TextMatrix(r_int_nrofil, 9)) + CDbl(.TextMatrix(r_int_nrofil, 10)), "###,##0.00")
               
               r_dbl_SumIni = r_dbl_SumIni + CDbl(.TextMatrix(r_int_nrofil, 9))
               r_dbl_SumProv = r_dbl_SumProv + CDbl(.TextMatrix(r_int_nrofil, 10))
               r_dbl_SumFin = r_dbl_SumFin + CDbl(.TextMatrix(r_int_nrofil, 11))
               r_int_nrofil = r_int_nrofil + 1
               .Rows = .Rows + 1
            End If
            
            g_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      .TextMatrix(r_int_nrofil, 2) = Format(r_dbl_SumTot1, "###,##0.00")
      .TextMatrix(r_int_nrofil, 3) = Format(r_dbl_SumTot2, "###,##0.00")
      .TextMatrix(r_int_nrofil, 4) = Format(r_dbl_SumTot3, "###,##0.00")
      .TextMatrix(r_int_nrofil, 5) = Format(r_dbl_SumTot4, "###,##0.00")
      .TextMatrix(r_int_nrofil, 6) = Format(r_dbl_SumTot5, "###,##0.00")
      .TextMatrix(r_int_nrofil, 7) = Format(r_dbl_SumTot, "###,##0.00")
      .TextMatrix(r_int_nrofil, 9) = Format(r_dbl_SumIni, "###,##0.00")
      .TextMatrix(r_int_nrofil, 10) = Format(r_dbl_SumProv, "###,##0.00")
      .TextMatrix(r_int_nrofil, 11) = Format(r_dbl_SumFin, "###,##0.00")
   End With
   
   MsgBox "Culmino proceso de búsqueda de información.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Function fs_Obtiene_SaldoPerAnterior(ByVal p_NumOpe As String) As Double
Dim r_rst_SalAnt     As ADODB.Recordset
Dim r_str_Cadena     As String
Dim r_str_MesAnt     As String
Dim r_str_AnoAnt     As String

   fs_Obtiene_SaldoPerAnterior = 0
   If CInt(l_str_PerMes) = 1 Then
      r_str_MesAnt = "12"
      r_str_AnoAnt = CInt(l_str_PerAno) - 1
   Else
      If CInt(l_str_PerMes) < 11 Then
         r_str_MesAnt = "0" & CStr(CInt(l_str_PerMes) - 1)
      Else
         r_str_MesAnt = CStr(CInt(l_str_PerMes) - 1)
      End If
      r_str_AnoAnt = CInt(l_str_PerAno)
   End If
   
   r_str_Cadena = gf_Query_InteresPBP(r_str_MesAnt, r_str_AnoAnt)
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SalAnt, 3) Then
      Exit Function
   End If
   
   If r_rst_SalAnt.BOF And r_rst_SalAnt.EOF Then
      r_rst_SalAnt.Close
      Set r_rst_SalAnt = Nothing
      MsgBox "No se encontraron datos del mes anterior para la operacion: " & Trim(p_NumOpe), vbInformation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   r_rst_SalAnt.MoveFirst
   Do While Not r_rst_SalAnt.EOF
      If p_NumOpe = r_rst_SalAnt!HIPCIE_NUMOPE Then
         fs_Obtiene_SaldoPerAnterior = r_rst_SalAnt!HIPCIE_DEVPBP * r_rst_SalAnt!NROMESES
         Exit Do
      End If
      r_rst_SalAnt.MoveNext
   Loop
   
   r_rst_SalAnt.Close
   Set r_rst_SalAnt = Nothing
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_str_Cadena     As String
Dim r_int_Contad     As Integer
Dim r_int_nrofil     As Integer
Dim r_int_Nindex     As Integer
Dim r_int_NroCol     As Integer
Dim r_dbl_SumDeb     As Double
Dim r_dbl_SumHab     As Double
Dim r_dbl_SumAju     As Double
Dim r_dbl_SumHabDol  As Double
Dim r_dbl_SumDebDol  As Double
Dim r_dbl_SumAjuDol  As Double
  
   'Prepara archivo excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_int_nrofil = 2
   
   '*******************
   'Reporte Interes PBP
   r_obj_Excel.Sheets(1).Name = "DATA PERIODO"
   With r_obj_Excel.Sheets(1)
   
      'Formato de titulos y columnas
      .Range("A" & r_int_nrofil & ":K" & r_int_nrofil & "").Merge
      .Range("A" & r_int_nrofil & ":K" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":K" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_nrofil & ":K" & r_int_nrofil & "").Font.Size = 14
      .Range("A" & r_int_nrofil & ":K" & r_int_nrofil & "").Font.Underline = xlUnderlineStyleSingle
      
      .Cells(r_int_nrofil, 1) = "REPORTE DE PROVISIONES DE INTERÉS PBP AL MES DE " & Trim(cmb_PerMes.Text) & " DEL " & Trim(ipp_PerAno.Text)
      .Columns("A").ColumnWidth = 12
      .Columns("B").ColumnWidth = 45
      .Columns("C").ColumnWidth = 9
      .Columns("D").ColumnWidth = 9
      .Columns("E").ColumnWidth = 9
      .Columns("F").ColumnWidth = 9
      .Columns("G").ColumnWidth = 9
      .Columns("H").ColumnWidth = 10
      .Columns("I").ColumnWidth = 10
      .Columns("J").ColumnWidth = 11
      .Columns("K").ColumnWidth = 11
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("D").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("E").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("H").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("H").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("J").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("K").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("L").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Cabecera Dolares
      r_int_nrofil = r_int_nrofil + 2
      .Cells(r_int_nrofil, 1) = "MONEDA"
      .Cells(r_int_nrofil, 2) = "DOLARES AMERICANOS"
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignLeft
      
      r_int_nrofil = r_int_nrofil + 2
      .Cells(r_int_nrofil, 1) = "OPERACION"
      .Cells(r_int_nrofil, 2) = "NOMBRE DE CLIENTE"
      For r_int_Nindex = 3 To 7
         .Cells(r_int_nrofil, r_int_Nindex) = "AJUSTE"
      Next r_int_Nindex
      .Cells(r_int_nrofil, 8) = "TOTAL"
      .Cells(r_int_nrofil, 9) = "GANA PBP"
      .Cells(r_int_nrofil, 10) = "SALDO ANT."
      .Cells(r_int_nrofil, 11) = "PROV. MES"
      .Cells(r_int_nrofil, 12) = "SALDO SIG."
      
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroCol = 7
      .Cells(r_int_nrofil, r_int_NroCol) = "'" & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      
      For r_int_Nindex = 1 To 4
         r_int_NroCol = r_int_NroCol - 1
         .Cells(r_int_nrofil, r_int_NroCol) = "'" & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
      Next r_int_Nindex
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).HorizontalAlignment = xlHAlignCenter
      r_int_nrofil = r_int_nrofil + 1
      .Cells(1, 1).Select
      
      'Informacion Dolares
      For r_int_Contad = 4 To grd_LisDolar.Rows + 1
         .Cells(r_int_nrofil, 1) = "'" & grd_LisDolar.TextMatrix(r_int_Contad - 2, 0)
         .Cells(r_int_nrofil, 2) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 1)
         .Cells(r_int_nrofil, 3) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 2)
         .Cells(r_int_nrofil, 4) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 3)
         .Cells(r_int_nrofil, 5) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 4)
         .Cells(r_int_nrofil, 6) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 5)
         .Cells(r_int_nrofil, 7) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 6)
         .Cells(r_int_nrofil, 8) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 7)
         .Cells(r_int_nrofil, 9) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 8)
         .Cells(r_int_nrofil, 10) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 9)
         .Cells(r_int_nrofil, 11) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 10)
         .Cells(r_int_nrofil, 12) = grd_LisDolar.TextMatrix(r_int_Contad - 2, 11)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Cells(r_int_nrofil - 1, 3).Font.Bold = True
      .Cells(r_int_nrofil - 1, 4).Font.Bold = True
      .Cells(r_int_nrofil - 1, 5).Font.Bold = True
      .Cells(r_int_nrofil - 1, 6).Font.Bold = True
      .Cells(r_int_nrofil - 1, 7).Font.Bold = True
      .Cells(r_int_nrofil - 1, 8).Font.Bold = True
      .Cells(r_int_nrofil - 1, 11).Font.Bold = True
      .Cells(r_int_nrofil - 1, 12).Font.Bold = True
      
      'Cabecera Soles
      r_int_nrofil = r_int_nrofil + 3
      .Cells(r_int_nrofil, 1) = "MONEDA"
      .Cells(r_int_nrofil, 2) = "SOLES"
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignLeft
      
      r_int_nrofil = r_int_nrofil + 2
      .Cells(r_int_nrofil, 1) = "OPERACION"
      .Cells(r_int_nrofil, 2) = "NOMBRE DE CLIENTE"
      For r_int_Nindex = 3 To 7
         .Cells(r_int_nrofil, r_int_Nindex) = "AJUSTE"
      Next r_int_Nindex
      .Cells(r_int_nrofil, 8) = "TOTAL"
      .Cells(r_int_nrofil, 9) = "GANA PBP"
      .Cells(r_int_nrofil, 10) = "SALDO ANT."
      .Cells(r_int_nrofil, 11) = "PROV. MES"
      .Cells(r_int_nrofil, 12) = "SALDO SIG."
      
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroCol = 7
      .Cells(r_int_nrofil, r_int_NroCol) = "'" & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yy"))
      
      For r_int_Nindex = 1 To 4
         r_int_NroCol = r_int_NroCol - 1
         .Cells(r_int_nrofil, r_int_NroCol) = "'" & UCase(Format$(DateAdd("m", -r_int_Nindex, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & "")), "mmm-yy"))
      Next r_int_Nindex
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).HorizontalAlignment = xlHAlignCenter
      r_int_nrofil = r_int_nrofil + 1
      .Cells(1, 1).Select
      
      'Informacion Soles
      For r_int_Contad = 4 To grd_LisSoles.Rows + 1
         .Cells(r_int_nrofil, 1) = "'" & grd_LisSoles.TextMatrix(r_int_Contad - 2, 0)
         .Cells(r_int_nrofil, 2) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 1)
         .Cells(r_int_nrofil, 3) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 2)
         .Cells(r_int_nrofil, 4) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 3)
         .Cells(r_int_nrofil, 5) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 4)
         .Cells(r_int_nrofil, 6) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 5)
         .Cells(r_int_nrofil, 7) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 6)
         .Cells(r_int_nrofil, 8) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 7)
         .Cells(r_int_nrofil, 9) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 8)
         .Cells(r_int_nrofil, 10) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 9)
         .Cells(r_int_nrofil, 11) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 10)
         .Cells(r_int_nrofil, 12) = grd_LisSoles.TextMatrix(r_int_Contad - 2, 11)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Cells(r_int_nrofil - 1, 3).Font.Bold = True
      .Cells(r_int_nrofil - 1, 4).Font.Bold = True
      .Cells(r_int_nrofil - 1, 5).Font.Bold = True
      .Cells(r_int_nrofil - 1, 6).Font.Bold = True
      .Cells(r_int_nrofil - 1, 7).Font.Bold = True
      .Cells(r_int_nrofil - 1, 8).Font.Bold = True
      .Cells(r_int_nrofil - 1, 11).Font.Bold = True
      .Cells(r_int_nrofil - 1, 12).Font.Bold = True
   End With
   
   '********************
   'Busca Tipo de cambio
   l_dbl_TipCam = 0
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT MAX(HIPCIE_TIPCAM) AS TIPO_CAMBIO"
   r_str_Cadena = r_str_Cadena & "  FROM CRE_HIPCIE  "
   r_str_Cadena = r_str_Cadena & " WHERE HIPCIE_PERMES = " & l_str_PerMes & " "
   r_str_Cadena = r_str_Cadena & "   AND HIPCIE_PERANO = " & l_str_PerAno & " "

   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_GenAux, 3) Then
      Exit Sub
   End If

   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      l_dbl_TipCam = CDbl(g_rst_GenAux!TIPO_CAMBIO)
   End If

   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing

   '******************
   'Asientos Contables
   r_obj_Excel.Sheets(2).Name = "ASIENTOS"
   With r_obj_Excel.Sheets(2)
      'Cabecera Soles
      r_int_nrofil = 2
      .Range("B" & r_int_nrofil & ":F" & r_int_nrofil & "").Merge
      .Cells(r_int_nrofil, 2) = "MONEDA : SOLES"
      .Cells(r_int_nrofil, 2).Font.Bold = True
      
      r_int_nrofil = r_int_nrofil + 1
      .Cells(r_int_nrofil, 2) = "CUENTA"
      .Cells(r_int_nrofil, 3) = "GLOSA"
      .Cells(r_int_nrofil, 4) = "FECHA CONTABLE"
      .Cells(r_int_nrofil, 5) = "TIPO CUENTA"
      .Cells(r_int_nrofil, 6) = "MONTO"
      .Cells(r_int_nrofil, 2).Font.Bold = True
      .Cells(r_int_nrofil, 3).Font.Bold = True
      .Cells(r_int_nrofil, 4).Font.Bold = True
      .Cells(r_int_nrofil, 5).Font.Bold = True
      .Cells(r_int_nrofil, 6).Font.Bold = True
      
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 20
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 12
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignRight
      
      'Informacion Soles
      r_dbl_SumDeb = 0
      r_dbl_SumHab = 0
      r_dbl_SumAju = 0
      For r_int_Contad = 4 To grd_LisSoles.Rows
         If CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10)) <> 0 Then
            r_int_nrofil = r_int_nrofil + 1
            If grd_LisSoles.TextMatrix(r_int_Contad - 2, 10) > 0 Then
               .Cells(r_int_nrofil, 5) = "HABER"
               If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
                  .Cells(r_int_nrofil, 2) = "'511401042501"
               Else
                  .Cells(r_int_nrofil, 2) = "'511401040601"
               End If
               r_dbl_SumHab = r_dbl_SumHab + CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10))
            Else
               .Cells(r_int_nrofil, 5) = "DEBE"
               If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
                  .Cells(r_int_nrofil, 2) = "'511401042501"
               Else
                  .Cells(r_int_nrofil, 2) = "'511401040601"
               End If
               r_dbl_SumDeb = r_dbl_SumDeb + CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10))
            End If
            If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
               .Cells(r_int_nrofil, 3) = "'" & grd_LisSoles.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-CME " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
            Else
               .Cells(r_int_nrofil, 3) = "'" & grd_LisSoles.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-MICASITA " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
            End If
            .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
            .Cells(r_int_nrofil, 6) = Format(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10), "###,###,##0.00")
         End If
      Next
      
      'Ajuste Soles
      r_int_nrofil = r_int_nrofil + 1
      .Cells(r_int_nrofil, 2) = "'151719010105" '"'151709090103"
      .Cells(r_int_nrofil, 3) = "APLIC. PBP-CME " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
      .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
      
      If r_dbl_SumDeb > r_dbl_SumHab Then
         .Cells(r_int_nrofil, 5) = "HABER"
         r_dbl_SumAju = r_dbl_SumDeb - r_dbl_SumHab
         .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju, "###,###,##0.00")
      ElseIf r_dbl_SumDeb < r_dbl_SumHab Then
         .Cells(r_int_nrofil, 5) = "DEBE"
         r_dbl_SumAju = r_dbl_SumHab - r_dbl_SumDeb
         .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju, "###,###,##0.00")
      End If
      
      'Cabecera Dolares
      r_int_nrofil = r_int_nrofil + 3
      .Range("B" & r_int_nrofil & ":F" & r_int_nrofil & "").Merge
      .Cells(r_int_nrofil, 2) = "MONEDA : DOLARES AMERICANOS"
      .Cells(r_int_nrofil, 2).Font.Bold = True
      
      r_int_nrofil = r_int_nrofil + 1
      .Cells(r_int_nrofil, 2) = "CUENTA"
      .Cells(r_int_nrofil, 3) = "GLOSA"
      .Cells(r_int_nrofil, 4) = "FECHA CONTABLE"
      .Cells(r_int_nrofil, 5) = "TIPO CUENTA"
      .Cells(r_int_nrofil, 6) = "MONTO"
      .Cells(r_int_nrofil, 2).Font.Bold = True
      .Cells(r_int_nrofil, 3).Font.Bold = True
      .Cells(r_int_nrofil, 4).Font.Bold = True
      .Cells(r_int_nrofil, 5).Font.Bold = True
      .Cells(r_int_nrofil, 6).Font.Bold = True
      
      'Informacion Dolares
      r_dbl_SumDeb = 0
      r_dbl_SumHab = 0
      r_dbl_SumAju = 0
      For r_int_Contad = 4 To grd_LisDolar.Rows
         If CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10)) <> 0 Then
            r_int_nrofil = r_int_nrofil + 1
            If grd_LisDolar.TextMatrix(r_int_Contad - 2, 10) > 0 Then
               .Cells(r_int_nrofil, 5) = "HABER"
               .Cells(r_int_nrofil, 2) = "'512401042401"
               r_dbl_SumHab = r_dbl_SumHab + Format(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10) * l_dbl_TipCam), "###,##0.00")
               r_dbl_SumHabDol = r_dbl_SumHabDol + CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10))
            Else
               .Cells(r_int_nrofil, 5) = "DEBE"
               .Cells(r_int_nrofil, 2) = "'512401042401"
               r_dbl_SumDeb = r_dbl_SumDeb + Format(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10) * l_dbl_TipCam), "###,##0.00")
               r_dbl_SumDebDol = r_dbl_SumDebDol + CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10))
            End If
            .Cells(r_int_nrofil, 3) = "'" & grd_LisDolar.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-CRC " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
            .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
            .Cells(r_int_nrofil, 6) = Format(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10) * l_dbl_TipCam), "###,##0.00")
         End If
      Next
      
      'Ajuste Dolares
      r_int_nrofil = r_int_nrofil + 1
      .Cells(r_int_nrofil, 2) = "'152719010105" '"'152709090103"
      .Cells(r_int_nrofil, 3) = "APLIC. PBP-CRC-" & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
      .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
      
      If r_dbl_SumDebDol > r_dbl_SumHabDol Then
         r_dbl_SumAjuDol = r_dbl_SumDebDol - r_dbl_SumHabDol
         r_dbl_SumAju = Format(r_dbl_SumAjuDol * l_dbl_TipCam, "###,##0.00")
         r_dbl_SumAju = Format(r_dbl_SumAjuDol, "###,##0.00")
         .Cells(r_int_nrofil, 5) = "HABER"
         .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju, "###,##0.00")
      ElseIf r_dbl_SumDebDol < r_dbl_SumHabDol Then
         r_dbl_SumAjuDol = r_dbl_SumHabDol - r_dbl_SumDebDol
         r_dbl_SumAju = Format(r_dbl_SumAjuDol * l_dbl_TipCam, "###,##0.00")
         r_dbl_SumAju = Format(r_dbl_SumAjuDol, "###,##0.00")
         .Cells(r_int_nrofil, 5) = "DEBE"
         .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju, "###,##0.00")
      End If
      
      'Diferencia de cambio
      'If r_dbl_SumAju <> Abs(r_dbl_SumDebDol - r_dbl_SumHabDol) Then
      If Format(r_dbl_SumAju - (Abs(r_dbl_SumDebDol - r_dbl_SumHabDol)), "######0.00") <> 0 Then
         r_int_nrofil = r_int_nrofil + 1
         If r_dbl_SumDeb > r_dbl_SumHabDol Then
            .Cells(r_int_nrofil, 2) = "'412804090101"
            .Cells(r_int_nrofil, 3) = "PERDIDA DIF. CAMBIO PBP-CRC-" & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
            .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
            .Cells(r_int_nrofil, 5) = "HABER"
            .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju - (r_dbl_SumDeb - r_dbl_SumHab), "###,###,##0.00")
         Else
            .Cells(r_int_nrofil, 2) = "'512804090101"
            .Cells(r_int_nrofil, 3) = "GANANCIA DIF. CAMBIO PBP-CRC-" & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
            .Cells(r_int_nrofil, 4) = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & "/" & l_str_PerMes & "/" & l_str_PerAno
            .Cells(r_int_nrofil, 5) = "HABER"
            .Cells(r_int_nrofil, 6) = Format(r_dbl_SumAju - (r_dbl_SumHabDol - r_dbl_SumDebDol), "###,###,##0.00")
         End If
      End If
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function gf_Query_InteresPBP(ByVal p_PerMes As String, ByVal p_PerAno As String) As String
   gf_Query_InteresPBP = ""
   gf_Query_InteresPBP = gf_Query_InteresPBP & " SELECT A.HIPCIE_NUMOPE, A.HIPCIE_DEVPBP, A.HIPCIE_TIPMON, "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "        TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS CLIENTE, "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "        NVL((SELECT MONTHS_BETWEEN (TO_DATE('" & p_PerAno & p_PerMes & "'||SUBSTRC(HIPCUO_FECVCT,7,2), 'YYYY/MM/DD'),TO_DATE(HIPCUO_FECVCT, 'YYYY/MM/DD')) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "               FROM (SELECT H.HIPCUO_NUMOPE AS NUMOPE,HIPCUO_FECVCT, TO_DATE(H.HIPCUO_FECVCT, 'YYYY/MM/DD') AS FECHA, "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                            SUBSTRC(H.HIPCUO_FECVCT,5,2) AS MES "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                       FROM CRE_HIPCUO H "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                      WHERE HIPCUO_TIPCRO = 2 "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                        AND TO_DATE(HIPCUO_FECVCT, 'YYYY/MM/DD') <= "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                (SELECT ADD_MONTHS(TO_DATE('" & p_PerAno & p_PerMes & "'||SUBSTRC(HIPCUO_FECVCT,7,2), 'YYYY/MM/DD'),6) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                   FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = H.HIPCUO_NUMOPE AND HIPCUO_TIPCRO = 2 AND ROWNUM <2) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                        AND TO_DATE(HIPCUO_FECVCT, 'YYYY/MM/DD') >= "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                (SELECT ADD_MONTHS(TO_DATE('" & p_PerAno & p_PerMes & "'||SUBSTRC(HIPCUO_FECVCT,7,2), 'YYYY/MM/DD'),-6) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                   FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = H.HIPCUO_NUMOPE AND HIPCUO_TIPCRO = 2 AND ROWNUM <2) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                        AND TO_DATE(HIPCUO_FECVCT, 'YYYY/MM/DD') <= "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                (SELECT TO_DATE('" & p_PerAno & p_PerMes & "'||SUBSTRC(HIPCUO_FECVCT,7,2), 'YYYY/MM/DD')"
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                                   FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = H.HIPCUO_NUMOPE AND HIPCUO_TIPCRO = 2 AND ROWNUM <2) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                     ORDER BY H.HIPCUO_NUMOPE, H.HIPCUO_FECVCT DESC) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "              WHERE NUMOPE = HIPCIE_NUMOPE "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                AND SUBSTRC(HIPCUO_FECVCT,5,2) NOT IN (" & p_PerMes & ", (CASE WHEN " & p_PerMes & " > 6 THEN " & p_PerMes & " -6 ELSE " & p_PerMes & "+6 END)) "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "            ),0) AS NROMESES, "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "        NVL((SELECT DETPBP_FLGPBP "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "               FROM CRE_DETPBP "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "              WHERE DETPBP_PERMES = " & p_PerMes & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                AND DETPBP_PERANO = " & p_PerAno & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                AND DETPBP_NUMOPE = HIPCIE_NUMOPE "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "           ),0) AS FLAGPBP, "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "        NVL((SELECT DETPBP_INTADE "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "               FROM CRE_DETPBP "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "              WHERE DETPBP_PERMES = " & p_PerMes & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                AND DETPBP_PERANO = " & p_PerAno & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "                AND DETPBP_NUMOPE = HIPCIE_NUMOPE "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "           ),0) AS INTADE "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "   FROM CRE_HIPCIE A "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "  INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPCIE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPCIE_NDOCLI "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "  WHERE A.HIPCIE_PERMES = " & p_PerMes & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "    AND A.HIPCIE_PERANO = " & p_PerAno & " "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "    AND A.HIPCIE_DEVPBP > 0 "
   gf_Query_InteresPBP = gf_Query_InteresPBP & "    ORDER BY HIPCIE_TIPMON, HIPCIE_NUMOPE "
End Function

Private Sub fs_GeneraAsientos()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Cadena        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_str_Origen        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_CtaCtb        As String
Dim r_str_FecCtb        As String
Dim r_str_TipNot        As String
Dim r_str_DetGlo        As String
Dim r_str_TipCta        As String
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_dbl_SumDeb        As Double
Dim r_dbl_SumHab        As Double
Dim r_dbl_SumAju        As Double
Dim r_dbl_SumDebDol     As Double
Dim r_dbl_SumHabDol     As Double
Dim r_dbl_SumAjuDol     As Double
Dim r_str_AsiSol        As String
Dim r_str_AsiDol        As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1035"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   '********************
   'Busca Tipo de cambio
   l_dbl_TipCam = 0
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT MAX(HIPCIE_TIPCAM) AS TIPO_CAMBIO"
   r_str_Cadena = r_str_Cadena & "  FROM CRE_HIPCIE  "
   r_str_Cadena = r_str_Cadena & " WHERE HIPCIE_PERMES = " & l_str_PerMes & " "
   r_str_Cadena = r_str_Cadena & "   AND HIPCIE_PERANO = " & l_str_PerAno & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      l_dbl_TipCam = CDbl(g_rst_GenAux!TIPO_CAMBIO)
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   '*****************************
   'Genera Asientos: Nuevos Soles
   '*****************************
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_FecCtb = CStr(ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno))) & "/" & Format(CInt(l_str_PerMes), "00") & "/" & l_str_PerAno
   
   'Obteniendo numero de asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, CInt(l_str_PerAno), CInt(l_str_PerMes), r_str_Origen, r_int_NumLib)
   r_str_AsiSol = r_int_NumAsi
   
   'Insertar en cabecera
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), l_dbl_TipCam, r_str_TipNot, "APLIC. PBP CME-MICASITA " & UCase(fs_Obtiene_NombreMes(CInt(l_str_PerMes))) & "-" & l_str_PerAno, r_str_FecCtb, "1")
   
   r_dbl_SumDeb = 0
   r_dbl_SumHab = 0
   r_dbl_SumAju = 0
   r_int_NumIte = 0
   
   For r_int_Contad = 4 To grd_LisSoles.Rows
      If CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10)) <> 0 Then
         r_str_TipCta = ""
         r_str_CtaCtb = ""
         r_str_DetGlo = ""
         r_dbl_ImpSol = 0
         r_dbl_ImpDol = 0
         r_int_NumIte = r_int_NumIte + 1
         
         If grd_LisSoles.TextMatrix(r_int_Contad - 2, 10) > 0 Then
            r_str_TipCta = "H"
            If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
               r_str_CtaCtb = "511401042501"
            Else
               r_str_CtaCtb = "511401040601"
            End If
            r_dbl_ImpSol = Abs(CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10)))
            r_dbl_SumHab = r_dbl_SumHab + r_dbl_ImpSol
         Else
            r_str_TipCta = "D"
            If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
               r_str_CtaCtb = "511401042501"
            Else
               r_str_CtaCtb = "511401040601"
            End If
            r_dbl_ImpSol = Abs(CDbl(grd_LisSoles.TextMatrix(r_int_Contad - 2, 10)))
            r_dbl_SumDeb = r_dbl_SumDeb + r_dbl_ImpSol
         End If
         
         If Mid(grd_LisSoles.TextMatrix(r_int_Contad - 2, 0), 1, 3) = "003" Then
            r_str_DetGlo = grd_LisSoles.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-CME " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
         Else
            r_str_DetGlo = grd_LisSoles.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-MICASITA " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
         End If
         
         'Insertar en detalle
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCtb), r_str_DetGlo, r_str_TipCta, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecCtb))
      End If
   Next
   
   'Ajuste Soles
   r_int_NumIte = r_int_NumIte + 1
   r_str_CtaCtb = "151719010105" '"151709090103"
   r_str_DetGlo = "APLIC. PBP-CME " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
   
   If r_dbl_SumDeb > r_dbl_SumHab Then
      r_str_TipCta = "H"
      r_dbl_SumAju = r_dbl_SumDeb - r_dbl_SumHab
      r_dbl_ImpSol = Format(r_dbl_SumAju, "###,###,##0.00")
   ElseIf r_dbl_SumDeb < r_dbl_SumHab Then
      r_str_TipCta = "D"
      r_dbl_SumAju = r_dbl_SumHab - r_dbl_SumDeb
      r_dbl_ImpSol = Format(r_dbl_SumAju, "###,###,##0.00")
   End If
   
   'Insertar en ajuste en detalle
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCtb), r_str_DetGlo, r_str_TipCta, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecCtb))
   
   '***********************************
   'Genera Asientos: Dolares Americanos
   '***********************************
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_FecCtb = CStr(ff_Ultimo_Dia_Mes(CInt(l_str_PerMes), CInt(l_str_PerAno))) & "/" & Format(CInt(l_str_PerMes), "00") & "/" & l_str_PerAno
   
   'Obteniendo numero de asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, CInt(l_str_PerAno), CInt(l_str_PerMes), r_str_Origen, r_int_NumLib)
   r_str_AsiDol = r_int_NumAsi
   
   'Insertar en cabecera
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), l_dbl_TipCam, r_str_TipNot, "APLIC. PBP CME-MICASITA " & UCase(fs_Obtiene_NombreMes(CInt(l_str_PerMes))) & "-" & l_str_PerAno, r_str_FecCtb, "1")
   
   r_dbl_SumDeb = 0
   r_dbl_SumHab = 0
   r_dbl_SumAju = 0
   r_int_NumIte = 0
   r_dbl_SumDebDol = 0
   r_dbl_SumHabDol = 0
   
   For r_int_Contad = 4 To grd_LisDolar.Rows
      If CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10)) <> 0 Then
         r_str_TipCta = ""
         r_str_CtaCtb = ""
         r_str_DetGlo = ""
         r_dbl_ImpSol = 0
         r_dbl_ImpDol = 0
         r_int_NumIte = r_int_NumIte + 1
         
         If grd_LisDolar.TextMatrix(r_int_Contad - 2, 10) > 0 Then
            r_str_TipCta = "H"
            r_str_CtaCtb = "512401042401"
            r_dbl_ImpSol = Format(Abs(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10))) * l_dbl_TipCam, "###,##0.00")
            r_dbl_SumHab = r_dbl_SumHab + r_dbl_ImpSol
            r_dbl_SumHabDol = r_dbl_SumHabDol + Abs(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10)))
         Else
            r_str_TipCta = "D"
            r_str_CtaCtb = "512401042401"
            r_dbl_ImpSol = Format(Abs(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10))) * l_dbl_TipCam, "###,##0.00")
            r_dbl_SumDeb = r_dbl_SumDeb + r_dbl_ImpSol
            r_dbl_SumDebDol = r_dbl_SumDebDol + Abs(CDbl(grd_LisDolar.TextMatrix(r_int_Contad - 2, 10)))
         End If
         r_str_DetGlo = grd_LisDolar.TextMatrix(r_int_Contad - 2, 0) & "/" & "APLIC. PBP-CRC " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
         
         'Insertar en detalle
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCtb), r_str_DetGlo, r_str_TipCta, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecCtb))
      End If
   Next
   
   'Ajuste Dolares
   r_int_NumIte = r_int_NumIte + 1
   r_str_CtaCtb = "152719010105" '"152709090103"
   r_str_DetGlo = "APLIC. PBP-CRC " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
   
   If r_dbl_SumDebDol > r_dbl_SumHabDol Then
      r_dbl_SumAjuDol = r_dbl_SumDebDol - r_dbl_SumHabDol
      r_dbl_SumAju = Format(r_dbl_SumAjuDol * l_dbl_TipCam, "###,##0.00")
      r_dbl_ImpSol = Format(r_dbl_SumAjuDol, "###,##0.00")
      r_str_TipCta = "H"
   ElseIf r_dbl_SumDeb < r_dbl_SumHab Then
      r_dbl_SumAjuDol = r_dbl_SumHabDol - r_dbl_SumDebDol
      r_dbl_SumAju = Format(r_dbl_SumAjuDol * l_dbl_TipCam, "###,##0.00")
      r_dbl_ImpSol = Format(r_dbl_SumAjuDol, "###,##0.00")
      r_str_TipCta = "D"
   End If
   
   'Insertar en ajuste en detalle
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCtb), r_str_DetGlo, r_str_TipCta, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecCtb))
   
   'Diferencia de cambio
   If r_dbl_SumAju <> Abs(r_dbl_SumDeb - r_dbl_SumHab) Then
      r_int_NumIte = r_int_NumIte + 1
      If r_dbl_SumDeb > r_dbl_SumHab Then
         r_str_CtaCtb = "412804090101"
         r_str_DetGlo = "PERDIDA DIF. CAMBIO PBP-CRC " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
         r_str_TipCta = "H"
         r_dbl_ImpSol = Format(r_dbl_SumAju - (r_dbl_SumDeb - r_dbl_SumHab), "###,###,##0.00")
      Else
         r_str_CtaCtb = "512804090101"
         r_str_DetGlo = "GANANCIA DIF. CAMBIO PBP-CRC " & UCase(Format$(CDate("01/" & l_str_PerMes & "/" & l_str_PerAno & ""), "mmm-yyyy"))
         r_str_TipCta = "H"
         r_dbl_ImpSol = Format(r_dbl_SumAju - (r_dbl_SumHab - r_dbl_SumDeb), "###,###,##0.00")
      End If
      
      'Insertar en diferencia de cambio
      Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(l_str_PerAno), CInt(l_str_PerMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCtb), r_str_DetGlo, r_str_TipCta, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecCtb))
   End If
   
   Call modprc_fs_Actualiza_Proceso(CInt(l_str_PerAno), CInt(l_str_PerMes), 1)
   
   MsgBox "Se generaron los asientos contables satisfactoriamente." & vbCrLf & " - Asiento en moneda Soles     : " & Trim(r_str_AsiSol) & vbCrLf & " - Asiento en moneda Dolares : " & Trim(r_str_AsiDol), vbInformation, modgen_g_str_NomPlt
End Sub

Private Function fs_Obtiene_NombreMes(ByVal p_Mes As Integer) As String
   fs_Obtiene_NombreMes = ""
   Select Case p_Mes
      Case 1: fs_Obtiene_NombreMes = "ENE"
      Case 2: fs_Obtiene_NombreMes = "FEB"
      Case 3: fs_Obtiene_NombreMes = "MAR"
      Case 4: fs_Obtiene_NombreMes = "ABR"
      Case 5: fs_Obtiene_NombreMes = "MAY"
      Case 6: fs_Obtiene_NombreMes = "JUN"
      Case 7: fs_Obtiene_NombreMes = "JUL"
      Case 8: fs_Obtiene_NombreMes = "AGO"
      Case 9: fs_Obtiene_NombreMes = "SET"
      Case 10: fs_Obtiene_NombreMes = "OCT"
      Case 11: fs_Obtiene_NombreMes = "NOV"
      Case 12: fs_Obtiene_NombreMes = "DIC"
   End Select
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
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
