VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbPbp_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   Icon            =   "GesCtb_frm_914.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel6 
      Height          =   735
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   1296
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Left            =   660
         TabIndex        =   9
         Top             =   60
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Proceso"
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   315
         Left            =   660
         TabIndex        =   10
         Top             =   360
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Contabilización de Asignación PBP"
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
         Left            =   80
         Picture         =   "GesCtb_frm_914.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   675
      Left            =   30
      TabIndex        =   11
      Top             =   750
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
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
      Begin VB.CommandButton cmd_Buscar 
         Height          =   585
         Left            =   30
         Picture         =   "GesCtb_frm_914.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Registros"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Limpia 
         Height          =   585
         Left            =   660
         Picture         =   "GesCtb_frm_914.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Datos"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   11230
         Picture         =   "GesCtb_frm_914.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   1275
         Picture         =   "GesCtb_frm_914.frx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exportar a Excel"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Proces 
         Enabled         =   0   'False
         Height          =   585
         Left            =   2520
         Picture         =   "GesCtb_frm_914.frx":1076
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Generar asientos automaticos"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmd_Detalle 
         Height          =   585
         Left            =   1890
         Picture         =   "GesCtb_frm_914.frx":1380
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ver Detalle"
         Top             =   60
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   810
      Left            =   30
      TabIndex        =   12
      Top             =   1440
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   1429
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
         ItemData        =   "GesCtb_frm_914.frx":17C2
         Left            =   1110
         List            =   "GesCtb_frm_914.frx":17C4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2265
      End
      Begin EditLib.fpLongInteger ipp_PerAno 
         Height          =   315
         Left            =   1110
         TabIndex        =   1
         Top             =   405
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
      Begin VB.Label Label10 
         Caption         =   "Mes:"
         Height          =   315
         Left            =   135
         TabIndex        =   14
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   315
         Left            =   135
         TabIndex        =   13
         Top             =   420
         Width           =   1365
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   5415
      Left            =   30
      TabIndex        =   15
      Top             =   2250
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   9551
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
      Begin Threed.SSPanel pnl_Tit_NomCli 
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   5700
         _Version        =   65536
         _ExtentX        =   10054
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Producto"
         ForeColor       =   16777215
         BackColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   285
         Left            =   10185
         TabIndex        =   17
         Top             =   90
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "  Seleccionar"
         ForeColor       =   16777215
         BackColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
         Begin VB.CheckBox chkSeleccionar 
            BackColor       =   &H00004000&
            Caption         =   "Check1"
            Height          =   255
            Left            =   1030
            TabIndex        =   18
            Top             =   20
            Width           =   255
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   285
         Left            =   5775
         TabIndex        =   19
         Top             =   90
         Width           =   2220
         _Version        =   65536
         _ExtentX        =   3916
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Capital Tramo Cofide/MVI"
         ForeColor       =   16777215
         BackColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   285
         Left            =   7980
         TabIndex        =   20
         Top             =   90
         Width           =   2220
         _Version        =   65536
         _ExtentX        =   3916
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Interés Tramo Cofide/MVI"
         ForeColor       =   16777215
         BackColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid grd_Listad 
         Height          =   4995
         Left            =   75
         TabIndex        =   21
         Top             =   330
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   8811
         _Version        =   393216
         Rows            =   4
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPbp_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer
Dim l_rst_RstPbp        As ADODB.Recordset

Private Sub cmd_Buscar_Click()
Dim r_str_Cadena     As String
Dim r_rst_Record     As ADODB.Recordset
Dim r_int_NumVec     As Integer
   
   If Trim(cmb_PerMes.Text) = "" Then
      MsgBox "Debe seleccionar el tipo de mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   l_int_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_int_PerAno = ipp_PerAno.Text
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
   
   If (grd_Listad.Rows = 0) Then
      Call cmd_Limpia_Click
   End If
   
   'Valida si contabilizacion ya fue procesada
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS NUM_EJEC "
   r_str_Cadena = r_str_Cadena & "  FROM CTB_PERPRO "
   r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(l_int_PerAno) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(l_int_PerMes) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = 3 "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 3) Then
      Exit Sub
   End If
   
   r_rst_Record.MoveFirst
   r_int_NumVec = r_rst_Record!NUM_EJEC
   
   r_rst_Record.Close
   Set r_rst_Record = Nothing
   
   If r_int_NumVec > 0 Then
      MsgBox "Período seleccionado ya fue contabilizado.", vbExclamation, modgen_g_str_NomPlt
      cmd_Proces.Enabled = False
      chkSeleccionar.Enabled = False
      chkSeleccionar.Value = 0
      Exit Sub
   End If
   
   'Verifica periodos pasados
   If l_int_PerAno <= 2015 And l_int_PerMes <= 4 Then
      MsgBox "Período seleccionado ya esta cerrado.", vbExclamation, modgen_g_str_NomPlt
      cmd_Proces.Enabled = False
      chkSeleccionar.Enabled = False
      chkSeleccionar.Value = 0
      Exit Sub
   End If
End Sub

Private Sub cmd_Limpia_Click()
   grd_Listad.Rows = 0
   cmb_PerMes.ListIndex = -1
   chkSeleccionar.Value = 0
   ipp_PerAno.Text = Year(date)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Detalle_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'NOMBRE DEL PRODUCTO
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = Trim(grd_Listad.Text)
   
   'CÓDIGO DEL PRODUCTO
   grd_Listad.Col = 4
   moddat_g_str_CodPrd = Trim(grd_Listad.Text)
   
   'MES
   moddat_g_int_EdaMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   'AÑO
   moddat_g_int_EdaAno = CStr(ipp_PerAno.Value)
      
   Call gs_RefrescaGrid(grd_Listad)
      
   frm_Pro_CtbPbp_02.Show 1
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 3) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionados registros para generar asientos automáticos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro de generar los asientos contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiento
   Call cmd_Limpia_Click
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   'Call fs_BuscaPeriodo
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
        
   grd_Listad.ColWidth(0) = 5675       'Producto
   grd_Listad.ColWidth(1) = 2195       'importe del capital COFIDE
   grd_Listad.ColWidth(2) = 2195       'importe del interés COFIDE
   grd_Listad.ColWidth(3) = 1330       'Seleccionar
   grd_Listad.ColWidth(4) = 0          'Codprod
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.Rows = 0
End Sub

Private Sub fs_BuscaPeriodo()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PERMES_CODANO, PERMES_CODMES "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND PERMES_TIPPER = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se pudo determinar el período actual.", vbInformation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   l_int_PerMes = g_rst_Princi!PERMES_CODMES
   l_int_PerAno = g_rst_Princi!PERMES_CODANO

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If cmb_PerMes.ListIndex > -1 Then
      If KeyAscii = 13 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   cmd_Proces.Enabled = Not p_Activa
   cmd_Detalle.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
   chkSeleccionar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
Dim swtMes              As String
Dim swtAno              As String
Dim r_dbl_MVCapAde      As Double
Dim r_dbl_MVIntAde      As Double
Dim r_str_CodProd       As String

   Call gs_LimpiaGrid(grd_Listad)
   
   swtMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   swtAno = CStr(ipp_PerAno.Value)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT B.HIPMAE_CODPRD, SUM(A.DETPBP_CAPCLI) AS CAPCLI, SUM(A.DETPBP_INTCLI) AS INTCLI, "
   g_str_Parame = g_str_Parame & "        SUM(A.DETPBP_CAPADE) AS CAPADE, SUM(A.DETPBP_INTADE) AS INTADE "
   g_str_Parame = g_str_Parame & "   FROM CRE_DETPBP A, CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & "  WHERE DETPBP_PERMES = " & swtMes & " "
   g_str_Parame = g_str_Parame & "    AND DETPBP_PERANO = " & swtAno & " "
   g_str_Parame = g_str_Parame & "    AND DETPBP_NUMOPE = HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "    AND DETPBP_FLGPBP = 1 "
   g_str_Parame = g_str_Parame & "  GROUP BY HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  ORDER BY HIPMAE_CODPRD "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_RstPbp, 3) Then
       Exit Sub
   End If

   If l_rst_RstPbp.BOF And l_rst_RstPbp.EOF Then
      l_rst_RstPbp.Close
      Set l_rst_RstPbp = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   l_rst_RstPbp.MoveFirst
   Do While Not l_rst_RstPbp.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      If l_rst_RstPbp!HIPMAE_CODPRD <> "001" And l_rst_RstPbp!HIPMAE_CODPRD <> "003" And l_rst_RstPbp!HIPMAE_CODPRD <> "004" And l_rst_RstPbp!HIPMAE_CODPRD <> "006" Then
         r_dbl_MVCapAde = r_dbl_MVCapAde + l_rst_RstPbp!CAPADE
         r_dbl_MVIntAde = r_dbl_MVIntAde + l_rst_RstPbp!INTADE
         r_str_CodProd = r_str_CodProd & " , " & l_rst_RstPbp!HIPMAE_CODPRD
         grd_Listad.Rows = grd_Listad.Rows - 1
         
      ElseIf l_rst_RstPbp!HIPMAE_CODPRD = "006" Then
         'PRODUCTO
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_Produc(l_rst_RstPbp!HIPMAE_CODPRD)

         'CAPITAL TRAMO CLIENTE
         grd_Listad.Col = 1
         grd_Listad.Text = Format(l_rst_RstPbp!CAPCLI, "###,###,###,##0.00")

         'INTERES TRAMO CLIENTE
         grd_Listad.Col = 2
         grd_Listad.Text = Format(l_rst_RstPbp!INTCLI, "###,###,###,##0.00")
         
         'CÓDIGO DEL PRODUCTO
         grd_Listad.Col = 4
         grd_Listad.Text = l_rst_RstPbp!HIPMAE_CODPRD

      Else
         'PRODUCTO
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_Produc(l_rst_RstPbp!HIPMAE_CODPRD)
         
         'CAPITAL TRAMO COFIDE/MVI
         grd_Listad.Col = 1
         grd_Listad.Text = Format(l_rst_RstPbp!CAPADE, "###,###,###,##0.00")
         
         'INTERES TRAMO COFIDE/MVI
         grd_Listad.Col = 2
         grd_Listad.Text = Format(l_rst_RstPbp!INTADE, "###,###,###,##0.00")
         
         'CÓDIGO DEL PRODUCTO
         grd_Listad.Col = 4
         grd_Listad.Text = l_rst_RstPbp!HIPMAE_CODPRD
      End If

      l_rst_RstPbp.MoveNext
   Loop
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   
   'PRODUCTO
   grd_Listad.Col = 0
   grd_Listad.Text = "CREDITO MIVIVIENDA"
     
   'CAPITAL TRAMO COFIDE/MVI
   grd_Listad.Col = 1
   grd_Listad.Text = Format(r_dbl_MVCapAde, "###,###,###,##0.00")
   
   'INTERES TRAMO COFIDE/MVI
   grd_Listad.Col = 2
   grd_Listad.Text = Format(r_dbl_MVIntAde, "###,###,###,##0.00")
   
   'CÓDIGO DEL PRODUCTO
   grd_Listad.Col = 4
   grd_Listad.Text = Trim(Mid(r_str_CodProd, 3))

   grd_Listad.Redraw = True
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   r_dbl_MVCapAde = 0
   r_dbl_MVIntAde = 0
         
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GeneraAsiento()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_dbl_TipSbs        As Double
Dim r_str_FecPbpC       As String
Dim r_str_FecPbpL       As String
Dim r_str_CtaCtb        As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_str_Codigo        As String
Dim r_int_Colum         As Integer
Dim r_int_Total         As Integer
Dim r_str_DebHab        As String
Dim r_str_Glosa         As String
Dim r_str_DesGlosa      As String
Dim r_dbl_Importe       As Double
Dim r_str_MtoCof        As String
Dim r_str_CapCof        As String
Dim r_str_IntCof        As String
Dim r_str_CapCof2       As String
Dim r_str_CapCof3       As String
Dim r_dbl_CapFMV        As Double
Dim r_dbl_CapMVPE       As Double
Dim r_dbl_CapMV         As Double
Dim r_str_Cadena        As String
Dim r_str_CodAux        As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, 1))
      If (grd_Listad.TextMatrix(r_int_Contad, 3) = "X") And (r_dbl_Importe > 0) Then
         '*************************************************
         'GENERACION DE ASIENTOS CONTABLES DEL PBP
         '*************************************************
         
         'Inicializa variables
         r_int_NumAsi = 0
         r_str_FecPbpC = Format(ff_Ultimo_Dia_Mes(IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))), CInt(ipp_PerAno.Text)), "00") & "/" & IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) & "/" & CInt(ipp_PerAno.Text)
         r_str_FecPbpL = moddat_g_str_FecSis
         r_str_Codigo = grd_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, 2, Format(r_str_FecPbpL, "yyyymmdd"), 1)
         r_str_DesGlosa = "APLICACION DEL PBP " & grd_Listad.TextMatrix(r_int_Contad, 0)
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
         r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
         
         'Insertar en CABECERA
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_DesGlosa), r_str_FecPbpC, "1")
         
         'Inicializa
         r_int_NumIte = 0
         
         Select Case r_str_Codigo
            'CREDITO CRC-PBP
            Case "001"
               r_str_MtoCof = "152719010110"
               r_str_CapCof = "142104240101"
               r_str_IntCof = "152719010105"
               
            'CREDITO CME
            Case "003"
               r_str_MtoCof = "151719010103"
               r_str_CapCof = "141104250101"
               r_str_IntCof = "151719010105"
               
            'CREDITO PROYECTO MIHOGAR
            Case "004"
               r_str_MtoCof = "261202010101"
               r_str_CapCof = "141104230101"
               
            'CREDITO MICASITA SOLES
            Case "006"
               r_str_MtoCof = "511401040601"
               r_str_CapCof = "141104060101"
               r_str_IntCof = "511401040601"
               
            'CREDITO MIVIVIENDA
            Case "007", "009", "010", "012", "013", "014", "015", "016", "017", "018"
               r_str_MtoCof = "191807010101"
               r_str_CapCof = "141104230102"
               r_str_CapCof2 = "141104230104"
               r_str_CapCof3 = "141104230103"
         End Select
         
         r_str_DebHab = "H"
         Select Case r_str_Codigo
            Case "001"
               r_int_Colum = 0
               r_int_Total = grd_Listad.Cols - 2
               
               Do While (r_int_Colum < r_int_Total)
                  If r_int_Colum = 1 Then
                     r_str_Glosa = "APLICACION PBP CRC - CAP"
                     r_str_CtaCtb = r_str_CapCof
                     r_str_DebHab = "H"
                  ElseIf r_int_Colum = 2 Then
                     r_str_Glosa = "APLICACION PBP CRC - INT"
                     r_str_CtaCtb = r_str_IntCof
                     r_str_DebHab = "H"
                  Else
                     r_str_Glosa = "APLICACION PBP CRC"
                     r_str_CtaCtb = r_str_MtoCof
                     r_str_DebHab = "D"
                  End If
                  
                  If r_int_Colum = 0 Then
                     r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, 1)) + CDbl(grd_Listad.TextMatrix(r_int_Contad, 2))
                  Else
                     r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, r_int_Colum))
                  End If
                  
                  If (r_dbl_Importe > 0) Then
                     r_int_NumIte = r_int_NumIte + 1
                     
                     If r_str_Codigo = "001" Then
                        r_dbl_MtoDol = Format(r_dbl_Importe, "###,###,##0.00")                     'Importe dolares
                     Else
                        r_dbl_MtoDol = Format(0, "###,###,##0.00")
                     End If
                     
                     r_dbl_MtoSol = Format(CDbl(r_dbl_MtoDol * r_dbl_TipSbs), "###,###,##0.00") 'Importe soles
                     
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecPbpC), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPbpL))
                     r_dbl_Importe = 0
                  End If
                  
                  r_int_Colum = r_int_Colum + 1
               Loop
               
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & " SELECT TOT_DOLDEB, TOT_DOLHAB "
               g_str_Parame = g_str_Parame & "   FROM CNTBL_ASIENTO "
               g_str_Parame = g_str_Parame & "  WHERE ORIGEN = '" & r_str_Origen & "' AND ANO = " & l_int_PerAno & " AND MES = " & l_int_PerMes & " "
               g_str_Parame = g_str_Parame & "    AND NRO_LIBRO = " & r_int_NumLib & " AND NRO_ASIENTO = " & r_int_NumAsi & " "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                   MsgBox "No se ejecutó correctamente la consulta.", vbInformation, modgen_g_str_NomPlt
                   Exit Sub
               End If
               
               If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
                  If g_rst_Princi!TOT_DOLDEB <> g_rst_Princi!TOT_DOLHAB Then
                     Call modprc_fs_Actualiza_TotDol(r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi)
                  End If
               End If
               
            Case "003", "004", "006"
               r_int_Colum = 0
               r_int_Total = grd_Listad.Cols - 2
               
               Do While (r_int_Colum < r_int_Total)
                  If r_int_Colum = 1 Then
                     r_str_Glosa = "APLICACION PBP " & IIf(r_str_Codigo = "003", "CME", IIf(r_str_Codigo = "004", "MIHOGAR", "MICASITA")) & " - CAP"     '"CAP. TRAMO COFIDE/MVI"
                     r_str_CtaCtb = r_str_CapCof
                     r_str_DebHab = "H"
                  ElseIf r_int_Colum = 2 And r_str_Codigo <> "004" Then
                     r_str_Glosa = "APLICACION PBP " & IIf(r_str_Codigo = "003", "CME", IIf(r_str_Codigo = "004", "MIHOGAR", "MICASITA")) & " - INT"     '"INT. TRAMO COFIDE/MVI"
                     r_str_CtaCtb = r_str_IntCof
                     r_str_DebHab = "H"
                  Else
                     If r_str_Codigo = "004" Then
                        r_str_Glosa = "APLICACION PBP " & IIf(r_str_Codigo = "003", "CME", IIf(r_str_Codigo = "004", "MIHOGAR", "MICASITA")) & " - CAP"  '"MTO CAP. TRAMO COFIDE/MVI"
                     Else
                        r_str_Glosa = "APLICACION PBP " & IIf(r_str_Codigo = "003", "CME", IIf(r_str_Codigo = "004", "MIHOGAR", "MICASITA"))             '"MTO CAP.+INT. TRAMO COFIDE/MVI"
                     End If
                     r_str_CtaCtb = r_str_MtoCof
                     r_str_DebHab = "D"
                  End If
                  
                  If r_int_Colum = 0 And r_str_Codigo <> "004" Then
                     r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, 1)) + CDbl(grd_Listad.TextMatrix(r_int_Contad, 2))
                  Else
                     If r_str_Codigo = "004" And r_int_Colum <> 2 Then
                        r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, 1))
                     ElseIf r_str_Codigo <> "004" Then
                        r_dbl_Importe = CDbl(grd_Listad.TextMatrix(r_int_Contad, r_int_Colum))
                     End If
                  End If
                           
                  If (r_dbl_Importe > 0) Then
                     r_int_NumIte = r_int_NumIte + 1
                     r_dbl_MtoSol = Format(r_dbl_Importe, "###,###,##0.00") 'Importe soles
                     
                     If r_str_Codigo = "001" Then
                        r_dbl_MtoDol = Format(CDbl(r_dbl_MtoSol / r_dbl_TipSbs), "###,###,##0.00")  'Importe dolares
                     Else
                        r_dbl_MtoDol = Format(0, "###,###,##0.00")
                     End If
                     
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecPbpC), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPbpL))
                     r_dbl_Importe = 0
                  End If
                  
                  r_int_Colum = r_int_Colum + 1
               Loop
               
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & " SELECT TOT_DOLDEB, TOT_DOLHAB "
               g_str_Parame = g_str_Parame & "   FROM CNTBL_ASIENTO "
               g_str_Parame = g_str_Parame & "  WHERE ORIGEN = '" & r_str_Origen & "' AND ANO = " & l_int_PerAno & " AND MES = " & l_int_PerMes & " "
               g_str_Parame = g_str_Parame & "    AND NRO_LIBRO = " & r_int_NumLib & " AND NRO_ASIENTO = " & r_int_NumAsi & " "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                   MsgBox "No se ejecutó correctamente la consulta.", vbInformation, modgen_g_str_NomPlt
                   Exit Sub
               End If
               
               If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
                  If g_rst_Princi!TOT_DOLDEB <> g_rst_Princi!TOT_DOLHAB Then
                     Call modprc_fs_Actualiza_TotDol(r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi)
                  End If
               End If
         End Select
        
         If InStr(r_str_Codigo, ",") <> 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & " SELECT B.HIPMAE_CODPRD, SUM(A.DETPBP_CAPCLI) AS CAPCLI, SUM(A.DETPBP_INTCLI) AS INTCLI, "
            g_str_Parame = g_str_Parame & "        SUM(A.DETPBP_CAPADE) AS CAPADE, SUM(A.DETPBP_INTADE) AS INTADE "
            g_str_Parame = g_str_Parame & "  FROM  CRE_DETPBP A, CRE_HIPMAE B "
            g_str_Parame = g_str_Parame & " WHERE  DETPBP_PERMES = " & IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) & " "
            g_str_Parame = g_str_Parame & "   AND  DETPBP_PERANO = " & CInt(ipp_PerAno.Text) & " "
            g_str_Parame = g_str_Parame & "   AND  DETPBP_NUMOPE = HIPMAE_NUMOPE "
            g_str_Parame = g_str_Parame & "   AND  DETPBP_FLGPBP = 1 "
            g_str_Parame = g_str_Parame & "   AND  HIPMAE_CODPRD IN (" & r_str_Codigo & ")"
            g_str_Parame = g_str_Parame & " GROUP  BY HIPMAE_CODPRD "
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                Exit Sub
            End If
            
            If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
            
            Do Until g_rst_Princi.EOF
               r_str_Codigo = g_rst_Princi!HIPMAE_CODPRD
               
               Select Case r_str_Codigo
                  Case "007", "009", "010", "012", "013", "014", "015", "016", "017", "018"
                     r_str_MtoCof = "191807010101"
                     r_str_CapCof = "141104230102"
                     r_str_CapCof2 = "141104230104"
                     r_str_CapCof3 = "141104230103"
               
                     If r_str_Codigo = "009" Then
                        r_dbl_CapFMV = r_dbl_CapFMV + g_rst_Princi!CAPADE
                     ElseIf r_str_Codigo = "010" Then
                        r_dbl_CapMVPE = r_dbl_CapMVPE + g_rst_Princi!CAPADE
                     Else
                        r_dbl_CapMV = r_dbl_CapMV + g_rst_Princi!CAPADE
                     End If
               End Select
               g_rst_Princi.MoveNext
            Loop
            
            If r_dbl_CapFMV >= 0 Or r_dbl_CapMVPE > 0 Or r_dbl_CapMV > 0 Then
               For r_int_Total = 1 To 4
                  r_dbl_Importe = 0
                  If r_int_Total = 1 Then
                     r_str_Glosa = "APLICACION PBP MIVIVIENDA"
                     r_str_CtaCtb = r_str_MtoCof
                     r_str_DebHab = "D"
                     r_dbl_Importe = CDbl(r_dbl_CapFMV + r_dbl_CapMVPE + r_dbl_CapMV)
                  ElseIf r_int_Total = 2 Then
                     r_str_Glosa = "APLICACION PBP MIVIVIENDA - CAP"
                     r_str_CtaCtb = r_str_CapCof
                     r_str_DebHab = "H"
                     r_dbl_Importe = CDbl(r_dbl_CapMV)
                  ElseIf r_int_Total = 3 Then
                     r_str_Glosa = "APLICACION PBP MIVIVIENDA - PER EXT CAP"
                     r_str_CtaCtb = r_str_CapCof2
                     r_str_DebHab = "H"
                     r_dbl_Importe = CDbl(r_dbl_CapMVPE)
                  ElseIf r_int_Total = 4 Then
                     r_str_Glosa = "APLICACION PBP MIVIVIENDA - UNI AND CAP"
                     r_str_CtaCtb = r_str_CapCof3
                     r_str_DebHab = "H"
                     r_dbl_Importe = CDbl(r_dbl_CapFMV)
                  End If
                  
                  If (r_dbl_Importe > 0) Then
                     r_int_NumIte = r_int_NumIte + 1
                     r_dbl_MtoSol = Format(r_dbl_Importe, "###,###,##0.00") 'Importe soles
                     r_dbl_MtoDol = Format(0, "###,###,##0.00") 'r_dbl_MtoDol = Format(CDbl(r_dbl_MtoSol / r_dbl_TipSbs), "###,###,##0.00")  'Importe * CONVERTIDO
                     
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecPbpC), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPbpL))
                  End If
               Next r_int_Total
            End If
            
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & " SELECT TOT_DOLDEB, TOT_DOLHAB FROM CNTBL_ASIENTO "
            g_str_Parame = g_str_Parame & " WHERE ORIGEN = '" & r_str_Origen & "' AND ANO = " & l_int_PerAno & " AND MES = " & l_int_PerMes & " "
            g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = " & r_int_NumLib & " AND NRO_ASIENTO = " & r_int_NumAsi & " "
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
               MsgBox "No se ejecutó correctamente la consulta.", vbInformation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
               If g_rst_Princi!TOT_DOLDEB <> g_rst_Princi!TOT_DOLHAB Then
                  Call modprc_fs_Actualiza_TotDol(r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi)
               End If
            End If
         
         End If
      End If
   Next r_int_Contad
   
   Call modprc_fs_Actualiza_Proceso(l_int_PerAno, l_int_PerMes, 3)
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros seleccionados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub modprc_fs_Actualiza_TotDol(ByVal Origen As String, ByVal PerAno As Integer, ByVal PerMes As Integer, ByVal r_int_NumLib As Integer, ByVal NumAsi As Integer)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " UPDATE CNTBL_ASIENTO_DET SET IMP_MOVDOL = (SELECT SUM(A.IMP_MOVDOL) "
   g_str_Parame = g_str_Parame & "                                              FROM CNTBL_ASIENTO_DET A "
   g_str_Parame = g_str_Parame & "                                             WHERE A.ORIGEN = '" & Origen & "'"
   g_str_Parame = g_str_Parame & "                                               AND A.ANO = " & PerAno & "  "
   g_str_Parame = g_str_Parame & "                                               AND A.MES = " & PerMes & " "
   g_str_Parame = g_str_Parame & "                                               AND A.NRO_LIBRO = " & r_int_NumLib & " "
   g_str_Parame = g_str_Parame & "                                               AND A.NRO_ASIENTO = " & NumAsi & " "
   g_str_Parame = g_str_Parame & "                                               AND A.FLAG_DEBHAB = 'H')"
   g_str_Parame = g_str_Parame & "                       WHERE ORIGEN = '" & Origen & "'"
   g_str_Parame = g_str_Parame & "                         AND ANO = " & PerAno & " "
   g_str_Parame = g_str_Parame & "                         AND MES = 1 AND NRO_LIBRO = " & r_int_NumLib & " "
   g_str_Parame = g_str_Parame & "                         AND NRO_ASIENTO = " & NumAsi & " "
   g_str_Parame = g_str_Parame & "                         AND FLAG_DEBHAB = 'D' "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACIÓN DE ASIENTOS DEL PBP: " & Trim(cmb_PerMes.Text) & " - " & ipp_PerAno.Text
      .Range(.Cells(2, 2), .Cells(2, 5)).Merge
      .Range(.Cells(2, 2), .Cells(2, 5)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 5)).Font.Size = 12
      
      .Range(.Cells(2, 2), .Cells(2, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(2, 2), .Cells(2, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Cells(5, 2) = "N°"
      .Cells(5, 3) = "PRODUCTO"
      .Cells(4, 4) = "TRAMO COFIDE"
      .Cells(5, 4) = "CAPITAL (S/.)"
      .Cells(5, 5) = "INTERES (S/.)"
      
      .Range(.Cells(4, 2), .Cells(5, 2)).Merge
      .Range(.Cells(4, 3), .Cells(5, 3)).Merge
      .Range(.Cells(4, 4), .Cells(4, 5)).Merge
      .Range(.Cells(4, 2), .Cells(4, 5)).Font.Bold = True
      .Range(.Cells(4, 4), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(4, 2), .Cells(5, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(4, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(4, 2), .Cells(5, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(5, 5)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(5, 5)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("B").VerticalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      .Columns("C").VerticalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("E").ColumnWidth = 15
     
      .Range(.Cells(5, 1), .Cells(5, 5)).Font.Name = "Calibri"
      .Range(.Cells(5, 1), .Cells(5, 5)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous

         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 5)).Font.Size = 10
                  
         .Cells(r_int_NumFil + 3, 2) = r_int_NumFil - 2
         .Cells(r_int_NumFil + 3, 3) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0))   'PRODUCTO
         .Cells(r_int_NumFil + 3, 4) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 1))         'CAPITAL TRAMO COFIDE
         .Cells(r_int_NumFil + 3, 5) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 2))         'INTERES TRAMO COFIDE
  
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 3) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 3) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
