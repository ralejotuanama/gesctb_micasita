VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbBbp_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10125
   Icon            =   "GesCtb_frm_923.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10120
      _Version        =   65536
      _ExtentX        =   17851
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
         TabIndex        =   10
         Top             =   60
         Width           =   10005
         _Version        =   65536
         _ExtentX        =   17648
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
            TabIndex        =   11
            Top             =   30
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Left            =   570
            TabIndex        =   12
            Top             =   270
            Width           =   5235
            _Version        =   65536
            _ExtentX        =   9234
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Contabilización de Desembolsos BBP"
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
            Picture         =   "GesCtb_frm_923.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   10005
         _Version        =   65536
         _ExtentX        =   17648
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9380
            Picture         =   "GesCtb_frm_923.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1280
            Picture         =   "GesCtb_frm_923.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   2500
            Picture         =   "GesCtb_frm_923.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Detail 
            Height          =   585
            Left            =   1890
            Picture         =   "GesCtb_frm_923.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_923.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   40
            Picture         =   "GesCtb_frm_923.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   4650
         Left            =   60
         TabIndex        =   14
         Top             =   2325
         Width           =   10005
         _Version        =   65536
         _ExtentX        =   17648
         _ExtentY        =   8202
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
         Begin Threed.SSPanel pnl_Tit_NumCar 
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   60
            Width           =   2100
            _Version        =   65536
            _ExtentX        =   3704
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. de Carta Cofide"
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
         Begin Threed.SSPanel pnl_Tit_FecDes 
            Height          =   285
            Left            =   2190
            TabIndex        =   16
            Top             =   60
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Desembolso"
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
            Height          =   4215
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   30
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   3990
            TabIndex        =   17
            Top             =   60
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operaciones"
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
         Begin Threed.SSPanel pnl_Tit_Selecc 
            Height          =   285
            Left            =   7890
            TabIndex        =   18
            Top             =   60
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
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
               Left            =   1050
               TabIndex        =   19
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_DesCof 
            Height          =   285
            Left            =   5790
            TabIndex        =   23
            Top             =   60
            Width           =   2100
            _Version        =   65536
            _ExtentX        =   3704
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Desembolso Cliente"
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   840
         Left            =   60
         TabIndex        =   20
         Top             =   1440
         Width           =   10005
         _Version        =   65536
         _ExtentX        =   17648
         _ExtentY        =   1482
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
            Top             =   105
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1200
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
            TabIndex        =   22
            Top             =   130
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   495
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbBbp_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer

Private Sub cmd_Buscar_Click()
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
   
   If (ipp_PerAno.Text = 2015 And cmb_PerMes.ListIndex >= 9) Or ipp_PerAno.Text > 2015 Then
      Screen.MousePointer = 11
      Call fs_BuscaDatos
      Call fs_Activa(False)
      If (grd_Listad.Rows = 0) Then
         Call cmd_Limpia_Click
      End If
      Screen.MousePointer = 0
   Else
      MsgBox "Los desembolsos se realizan a partir de Octubre 2015", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
End Sub

Private Sub cmd_Limpia_Click()
    cmb_PerMes.ListIndex = -1
    ipp_PerAno.Text = Year(date)
    grd_Listad.Rows = 0
    Call fs_Activa(True)
    Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
    cmb_PerMes.Enabled = p_Activa
    ipp_PerAno.Enabled = p_Activa
    cmd_Buscar.Enabled = p_Activa
    cmd_ExpExc.Enabled = Not p_Activa
    cmd_Detail.Enabled = Not p_Activa
    cmd_Proces.Enabled = Not p_Activa
    grd_Listad.Enabled = Not p_Activa
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 4) = "X" Then
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
   Call fs_GeneraAsiento_DesBBP
   Call fs_BuscaDatos
   Screen.MousePointer = 0
   
   If (grd_Listad.Rows = 0) Then
       Call cmd_Limpia_Click
   End If
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

Private Sub cmd_Detail_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = CStr(grd_Listad.Text)
   
   grd_Listad.Col = 1
   moddat_g_str_FecDes = Mid(Trim(CStr(grd_Listad.Text)), 7, 4) & Mid(Trim(CStr(grd_Listad.Text)), 4, 2) & Mid(Trim(CStr(grd_Listad.Text)), 1, 2)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Pro_CtbBbp_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_BuscaPeriodo
   Call gs_CentraForm(Me)
   Call cmd_Limpia_Click
   
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub fs_Limpia()
   grd_Listad.Cols = 5
   grd_Listad.ColWidth(0) = 2100
   grd_Listad.ColWidth(1) = 1800
   grd_Listad.ColWidth(2) = 1800
   grd_Listad.ColWidth(3) = 2100
   grd_Listad.ColWidth(4) = 1300
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_BuscaDatos()
Dim r_dbl_ComDes     As Double
Dim r_dbl_SumDes     As Double
Dim r_dbl_SumPre     As Double
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Dim swtFechaInicio As Date
   Dim swtFechaFin As Date
   Dim swtMes As String
   Dim swtNextFecha As Date
   Dim swtObect As Object
   
   swtMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   swtFechaInicio = Trim("01-" & swtMes & "-" & CStr(ipp_PerAno.Value))
   swtNextFecha = DateAdd("m", 1, swtFechaInicio)
   swtFechaFin = DateAdd("d", -1, swtNextFecha)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT NUM_CARTA, FEC_DESEMBOLSO, COUNT(*) AS NRO_OPERACIONES, SUM(DES_CLIENTE) AS DES_CLIENTE "
   g_str_Parame = g_str_Parame & "  FROM (SELECT EVACOF_NUMCAR      AS NUM_CARTA, "
   g_str_Parame = g_str_Parame & "               EVACOF_NUMSOL      AS NUM_SOLICITUD, "
   g_str_Parame = g_str_Parame & "               EVACOF_FECDES      AS FEC_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "               SOLMAE_FMVBBP      AS DES_CLIENTE "
   g_str_Parame = g_str_Parame & "          FROM TRA_EVACOF "
   g_str_Parame = g_str_Parame & "         INNER JOIN CRE_HIPMAE ON HIPMAE_NUMSOL = EVACOF_NUMSOL AND HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "         INNER JOIN CRE_SOLMAE ON EVACOF_NUMSOL = SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "         WHERE ((EVACOF_FLGCNT = 1) OR (EVACOF_FLGCNT IS NULL)) "
   g_str_Parame = g_str_Parame & "           AND (EVACOF_FECDES >= " & Format(swtFechaInicio, "yyyymmdd") & ") "
   g_str_Parame = g_str_Parame & "           AND (EVACOF_FECDES <= " & Format(swtFechaFin, "yyyymmdd") & ") "
   g_str_Parame = g_str_Parame & "           AND SUBSTR(HIPMAE_NUMOPE,1,3) IN ('021','022','023') "
   g_str_Parame = g_str_Parame & "         ORDER BY EVACOF_FECDES) "
   g_str_Parame = g_str_Parame & " GROUP BY NUM_CARTA, FEC_DESEMBOLSO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se encontraron nuevos números de carta BBP.", vbInformation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_dbl_ComDes = 0
   r_dbl_SumDes = 0
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      If Not IsNull(g_rst_Princi!NUM_CARTA) Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!NUM_CARTA)
         
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FEC_DESEMBOLSO))
         
         grd_Listad.Col = 2
         grd_Listad.Text = Trim(g_rst_Princi!NRO_OPERACIONES)
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(g_rst_Princi!DES_CLIENTE, "###,###,##0.00")
       
         grd_Listad.Col = 4
         grd_Listad.Text = ""
      End If
      g_rst_Princi.MoveNext
   Loop
   
   If grd_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_Listad)
   End If
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_BuscaPeriodo()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PERMES_CODANO, PERMES_CODMES "
   g_str_Parame = g_str_Parame & "  FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND PERMES_TIPPER = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      cmd_Proces.Enabled = False
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se pudo determinar el período actual.", vbInformation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      cmd_Proces.Enabled = False
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   l_int_PerMes = g_rst_Princi!PERMES_CODMES
   l_int_PerAno = g_rst_Princi!PERMES_CODANO
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GeneraAsiento_DesBBP()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_dbl_TipSbs        As Double
Dim r_str_NroOpe        As String
Dim r_str_NumCar        As String
Dim r_str_FecDes        As String
Dim r_str_NumOpe        As String
Dim r_str_CtaDeb        As String
Dim r_str_CtaHab        As String
Dim r_dbl_DesBBP        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_DesBbpAcu     As Double

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1030"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      If grd_Listad.TextMatrix(r_int_Contad, 4) = "X" Then
         '*************************************************
         'GENERACION DE ASIENTOS CONTABLES POR CARTA BBP
         '*************************************************
         'Inicializa variables
         r_int_NumAsi = 0
         r_str_NumCar = grd_Listad.TextMatrix(r_int_Contad, 0)
         r_str_FecDes = grd_Listad.TextMatrix(r_int_Contad, 1)
         r_str_NroOpe = grd_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, 2, Format(r_str_FecDes, "yyyymmdd"), 1)
         
         'Consulta operaciones asociadas al numero de carta BBP
         g_str_Parame = " "
         g_str_Parame = g_str_Parame & "SELECT TRIM(B.SOLMAE_TITTDO)||'-'||TRIM(B.SOLMAE_TITNDO) AS DOCUMENTO_CLIENTE, "
         g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
         g_str_Parame = g_str_Parame & "       B.SOLMAE_FMVBBP           AS MONTO_DESEMBOLSO, "
         g_str_Parame = g_str_Parame & "       NVL(D.HIPMAE_NUMOPE, '-') AS NRO_OPERACION "
         g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF A "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.EVACOF_NUMSOL "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.SOLMAE_TITTDO AND C.DATGEN_NUMDOC = B.SOLMAE_TITNDO "
         g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPMAE D ON D.HIPMAE_NUMSOL = B.SOLMAE_NUMERO "
         g_str_Parame = g_str_Parame & " WHERE TRIM(A.EVACOF_NUMCAR) = '" & Trim(r_str_NumCar) & "' "
         g_str_Parame = g_str_Parame & "   AND EVACOF_FECDES = " & Format(r_str_FecDes, "yyyymmdd") & " "
         g_str_Parame = g_str_Parame & " ORDER BY NOMBRE_CLIENTE "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            MsgBox "No se encontraron operaciones de la carta BBP " & Trim(r_str_NumCar) & ".", vbExclamation, modgen_g_str_NomPlt
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
            Exit Sub
         End If
         
         'valida operaciones
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            If IsNull(g_rst_Princi!MONTO_DESEMBOLSO) Or g_rst_Princi!NRO_OPERACION = "-" Then
               MsgBox "Faltan desembolsarse operaciones, no se puede generar asiento contable.", vbCritical, modgen_g_str_NomPlt
               g_rst_Princi.Close
               Set g_rst_Princi = Nothing
               Exit Sub
            End If
            
            g_rst_Princi.MoveNext
         Loop
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
         r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
         
         'Insertar en CABECERA
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Mid(Trim(r_str_NumCar) & " / " & "DESEMBOLSOS DE BBP PARA FONDO MI VIVIENDA", 1, 60), r_str_FecDes, "1")
         
         'Inicializa
         r_int_NumIte = 1
         r_dbl_DesBbpAcu = 0
         
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            'Inicializa
            r_str_NumOpe = Trim(g_rst_Princi!NRO_OPERACION)
            r_dbl_DesBBP = Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")
            r_dbl_MtoDol = Format(0, "###,###,##0.00")
            r_str_CtaHab = "291807010113"
            
            'Insertar en DETALLE
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FecDes), Mid(CStr(r_str_NumOpe) & " - " & Trim(g_rst_Princi!NOMBRE_CLIENTE) & " - " & Trim(r_str_NumCar), 1, 60), "H", r_dbl_DesBBP, r_dbl_MtoDol, 1, CDate(r_str_FecDes))
            
            r_dbl_DesBbpAcu = r_dbl_DesBbpAcu + Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")
            r_int_NumIte = r_int_NumIte + 1
            g_rst_Princi.MoveNext
         Loop
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         'Insertar en DETALLE (suma desembolsos BBP)
         r_str_CtaDeb = "111301060102"
         r_dbl_MtoDol = Format(0, "###,###,##0.00")
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FecDes), Mid("AB/ Fondo Mivivienda/BBP - " & Trim(r_str_NumCar), 1, 60), "D", r_dbl_DesBbpAcu, r_dbl_MtoDol, 1, CDate(r_str_FecDes))
         
         'Actualiza flag de contabilizacion para las operacion con el numero de carta
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE TRA_EVACOF "
         g_str_Parame = g_str_Parame & "   SET EVACOF_FLGCNT = 2 "
         g_str_Parame = g_str_Parame & " WHERE EVACOF_NUMCAR  = '" & Trim(r_str_NumCar) & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros seleccionados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_GeneraAsiento_DesBBP_OLD()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_dbl_TipSbs        As Double
Dim r_str_NroOpe        As String
Dim r_str_NumCar        As String
Dim r_str_FecDes        As String
Dim r_str_NumOpe        As String
Dim r_str_CtaDeb        As String
Dim r_str_CtaHab        As String
Dim r_dbl_DesBBP        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_DesBbpAcu     As Double

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1030"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      If grd_Listad.TextMatrix(r_int_Contad, 4) = "X" Then
         '*************************************************
         'GENERACION DE ASIENTOS CONTABLES POR CARTA BBP
         '*************************************************
         'Inicializa variables
         r_int_NumAsi = 0
         r_str_NumCar = grd_Listad.TextMatrix(r_int_Contad, 0)
         r_str_FecDes = grd_Listad.TextMatrix(r_int_Contad, 1)
         r_str_NroOpe = grd_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, 2, Format(r_str_FecDes, "yyyymmdd"), 1)
         
         'Consulta operaciones asociadas al numero de carta BBP
         g_str_Parame = " "
         g_str_Parame = g_str_Parame & "SELECT TRIM(B.SOLMAE_TITTDO)||'-'||TRIM(B.SOLMAE_TITNDO) AS DOCUMENTO_CLIENTE, "
         g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
         g_str_Parame = g_str_Parame & "       B.SOLMAE_FMVBBP           AS MONTO_DESEMBOLSO, "
         g_str_Parame = g_str_Parame & "       NVL(D.HIPMAE_NUMOPE, '-') AS NRO_OPERACION "
         g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF A "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.EVACOF_NUMSOL "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.SOLMAE_TITTDO AND C.DATGEN_NUMDOC = B.SOLMAE_TITNDO "
         g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPMAE D ON D.HIPMAE_NUMSOL = B.SOLMAE_NUMERO "
         g_str_Parame = g_str_Parame & " WHERE TRIM(A.EVACOF_NUMCAR) = '" & Trim(r_str_NumCar) & "' "
         g_str_Parame = g_str_Parame & "   AND EVACOF_FECDES = " & Format(r_str_FecDes, "yyyymmdd") & " "
         g_str_Parame = g_str_Parame & " ORDER BY NOMBRE_CLIENTE "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            MsgBox "No se encontraron operaciones de la carta BBP " & Trim(r_str_NumCar) & ".", vbExclamation, modgen_g_str_NomPlt
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
            Exit Sub
         End If
         
         'valida operaciones
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            If IsNull(g_rst_Princi!MONTO_DESEMBOLSO) Or g_rst_Princi!NRO_OPERACION = "-" Then
               MsgBox "Faltan desembolsarse operaciones, no se puede generar asiento contable.", vbCritical, modgen_g_str_NomPlt
               g_rst_Princi.Close
               Set g_rst_Princi = Nothing
               Exit Sub
            End If
            
            g_rst_Princi.MoveNext
         Loop
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
         r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
         
         'Insertar en CABECERA
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_NumCar) & " / " & "DESEMBOLSOS DE BBP PARA FONDO MI VIVIENDA", r_str_FecDes, "1")
         
         'Inicializa
         r_int_NumIte = 1
         r_dbl_DesBbpAcu = 0
         
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            'Inicializa
            r_str_NumOpe = Trim(g_rst_Princi!NRO_OPERACION)
            r_dbl_DesBBP = Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")
            r_dbl_MtoDol = Format(0, "###,###,##0.00")
            r_str_CtaHab = "291807010113"
            
            'Insertar en DETALLE
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FecDes), CStr(r_str_NumOpe) & " - " & Trim(g_rst_Princi!NOMBRE_CLIENTE) & " - " & Trim(r_str_NumCar), "H", r_dbl_DesBBP, r_dbl_MtoDol, 1, CDate(r_str_FecDes))
            
            r_dbl_DesBbpAcu = r_dbl_DesBbpAcu + Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")
            r_int_NumIte = r_int_NumIte + 1
            g_rst_Princi.MoveNext
         Loop
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         'Insertar en DETALLE (suma desembolsos BBP)
         r_str_CtaDeb = "111301060102"
         r_dbl_MtoDol = Format(0, "###,###,##0.00")
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FecDes), "AB/ Fondo Mivivienda/BBP - " & Trim(r_str_NumCar), "D", r_dbl_DesBbpAcu, r_dbl_MtoDol, 1, CDate(r_str_FecDes))
         
         'Actualiza flag de contabilizacion para las operacion con el numero de carta
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE TRA_EVACOF "
         g_str_Parame = g_str_Parame & "   SET EVACOF_FLGCNT = 2 "
         g_str_Parame = g_str_Parame & " WHERE EVACOF_NUMCAR  = '" & Trim(r_str_NumCar) & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros seleccionados." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE DESEMBOLSOS BBP"
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(5, 2) = "NRO. CARTA BBP"
      .Cells(5, 3) = "FECHA DESEMBOLSO"
      .Cells(5, 4) = "NRO. OPERACIONES"
      .Cells(5, 5) = "DESEMBOLSO CLIENTE"
      
      .Range(.Cells(5, 2), .Cells(5, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 5)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 5)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 20
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 20
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 20
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 20
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 5)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 5)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0))
         .Cells(r_int_NumFil + 3, 3) = grd_Listad.TextMatrix(r_int_NumFil - 3, 1)
         .Cells(r_int_NumFil + 3, 4) = grd_Listad.TextMatrix(r_int_NumFil - 3, 2)
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 4
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub chkSeleccionar_Click()
 Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 4) = ""
         Next r_Fila
      End If
   
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 4) = "X"
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub
