VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbPrv_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12885
   Icon            =   "GesCtb_frm_918.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel9 
      Height          =   8115
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12885
      _Version        =   65536
      _ExtentX        =   22728
      _ExtentY        =   14314
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
         Height          =   735
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Contabilización de Provisiones"
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
            Picture         =   "GesCtb_frm_918.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   12
         Top             =   840
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
         Begin VB.CommandButton cmd_Detalle 
            Height          =   585
            Left            =   1890
            Picture         =   "GesCtb_frm_918.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Ver Detalle"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Enabled         =   0   'False
            Height          =   585
            Left            =   2520
            Picture         =   "GesCtb_frm_918.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Generar asientos automaticos"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1275
            Picture         =   "GesCtb_frm_918.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12100
            Picture         =   "GesCtb_frm_918.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_918.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_918.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   810
         Left            =   60
         TabIndex        =   13
         Top             =   1560
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
            ItemData        =   "GesCtb_frm_918.frx":17C2
            Left            =   1080
            List            =   "GesCtb_frm_918.frx":17C4
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1080
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
         Begin VB.Label Label1 
            Caption         =   "Año:"
            Height          =   315
            Left            =   135
            TabIndex        =   15
            Top             =   420
            Width           =   1365
         End
         Begin VB.Label Label10 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   135
            TabIndex        =   14
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5655
         Left            =   60
         TabIndex        =   16
         Top             =   2400
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   90
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2760
            TabIndex        =   18
            Top             =   90
            Width           =   5160
            _Version        =   65536
            _ExtentX        =   9102
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Glosa"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   10080
            TabIndex        =   19
            Top             =   90
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber (S/.)"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   7920
            TabIndex        =   20
            Top             =   90
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe (S/.)"
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
            Height          =   4815
            Left            =   90
            TabIndex        =   21
            Top             =   360
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   8493
            _Version        =   393216
            Rows            =   15
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Haber 
            Height          =   315
            Left            =   10140
            TabIndex        =   22
            Top             =   5220
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Debe 
            Height          =   315
            Left            =   7920
            TabIndex        =   23
            Top             =   5220
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPrv_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer

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
   Me.pnl_Debe.Caption = 0
   Me.pnl_Haber.Caption = 0
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
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = 2 "
   
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
      Exit Sub
   End If
   
   'Verifica periodos pasados
   If l_int_PerAno <= 2015 And l_int_PerMes <= 4 Then
      MsgBox "Período seleccionado ya esta cerrado.", vbExclamation, modgen_g_str_NomPlt
      cmd_Proces.Enabled = False
      Exit Sub
   End If
End Sub

Private Sub cmd_Limpia_Click()
   grd_Listad.Rows = 0
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PerMes)
   Me.pnl_Debe.Caption = 0
   Me.pnl_Haber.Caption = 0
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
   
   'CUENTA
   grd_Listad.Col = 0
   moddat_g_str_DesMod = Trim(grd_Listad.Text)
   
   'GLOSA
   grd_Listad.Col = 1
   moddat_g_str_Descri = Trim(grd_Listad.Text)
   
   'MONEDA
   grd_Listad.Col = 4
   moddat_g_str_Moneda = Trim(grd_Listad.Text)
   
   'MES ANTERIOR
   grd_Listad.Col = 5
   moddat_g_dbl_MtoPre = Trim(grd_Listad.Text)
   
   'MES ACTUAL
   grd_Listad.Col = 6
   moddat_g_dbl_SalCap = IIf(Trim(grd_Listad.Text) = "", 0, Trim(grd_Listad.Text))
   
   'AJUSTE
   grd_Listad.Col = 7
   moddat_g_dbl_IngDec = IIf(Trim(grd_Listad.Text) = "", 0, Trim(grd_Listad.Text))
   
   'TIPO CAMBIO
   grd_Listad.Col = 8
   moddat_g_dbl_TasInt = IIf(Trim(grd_Listad.Text) = "", 0, Trim(grd_Listad.Text))
   
   'Mes
   moddat_g_str_CodMes = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   'AÑO
   moddat_g_str_CodAno = CStr(ipp_PerAno.Value)
   
   Call gs_RefrescaGrid(grd_Listad)
   frm_Pro_CtbPrv_02.Show 1
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida DEBE = HABER
   If pnl_Debe.Caption <> pnl_Haber.Caption Then
      MsgBox "El Monto Debe no es igual al Monto Haber. No se puede generar asientos automáticos.", vbInformation, modgen_g_str_NomPlt
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
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   
   grd_Listad.ColWidth(0) = 2645       'CUENTA
   grd_Listad.ColWidth(1) = 5150       'GLOSA
   grd_Listad.ColWidth(2) = 2195       'DEBE
   grd_Listad.ColWidth(3) = 2180       'HABER
   grd_Listad.ColWidth(4) = 0          'MONEDA
   grd_Listad.ColWidth(5) = 0          'MONTO MES ANTERIOR
   grd_Listad.ColWidth(6) = 0          'MONTO MES ACTUAL
   grd_Listad.ColWidth(7) = 0          'AJUSTE
   grd_Listad.ColWidth(8) = 0          'TIPO CAMBIO
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
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

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   cmd_Proces.Enabled = Not p_Activa
   cmd_Detalle.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
Dim r_int_MesCie        As String
Dim r_int_AnoCie        As String
Dim r_int_Cont          As Integer
Dim r_str_NomProv       As String
Dim r_str_CamTbl        As String
Dim r_str_Cond          As String
Dim l_str_CodEmp        As String

   Call gs_LimpiaGrid(grd_Listad)
     
   '*** INICIALIZA VARIABLES
   r_int_MesCie = IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) 'CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_AnoCie = CInt(ipp_PerAno.Text)
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & " WHERE EMPGRP_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_str_CodEmp = g_rst_Princi!EMPGRP_CODIGO
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_PROVCRED("
   g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
   g_str_Parame = g_str_Parame & CInt(r_int_MesCie) & ", "
   g_str_Parame = g_str_Parame & CInt(r_int_AnoCie) & ", "
   g_str_Parame = g_str_Parame & "'" & "REPORTE DE PROVISIONES" & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "')"
        
   'Ejecuta consulta
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_RPT_PROVCRED.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'CUENTA
      grd_Listad.Col = 0
      If g_rst_Princi!Cuenta = "541401010102" Then
         grd_Listad.Text = "541401010101"
      ElseIf g_rst_Princi!Cuenta = "542401010102" Then
         grd_Listad.Text = "542401010101"
      Else
         grd_Listad.Text = g_rst_Princi!Cuenta
      End If
      
      'GLOSA
      grd_Listad.Col = 1
      grd_Listad.Text = g_rst_Princi!DESCRIPCION
      
      'DEBE
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!DEBE, "###,###,###,##0.00")
      
      'HABER
      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!HABER, "###,###,###,##0.00")
      
      'MONEDA
      grd_Listad.Col = 4
      grd_Listad.Text = g_rst_Princi!Moneda
      
      'MONTO ANTERIOR
      grd_Listad.Col = 5
      grd_Listad.Text = Format(g_rst_Princi!MES_ANTERIOR, "###,###,###,##0.00")
      
      'MONTO ACTUAL
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!MES_ACTUAL, "###,###,###,##0.00")
      
      'AJUSTE
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!AJUSTE, "###,###,###,##0.00")
      
      'TIPO CAMBIO
      grd_Listad.Col = 8
      grd_Listad.Text = g_rst_Princi!TIPOCAMBIO
      
      g_rst_Princi.MoveNext
   Loop
   
   Call Sumar_Columnas
   
   grd_Listad.Redraw = True
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
           
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub Sumar_Columnas()
Dim r_dbl_TotDebe       As Double
Dim r_dbl_TotHaber      As Double
Dim r_int_Fila          As Integer

    For r_int_Fila = 0 To grd_Listad.Rows - 1
        r_dbl_TotDebe = r_dbl_TotDebe + IIf(grd_Listad.TextMatrix(r_int_Fila, 2) = "", 0, grd_Listad.TextMatrix(r_int_Fila, 2))
        r_dbl_TotHaber = r_dbl_TotHaber + IIf(grd_Listad.TextMatrix(r_int_Fila, 3) = "", 0, grd_Listad.TextMatrix(r_int_Fila, 3))
    Next r_int_Fila
    
    pnl_Debe.Caption = Format(r_dbl_TotDebe, "###,###,##0.00")
    pnl_Haber.Caption = Format(r_dbl_TotHaber, "###,###,##0.00")
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
Dim r_str_DebHab        As String
Dim r_str_Glosa         As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_Importe       As Double
Dim r_dbl_TipCam        As Double
Dim r_int_NumTipMon     As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   r_str_Origen = "LM"
   r_str_TipNot = "E"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   r_int_NumAsi = 0 'Inicializa variables
   r_int_NumIte = 0
      
   'Obteniendo Nro. de Asiento (único)
   If grd_Listad.Rows > 0 Then
      r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
      r_str_AsiGen = CStr(r_int_NumAsi)
      r_str_FecPbpC = Format(ff_Ultimo_Dia_Mes(IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))), CInt(ipp_PerAno.Text)), "00") & "/" & IIf(Len(Trim(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) = 1, "0" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))) & "/" & CInt(ipp_PerAno.Text)
      r_str_FecPbpL = moddat_g_str_FecSis
              
      r_dbl_TipCam = grd_Listad.TextMatrix(0, 8)
      r_str_Glosa = "ASIENTO PROVISIONES " & Right(l_int_PerAno, 4) & " - " & Right("00" & l_int_PerMes, 2)
   
      'Insertar en CABECERA
      Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPbpC, "1")
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      '*************************************************
      'GENERACION DE ASIENTOS CONTABLES DE PROVISIONES
      '*************************************************
      If grd_Listad.TextMatrix(r_int_Contad, 2) > 0 Or grd_Listad.TextMatrix(r_int_Contad, 3) > 0 Then
         If grd_Listad.TextMatrix(r_int_Contad, 2) > 0 Then r_dbl_Importe = grd_Listad.TextMatrix(r_int_Contad, 2): r_str_DebHab = "D"
         If grd_Listad.TextMatrix(r_int_Contad, 3) > 0 Then r_dbl_Importe = grd_Listad.TextMatrix(r_int_Contad, 3): r_str_DebHab = "H"
         
         r_str_CtaCtb = grd_Listad.TextMatrix(r_int_Contad, 0)
         r_str_Glosa = grd_Listad.TextMatrix(r_int_Contad, 1)
         r_int_NumTipMon = grd_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_TipCam = grd_Listad.TextMatrix(r_int_Contad, 8)
         
         If (r_dbl_Importe > 0) Then
            r_int_NumIte = r_int_NumIte + 1
             
            If r_int_NumTipMon = 2 Then
               r_dbl_MtoSol = Format(r_dbl_Importe, "###,###,##0.00")
               r_dbl_MtoDol = Format(CDbl(r_dbl_MtoSol / r_dbl_TipCam), "###,###,##0.00")
            Else
               r_dbl_MtoSol = Format(r_dbl_Importe, "###,###,##0.00")
               r_dbl_MtoDol = Format(0, "###,###,##0.00")
            End If
             
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecPbpC), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPbpL))
            r_dbl_Importe = 0
         End If
     End If
   Next r_int_Contad
   
   Call modprc_fs_Actualiza_Proceso(l_int_PerAno, l_int_PerMes, 2)
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
      .Cells(2, 2) = "CONTABILIZACIÓN DE ASIENTOS DE PROVISIONES"
      .Range(.Cells(2, 2), .Cells(2, 5)).Merge
      .Range(.Cells(2, 2), .Cells(2, 5)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 2) = "CUENTA"
      .Cells(4, 3) = "GLOSA"
      .Cells(4, 4) = "DEBE (S/.)"
      .Cells(4, 5) = "HABER (S/.)"
            
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 35
      .Columns("B").HorizontalAlignment = xlHAlignCenter 'xlHAlignCenter
      .Columns("C").ColumnWidth = 46 '26
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 24
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 24
      .Columns("E").HorizontalAlignment = xlHAlignRight
            
      .Range(.Cells(4, 2), .Cells(4, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 5)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(3, 5)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(3, 5)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
            .Cells(r_int_NumFil + 2, 2) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0))   'CUENTA
            .Cells(r_int_NumFil + 2, 3) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 1))         'GLOSA
            .Cells(r_int_NumFil + 2, 4) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 2))         'DEBE
            .Cells(r_int_NumFil + 2, 5) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 3))         'HABER
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If cmb_PerMes.ListIndex > -1 Then
      If KeyAscii = 13 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub
