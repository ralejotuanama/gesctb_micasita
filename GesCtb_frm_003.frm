VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   7845
   ClientTop       =   5910
   ClientWidth     =   7200
   Icon            =   "GesCtb_frm_003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   6376
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   2115
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   3731
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
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   5895
         End
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   1080
            Width           =   1995
         End
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todos las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1140
            TabIndex        =   4
            Top             =   1380
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1140
            TabIndex        =   5
            Top             =   1740
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   1770
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   915
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Height          =   570
            Left            =   630
            TabIndex        =   11
            Top             =   45
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   1005
            _StockProps     =   15
            Caption         =   "Reporte de Créditos Hipotecarios Desembolsados"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6600
            Top             =   30
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_003.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Left            =   6480
            Picture         =   "GesCtb_frm_003.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_003.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_003.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_Empres_Click()
   If chk_Empres.Value = 1 Then
      cmb_Empres.ListIndex = -1
      cmb_Empres.Enabled = False
      If cmb_TipPro.Enabled Then
         Call gs_SetFocus(cmb_TipPro)
      Else
         Call gs_SetFocus(ipp_FecIni)
      End If
   ElseIf chk_Empres.Value = 0 Then
      cmb_Empres.Enabled = True
      Call gs_SetFocus(cmb_Empres)
   End If
End Sub

Private Sub chk_TipPro_Click()
   If chk_TipPro.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_TipPro.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_TipPro.Enabled Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
Dim r_str_TIPMON As String
      
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Proceso
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_CREDES WHERE "
   g_str_Parame = g_str_Parame & "CREDES_NOMRPT = 'CTB_RPTSOL_01.RPT' AND "
   g_str_Parame = g_str_Parame & "CREDES_TERCRE ='" & modgen_g_str_NombPC & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE, CLI_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCLI = DATGEN_TIPDOC "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCLI = DATGEN_NUMDOC "
   If chk_Empres.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND HIPMAE_PROCRE = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' "
   End If
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 11
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_TIPMON = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
        
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_CREDES("
         g_str_Parame = g_str_Parame & "CREDES_NOMRPT, "
         g_str_Parame = g_str_Parame & "CREDES_FECCRE, "
         g_str_Parame = g_str_Parame & "CREDES_HORCRE, "
         g_str_Parame = g_str_Parame & "CREDES_TERCRE, "
         g_str_Parame = g_str_Parame & "CREDES_NUMOPE, "
         g_str_Parame = g_str_Parame & "CREDES_FECINI, "
         g_str_Parame = g_str_Parame & "CREDES_FECFIN, "
         g_str_Parame = g_str_Parame & "CREDES_CODPRD, "
         g_str_Parame = g_str_Parame & "CREDES_TIPMON, "
         g_str_Parame = g_str_Parame & "CREDES_TIPDOC, "
         g_str_Parame = g_str_Parame & "CREDES_NUMDOC, "
         g_str_Parame = g_str_Parame & "CREDES_APEPAT, "
         g_str_Parame = g_str_Parame & "CREDES_APEMAT, "
         g_str_Parame = g_str_Parame & "CREDES_NOMBRE, "
         g_str_Parame = g_str_Parame & "CREDES_MTOPRE, "
         g_str_Parame = g_str_Parame & "CREDES_TOTPRE, "
         g_str_Parame = g_str_Parame & "CREDES_INTCAP, "
         g_str_Parame = g_str_Parame & "CREDES_TASINT, "
         g_str_Parame = g_str_Parame & "CREDES_NUMCUO, "
         g_str_Parame = g_str_Parame & "CREDES_PERGRA, "
         g_str_Parame = g_str_Parame & "CREDES_FECDES, "
         g_str_Parame = g_str_Parame & "CREDES_FECACT, "
         g_str_Parame = g_str_Parame & "CREDES_IMPNCO, "
         g_str_Parame = g_str_Parame & "CREDES_IMPCON, "
         g_str_Parame = g_str_Parame & "CREDES_IMPDES, "
         g_str_Parame = g_str_Parame & "CREDES_COSEFE, "
         g_str_Parame = g_str_Parame & "CREDES_EMPRES, "
         g_str_Parame = g_str_Parame & "CREDES_MTOCVT, "
         g_str_Parame = g_str_Parame & "CREDES_APOPRO, "
         g_str_Parame = g_str_Parame & "CREDES_PORINI) "
                           
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'CTB_RPTSOL_01.RPT', "
         g_str_Parame = g_str_Parame & l_str_Fecha & ", "
         g_str_Parame = g_str_Parame & l_str_Hora & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_NUMOPE & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_CODPRD & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_TIPMON & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TDOCLI & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_NDOCLI & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_ApePat & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_ApeMat & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_Nombre & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_MTOPRE & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TOTPRE & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_INTCAP & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TASINT & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_NUMCUO & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_PERGRA & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_FECDES & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_FECACT & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPNCO & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPCON & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPDES & ","
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_COSEFE & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_ConsultaEmpGrp(g_rst_Princi!HIPMAE_PROCRE) & "',"
      
         If g_rst_Princi!HIPMAE_MONEDA = 2 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_CVTDOL & ", "
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_APODOL & ", "
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!HIPMAE_APODOL / g_rst_Princi!HIPMAE_CVTDOL * 100, "##0.00") & ") "
         Else
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_CVTSOL & ", "
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_APOSOL & ", "
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!HIPMAE_APOSOL / g_rst_Princi!HIPMAE_CVTSOL * 100, "##0.00") & ") "
         End If
                
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_PRODUC"
   crp_Imprim.DataFiles(1) = "RPT_CREDES"
   crp_Imprim.SelectionFormula = "{RPT_CREDES.CREDES_NOMRPT} = 'CTB_RPTSOL_01.RPT' " & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_CREDES.CREDES_TERCRE} = '" & modgen_g_str_NombPC & "'"
      
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_01.RPT"
      
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
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
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = -1
   chk_Empres.Value = 0
   cmb_TipPro.ListIndex = -1
   chk_TipPro.Value = 0
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConAux     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(B.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_NUMOPE AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(J.SUBPRD_DESCRI) AS SUB_PRODUCTO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_NUMSOL AS SOLICITUD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_TIPDOC)||'-'||TRIM(C.DATGEN_NUMDOC) AS DOCUM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_FECACT AS FEC_ACTIVACION, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_FECDES AS FEC_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_MONEDA AS TIPO_MONEDA, "
   g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS MONEDA_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_IMPDES AS MTO_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_MTOPRE AS MTO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_INTCAP AS INT_CAPITALIZADO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_TOTPRE AS TOT_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_IMPNCO AS TOT_PREST_TNC, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_IMPCON AS TOT_PREST_TC, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_TASINT AS TASA_INTERES, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_COSEFE AS COSTO_EFECTIVO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_NUMCUO AS NUM_CUOTAS, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_PERGRA AS PERIODO_GRACIA, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_MONEDA, 2, HIPMAE_CVTDOL, HIPMAE_CVTSOL) AS VAL_COMPRAVENTA, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_MONEDA, 2, HIPMAE_APODOL, HIPMAE_APOSOL) AS MTO_INICIAL, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_MONEDA,1, O.SOLMAE_APOPRO_SOL, O.SOLMAE_APOPRO_DOL) - O.SOLMAE_FMVBBP - O.SOLMAE_PBPMTO - O.SOLMAE_BMSMTO - O.SOLMAE_AFPMTO  AS APORTE_PROPIO, "
   g_str_Parame = g_str_Parame & "       O.SOLMAE_PBPMTO AS MTO_PBP, O.SOLMAE_FMVBBP AS MTO_BBP, O.SOLMAE_BMSMTO AS MTO_BMS, O.SOLMAE_AFPMTO AS MTO_AFP, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_MONEDA,1, O.SOLMAE_COMVTA_SOL, O.SOLMAE_COMVTA_DOL) - "
   g_str_Parame = g_str_Parame & "       (DECODE(A.HIPMAE_MONEDA,1, O.SOLMAE_APOPRO_SOL, O.SOLMAE_APOPRO_DOL) - O.SOLMAE_FMVBBP - O.SOLMAE_PBPMTO - O.SOLMAE_BMSMTO - O.SOLMAE_AFPMTO) - "
   g_str_Parame = g_str_Parame & "       O.SOLMAE_PBPMTO - O.SOLMAE_FMVBBP - O.SOLMAE_BMSMTO - O.SOLMAE_AFPMTO  AS MONTO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_CONHIP AS CONSEJERO, "
   g_str_Parame = g_str_Parame & "       TRIM(E.PARDES_DESCRI) AS ESTADO_CREDITO, "
   g_str_Parame = g_str_Parame & "       TRIM(F.PARDES_DESCRI) AS TIPO_GARANTIA, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN H.SOLINM_TABPRY IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "             CASE WHEN H.SOLINM_TABPRY = 2 THEN"
   g_str_Parame = g_str_Parame & "                  CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "                       CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN"
   g_str_Parame = g_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   g_str_Parame = g_str_Parame & "                       ELSE"
   g_str_Parame = g_str_Parame & "                            CASE WHEN LENGTH (H.SOLINM_PRYNOM) > 0 THEN TRIM(H.SOLINM_PRYNOM) END"
   g_str_Parame = g_str_Parame & "                        END"
   g_str_Parame = g_str_Parame & "                  ELSE"
   g_str_Parame = g_str_Parame & "                       CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN"
   g_str_Parame = g_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   g_str_Parame = g_str_Parame & "                       ELSE"
   g_str_Parame = g_str_Parame & "                            CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "                              TRIM(H.SOLINM_PRYNOM)"
   g_str_Parame = g_str_Parame & "                            ELSE ''"
   g_str_Parame = g_str_Parame & "                             END"
   g_str_Parame = g_str_Parame & "                        END"
   g_str_Parame = g_str_Parame & "                   END"
   g_str_Parame = g_str_Parame & "             ELSE"
   g_str_Parame = g_str_Parame & "                   CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "                        (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   g_str_Parame = g_str_Parame & "                   ELSE"
   g_str_Parame = g_str_Parame & "                        CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "                          TRIM(H.SOLINM_PRYNOM)"
   g_str_Parame = g_str_Parame & "                        ELSE"
   g_str_Parame = g_str_Parame & "                          ''"
   g_str_Parame = g_str_Parame & "                         END"
   g_str_Parame = g_str_Parame & "                    END"
   g_str_Parame = g_str_Parame & "              END"
   g_str_Parame = g_str_Parame & "       ELSE"
   g_str_Parame = g_str_Parame & "              CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   g_str_Parame = g_str_Parame & "                (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   g_str_Parame = g_str_Parame & "              ELSE"
   g_str_Parame = g_str_Parame & "                ''"
   g_str_Parame = g_str_Parame & "               END"
   g_str_Parame = g_str_Parame & "       END) AS NOMBRE_PROYECTO,"
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN H.SOLINM_TIPDOC_PRO = 7 THEN TRIM(K.DATGEN_RAZSOC) ELSE TRIM(H.SOLINM_RAZSOC_PRO) END, '-') AS NOM_PROMOTOR,"
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN H.SOLINM_TIPDOC_CON = 7 THEN TRIM(L.DATGEN_RAZSOC) ELSE TRIM(H.SOLINM_RAZSOC_CON) END, '-') AS NOM_CONSTRUCTOR,"
   g_str_Parame = g_str_Parame & "       A.HIPMAE_TASCOF INTERES_COFIDE, "
   g_str_Parame = g_str_Parame & "       TRIM(N.PARDES_DESCRI) AS ESTADO_CIVIL,"
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_NDOCYG,NULL,'',A.HIPMAE_TDOCYG||'-'||A.HIPMAE_NDOCYG) AS NRO_DOC_CYG,"
   g_str_Parame = g_str_Parame & "       TRIM(M.DATGEN_APEPAT)||' '||TRIM(M.DATGEN_APEMAT)||' '||TRIM(M.DATGEN_NOMBRE) AS NOMBRE_CYG, NVL(O.SOLMAE_MTOGCI,0) AS GASTOS_CIERRE "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE O ON O.SOLMAE_NUMERO = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = A.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES N ON N.PARDES_CODGRP= 205 AND N.PARDES_CODITE = C.DATGEN_ESTCIV "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPMAE_MONEDA "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 027 AND E.PARDES_CODITE = A.HIPMAE_SITUAC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 241 AND F.PARDES_CODITE = A.HIPMAE_TIPGAR "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SUBPRD J ON J.SUBPRD_CODPRD = A.HIPMAE_CODPRD AND J.SUBPRD_CODSUB = A.HIPMAE_CODSUB "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM H ON H.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN K ON K.DATGEN_EMPTDO = H.SOLINM_TIPDOC_PRO AND K.DATGEN_EMPNDO = H.SOLINM_NUMDOC_PRO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN L ON L.DATGEN_EMPTDO = H.SOLINM_TIPDOC_CON AND L.DATGEN_EMPNDO = H.SOLINM_NUMDOC_CON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CLI_DATGEN M ON M.DATGEN_TIPDOC = A.HIPMAE_TDOCYG AND M.DATGEN_NUMDOC = A.HIPMAE_NDOCYG "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) AND "
   
   If chk_Empres.Value = 0 Then
      g_str_Parame = g_str_Parame & "      A.HIPMAE_PROCRE = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "      A.HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   g_str_Parame = g_str_Parame & "      A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "      A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "SUB-PRODUCTO"
      .Cells(1, 4) = "OPERACION"
      .Cells(1, 5) = "DOC. IDENTIDAD TIT."
      .Cells(1, 6) = "NOMBRE CLIENTE TIT."
      .Cells(1, 7) = "ESTADO CIVIL"
      .Cells(1, 8) = "DOC. IDENTIDAD CYG."
      .Cells(1, 9) = "NOMBRE CLIENTE CYG"
      .Cells(1, 10) = "F. ACTIV."
      .Cells(1, 11) = "F. DESEMB."
      .Cells(1, 12) = "TIP. DE MONEDA"
      .Cells(1, 13) = "V. COMPRA-VENTA"
      .Cells(1, 14) = "APORTE PROPIO"
      .Cells(1, 15) = "MTO. PBP"
      .Cells(1, 16) = "MTO. BBP"
      .Cells(1, 17) = "MTO. BMS"
      .Cells(1, 18) = "MTO. AFP"
      .Cells(1, 19) = "MTO. DESEMB."
      .Cells(1, 20) = "MTO. PRESTAMO"
      
      .Cells(1, 21) = "GASTOS CIERRE"
      .Cells(1, 22) = "INT. CAPIT."
      .Cells(1, 23) = "TOTAL PREST."
      .Cells(1, 24) = "M. PREST. T.N.C."
      .Cells(1, 25) = "M. PREST. T.C."
      .Cells(1, 26) = "%"
      .Cells(1, 27) = "TASA INT."
      .Cells(1, 28) = "TASA PONDERADA"
      .Cells(1, 29) = "COSTO EFECTIVO"
      .Cells(1, 30) = "TASA PASIVA"
      
      .Cells(1, 31) = "CUOTAS"
      .Cells(1, 32) = "P.G."
      .Cells(1, 33) = "% INICIAL"
      .Cells(1, 34) = "CONSEJERO HIPOT."
      .Cells(1, 35) = "ESTADO CREDITO"
      .Cells(1, 36) = "COD.EXCEP.(CREDITOS)"
      .Cells(1, 37) = "FORMA DE DESEMBOLSO"
      .Cells(1, 38) = "# CHEQUE"
      .Cells(1, 39) = "BANCO EMISOR"
      .Cells(1, 40) = "TIPO DE GARANTIA"
      .Cells(1, 41) = "NOMBRE DEL PROYECTO"
      .Cells(1, 42) = "NOMBRE DEL PROMOTOR"
      .Cells(1, 43) = "NOMBRE DEL CONSTRUCTOR"
      
      .Range(.Cells(1, 1), .Cells(1, 43)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 43)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 50
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 70
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 19
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 45
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 20
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 45
      .Columns("J").ColumnWidth = 13
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 22
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 18
      .Columns("N").ColumnWidth = 15
      .Columns("O").ColumnWidth = 15
      .Columns("P").ColumnWidth = 15
      .Columns("Q").ColumnWidth = 15
      .Columns("R").ColumnWidth = 15
      .Columns("S").ColumnWidth = 0  '14
      .Columns("T").ColumnWidth = 0  '16
      .Columns("U").ColumnWidth = 14 'GASTOS CIERRE
      .Columns("V").ColumnWidth = 12
      .Columns("W").ColumnWidth = 13
      .Columns("X").ColumnWidth = 15
      .Columns("Y").ColumnWidth = 14
      .Columns("Z").ColumnWidth = 11
      .Columns("Z").NumberFormat = "0.00%"
      .Columns("AA").ColumnWidth = 10
      .Columns("AA").NumberFormat = "0.00%"
      .Columns("AB").ColumnWidth = 17
      .Columns("AB").NumberFormat = "0.00%"
      .Columns("AC").ColumnWidth = 16
      .Columns("AD").ColumnWidth = 13
      .Columns("AD").HorizontalAlignment = xlHAlignRight
      .Columns("AE").ColumnWidth = 9
      .Columns("AE").ColumnWidth = 10
      .Columns("AF").ColumnWidth = 14
      .Columns("AG").ColumnWidth = 13
      .Columns("AG").HorizontalAlignment = xlHAlignCenter
      .Columns("AH").ColumnWidth = 18
      .Columns("AH").HorizontalAlignment = xlHAlignCenter
      .Columns("AI").ColumnWidth = 22
      .Columns("AI").HorizontalAlignment = xlHAlignCenter
      .Columns("AJ").ColumnWidth = 40
      .Columns("AJ").HorizontalAlignment = xlHAlignCenter
      .Columns("AK").ColumnWidth = 30
      .Columns("AK").HorizontalAlignment = xlHAlignCenter
      .Columns("AL").ColumnWidth = 30
      .Columns("AL").HorizontalAlignment = xlHAlignCenter
      .Columns("AM").ColumnWidth = 33
      .Columns("AM").HorizontalAlignment = xlHAlignCenter
      .Columns("AN").ColumnWidth = 50
      .Columns("AN").HorizontalAlignment = xlHAlignCenter
      .Columns("AO").ColumnWidth = 70
      .Columns("AO").HorizontalAlignment = xlHAlignCenter
      .Columns("AP").ColumnWidth = 70
      .Columns("AP").HorizontalAlignment = xlHAlignCenter
      .Columns("AQ").ColumnWidth = 70
      .Columns("AQ").HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SUB_PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!DOCUM_CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!Nom_Cliente)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(CStr(g_rst_Princi!ESTADO_CIVIL & ""))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(CStr(g_rst_Princi!NRO_DOC_CYG & ""))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(CStr(g_rst_Princi!NOMBRE_CYG & ""))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_ACTIVACION)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_DESEMBOLSO)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!MONEDA_PRESTAMO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!VAL_COMPRAVENTA, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!APORTE_PROPIO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!MTO_PBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!MTO_BBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!MTO_BMS, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!MTO_AFP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!MTO_DESEMBOLSO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!MONTO_PRESTAMO, "###,###,##0.00") 'MTO_PRESTAMO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!GASTOS_CIERRE, "###,###,##0.00") 'GASTOS CIERRE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!INT_CAPITALIZADO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!TOT_PRESTAMO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!TOT_PREST_TNC, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!TOT_PREST_TC, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = 0
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Format(g_rst_Princi!TASA_INTERES, "#0.00") / 100
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = 0
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Format(g_rst_Princi!COSTO_EFECTIVO, "#0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Format(g_rst_Princi!INTERES_COFIDE, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = g_rst_Princi!NUM_CUOTAS
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = g_rst_Princi!PERIODO_GRACIA
      
      'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!APORTE_PROPIO / g_rst_Princi!VAL_COMPRAVENTA * 100, "##0.00") & "%"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format(g_rst_Princi!MTO_INICIAL / g_rst_Princi!VAL_COMPRAVENTA * 100, "##0.00") & "%"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Trim(g_rst_Princi!CONSEJERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Trim(g_rst_Princi!ESTADO_CREDITO)
      
      'OBTIENE NUMERO DE EXCEPCION SI LA HUBIERA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TRA_SEGEXC "
      g_str_Parame = g_str_Parame & " WHERE SEGEXC_NUMSOL = '" & g_rst_Princi!SOLICITUD & "' "
      g_str_Parame = g_str_Parame & "   AND SEGEXC_CODINS = 21"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = g_rst_GenAux!SEGEXC_MOTEXC
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = ""
      End If
      '-------------------------------------------------------------------
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
      g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & g_rst_Princi!OPERACION & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         If CInt(g_rst_GenAux!HIPDES_TIPDES) > 0 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = moddat_gf_Consulta_ParDes("226", g_rst_GenAux!HIPDES_TIPDES)
            
            If g_rst_GenAux!HIPDES_TIPDES = 1 Then
               If Len(Trim(g_rst_GenAux!HIPDES_CHECGO & "")) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = Trim(g_rst_GenAux!HIPDES_CHECGO & "")
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = moddat_gf_Consulta_ParDes("516", g_rst_GenAux!HIPDES_BANCGO & "")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 38) = "CHEQUE NO EMITIDO"
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 39) = ""
               End If
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = ""
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 37) = ""
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 40) = Trim(g_rst_Princi!TIPO_GARANTIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 41) = Trim(g_rst_Princi!NOMBRE_PROYECTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 42) = Trim(g_rst_Princi!NOM_PROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 43) = Trim(g_rst_Princi!NOM_CONSTRUCTOR)
      '-------------------------------------------------------------------
      'Trim(g_rst_Princi!NOMBRE_PROYECTO)
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19).Select '13
   r_obj_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 2 & "]C:R[-1]C)"

   For r_int_ConAux = 2 To r_int_ConVer - 1
       r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 26) = r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 19) / r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19)
       r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 28) = r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 26) * r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 27)
   Next r_int_ConAux

   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26).Select
   r_obj_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 2 & "]C:R[-1]C)"
   
   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28).Select
   r_obj_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 2 & "]C:R[-1]C)"
   
   r_obj_Excel.ActiveSheet.Cells(1, 1).Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Imprim)
    End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_FecFin)
    End If
End Sub
