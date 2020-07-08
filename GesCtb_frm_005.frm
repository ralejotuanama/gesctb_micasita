VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   4425
   ClientTop       =   2505
   ClientWidth     =   5010
   Icon            =   "GesCtb_frm_005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3075
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   5424
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
         TabIndex        =   8
         Top             =   60
         Width           =   4965
         _Version        =   65536
         _ExtentX        =   8758
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
            TabIndex        =   9
            Top             =   30
            Width           =   4275
            _Version        =   65536
            _ExtentX        =   7541
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de Hipotecas y Bloqueos Registrales"
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
            TabIndex        =   10
            Top             =   315
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Por Producto"
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
            Picture         =   "GesCtb_frm_005.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   4965
         _Version        =   65536
         _ExtentX        =   8758
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_005.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_005.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4350
            Picture         =   "GesCtb_frm_005.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1230
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1515
         Left            =   30
         TabIndex        =   12
         Top             =   1470
         Width           =   4965
         _Version        =   65536
         _ExtentX        =   8758
         _ExtentY        =   2672
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
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3795
         End
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1080
            TabIndex        =   1
            Top             =   390
            Width           =   1995
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   750
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
            Left            =   1080
            TabIndex        =   3
            Top             =   1110
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
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   1140
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String
Dim r_str_TipGar        As String
Dim r_str_MonGar        As String
Dim r_str_TipDoc        As String
Dim r_str_SedReg        As String

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

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
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
Dim r_str_TipMon As String
      
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
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_HIPGAR  "
   g_str_Parame = g_str_Parame & " WHERE HIPGAR_NOMRPT ='CTB_RPTSOL_03.RPT' "
   g_str_Parame = g_str_Parame & "   AND HIPGAR_TERCRE ='" & modgen_g_str_NombPC & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_HIPMAE B, CRE_HIPGAR A WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = HIPGAR_NUMOPE AND "
   
   'Si no escogio todos los Consejeros Hipotecarios
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
         r_str_TipGar = moddat_gf_Consulta_ParDes("030", CStr(g_rst_Princi!HIPGAR_BIEGAR))
         r_str_TipMon = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPGAR_TIPMON))
         r_str_MonGar = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONGAR))
         r_str_TipDoc = moddat_gf_Consulta_ParDes("026", CStr(g_rst_Princi!HIPGAR_TDOREG))
         r_str_SedReg = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!HIPGAR_SEDREG))
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_HIPGAR("
         g_str_Parame = g_str_Parame & "HIPGAR_NOMRPT, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECCRE, "
         g_str_Parame = g_str_Parame & "HIPGAR_HORCRE, "
         g_str_Parame = g_str_Parame & "HIPGAR_TERCRE, "
         g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE, "
         g_str_Parame = g_str_Parame & "HIPGAR_CODPRD, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECINI, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECFIN, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECDES, "
         g_str_Parame = g_str_Parame & "HIPGAR_BIEGAR, "
         g_str_Parame = g_str_Parame & "HIPGAR_MONEDA, "
         g_str_Parame = g_str_Parame & "HIPGAR_MTOPRE, "
         g_str_Parame = g_str_Parame & "HIPGAR_MONGAR, "
         g_str_Parame = g_str_Parame & "HIPGAR_MTOGAR, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECCON, "
         g_str_Parame = g_str_Parame & "HIPGAR_FECREC, "
         g_str_Parame = g_str_Parame & "HIPGAR_SEDREG, "
         g_str_Parame = g_str_Parame & "HIPGAR_TIPDOC, "
         g_str_Parame = g_str_Parame & "HIPGAR_TOMO, "
         g_str_Parame = g_str_Parame & "HIPGAR_FOJAS, "
         g_str_Parame = g_str_Parame & "HIPGAR_LIBRO) "
                  
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "CTB_RPTSOL_03.RPT" & "', "
         g_str_Parame = g_str_Parame & l_str_Fecha & ", "
         g_str_Parame = g_str_Parame & l_str_Hora & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_NUMOPE & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_CODPRD & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_FECDES & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_TipGar & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_TipMon & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_MTOPRE & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_MonGar & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_MTOGAR & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPGAR_FECCON & ", "
         'g_str_Parame = g_str_Parame & g_rst_Princi.Fields("A.SEGFECCRE") & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!SEGFECCRE & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_SedReg & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_TipDoc & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPGAR_NUMTOM & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPGAR_NUMFOJ & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPGAR_NUMLIB & "') "
                           
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
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_HIPGAR"
   'crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   'crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".RPT_SOLTRA"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_03.RPT"
   
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
   
   Call gs_SetFocus(cmb_TipPro)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia()
   cmb_TipPro.ListIndex = -1
   chk_TipPro.Value = 0
   
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A"
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPGAR B ON B.HIPGAR_NUMOPE = A.HIPMAE_NUMOPE AND B.HIPGAR_BIEGAR = 1"
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2 "
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " ORDER BY HIPMAE_CODPRD ASC, HIPMAE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "OPERACION"
      .Cells(1, 4) = "F.DESEMBOLSO"
      .Cells(1, 5) = "TIPO GARANTIA"
      .Cells(1, 6) = "TIPO MONEDA"
      .Cells(1, 7) = "MTO. PRESTAMO S/."
      .Cells(1, 8) = "MTO. PRESTAMO US$."
      .Cells(1, 9) = "MONEDA GARANTIA"
      .Cells(1, 10) = "MTO. GARANTIA S/."
      .Cells(1, 11) = "MTO. GARANTIA US$."
      .Cells(1, 12) = "F.CONSTITUCION"
      .Cells(1, 13) = "F.RECEPCION MICASITA"
      .Cells(1, 14) = "SEDE REGISTRAL"
      .Cells(1, 15) = "TIPO DOCUMENTO REGISTRAL"
       
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 20
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 30
      .Columns("H").ColumnWidth = 20
      .Columns("I").ColumnWidth = 20
      .Columns("J").ColumnWidth = 20
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 20
      .Columns("M").ColumnWidth = 20
      .Columns("N").ColumnWidth = 20
      .Columns("O").ColumnWidth = 20
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      If IsNull(g_rst_Princi!HIPGAR_BIEGAR) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Consulta_ParDes("030", CStr(g_rst_Princi!HIPGAR_BIEGAR))
      End If
      If IsNull(g_rst_Princi!HIPGAR_TIPMON) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPGAR_TIPMON))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPMAE_INTCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPMAE_TOTPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(g_rst_Princi!HIPMAE_PLAANO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPMAE_SALCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_UlTVCT)))
      If g_rst_Princi!HIPMAE_ULTPAG > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_ULTPAG)))
      End If
      If g_rst_Princi!HIPMAE_VCTANT > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_VCTANT)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = g_rst_Princi!HIPMAE_DIAMOR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
      If g_rst_Princi!HIPMAE_MONGAR = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      End If
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

