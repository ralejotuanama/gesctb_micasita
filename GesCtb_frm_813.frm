VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   9045
   ClientTop       =   3060
   ClientWidth     =   7950
   Icon            =   "GesCtb_frm_813.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   6588
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
         TabIndex        =   7
         Top             =   60
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
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
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   270
            Width           =   5235
            _Version        =   65536
            _ExtentX        =   9234
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Contabilización de Desembolso de Créditos Hipotecarios"
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
            Picture         =   "GesCtb_frm_813.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   780
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
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
            Left            =   7260
            Picture         =   "GesCtb_frm_813.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_813.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_813.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   855
         Left            =   30
         TabIndex        =   11
         Top             =   2280
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1508
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1500
            TabIndex        =   1
            Top             =   90
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
            Left            =   1500
            TabIndex        =   2
            Top             =   450
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
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   450
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   14
         Top             =   3180
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_BarPro 
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   16777215
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
            FloodType       =   1
            FloodColor      =   49152
            Font3D          =   2
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   16
         Top             =   1470
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1349
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6285
         End
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   1530
            TabIndex        =   17
            Top             =   390
            Width           =   6285
            _Version        =   65536
            _ExtentX        =   11086
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Período:"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim p_FecIni            As String
Dim p_FecFin            As String
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmd_Proces_Click()
Dim r_str_PriDia        As String
Dim r_str_UltDia        As String
Dim r_str_PerMes        As String
Dim r_str_PerAno        As String
Dim r_str_FecAct        As String
Dim r_str_PerIni        As String
Dim r_str_PerFin        As String
Dim r_str_CtbIni        As String
Dim r_str_CtbFin        As String
Dim r_rst_PerMes        As ADODB.Recordset
      
   'Fecha de Movimiento
   p_FecIni = ipp_FecIni.Text
   p_FecFin = ipp_FecFin.Text
   r_str_FecAct = date
   
   r_str_PriDia = "01" & "/" & Mid(ipp_FecIni, 4, 2) & "/" & Mid(ipp_FecIni, 7, 4)
   r_str_UltDia = ff_Ultimo_Dia_Mes(Mid(ipp_FecIni, 4, 2), Mid(ipp_FecIni, 7, 4)) & "/" & Mid(ipp_FecIni, 4, 2) & "/" & Mid(ipp_FecIni, 7, 4)
   r_str_PerMes = Mid(ipp_FecIni, 4, 2)
   r_str_PerAno = Mid(ipp_FecIni, 7, 4)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_PERMES "
   g_str_Parame = g_str_Parame & " WHERE PERMES_SITUAC =  1 "
   g_str_Parame = g_str_Parame & " ORDER BY PERMES_CODEMP, PERMES_TIPPER, PERMES_CODANO, PERMES_CODMES ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_PerMes = Format(g_rst_Princi!PERMES_CODMES, "00")
         r_str_PerAno = Format(g_rst_Princi!PERMES_CODANO, "0000")
         
         If g_rst_Princi!PERMES_TIPPER = 1 Then
            r_str_PerIni = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!PERMES_FECINI))))
            r_str_PerFin = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!PERMES_FECFIN))))
         ElseIf g_rst_Princi!PERMES_TIPPER = 2 Then
            r_str_CtbIni = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!PERMES_FECINI))))
            r_str_CtbFin = CDate(gf_FormatoFecha(CStr(Trim(g_rst_Princi!PERMES_FECFIN))))
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   'Validacion del Ingreso de Tipo de Cambio
   If Not fs_ValidacionTipoCambio(p_FecIni, p_FecFin, r_str_CtbFin) Then
      MsgBox "Debe de Ingresar el Tipo de Cambio SBS o Sunat.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (modtac_gf_ValidaTipCamDia_2(r_str_FecAct, p_FecFin) = 0) Then
      MsgBox "Debe de Ingresar el Tipo de Cambio SBS o Sunat.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   If Not (CDate(ipp_FecIni.Text) >= CDate(r_str_PerIni) And CDate(ipp_FecIni.Text) <= CDate(r_str_PerFin)) Then
      MsgBox "El rango de Fechas no corresponde a las Fechas Operativas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   If Not (CDate(ipp_FecFin.Text) >= CDate(r_str_PerIni) And CDate(ipp_FecFin.Text) <= CDate(r_str_PerFin)) Then
      MsgBox "El rango de Fechas no corresponde a las Fechas Operativas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
      
   If MsgBox("¿Está seguro de contabilizar los desembolsos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   cmd_Proces.Enabled = False
   cmd_ExpExc.Enabled = False
      
   'Proceso de Desembolso Creditos Hipotecarios
   Call modprc_ctbp1014(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, p_FecIni, p_FecFin, r_str_PerIni, r_str_PerFin, r_str_PerMes, r_str_PerAno, r_str_CtbIni, r_str_CtbFin, pnl_BarPro)
    
   Screen.MousePointer = 0
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   cmd_Proces.Enabled = True
   cmd_ExpExc.Enabled = True
   
   'Exportación al Crystal Report
   'Screen.MousePointer = 11
   'crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   'crp_Imprim.DataFiles(0) = "CNTBL_ASIENTO"
   'crp_Imprim.DataFiles(1) = "CNTBL_ASIENTO_DET"
     
   'Se selecciona la formula
   'crp_Imprim.SelectionFormula = ""
   
   'Se realiza la validación para codigo de instancia y fechas
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.ORIGEN} = {CNTBL_ASIENTO_DET.ORIGEN} AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.ANO} = {CNTBL_ASIENTO_DET.ANO} AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.MES} = {CNTBL_ASIENTO_DET.MES} AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.NRO_LIBRO} = {CNTBL_ASIENTO_DET.NRO_LIBRO} AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.NRO_ASIENTO} = {CNTBL_ASIENTO_DET.NRO_ASIENTO} AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.TIPO_NOTA} = 'O' AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.FECHA_CNTBL} >= #" & Format(CDate(p_FecIni), "mm/dd/yyyy") & "# AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO.FECHA_CNTBL} <= #" & Format(CDate(p_FecFin), "mm/dd/yyyy") & "# AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.ANO} = " & r_str_PerAno & " AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.MES} = " & r_str_PerMes & " AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.NRO_LIBRO} = 7 AND "
   'crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CNTBL_ASIENTO_DET.NRO_DOCREF3} = '1' "
              
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   'crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTAMO_01.RPT"
   
   'Se le envia el destino a una ventana de crystal report
   'crp_Imprim.Destination = crptToWindow
   'crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   'Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   p_FecIni = ipp_FecIni.Text
   p_FecFin = ipp_FecFin.Text
    
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar el proceso de Desembolso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
         
   Screen.MousePointer = 11
   Call fs_GenExc
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
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_FecIni)
    End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_FecFin)
    End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Proces)
    End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_ConVer        As Integer
Dim p_fecMes            As String
Dim p_FecAno            As String
Dim r_str_PriDia        As String
Dim r_str_UltDia        As String
Dim r_str_PerMes        As String
Dim r_str_PerAno        As String
Dim r_str_FecAct        As String
Dim r_str_PerIni        As String
Dim r_str_PerFin        As String
Dim r_str_CtbIni        As String
Dim r_str_CtbFin        As String
Dim r_rst_PerMes        As ADODB.Recordset
   
   p_fecMes = Mid(p_FecIni, 4, 2)
   p_FecAno = Mid(p_FecIni, 7, 4)
   p_FecIni = ipp_FecIni.Text
   p_FecFin = ipp_FecFin.Text
   r_str_FecAct = date
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.ORIGEN AS NUMSUC, A.NRO_LIBRO AS NROLIB, A.NRO_ASIENTO AS NROASI, A.FEC_REGISTRO AS FECREG, "
   g_str_Parame = g_str_Parame & "       A.COD_USR AS CODUSR, B.ITEM AS CODREG, B.CNTA_CTBL AS CTACTB, B.FECHA_CNTBL AS FECCTB, B.DET_GLOSA AS DESGLO, "
   g_str_Parame = g_str_Parame & "       B.FLAG_DEBHAB AS DEBHAB, B.IMP_MOVSOL AS MOVSOL, B.IMP_MOVDOL AS MOVDOL, B.CTB_FECAUX AS FECAUX "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO A, CNTBL_ASIENTO_DET B "
   g_str_Parame = g_str_Parame & " WHERE A.ORIGEN = B.ORIGEN "
   g_str_Parame = g_str_Parame & "   AND A.ANO = B.ANO "
   g_str_Parame = g_str_Parame & "   AND A.MES = B.MES "
   g_str_Parame = g_str_Parame & "   AND A.NRO_LIBRO = B.NRO_LIBRO "
   g_str_Parame = g_str_Parame & "   AND A.NRO_ASIENTO = B.NRO_ASIENTO "
   g_str_Parame = g_str_Parame & "   AND A.TIPO_NOTA = 'O' "
   g_str_Parame = g_str_Parame & "   AND B.ANO = " & p_FecAno & " "
   g_str_Parame = g_str_Parame & "   AND B.MES = " & p_fecMes & " "
   g_str_Parame = g_str_Parame & "   AND B.NRO_LIBRO = 6  "
   g_str_Parame = g_str_Parame & "   AND B.NRO_DOCREF3 = '1' "
   g_str_Parame = g_str_Parame & "   AND A.FECHA_CNTBL >= to_date( '" & p_FecIni & "', 'dd/mm/yyyy') "
   g_str_Parame = g_str_Parame & "   AND A.FECHA_CNTBL <= to_date( '" & p_FecFin & "', 'dd/mm/yyyy') "
   g_str_Parame = g_str_Parame & " ORDER BY B.ANO, B.MES, B.NRO_LIBRO, SUBSTR(DESGLO,1,10), B.NRO_ASIENTO, CODREG ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ORIGEN"
      .Cells(1, 2) = "NRO LIBRO"
      .Cells(1, 3) = "NRO ASIENTO"
      .Cells(1, 4) = "FECHA DE REGISTRO"
      .Cells(1, 5) = "USUARIO"
      .Cells(1, 6) = "ITEM"
      .Cells(1, 7) = "CUENTA CONTABLE"
      .Cells(1, 8) = "FECHA CONTABLE"
      .Cells(1, 9) = "NRO OPERACION"
      .Cells(1, 10) = "DESCRIPCION"
      .Cells(1, 11) = "FLAG DEBE/HABER"
      .Cells(1, 12) = "IMPOR. SOLES"
      .Cells(1, 13) = "IMPOR. DOLARES"
      .Cells(1, 15) = "FECHA AUXILIAR"
       
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 9
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 10
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 13
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 10
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 5
      .Columns("G").ColumnWidth = 18
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 17
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 45
      .Columns("K").ColumnWidth = 17
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 13
      .Columns("M").ColumnWidth = 16
      .Columns("O").ColumnWidth = 15
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "@"
      .Columns("I").NumberFormat = "@"
      .Columns("L").NumberFormat = "#,##0.00"
      .Columns("M").NumberFormat = "#,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
      
   Do While Not g_rst_Princi.EOF
      'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = g_rst_Princi!NUMSUC
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!NROLIB
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!NROASI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CDate(g_rst_Princi!FECREG)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!CODUSR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!CODREG
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!CtaCtb)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(g_rst_Princi!FECCTB)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Left(Trim(g_rst_Princi!DESGLO), 10)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!DESGLO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!DEBHAB)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!MOVSOL, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!MOVDOL, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CDate(g_rst_Princi!FECAUX)
            
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function fs_ValidacionTipoCambio(ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_CtbFin As String) As Boolean
Dim r_str_CadEje        As String
Dim r_str_FecDes        As String
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_int_FlgCam        As Integer
   
   fs_ValidacionTipoCambio = False
   r_int_FlgCam = 0
   r_str_CadEje = "SELECT A.CAJMOV_FECMOV, A.CAJMOV_FECDEP, A.CAJMOV_NUMMOV, A.CAJMOV_MONPAG, A.CAJMOV_NUMOPE, B.HIPMAE_FECDES " & _
                  "  FROM OPE_CAJMOV A, CRE_HIPMAE B " & _
                  " WHERE A.CAJMOV_TIPMOV = 1103 AND A.CAJMOV_CTBFLG = 0 AND " & _
                  "       A.CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND " & _
                  "       A.CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " AND " & _
                  "       A.CAJMOV_NUMOPE = B.HIPMAE_NUMOPE " & _
                  " ORDER BY CAJMOV_NUMMOV ASC "
   
   If Not gf_EjecutaSQL(r_str_CadEje, g_rst_Princi, 3) Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_FecDes = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
         r_dbl_TipSbs = fs_ObtieneTipCamDia_3(2, CStr(g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecDes, "yyyymmdd"), 1)
         r_dbl_TipSun = fs_ObtieneTipCamDia_2(2, CStr(g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecDes, "yyyymmdd"), 2)
         
         If r_dbl_TipSbs = 0 Then
            r_int_FlgCam = 1
            MsgBox "Falta ingresar tipo de cambio SBS para la fecha: " & r_str_FecDes, vbExclamation, modgen_g_str_NomPlt
         End If
         If r_dbl_TipSun = 0 Then
            r_int_FlgCam = 1
            MsgBox "Falta ingresar tipo de cambio SUNAT para la fecha: " & r_str_FecDes, vbExclamation, modgen_g_str_NomPlt
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If r_int_FlgCam = 0 Then
      fs_ValidacionTipoCambio = True
   End If
End Function

'Tipo de Cambio SBS
Private Function fs_ObtieneTipCamDia_3(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   fs_ObtieneTipCamDia_3 = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CALENDARIO "
   g_str_Parame = g_str_Parame & " WHERE FECHA = to_date(" & p_FecDia & ",'yyyy/mm/dd') "
   g_str_Parame = g_str_Parame & " ORDER BY FECHA DESC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If IsNull(g_rst_Genera!PROM_SBS) Then
         fs_ObtieneTipCamDia_3 = 0
      Else
         fs_ObtieneTipCamDia_3 = g_rst_Genera!PROM_SBS
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

'Tipo de Cambio Sunat
Private Function fs_ObtieneTipCamDia_2(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   fs_ObtieneTipCamDia_2 = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CALENDARIO "
   g_str_Parame = g_str_Parame & " WHERE FECHA = to_date(" & p_FecDia & ",'yyyy/mm/dd') "
   g_str_Parame = g_str_Parame & " ORDER BY FECHA DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If p_TipTip = 1 Then
         fs_ObtieneTipCamDia_2 = g_rst_Genera!PROM_SBS
      ElseIf p_TipTip = 2 Then
         fs_ObtieneTipCamDia_2 = g_rst_Genera!CMP_DOL_PROM
      Else
         fs_ObtieneTipCamDia_2 = 0
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

