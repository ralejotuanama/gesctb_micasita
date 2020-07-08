VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_EntRen_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "GesCtb_frm_215.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3210
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7245
      _Version        =   65536
      _ExtentX        =   12779
      _ExtentY        =   5662
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
         TabIndex        =   7
         Top             =   60
         Width           =   6990
         _Version        =   65536
         _ExtentX        =   12330
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   300
            Left            =   660
            TabIndex        =   8
            Top             =   180
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Devolución"
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
            Picture         =   "GesCtb_frm_215.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   9
         Top             =   780
         Width           =   6990
         _Version        =   65536
         _ExtentX        =   12330
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_215.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   6390
            Picture         =   "GesCtb_frm_215.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1575
         Left            =   60
         TabIndex        =   10
         Top             =   1500
         Width           =   6990
         _Version        =   65536
         _ExtentX        =   12330
         _ExtentY        =   2778
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
         Begin VB.TextBox txt_NumOper 
            Height          =   315
            Left            =   1530
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1080
            Width           =   2300
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   1530
            TabIndex        =   0
            Top             =   420
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
         Begin EditLib.fpDoubleSingle ipp_ImpDev 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   750
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
            ButtonStyle     =   0
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin Threed.SSPanel pnl_TipCambio 
            Height          =   315
            Left            =   5340
            TabIndex        =   5
            Top             =   420
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
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
            Alignment       =   4
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   4260
            TabIndex        =   15
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro Operación:"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   1125
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago:"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   810
            Width           =   570
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_EntRen_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 0 Then
      'consultar
      pnl_Titulo.Caption = "Registro de Devolución - Consulta"
      cmd_Grabar.Visible = False
      Call fs_Desabilitar
      Call fs_CargarDatos
   ElseIf moddat_g_int_FlgGrb = 1 Then
      'adicion
      pnl_Titulo.Caption = "Registro de Devolución - Adición"
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Limpia()
   ipp_FecPag.Text = moddat_g_str_FecSis
   pnl_TipCambio.Caption = "0.000000" & " "
   Call ipp_FecPag_LostFocus
   ipp_ImpDev.Text = "0.00"
   txt_NumOper.Text = ""
End Sub

Private Sub fs_Desabilitar()
   ipp_FecPag.Enabled = False
   ipp_ImpDev.Enabled = False
   txt_NumOper.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
    If CDbl(ipp_ImpDev.Text) <= 0 Then
        MsgBox "Debe de ingresar el importe a devolver.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(ipp_ImpDev)
        Exit Sub
    End If
    
    If Trim(txt_NumOper.Text) = "" Then
        MsgBox "Debe de ingresar un nro de operación.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_NumOper)
        Exit Sub
    End If
    
    If Format(ipp_FecPag.Text, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
        MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
             
        If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
           Call gs_SetFocus(ipp_FecPag)
           Exit Sub
        End If
    End If
    
    If CDbl(pnl_TipCambio.Caption) = 0 Then
       MsgBox "Tiene que registrar el tipo de cambio sbs del día.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FecPag)
       Exit Sub
    End If
        
   If fs_ValidaPeriodo(ipp_FecPag.Text) = False Then
      Exit Sub
   End If

'    If (Format(ipp_FecPag.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'        Format(ipp_FecPag.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'        MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'        Call gs_SetFocus(ipp_FecPag)
'        Exit Sub
'    End If

   '--ipp_FecPag.Text
'   If Format(moddat_g_str_FecSis, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > modctb_int_PerAno & Format(modctb_int_PerMes, "00") & Format(moddat_g_int_PerLim, "00")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FecPag)
'          Exit Sub
'      End If
'      MsgBox "Los asiento a generar perteneceran al periodo anterior.", vbExclamation, modgen_g_str_NomPlt
'   Else
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FecPag)
'          Exit Sub
'      End If
'   End If
    
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Public Sub fs_Grabar()
Dim r_str_AsiGen   As String
  
   r_str_AsiGen = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_ENTREN_DEV ( "
   g_str_Parame = g_str_Parame & moddat_g_str_Codigo & ", " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & Format(ipp_FecPag.Text, "yyyymmdd") & ", " 'CAJCHC_FECCAJ
   g_str_Parame = g_str_Parame & moddat_g_str_CodMod & ", "  'CAJCHC_CODMON
   g_str_Parame = g_str_Parame & CDbl(pnl_TipCambio.Caption) & ", " 'CAJCHC_TIPCAM
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpDev.Text) & ", " 'CAJCHC_IMPORT
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumOper.Text) & "', " 'CAJCHC_NUMOPE
   '-----------------------------------
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "1) " 'insert
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If g_rst_Genera!RESUL = 1 Then
      Call fs_GeneraAsiento(moddat_g_str_Codigo, g_rst_Genera!Item, r_str_AsiGen)
      MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
             "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
                 
      Call frm_Ctb_EntRen_01.fs_BuscarCaja
      Call frm_Ctb_EntRen_03.fs_BuscarCaja
      Screen.MousePointer = 0
      Unload Me
   End If
  
End Sub

Private Sub fs_GeneraAsiento(ByVal p_Codigo As String, ByVal p_SubCod As Long, ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_Import        As Double
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_CtaHab        As String
Dim r_str_CtaDeb        As String
Dim r_str_CadAux        As String
Dim r_int_PerAno        As Integer
Dim r_int_PerMes        As Integer
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   r_str_CtaHab = ""
   r_str_CtaDeb = ""

   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(ipp_FecPag.Text, "yyyymmdd")
   r_str_FecPrPgoL = ipp_FecPag.Text
      
   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecPag.Text, "yyyymmdd"), 1)
      
   r_str_Glosa = "ER" & p_Codigo & "/" & "DEVOLUCION/" & Trim(txt_NumOper.Text) & "/" & _
                 Trim(frm_Ctb_EntRen_01.grd_Listad.TextMatrix(frm_Ctb_EntRen_01.grd_Listad.Row, 17))
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
         
   r_int_PerMes = modctb_int_PerMes 'Month(ipp_FecPag.Text)
   r_int_PerAno = modctb_int_PerAno 'Year(ipp_FecPag.Text)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_MtoSol = 0
   r_dbl_MtoDol = 0
   If CInt(moddat_g_str_CodMod) = 1 Then
      'Entrega a rendir Soles:
      r_dbl_MtoSol = CDbl(ipp_ImpDev.Text)
      r_dbl_MtoDol = Format(CDbl(CDbl(ipp_ImpDev.Text) / r_dbl_TipSbs), "###,###,##0.00")
      r_str_CtaDeb = "111301060102"
      r_str_CtaHab = "191807020101"
   Else
      'Entrega a rendir dólares:
      r_dbl_MtoSol = Format(CDbl(CDbl(ipp_ImpDev.Text) * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      r_dbl_MtoDol = CDbl(ipp_ImpDev.Text)
      r_str_CtaDeb = "112301060102"
      r_str_CtaHab = "192807020101"
   End If
   
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, r_str_CtaDeb, CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                                        
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, r_str_CtaHab, CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
   p_AsiGen = r_str_AsiGen
   
   'Actualiza flag de contabilizacion
   r_str_CadAux = ""
   r_str_CadAux = r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATCNT = '" & r_str_CadAux & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 4 "
   g_str_Parame = g_str_Parame & "   AND CAJCHC_NUMERO = " & p_SubCod
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
      
End Sub

Private Sub fs_CargarDatos()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT CAJCHC_FECCAJ, CAJCHC_TIPCAM, CAJCHC_IMPORT, CAJCHC_NUMOPE  "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_CAJCHC  "
   g_str_Parame = g_str_Parame & "   WHERE CajChc_CodCaj = " & CLng(moddat_g_str_CodIte)
   g_str_Parame = g_str_Parame & "     AND CAJCHC_TIPTAB = 4  " 'DEVOLUCION
   g_str_Parame = g_str_Parame & "     AND CAJCHC_NUMERO =  " & moddat_g_dbl_IngDec 'ITEM
   g_str_Parame = g_str_Parame & "     AND CAJCHC_SITUAC = 1  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      ipp_FecPag.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)
      ipp_ImpDev.Text = Format(g_rst_Princi!CajChc_Import, "###,###,##0.00")
      txt_NumOper.Text = Trim(g_rst_Princi!CAJCHC_NUMOPE & "")
      pnl_TipCambio.Caption = Format(g_rst_Princi!CAJCHC_TIPCAM, "###,###,##0.000000") & " "
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_FecPag_LostFocus()
   'TIPO CAMBIO SBS(2) - VENTA(1)
   pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FecPag.Text, "yyyymmdd"), 1)
   pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpDev)
   End If
End Sub

Private Sub ipp_ImpDev_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NumOper)
   End If
End Sub

Private Sub txt_NumOper_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

