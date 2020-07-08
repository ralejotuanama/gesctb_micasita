VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CajChc_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "GesCtb_frm_189.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   4110
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   7250
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
         TabIndex        =   9
         Top             =   60
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
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
            TabIndex        =   10
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Caja chica"
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
            Picture         =   "GesCtb_frm_189.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7620
            Picture         =   "GesCtb_frm_189.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_189.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2385
         Left            =   60
         TabIndex        =   12
         Top             =   1500
         Width           =   8235
         _Version        =   65536
         _ExtentX        =   14526
         _ExtentY        =   4207
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
         Begin VB.CheckBox chk_RemAnt 
            Caption         =   "El reembolso corresponde a una rendición del mes pasado"
            Height          =   375
            Left            =   3600
            TabIndex        =   4
            Top             =   1560
            Width           =   4515
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   960
            Width           =   1600
         End
         Begin VB.ComboBox cmb_Respon 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1950
            Width           =   6330
         End
         Begin Threed.SSPanel pnl_NumCaja 
            Height          =   315
            Left            =   1710
            TabIndex        =   17
            Top             =   300
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
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
         Begin EditLib.fpDateTime ipp_FchCaj 
            Height          =   315
            Left            =   1710
            TabIndex        =   0
            Top             =   630
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
         Begin EditLib.fpDoubleSingle ipp_ImpAsig 
            Height          =   315
            Left            =   1710
            TabIndex        =   2
            Top             =   1290
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
         Begin EditLib.fpDoubleSingle ipp_ImpRem 
            Height          =   315
            Left            =   1710
            TabIndex        =   3
            Top             =   1620
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto Reembolsado:"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Monto Asignado:"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   1050
            Width           =   630
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
            TabIndex        =   16
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Caja:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Caja:"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   690
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   2010
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CajChc_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Respon()      As moddat_tpo_Genera
Dim l_int_PerAno        As Integer
Dim l_int_PerMes        As Integer

Private Sub Form_Load()
Dim r_int_NumDet  As Integer
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then 'consultar
      pnl_Titulo.Caption = "Registro de Caja Chica - Consultar"
      cmd_Grabar.Visible = False
      Call fs_Cargar_Datos(r_int_NumDet)
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then 'insertar
      pnl_Titulo.Caption = "Registro de Caja Chica - Adicionar"
   ElseIf moddat_g_int_FlgGrb = 2 Then 'modificar
      pnl_Titulo.Caption = "Registro de Caja Chica - Modificar"
      Call fs_Cargar_Datos(r_int_NumDet)
      Call fs_Desabilitar
      cmb_Respon.Enabled = True
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   
   If moddat_g_int_FlgGrb = 1 Then 'insertar
      Call moddat_gs_Carga_EjecMC(cmb_Respon, l_arr_Respon, 132, 1)
   Else 'editar y insertar
      Call moddat_gs_Carga_EjecMC(cmb_Respon, l_arr_Respon, 132, 2)
   End If
End Sub

Private Sub fs_Limpiar()
   pnl_NumCaja.Caption = ""
   ipp_FchCaj.Text = moddat_g_str_FecSis
   cmb_Moneda.ListIndex = 0
   ipp_ImpAsig.Text = "0.00"
   ipp_ImpRem.Text = "0.00"
   cmb_Respon.ListIndex = -1
   chk_RemAnt.Value = 0
End Sub

Private Sub fs_Desabilitar()
   ipp_FchCaj.Enabled = False
   cmb_Moneda.Enabled = False
   ipp_ImpAsig.Enabled = False
   ipp_ImpRem.Enabled = False
   cmb_Respon.Enabled = False
   chk_RemAnt.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Grabar_Click()
    If (cmb_Moneda.ListIndex = -1) Then
        MsgBox "Debe seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Moneda)
        Exit Sub
    End If
        
    If CDbl(ipp_ImpAsig.Text) <= 0 Then
        MsgBox "El monto asignado debe de ser mayor a cero.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(ipp_ImpAsig)
        Exit Sub
    End If
    
    If CDbl(ipp_ImpRem.Text) > CDbl(ipp_ImpAsig.Text) Then
        MsgBox "El reembolso no puede ser mayor al monto asignado.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(ipp_ImpRem)
        Exit Sub
    End If
    
    If cmb_Respon.ListIndex = -1 Then
        MsgBox "Debe de seleccionar el responsable.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_Respon)
        Exit Sub
    End If
    
    If Format(ipp_FchCaj.Text, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
       MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
             
       If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Call gs_SetFocus(ipp_FchCaj)
          Exit Sub
       End If
    End If
    
'    If (Format(ipp_FchCaj.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'        Format(ipp_FchCaj.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'        MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'        Call gs_SetFocus(ipp_FchCaj)
'        Exit Sub
'    End If

   '--ipp_FchCaj.Text
'   If Format(moddat_g_str_FecSis, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > modctb_int_PerAno & Format(modctb_int_PerMes, "00") & Format(moddat_g_int_PerLim, "00")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FchCaj)
'          Exit Sub
'      End If
'      MsgBox "Los asiento a generar perteneceran al periodo anterior.", vbExclamation, modgen_g_str_NomPlt
'   Else
'      If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'          Format(moddat_g_str_FecSis, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'          MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(ipp_FchCaj)
'          Exit Sub
'      End If
'   End If

    If fs_ValidaPeriodo(ipp_FchCaj.Text) = False Then
       Exit Sub
    End If
   
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar()
Dim r_str_AsiGen   As String
Dim r_str_CodGen  As String

   r_str_AsiGen = ""
   r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 9)
   Else
      r_str_CodGen = Trim(pnl_NumCaja.Caption)
   End If
   
   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC ( "
   g_str_Parame = g_str_Parame & CLng(r_str_CodGen) & ", " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & "'" & Format(ipp_FchCaj.Text, "yyyymmdd") & "', " 'CAJCHC_FECCAJ
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'CAJCHC_CODMON
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpAsig.Text) & ", " 'CAJCHC_IMPORT
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpRem.Text) & ", "  'CAJCHC_IMPREM
   g_str_Parame = g_str_Parame & "'" & l_arr_Respon(cmb_Respon.ListIndex + 1).Genera_Codigo & "', "  'CAJCHC_RESPON
   g_str_Parame = g_str_Parame & chk_RemAnt.Value & ", "  'CAJCHC_REMANT_EST
   g_str_Parame = g_str_Parame & "1, "  'CAJCHC_SITUAC
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (g_rst_Genera!RESUL = 1) Then
       If CDbl(ipp_ImpRem.Text) > 0 Then
          If chk_RemAnt.Value = 0 Then
             Call fs_GeneraAsiento_1(Format(g_rst_Genera!CODIGO, "0000000000"), r_str_AsiGen)
          Else
             Call fs_GeneraAsiento_2(Format(g_rst_Genera!CODIGO, "0000000000"), r_str_AsiGen)
          End If
          MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
                 "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
       Else
          MsgBox "Los datos se grabaron correctamente." & vbCrLf & _
                 "No se genero el asiento por que el monto reembolsado es cero.", vbInformation, modgen_g_str_NomPlt
       End If
       Call frm_Ctb_CajChc_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_CajChc_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 3) Then
       MsgBox "El Importe no puede ser menor al total de su detalle: " & Format(g_rst_Genera!TOTDET, "###,###,##0.00") & "", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_ImpAsig)
       Screen.MousePointer = 0
   End If
End Sub

Private Sub fs_GeneraAsiento_1(ByVal p_Codigo As String, ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
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

   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(ipp_FchCaj.Text, "yyyymmdd")
   r_str_FecPrPgoL = ipp_FchCaj.Text
   
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(3, 2, Format(ipp_FchCaj.Text, "yyyymmdd"), 1)
   
   r_str_Glosa = "REEMBOLSO CAJA CHICA " & p_Codigo
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
   
   'l_int_PerMes = modctb_int_PerMes 'Month(ipp_FchCaj.Text)
   'l_int_PerAno = modctb_int_PerAno 'Year(ipp_FchCaj.Text)
   l_int_PerMes = Month(ipp_FchCaj.Text)
   l_int_PerAno = Year(ipp_FchCaj.Text)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text)
   Else
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text * r_dbl_TipSbs) 'Importe * CONVERTIDO
      r_dbl_ImpDol = CDbl(ipp_ImpRem.Text)
   End If
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, "111701010101", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
                                        
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, "111301060102", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
   p_AsiGen = r_str_AsiGen
   
   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATCNT = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Private Sub fs_GeneraAsiento_2(ByVal p_Codigo As String, ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_FecPrPgoL     As String
Dim r_str_FecPrPgoC     As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   r_int_NumAsi = 0
   
   r_str_Glosa = Mid(Trim("REEMBOLSO CAJA CHICA " & p_Codigo), 1, 60)
   '-------------GENERACION DEL ASIENTO 1-------------------------------------------
   r_str_FecPrPgoL = DateAdd("d", -1, "01/" & Format(ipp_FchCaj.Text, "mm/yyyy"))
   r_str_FecPrPgoC = Format(r_str_FecPrPgoL, "yyyymmdd")
   'l_int_PerMes = modctb_int_PerMes 'Month(r_str_FecPrPgoL)
   'l_int_PerAno = modctb_int_PerAno 'Year(r_str_FecPrPgoL)
   l_int_PerMes = Month(r_str_FecPrPgoL)
   l_int_PerAno = Year(r_str_FecPrPgoL)
   
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(3, 2, Format(r_str_FecPrPgoL, "yyyymmdd"), 1)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text)
   Else
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text * r_dbl_TipSbs) 'Importe * CONVERTIDO
      r_dbl_ImpDol = CDbl(ipp_ImpRem.Text)
   End If
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, "111701010101", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
                                        
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, "291807010101", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
                                           
   p_AsiGen = CStr(r_int_NumAsi)
   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATCNT = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
   '-------------GENERACION DEL ASIENTO 2-------------------------------------------
   r_str_FecPrPgoL = ipp_FchCaj.Text
   r_str_FecPrPgoC = Format(r_str_FecPrPgoL, "yyyymmdd")
   l_int_PerMes = Month(r_str_FecPrPgoL)
   l_int_PerAno = Year(r_str_FecPrPgoL)
   
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(3, 2, Format(r_str_FecPrPgoL, "yyyymmdd"), 1)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_ImpSol = 0
   r_dbl_ImpDol = 0
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text)
   Else
      r_dbl_ImpSol = CDbl(ipp_ImpRem.Text * r_dbl_TipSbs) 'Importe * CONVERTIDO
      r_dbl_ImpDol = CDbl(ipp_ImpRem.Text)
   End If
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, "291807010101", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
                                        
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, "111301060102", CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FecPrPgoL))
                                           
   p_AsiGen = p_AsiGen & " - " & CStr(r_int_NumAsi)
   'Actualiza flag de contabilizacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATCNT_2 = '" & r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub
Private Sub fs_Cargar_Datos(ByRef p_NumDet As Integer)
Dim r_int_Contad As Integer

   Call gs_SetFocus(ipp_FchCaj)
   p_NumDet = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_CODMON, CAJCHC_FLGPRC_2,  "
   g_str_Parame = g_str_Parame & "       A.CAJCHC_IMPORT, A.CAJCHC_IMPORT_2, A.CAJCHC_RESPON,  "
   g_str_Parame = g_str_Parame & "       NVL((SELECT COUNT(*) FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "             WHERE X.CAJDET_CODCAJ = A.CAJCHC_CODCAJ AND X.CAJDET_TIPTAB = 1 AND  X.CAJDET_SITUAC = 1),0) AS DEPENDENCIAS,  "
   g_str_Parame = g_str_Parame & "       NVL((SELECT SUM(X.CAJDET_DEB_PPG1 + X.CAJDET_HAB_PPG1) FROM CNTBL_CAJCHC_DET X "
   g_str_Parame = g_str_Parame & "             WHERE X.CAJDET_CODCAJ = A.CAJCHC_CODCAJ AND X.CAJDET_TIPTAB = 1 AND X.CAJDET_SITUAC = 1),0) AS TOTAL_DET  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & " WHERE A.CAJCHC_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "   AND A.CAJCHC_TIPTAB = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_NumCaja.Caption = Format(g_rst_Princi!CajChc_CodCaj, "0000000000")
      ipp_FchCaj.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!CAJCHC_CODMON)
      ipp_ImpAsig.Text = g_rst_Princi!CajChc_Import
      ipp_ImpRem.Text = g_rst_Princi!CAJCHC_IMPORT_2
      cmb_Respon.ListIndex = gf_Busca_Arregl(l_arr_Respon, g_rst_Princi!CajChc_Respon) - 1
      p_NumDet = g_rst_Princi!DEPENDENCIAS
      chk_RemAnt.Value = g_rst_Princi!CAJCHC_FLGPRC_2
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_FchCaj_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpAsig)
   End If
End Sub

Private Sub ipp_FchCaj_LostFocus()
   If (Format(ipp_FchCaj.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(ipp_FchCaj.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub ipp_ImpAsig_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_ImpRem)
   End If
End Sub

Private Sub cmb_Respon_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_ImpRem_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(chk_RemAnt)
   End If
End Sub

Private Sub chk_RemAnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Respon)
   End If
End Sub

