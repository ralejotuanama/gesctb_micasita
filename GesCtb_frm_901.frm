VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_MtoItf_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3690
   ClientLeft      =   4005
   ClientTop       =   3885
   ClientWidth     =   9765
   Icon            =   "GesCtb_frm_901.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
      _ExtentY        =   6906
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
         TabIndex        =   12
         Top             =   30
         Width           =   9705
         _Version        =   65536
         _ExtentX        =   17119
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   400
            Left            =   600
            TabIndex        =   13
            Top             =   60
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   706
            _StockProps     =   15
            Caption         =   "Mantenimiento de ITF"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.15
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
            Picture         =   "GesCtb_frm_901.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   750
         Width           =   9705
         _Version        =   65536
         _ExtentX        =   17119
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
            Height          =   585
            Left            =   9090
            Picture         =   "GesCtb_frm_901.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_901.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   795
         Left            =   30
         TabIndex        =   15
         Top             =   2340
         Width           =   9705
         _Version        =   65536
         _ExtentX        =   17119
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_TipMov 
            Height          =   315
            Left            =   6030
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   3600
         End
         Begin VB.ComboBox cmb_TipOpe 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   2400
         End
         Begin VB.ComboBox cmb_TipDec 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   2400
         End
         Begin EditLib.fpDateTime ipp_FecDep 
            Height          =   315
            Left            =   6030
            TabIndex        =   7
            Top             =   390
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin VB.Label Label9 
            Caption         =   "Tipo Movimiento:"
            Height          =   315
            Left            =   4590
            TabIndex        =   19
            Top             =   90
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Movimiento:"
            Height          =   315
            Left            =   4590
            TabIndex        =   17
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Declarante:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   20
         Top             =   1470
         Width           =   9705
         _Version        =   65536
         _ExtentX        =   17119
         _ExtentY        =   1455
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
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   6060
            MaxLength       =   12
            TabIndex        =   3
            Top             =   90
            Width           =   1800
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   2400
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   2
            Top             =   420
            Width           =   2400
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   4590
            TabIndex        =   26
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   450
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   465
         Left            =   30
         TabIndex        =   23
         Top             =   3180
         Width           =   9705
         _Version        =   65536
         _ExtentX        =   17119
         _ExtentY        =   820
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
         Begin EditLib.fpDoubleSingle ipp_MtoItf 
            Height          =   315
            Left            =   6030
            TabIndex        =   9
            Top             =   60
            Width           =   1800
            _Version        =   196608
            _ExtentX        =   3175
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
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
         Begin EditLib.fpDoubleSingle ipp_MtoSol 
            Height          =   315
            Left            =   1860
            TabIndex        =   8
            Top             =   60
            Width           =   1800
            _Version        =   196608
            _ExtentX        =   3175
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
         Begin VB.Label Label2 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   4620
            TabIndex        =   25
            Top             =   90
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Monto:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   90
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_MtoItf_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipDec_Click()
   Call gs_SetFocus(cmb_TipOpe)
End Sub

Private Sub cmb_TipDoc_Click()

   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumCom.MaxLength = 8
         Case Else:  txt_NumCom.MaxLength = 12
      End Select
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
      txt_NumCom.Enabled = False
   Else
      txt_NumCom.Enabled = True
   End If
   
   Call gs_SetFocus(txt_NumDoc)
   txt_NumCom.Text = ""
   txt_NumDoc.Text = ""
   
End Sub

Private Sub cmb_TipMov_Click()
   Call gs_SetFocus(ipp_FecDep)
End Sub

Private Sub cmb_TipOpe_Click()
   'If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 1 Then
      'cmb_TipMov.Enabled = False
   '   Call gs_SetFocus(ipp_FecDep)
   'Else
      'cmb_TipMov.Enabled = True
      Call gs_SetFocus(cmb_TipMov)
   'End If
End Sub

Private Sub cmd_Grabar_Click()

   Dim r_dbl_TipCam        As Double
   Dim r_dbl_Porcen        As Double
   Dim r_int_NroCom        As Long
   Dim r_str_Numero        As String

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:
            If Len(Trim(txt_NumDoc.Text)) < 8 Then
               MsgBox "Debe ingresar un Número de Documento de 8 dígitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
         Case Else:
            If Len(Trim(txt_NumDoc.Text)) < 8 Then
               MsgBox "Debe ingresar un Número de Documento de 12 dígitos.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
      End Select
   End If
   
   If Trim(txt_NumDoc.Text) = "" Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
      If Trim(txt_NumCom.Text) = "" Then
         MsgBox "Debe ingresar el Número del Comprobante.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumCom)
         Exit Sub
      End If
   End If
      
   If cmb_TipDec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Declarante.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDec)
      Exit Sub
   End If
   
   If cmb_TipOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOpe)
      Exit Sub
   End If
   
   If cmb_TipMov.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMov)
      Exit Sub
   End If
   
   If ipp_MtoSol.Text = 0 Or ipp_MtoSol.Text = "" Then
      MsgBox "Debe ingresar el Monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoSol)
      Exit Sub
   End If
   
   If ipp_MtoItf.Text = "" Then
      MsgBox "Debe ingresar el ITF.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoItf)
      Exit Sub
   End If
   
   
   If moddat_g_int_FlgGrb = 1 Then
   
      If MsgBox("¿Está seguro que desea realizar el ingreso manual del ITF?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
         r_int_NroCom = modsec_gf_BusMov(Format(Mid(ipp_FecDep.Text, 4, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000"))
      Else
         r_int_NroCom = Trim(txt_NumCom.Text)
      End If
      
      r_dbl_TipCam = ff_TipCam(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      r_dbl_Porcen = ff_Porcen(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         r_str_Numero = ff_BuscarNumero(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex), cmb_TipMov.ItemData(cmb_TipMov.ListIndex))
      Else
         r_str_Numero = ""
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         If cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) = 1 Then
            If Len(r_str_Numero) < 12 Then
               MsgBox "El documento ingresado no presenta Número de Solicitud .", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         Else
            If Len(r_str_Numero) < 10 Then
               MsgBox "El documento ingresado no presenta Número de Operación .", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         End If
      End If
               
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_DETITF("
      g_str_Parame = g_str_Parame & "DETITF_PERMES, "
      g_str_Parame = g_str_Parame & "DETITF_PERANO, "
      g_str_Parame = g_str_Parame & "DETITF_TIPDOC, "
      g_str_Parame = g_str_Parame & "DETITF_TIPDEC, "
      g_str_Parame = g_str_Parame & "DETITF_FECMOV, "
      g_str_Parame = g_str_Parame & "DETITF_NUMDOC, "
      g_str_Parame = g_str_Parame & "DETITF_TIPMOV, "
      g_str_Parame = g_str_Parame & "DETITF_NROCOM, "
      g_str_Parame = g_str_Parame & "DETITF_TIPCOD, "
      g_str_Parame = g_str_Parame & "DETITF_ITFPOR, "
      g_str_Parame = g_str_Parame & "DETITF_MTOORG, "
      g_str_Parame = g_str_Parame & "DETITF_ITFORG, "
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL, "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL, "
      g_str_Parame = g_str_Parame & "DETITF_MTODOL, "
      g_str_Parame = g_str_Parame & "DETITF_ITFDOL, "
      g_str_Parame = g_str_Parame & "DETITF_OPEREF, "
      g_str_Parame = g_str_Parame & "DETITF_TIPMON, "
      g_str_Parame = g_str_Parame & "DETITF_TIPCAM, "
      g_str_Parame = g_str_Parame & "DETITF_MANUAL) "
                                 
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & ", "
      g_str_Parame = g_str_Parame & Format(Right(ipp_FecDep.Text, 4), "0000") & ", "
      g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
      g_str_Parame = g_str_Parame & cmb_TipDec.ItemData(cmb_TipDec.ListIndex) & ", "                             ' 1 - DECLARANTE / 2 - EXTORNO
      g_str_Parame = g_str_Parame & Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00") & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_TipMov.Text & "', "
      g_str_Parame = g_str_Parame & "'" & r_int_NroCom & "', "
      g_str_Parame = g_str_Parame & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "                              '1 - SOLICITUD / 2 - OPERACION
      g_str_Parame = g_str_Parame & r_dbl_Porcen & ","
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ","
      Else
         g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text, "###########0.00") & ","
      End If
      
      g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text, "###########0.00") & ","
      g_str_Parame = g_str_Parame & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ","
      g_str_Parame = g_str_Parame & "'" & r_str_Numero & "', "
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & 2 & ", "
      Else
         g_str_Parame = g_str_Parame & 1 & ", "
      End If
      
      g_str_Parame = g_str_Parame & r_dbl_TipCam & ","
      g_str_Parame = g_str_Parame & 1 & ")"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      
      MsgBox "Se realizó el ingreso manual.", vbInformation, modgen_g_str_NomPlt
   
   ElseIf moddat_g_int_FlgGrb = 2 Then
   
      If MsgBox("¿Está seguro que desea modificar el ingreso manual del ITF?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      
      r_dbl_TipCam = ff_TipCam(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      r_dbl_Porcen = ff_Porcen(Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"), Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00"))
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 6 Then
         r_str_Numero = ff_BuscarNumero(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex), cmb_TipMov.ItemData(cmb_TipMov.ListIndex))
      Else
         r_str_Numero = ""
      End If
               
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CTB_DETITF SET "

      g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & cmb_TipDec.ItemData(cmb_TipDec.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "DETITF_FECMOV = " & Format(Right(ipp_FecDep.Text, 4), "0000") & Format(Mid(ipp_FecDep.Text, 4, 2), "00") & Format(Left(ipp_FecDep.Text, 2), "00") & ", "

      g_str_Parame = g_str_Parame & "DETITF_TIPMOV = '" & cmb_TipMov.Text & "', "
      g_str_Parame = g_str_Parame & "DETITF_NROCOM = '" & Trim(txt_NumCom.Text) & "', "
      g_str_Parame = g_str_Parame & "DETITF_TIPCOD = " & cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFPOR = " & r_dbl_Porcen & ", "
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & "DETITF_MTOORG = " & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & "DETITF_ITFORG = " & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ", "
      Else
         g_str_Parame = g_str_Parame & "DETITF_MTOORG = " & Format(ipp_MtoSol.Text, "###########0.00") & ", "
         g_str_Parame = g_str_Parame & "DETITF_ITFORG = " & Format(ipp_MtoItf.Text, "###########0.00") & ", "
      End If
      
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & Format(ipp_MtoSol.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & Format(ipp_MtoItf.Text, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_MTODOL = " & Format(ipp_MtoSol.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_ITFDOL = " & Format(ipp_MtoItf.Text / r_dbl_TipCam, "###########0.00") & ", "
      g_str_Parame = g_str_Parame & "DETITF_OPEREF = '" & r_str_Numero & "', "
      
      If Left(r_str_Numero, 3) = "001" Or Left(r_str_Numero, 3) = "002" Or Left(r_str_Numero, 3) = "006" Then
         g_str_Parame = g_str_Parame & "DETITF_TIPMON = " & 2 & ", "
      Else
         g_str_Parame = g_str_Parame & "DETITF_TIPMON = " & 1 & ", "
      End If
            
      g_str_Parame = g_str_Parame & "DETITF_TIPCAM = " & r_dbl_TipCam & " "
      
      g_str_Parame = g_str_Parame & "WHERE "
      g_str_Parame = g_str_Parame & "DETITF_PERMES = " & Mid(modsec_g_str_Period, 6, 2) & " AND "
      g_str_Parame = g_str_Parame & "DETITF_PERANO = " & Mid(modsec_g_str_Period, 1, 4) & " AND "
      g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & moddat_g_str_TipDoc & " AND "
      g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & moddat_g_str_NumDoc & " AND "
      g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & moddat_g_int_TipCli & " AND "
      g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & modsec_g_dbl_MtoSol & " AND "
      g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & modsec_g_dbl_ITFSol & "  "
      
                                             
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      
      MsgBox "Se modificó el ingreso manual.", vbInformation, modgen_g_str_NomPlt
   
   End If
   
   Call cmd_Salida_Click

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_Limpia
   ElseIf moddat_g_int_FlgGrb = 2 Then
      Call ff_BusITF
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()

   ipp_FecDep.Text = Format(Now, "DD/MM/YYYY")
   'cmb_TipMov.Enabled = False
   txt_NumCom.Enabled = False
   
   cmb_TipDoc.Clear
   
   cmb_TipDoc.AddItem "DNI"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(1)
   
   cmb_TipDoc.AddItem "CARNE DE EXTRANJERIA"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(2)
   
   cmb_TipDoc.AddItem "PASAPORTE"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(5)
   
   cmb_TipDoc.AddItem "RUC"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(6)
   
   cmb_TipDec.Clear
   
   cmb_TipDec.AddItem "DECLARANTE"
   cmb_TipDec.ItemData(cmb_TipDec.NewIndex) = CInt(1)
   
   cmb_TipDec.AddItem "EXTORNO"
   cmb_TipDec.ItemData(cmb_TipDec.NewIndex) = CInt(2)
   
   cmb_TipOpe.Clear
   
   cmb_TipOpe.AddItem "SOLICITUD"
   cmb_TipOpe.ItemData(cmb_TipOpe.NewIndex) = CInt(1)
   
   cmb_TipOpe.AddItem "OPERACION"
   cmb_TipOpe.ItemData(cmb_TipOpe.NewIndex) = CInt(2)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMov, 1, "301")
      
End Sub

Private Sub fs_Limpia()

   cmb_TipDoc.ListIndex = -1
   cmb_TipDec.ListIndex = -1
   cmb_TipOpe.ListIndex = -1
   cmb_TipMov.ListIndex = -1
   
End Sub

Private Sub ipp_FecDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoSol)
   End If
End Sub

Private Sub ipp_MtoItf_GotFocus()
   Call gs_SelecTodo(ipp_MtoItf)
End Sub

Private Sub ipp_MtoItf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_MtoSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoItf)
   End If
End Sub

Private Sub ipp_MtoSol_GotFocus()
   Call gs_SelecTodo(ipp_MtoSol)
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub txt_NumCom_GotFocus()
   Call gs_SelecTodo(txt_NumCom)
End Sub

Private Sub txt_NumCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If

End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6 Then
         Call gs_SetFocus(txt_NumCom)
      Else
         Call gs_SetFocus(cmb_TipDec)
      End If
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 5:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 6:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Function ff_TipCam(ByVal p_FecPag As String, ByVal p_CtbFin As String) As Double
   Dim r_str_FecPag As String
            
   If CDate(gf_FormatoFecha(CStr(p_FecPag))) > CDate(gf_FormatoFecha(CStr(p_CtbFin))) Then
      r_str_FecPag = CDate(gf_FormatoFecha(CStr(p_CtbFin)))
   Else
      r_str_FecPag = CDate(gf_FormatoFecha(CStr(p_FecPag)))
   End If
            
            
   g_str_Parame = "SELECT FECHA, VTA_DOL_PROM, CMP_DOL_PROM FROM CALENDARIO WHERE "
   g_str_Parame = g_str_Parame & "FECHA = to_date ('" & r_str_FecPag & "','DD/MM/YYYY')"
   g_str_Parame = g_str_Parame & "ORDER BY FECHA DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
      
         If Trim(g_rst_Listas!Fecha) = Format(r_str_FecPag, "dd/mm/yyyy") Then
            ff_TipCam = g_rst_Listas!CMP_DOL_PROM
            g_rst_Listas.MoveNext
         Else
            ff_TipCam = 0
            g_rst_Listas.MoveNext
         End If
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function

Private Function ff_Porcen(ByVal p_FecIni As String, ByVal p_FecFin As String) As Double
      
   ff_Porcen = 0
      
   g_str_Parame = "SELECT * FROM OPE_TABITF WHERE "
   g_str_Parame = g_str_Parame & "TABITF_FECINI <= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "TABITF_FECFIN >= " & p_FecFin & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
                                             
         ff_Porcen = g_rst_Listas!TABITF_PORCEN
                
         g_rst_Listas.MoveNext
         
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function


Private Function ff_BuscarNumero(ByVal p_TdoCli As String, ByVal p_ndocli As String, ByVal p_TipPag As Integer, ByVal p_TipMov As Integer) As String

   ff_BuscarNumero = ""
   
   If p_TipPag = 2 Then
   
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & p_TdoCli & " AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & p_ndocli & "' AND "
      g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 9 )"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!HIPMAE_NUMOPE)
      End If
      
   ElseIf p_TipPag = 1 Then
   
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & p_TdoCli & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = " & p_ndocli & "  "
      
      If Left(p_TipMov, 2) <> 21 Then
         g_str_Parame = g_str_Parame & "AND SOLMAE_SITUAC = 1 "
      Else
         g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!SOLMAE_NUMERO)
      End If
   
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function

Private Sub ff_BusITF()

   Dim r_int_Contad As Integer
         
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_PERMES = " & Mid(modsec_g_str_Period, 6, 2) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO = " & Mid(modsec_g_str_Period, 1, 4) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & moddat_g_str_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & moddat_g_str_NumDoc & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDEC = " & moddat_g_int_TipCli & " AND "
   g_str_Parame = g_str_Parame & "DETITF_MTOSOL = " & modsec_g_dbl_MtoSol & " AND "
   g_str_Parame = g_str_Parame & "DETITF_ITFSOL = " & modsec_g_dbl_ITFSol & "  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      'cmb_TipDoc.ListIndex = g_rst_Listas!DETITF_TIPDOC
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Listas!DETITF_TIPDOC)
      txt_NumDoc.Text = g_rst_Listas!DETITF_NUMDOC
      txt_NumCom.Text = g_rst_Listas!DETITF_NROCOM
      'cmb_TipDec.ListIndex = g_rst_Listas!DETITF_TIPDEC
      Call gs_BuscarCombo_Item(cmb_TipDec, g_rst_Listas!DETITF_TIPDEC)
      'cmb_TipOpe.ListIndex = g_rst_Listas!DETITF_TIPCOD
      Call gs_BuscarCombo_Item(cmb_TipOpe, g_rst_Listas!DETITF_TIPCOD)
      'cmb_TipMov.ListIndex = g_rst_Listas!DETITF_TIPMOV
      
      For r_int_Contad = 1 To cmb_TipMov.ListCount Step 1
         cmb_TipMov.ListIndex = r_int_Contad
         If cmb_TipMov.Text = Trim(g_rst_Listas!DETITF_TIPMOV) Then
            Exit For
         End If
      
      Next
      
      'Call gs_BuscarCombo_Item(cmb_TipMov, g_rst_Listas!DETITF_TIPMOV)
      
      ipp_FecDep.Text = Mid(g_rst_Listas!DETITF_FECMOV, 7, 2) & "/" & Mid(g_rst_Listas!DETITF_FECMOV, 5, 2) & "/" & Mid(g_rst_Listas!DETITF_FECMOV, 1, 4)
      ipp_MtoSol.Text = g_rst_Listas!DETITF_MTOSOL
      ipp_MtoItf.Text = g_rst_Listas!DETITF_ITFSOL

   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Sub




