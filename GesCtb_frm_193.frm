VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_EntRen_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "GesCtb_frm_193.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7350
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8235
      _Version        =   65536
      _ExtentX        =   14526
      _ExtentY        =   12965
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
         TabIndex        =   17
         Top             =   60
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
            TabIndex        =   18
            Top             =   180
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Entregas a Rendir"
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
            Picture         =   "GesCtb_frm_193.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   19
         Top             =   780
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
            Picture         =   "GesCtb_frm_193.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7320
            Picture         =   "GesCtb_frm_193.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2550
         Left            =   60
         TabIndex        =   20
         Top             =   1500
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   4498
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
         Begin VB.ComboBox cmb_TipoPago 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   750
            Width           =   1600
         End
         Begin VB.TextBox txt_NumOpera 
            Height          =   315
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1740
            Width           =   1500
         End
         Begin VB.TextBox txt_Glosa 
            Height          =   315
            Left            =   1710
            MaxLength       =   60
            TabIndex        =   7
            Top             =   2070
            Width           =   5910
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1410
            Width           =   2880
         End
         Begin Threed.SSPanel pnl_NumCaja 
            Height          =   315
            Left            =   1710
            TabIndex        =   0
            Top             =   420
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
            TabIndex        =   2
            Top             =   1080
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
            TabIndex        =   5
            Top             =   1745
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
            Left            =   6120
            TabIndex        =   3
            Top             =   1080
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Pago:"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro Operación:"
            Height          =   195
            Left            =   4950
            TabIndex        =   28
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio:"
            Height          =   195
            Left            =   4950
            TabIndex        =   27
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   2130
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Entrega:"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   1140
            Width           =   1320
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Entrega:"
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos del pago"
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
            TabIndex        =   23
            Top             =   90
            Width           =   1305
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   1470
            Width           =   630
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Monto Asignado:"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   1800
            Width           =   1200
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1215
         Left            =   60
         TabIndex        =   29
         Top             =   4095
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   2143
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   420
            Width           =   5910
         End
         Begin VB.ComboBox cmb_Respon 
            Height          =   315
            ItemData        =   "GesCtb_frm_193.frx":0B9A
            Left            =   1710
            List            =   "GesCtb_frm_193.frx":0B9C
            TabIndex        =   9
            Text            =   "cmb_Respon"
            Top             =   755
            Width           =   5910
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
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
            TabIndex        =   32
            Top             =   90
            Width           =   1110
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   800
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1905
         Left            =   60
         TabIndex        =   33
         Top             =   5355
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   3351
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
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1100
            Width           =   2880
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1440
            Width           =   2880
         End
         Begin VB.ComboBox cmb_Benefi 
            Height          =   315
            ItemData        =   "GesCtb_frm_193.frx":0B9E
            Left            =   1710
            List            =   "GesCtb_frm_193.frx":0BA0
            TabIndex        =   11
            Text            =   "cmb_Benefi"
            Top             =   750
            Width           =   5910
         End
         Begin VB.ComboBox cmb_TipDoc_Bnf 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   5910
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   1470
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   1140
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Beneficiario:"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   795
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Beneficiario"
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
            TabIndex        =   34
            Top             =   90
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_EntRen_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_IniFrm       As Boolean
Dim l_arr_PrvRsp()     As moddat_tpo_Genera
Dim l_arr_PrvBnf()     As moddat_tpo_Genera
Dim l_arr_CtaCteSol()  As moddat_tpo_Genera
Dim l_arr_CtaCteDol()  As moddat_tpo_Genera

Private Sub cmb_Banco_Click()
Dim r_str_Cadena  As String
Dim r_int_Contar  As Integer

   cmb_CtaCte.Clear
   r_str_Cadena = ""
   lbl_Cuenta.Caption = "Cuenta:"
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   If (cmb_Moneda.ListIndex = 0) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
   
   If (cmb_Moneda.ListIndex = 1) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_Import  As Double
    r_dbl_Import = 0

    If (cmb_TipoPago.ListIndex = -1) Then
        MsgBox "Debe seleccione un tipo de pago.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipoPago)
        Exit Sub
    End If
    
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
    
    If Len(Trim(txt_Glosa.Text)) = 0 Then
        MsgBox "Tiene que Ingresa una glosa.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_Glosa)
        Exit Sub
    End If
        
    'validacion del responsable
    If cmb_TipDoc.ListIndex = -1 Then
        MsgBox "Debe de seleccionar el tipo documento del responsable.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipDoc)
        Exit Sub
    End If
            
    If fs_Valida_LstPrv(cmb_TipDoc, cmb_Respon, "responsable", l_arr_PrvRsp) = False Then
       Exit Sub
    End If
    
    'validacion del beneficiario
    If cmb_TipDoc_Bnf.ListIndex = -1 Then
        MsgBox "Debe de seleccionar el tipo documento del beneficiario.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipDoc_Bnf)
        Exit Sub
    End If
    
    If fs_Valida_LstPrv(cmb_TipDoc_Bnf, cmb_Benefi, "beneficiario", l_arr_PrvBnf) = False Then
       Exit Sub
    End If
    
    If cmb_Banco.ListIndex = -1 Then
       If MsgBox("¿No ha seleccionado un banco, desea continuar.?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
    End If
    
    If cmb_Banco.ListIndex > -1 Then
       If cmb_CtaCte.ListIndex = -1 Then
          MsgBox "Tiene que seleccionar una cuenta cuenta.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_CtaCte)
          Exit Sub
      End If
    End If
        
    If Format(ipp_FchCaj.Text, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
       MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
             
       If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Call gs_SetFocus(ipp_FchCaj)
          Exit Sub
       End If
    End If
    
    If CDbl(pnl_TipCambio.Caption) = 0 Then
       MsgBox "Tiene que registrar el tipo de cambio sbs del día.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FchCaj)
       Exit Sub
    End If
   
    '(detraccion)
    r_dbl_Import = 0
    r_dbl_Import = CDbl(ipp_ImpAsig.Text) 'SOLES
    If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 2 Then
       'DOLARES
       r_dbl_Import = CDbl(ipp_ImpAsig.Text) * CDbl(pnl_TipCambio.Caption)
    End If
   
    '1. Si el importe es mayor a S/.700 (Si es dólares convertido al tipo de cambio) debe mostrar la siguiente pregunta:
    If r_dbl_Import > 700 Then
       If MsgBox("El importe a pagar puede estar sujeto a la detracción del IGV. ¿desea continuar?, Si no está seguro por favor consulte a Contabilidad.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
    End If

   If fs_ValidaPeriodo(ipp_FchCaj.Text) = False Then
      Exit Sub
   End If
   
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

'    If (Format(ipp_FchCaj.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'        Format(ipp_FchCaj.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'        MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'        Call gs_SetFocus(ipp_FchCaj)
'        Exit Sub
'    End If
    
    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    Call fs_Grabar
    Screen.MousePointer = 0
End Sub

Private Function fs_Valida_LstPrv(p_ComboTip As ComboBox, p_ComboNom As ComboBox, p_MsgNom As String, p_Arregl() As moddat_tpo_Genera) As Boolean
Dim r_int_Contar  As Integer
Dim r_bol_Estado  As Boolean
   
   fs_Valida_LstPrv = True
   r_bol_Estado = True
   
   If Len(Trim(p_ComboNom.Text)) = 0 Then
       MsgBox "Tiene que ingresar un " & p_MsgNom & ".", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(p_ComboNom)
       r_bol_Estado = False 'Exit Sub
   Else
       If (fs_ValNumDoc(p_ComboTip, p_ComboNom) = False) Then
           r_bol_Estado = False 'Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(p_ComboNom.Text), "-") > 0 Then
              For r_int_Contar = 1 To UBound(p_Arregl)
                  If Trim(Mid(p_ComboNom.Text, 1, InStr(Trim(p_ComboNom.Text), "-") - 1)) = Trim(p_Arregl(r_int_Contar).Genera_Codigo) Then
                     r_bol_Estado = True
                     Exit For
                  End If
              Next
           End If
           If r_bol_Estado = False Then
              MsgBox "El " & p_MsgNom & " no se encuentra en la lista.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(p_ComboNom)
              'Exit Sub
           End If
       End If
   End If
   
   fs_Valida_LstPrv = r_bol_Estado
End Function

Private Function fs_ValNumDoc(p_ComboTip As ComboBox, p_ComboNom As ComboBox) As Boolean
Dim r_str_NumDoc  As String
Dim r_bol_Estado  As Boolean

   fs_ValNumDoc = True
   r_str_NumDoc = ""

   r_str_NumDoc = fs_NumDoc(p_ComboNom.Text, p_ComboTip)
   If (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 1) Then 'DNI - 8
       If Len(Trim(r_str_NumDoc)) <> 8 Then
          MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   ElseIf (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 6) Then 'RUC - 11
       If Not gf_Valida_RUC(Trim(r_str_NumDoc), Mid(Trim(r_str_NumDoc), Len(Trim(r_str_NumDoc)), 1)) Then
          MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   Else 'OTROS
       If Len(Trim(p_ComboNom.Text)) = 0 Then
          MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(p_ComboNom)
          fs_ValNumDoc = False
       End If
   End If
   
End Function

Private Function fs_NumDoc(p_Cadena As String, p_ComboTip As ComboBox) As String
   fs_NumDoc = ""
   If (p_ComboTip.ListIndex > -1) Then
      If (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (p_ComboTip.ItemData(p_ComboTip.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
           If InStr(1, p_Cadena, "-") <= 0 Then
              Exit Function
           End If
           fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
      End If
   End If
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_int_NumDet  As Integer
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   l_dbl_IniFrm = False
   
   Call fs_Inicia
   Call fs_Limpiar
   
   If moddat_g_int_FlgGrb = 0 Then 'consultar
      pnl_Titulo.Caption = "Registro de Entregas a Rendir - Consultar"
      cmd_Grabar.Visible = False
      Call fs_Cargar_Datos(r_int_NumDet)
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then 'insertar
      pnl_Titulo.Caption = "Registro de Entregas a Rendir - Adicionar"
   ElseIf moddat_g_int_FlgGrb = 2 Then 'modificar
      pnl_Titulo.Caption = "Registro de Entregas a Rendir - Modificar"
      Call fs_Cargar_Datos(r_int_NumDet)
      Call fs_Desabilitar
   End If
   
   txt_NumOpera.Enabled = False
   
   l_dbl_IniFrm = True
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call fs_CargaMntPardes(cmb_TipDoc, "118", 1) 'RESPONSABLE
   Call fs_CargaMntPardes(cmb_TipDoc_Bnf, "118", 2) 'BENEFICIARIO
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipoPago, 1, "138")
End Sub

Private Sub fs_Limpiar()
   pnl_NumCaja.Caption = ""
   pnl_TipCambio.Caption = "0.000000" & " "
   ipp_FchCaj.Text = moddat_g_str_FecSis
   Call ipp_FchCaj_LostFocus 'TIPO CAMBIO SBS(2) - VENTA(1)
      
   ipp_ImpAsig.Text = "0.00"
   txt_Glosa.Text = ""
   txt_NumOpera.Text = ""
   cmb_TipoPago.ListIndex = -1
   
   cmb_TipDoc.ListIndex = 0
   cmb_Respon.Text = ""
   cmb_TipDoc_Bnf.ListIndex = 0
   cmb_Benefi.Text = ""
   
   cmb_Moneda.ListIndex = 0
   cmb_Banco.Clear
   cmb_CtaCte.Clear
End Sub

Private Sub fs_Desabilitar()
   ipp_FchCaj.Enabled = False
   cmb_Moneda.Enabled = False
   ipp_ImpAsig.Enabled = False
   cmb_TipDoc.Enabled = False
   cmb_Respon.Enabled = False
   txt_Glosa.Enabled = False
   txt_NumOpera.Enabled = False
   
   cmb_TipoPago.Enabled = False
   cmb_TipDoc_Bnf.Enabled = False
   cmb_Benefi.Enabled = False
   cmb_Banco.Enabled = False
   cmb_CtaCte.Enabled = False
End Sub

Private Sub fs_Grabar()
Dim r_str_AsiGen   As String
Dim r_str_CodGen  As String

   r_str_AsiGen = ""
   r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 5)
   Else
      r_str_CodGen = Trim(pnl_NumCaja.Caption)
   End If
   
   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_ENTREN ( "
   g_str_Parame = g_str_Parame & r_str_CodGen & ", " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & "'" & Format(ipp_FchCaj.Text, "yyyymmdd") & "', " 'CAJCHC_FECCAJ
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", " 'CAJCHC_CODMON
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpAsig.Text) & ", " 'CAJCHC_IMPORT
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'CAJCHC_TIPDOC
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Respon.Text, cmb_TipDoc) & "', " 'CAJCHC_NUMDOC
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Glosa.Text) & "', " 'CAJCHC_DESCRI
   g_str_Parame = g_str_Parame & CDbl(pnl_TipCambio.Caption) & ", " 'CAJCHC_TIPCAM
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumOpera.Text) & "', " 'CAJCHC_NUMOPE
   g_str_Parame = g_str_Parame & "1, "  'CAJCHC_SITUAC
   g_str_Parame = g_str_Parame & cmb_TipoPago.ItemData(cmb_TipoPago.ListIndex) & ", "
   g_str_Parame = g_str_Parame & cmb_TipDoc_Bnf.ItemData(cmb_TipDoc_Bnf.ListIndex) & ", "
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Benefi.Text, cmb_TipDoc_Bnf) & "', "
   If cmb_Banco.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "null, "
   Else
      g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', "
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
   
       Call fs_GeneraAsiento(g_rst_Genera!CODIGO, r_str_AsiGen)
       MsgBox "Se culminó proceso de generación de asientos contables." & vbCrLf & _
              "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
                 
       Call frm_Ctb_EntRen_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
       Call frm_Ctb_EntRen_01.fs_BuscarCaja
       Screen.MousePointer = 0
       Unload Me
   End If
End Sub

Private Sub fs_Cargar_Datos(ByRef p_NumDet As Integer)
Dim r_int_Contad As Integer

   Call gs_SetFocus(ipp_FchCaj)
   p_NumDet = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_CODMON, CAJCHC_TIPCAM,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_IMPORT, A.CAJCHC_TIPDOC, A.CAJCHC_NUMDOC, A.CAJCHC_DESCRI,  "
   g_str_Parame = g_str_Parame & "        NVL((SELECT COUNT(*) FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "              WHERE X.CAJDET_CODCAJ = A.CAJCHC_CODCAJ AND X.CAJDET_TIPTAB = 2 AND X.CAJDET_SITUAC = 1),0) AS DEPENDENCIAS,  "
   g_str_Parame = g_str_Parame & "        NVL((SELECT SUM(X.CAJDET_DEB_PPG1 + X.CAJDET_HAB_PPG1) FROM CNTBL_CAJCHC_DET X   "
   g_str_Parame = g_str_Parame & "              WHERE X.CAJDET_CODCAJ = A.CAJCHC_CODCAJ AND X.CAJDET_TIPTAB = 2 AND X.CAJDET_SITUAC = 1),0) AS TOTAL_DET,  "
   g_str_Parame = g_str_Parame & "        CAJCHC_NUMOPE,  "
   g_str_Parame = g_str_Parame & "        CAJCHC_TIPPAG, CAJCHC_TIPDOC_2, CAJCHC_NUMDOC_2, CAJCHC_CODBCO_2 , CAJCHC_CTACRR_2  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_CODCAJ = '" & moddat_g_str_Codigo & "'  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB = 2  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_NumCaja.Caption = Format(g_rst_Princi!CajChc_CodCaj, "00000000")
      ipp_FchCaj.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!CAJCHC_CODMON)
      ipp_ImpAsig.Text = Format(g_rst_Princi!CajChc_Import, "###,###,##0.00")
      pnl_TipCambio.Caption = Format(g_rst_Princi!CAJCHC_TIPCAM, "###,###,##0.000000") & " "
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!CAJCHC_TIPDOC)
      cmb_Respon.ListIndex = fs_ComboIndex(cmb_Respon, g_rst_Princi!CAJCHC_NUMDOC & "", 0)
      txt_NumOpera.Text = g_rst_Princi!CAJCHC_NUMOPE & ""
      txt_Glosa.Text = g_rst_Princi!CajChc_Descri & ""
                  
      If Trim(g_rst_Princi!cajchc_TipPag & "") <> "" Then
         Call gs_BuscarCombo_Item(cmb_TipoPago, g_rst_Princi!cajchc_TipPag)
      End If
      If Trim(g_rst_Princi!CAJCHC_TIPDOC_2 & "") <> "" Then
         Call gs_BuscarCombo_Item(cmb_TipDoc_Bnf, g_rst_Princi!CAJCHC_TIPDOC_2)
      Else
         cmb_TipDoc_Bnf.ListIndex = -1
      End If
      If Trim(g_rst_Princi!CAJCHC_NUMDOC_2 & "") <> "" Then
         cmb_Benefi.ListIndex = fs_ComboIndex(cmb_Benefi, g_rst_Princi!CAJCHC_NUMDOC_2 & "", 0)
      End If
      If Trim(g_rst_Princi!CAJCHC_CODBCO_2 & "") <> "" Then
         Call gs_BuscarCombo_Item(cmb_Banco, g_rst_Princi!CAJCHC_CODBCO_2)
      End If
      If Trim(g_rst_Princi!CAJCHC_CTACRR_2 & "") <> "" Then
         Call gs_BuscarCombo_Text(cmb_CtaCte, g_rst_Princi!CAJCHC_CTACRR_2, -1)
         'cmb_CtaCte.Text = g_rst_Princi!CAJCHC_CTACRR_2
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function fs_ComboIndex(p_Combo As ComboBox, Cadena As String, p_Tipo As Integer) As Integer
Dim r_int_Contad As Integer

   fs_ComboIndex = -1
   For r_int_Contad = 0 To p_Combo.ListCount - 1
       If Trim(Cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, InStr(Trim(p_Combo.List(r_int_Contad)), "-") - 1)) Then
          fs_ComboIndex = r_int_Contad
          Exit For
       End If
   Next
End Function

Private Sub fs_GeneraAsiento(ByVal p_Codigo As String, ByRef p_AsiGen As String)
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
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   r_str_CtaHab = ""
   r_str_CtaDeb = ""

   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(ipp_FchCaj.Text, "yyyymmdd")
   r_str_FecPrPgoL = ipp_FchCaj.Text
      
   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FchCaj.Text, "yyyymmdd"), 1)
      
   r_str_Glosa = "ER" & p_Codigo & "/" & Trim(Mid(cmb_Respon.Text, InStr(1, Trim(cmb_Respon.Text), "-") + 1, Len(Trim(cmb_Respon.Text))))
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
   
   r_int_PerMes = modctb_int_PerMes 'Month(ipp_FchCaj.Text)
   r_int_PerAno = modctb_int_PerAno 'Year(ipp_FchCaj.Text)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_MtoSol = 0
   r_dbl_MtoDol = 0
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      'Entrega a rendir Soles:
      r_dbl_MtoSol = CDbl(ipp_ImpAsig.Text)
      r_dbl_MtoDol = Format(CDbl(CDbl(ipp_ImpAsig.Text) / r_dbl_TipSbs), "###,###,##0.00")
      r_str_CtaDeb = "191807020101"
      r_str_CtaHab = "251419010109"
   Else
      'Entrega a rendir dólares:
      r_dbl_MtoSol = Format(CDbl(CDbl(ipp_ImpAsig.Text) * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      r_dbl_MtoDol = CDbl(ipp_ImpAsig.Text)
      r_str_CtaDeb = "192807020101"
      r_str_CtaHab = "252419010109"
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
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
   
   'Enviar a la tabla de autorizaciones
   If cmb_TipoPago.ItemData(cmb_TipoPago.ListIndex) = 1 Then
      'SOLO LOS PAGOS DE TIPO ANTICIPOS(1)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
      g_str_Parame = g_str_Parame & " " & CLng(p_Codigo) & ", " 'COMAUT_CODOPE
      g_str_Parame = g_str_Parame & " " & Format(ipp_FchCaj.Text, "yyyymmdd") & ", " 'COMAUT_FECOPE
      g_str_Parame = g_str_Parame & " " & cmb_TipDoc_Bnf.ItemData(cmb_TipDoc_Bnf.ListIndex) & ", "  'COMAUT_TIPDOC
      g_str_Parame = g_str_Parame & " '" & fs_NumDoc(cmb_Benefi.Text, cmb_TipDoc_Bnf) & "', "   'COMAUT_NUMDOC
      g_str_Parame = g_str_Parame & " " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "      'COMAUT_CODMON
      g_str_Parame = g_str_Parame & " " & CDbl(ipp_ImpAsig.Text) & ", " 'COMAUT_IMPPAG
      If cmb_Banco.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "null, "  'COMAUT_CODBNC
      Else
         g_str_Parame = g_str_Parame & " " & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "  'COMAUT_CODBNC
      End If
      g_str_Parame = g_str_Parame & " '" & Trim(cmb_CtaCte.Text) & "', " 'COMAUT_CTACRR
      g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB
      g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
      g_str_Parame = g_str_Parame & " '" & Trim(txt_Glosa.Text) & "',  " 'COMAUT_DESCRP
      g_str_Parame = g_str_Parame & " 1,  " 'COMAUT_TIPOPE
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'SEGUSUCRE
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'SEGPLTCRE
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'SEGSUCCRE
                                                                                                                                                                                                                      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   'RESPONSABLE
   Call fs_CargarPrv(cmb_TipDoc, cmb_Respon, 1)
End Sub

Private Sub cmb_TipDoc_Bnf_Click()
   'BENEFICIARIO
   Call fs_CargarPrv(cmb_TipDoc_Bnf, cmb_Benefi, 2)
End Sub

Private Sub fs_CargarPrv(p_Combo_Tdoc As ComboBox, p_Combo_Nom As ComboBox, p_Tipo As Integer)
   If p_Tipo = 1 Then
      ReDim l_arr_PrvRsp(0) 'RESPONSABLE(1)
   Else
      ReDim l_arr_PrvBnf(0) 'BENEFICIARIO(2)
      ReDim l_arr_CtaCteSol(0)
      ReDim l_arr_CtaCteDol(0)
      cmb_Banco.Clear
      cmb_CtaCte.Clear
   End If
   p_Combo_Nom.Clear
   p_Combo_Nom.Text = ""
   If (p_Combo_Tdoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & p_Combo_Tdoc.ItemData(p_Combo_Tdoc.ListIndex)
   If moddat_g_int_FlgGrb = 1 Then 'INSERT
      g_str_Parame = g_str_Parame & " AND A.MAEPRV_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo_Nom.AddItem Trim(g_rst_Genera!MAEPRV_NUMDOC & "") & " - " & Trim(g_rst_Genera!MaePrv_RazSoc & "")
      If p_Tipo = 1 Then
         'RESPONSABLE
         ReDim Preserve l_arr_PrvRsp(UBound(l_arr_PrvRsp) + 1)
         l_arr_PrvRsp(UBound(l_arr_PrvRsp)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
         l_arr_PrvRsp(UBound(l_arr_PrvRsp)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      Else
         'BENEFICIARIO
         ReDim Preserve l_arr_PrvBnf(UBound(l_arr_PrvBnf) + 1)
         l_arr_PrvBnf(UBound(l_arr_PrvBnf)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
         l_arr_PrvBnf(UBound(l_arr_PrvBnf)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      End If
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_CargaMntPardes(p_Combo As ComboBox, ByVal p_CodGrp As String, p_TipPer As Integer)
   'RESPONSABLE = 1
   'BENEFICIARIO = 2
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES A "
   g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   If p_TipPer = 1 Then
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000007') "
   Else
      g_str_Parame = g_str_Parame & " AND A.PARDES_CODITE IN ('000001','000004','000006','000007') "
   End If
   g_str_Parame = g_str_Parame & "   AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!PARDES_CODITE)
            
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmb_Benefi_Click()
   Call fs_Buscar_Ctas
End Sub

Private Sub fs_Buscar_Ctas()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   r_str_NumDoc = ""
   
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc_Bnf.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc_Bnf)
          Exit Sub
       End If
       If cmb_Benefi.ListIndex = -1 Then
          Exit Sub
       End If
      
       If (fs_ValNumDoc(cmb_TipDoc_Bnf, cmb_Benefi) = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Benefi.Text, cmb_TipDoc_Bnf)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_CODBNC_MN1, A.MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN1, A.MAEPRV_CODBNC_MN2, A.MAEPRV_CTACRR_MN2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN2, A.MAEPRV_CODBNC_MN3, A.MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN3, A.MAEPRV_CODBNC_DL1, A.MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL1, A.MAEPRV_CODBNC_DL2, A.MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL2, A.MAEPRV_CODBNC_DL3, A.MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc_Bnf.ItemData(cmb_TipDoc_Bnf.ListIndex)
   g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      MsgBox "No se ha encontrado el beneficiario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Benefi)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   If (moddat_g_int_FlgGrb = 1) Then
       If (g_rst_GenAux!MAEPRV_CONDIC = 2) Then
          MsgBox "El beneficiario se encuentra en condición de NO HABIDO, revisar sunat.", vbExclamation, modgen_g_str_NomPlt
          g_rst_GenAux.Close
          Set g_rst_GenAux = Nothing
          Exit Sub
       End If
       'Call gs_SetFocus(txt_Descrip)
   End If
      
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)

   If (g_rst_GenAux!MAEPRV_CODBNC_MN1 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN1, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN2 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN2)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN2, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN3 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN3)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN3, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   
   If (g_rst_GenAux!MAEPRV_CODBNC_DL1 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL1, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL2 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL2)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL2, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL3 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL3)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL3, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   
   Call fs_CargarBancos
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Sub fs_CargarBancos()
Dim r_bol_Estado   As Boolean
Dim r_int_File     As Integer
Dim r_int_Contar   As Integer

   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   'soles
   If (cmb_Moneda.ListIndex = 0) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteSol(r_int_Contar).Genera_Codigo)
           End If
       Next
   End If
   'dolares
   If (cmb_Moneda.ListIndex = 1) Then
       For r_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteDol(r_int_Contar).Genera_Codigo)
           End If
       Next
   End If
End Sub

Private Sub cmb_Moneda_Click()
   Call fs_CargarBancos
End Sub

Private Sub ipp_FchCaj_LostFocus()
   'TIPO CAMBIO SBS(2) - VENTA(1)
   pnl_TipCambio.Caption = moddat_gf_ObtieneTipCamDia(2, 2, Format(ipp_FchCaj.Text, "yyyymmdd"), 1)
   pnl_TipCambio.Caption = Format(pnl_TipCambio.Caption, "###,###,##0.000000") & " "
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

Private Sub ipp_ImpAsig_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       If txt_NumOpera.Enabled = False Then
          Call gs_SetFocus(txt_Glosa)
       Else
          Call gs_SetFocus(txt_NumOpera)
       End If
   End If
End Sub

Private Sub txt_Glosa_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipDoc)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_NumOpera_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Glosa)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Respon_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_TipDoc_Bnf)
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Respon)
   End If
End Sub

Private Sub cmb_TipoPago_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FchCaj)
   End If
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_TipDoc_Bnf_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Benefi)
   End If
End Sub

Private Sub cmb_Benefi_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

