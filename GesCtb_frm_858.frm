VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_33 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "GesCtb_frm_858.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5160
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6885
      _Version        =   65536
      _ExtentX        =   12144
      _ExtentY        =   9102
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
         TabIndex        =   18
         Top             =   30
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   210
            Left            =   630
            TabIndex        =   19
            Top             =   150
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   370
            _StockProps     =   15
            Caption         =   "Reporte de Control de Flujo de Caja"
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
            Picture         =   "GesCtb_frm_858.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   16
         Top             =   750
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
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
            Left            =   6180
            Picture         =   "GesCtb_frm_858.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_858.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1920
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
         _ExtentY        =   3387
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
         Begin VB.CheckBox chk_Moneda 
            Caption         =   "Todas las Monedas"
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   840
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipMoneda 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   500
            Width           =   5055
         End
         Begin VB.CheckBox chk_Product 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   1560
            Width           =   1995
         End
         Begin VB.ComboBox cmd_TipProducto 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1200
            Width           =   5055
         End
         Begin VB.ComboBox cmb_TipReporte 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   5055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   525
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Producto:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reporte:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   195
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1035
         Left            =   30
         TabIndex        =   14
         Top             =   3390
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
         _ExtentY        =   1834
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
         Begin VB.ComboBox cmb_PerInicial 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1785
         End
         Begin VB.ComboBox cmb_PerFinal 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1785
         End
         Begin EditLib.fpLongInteger ipp_AnoInicial 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   570
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
         Begin EditLib.fpLongInteger ipp_AnoFinal 
            Height          =   315
            Left            =   4680
            TabIndex        =   8
            Top             =   570
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Año Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Final:"
            Height          =   195
            Left            =   3600
            TabIndex        =   23
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Año Final:"
            Height          =   195
            Left            =   3600
            TabIndex        =   22
            Top             =   630
            Width           =   705
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   4455
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
         _ExtentY        =   1129
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
         Begin EditLib.fpDateTime ipp_FecInicial 
            Height          =   315
            Left            =   1440
            TabIndex        =   9
            Top             =   180
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
         Begin EditLib.fpDateTime ipp_FecFinal 
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Top             =   180
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   3600
            TabIndex        =   27
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   225
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_FecIni   As String
Dim l_str_FecFin   As String
Dim l_int_TipRpt   As Integer
Dim l_int_TipMon   As Integer
Dim l_int_TipPrd   As Integer
Dim l_str_RptNom   As String
Dim l_obj_Excel    As Excel.Application
   
Private Sub chk_Moneda_Click()
   If chk_Moneda.Value = 1 Then
      cmb_TipMoneda.ListIndex = -1
      cmb_TipMoneda.Enabled = False
   ElseIf chk_Moneda.Value = 0 Then
      cmb_TipMoneda.Enabled = True
      Call gs_SetFocus(cmb_TipMoneda)
   End If
   Call fs_GeneraProd
End Sub

Private Sub chk_Product_Click()
   If chk_Product.Value = 1 Then
      cmd_TipProducto.ListIndex = -1
      cmd_TipProducto.Enabled = False
   ElseIf chk_Product.Value = 0 Then
      cmd_TipProducto.Enabled = True
      Call gs_SetFocus(cmd_TipProducto)
   End If
End Sub

Private Sub cmb_TipReporte_Click()
    cmb_PerInicial.Enabled = False
    ipp_AnoInicial.Enabled = False
    cmb_PerFinal.Enabled = False
    ipp_AnoFinal.Enabled = False
    ipp_FecInicial.Enabled = False
    ipp_FecFinal.Enabled = False
    
    ipp_FecInicial.AllowNull = True
    ipp_FecFinal.AllowNull = True
    ipp_FecInicial.Text = ""
    ipp_FecFinal.Text = ""
    
    If (cmb_TipReporte.ListIndex <> 0 And cmb_TipReporte.ListIndex <> 1) Then
        cmb_PerInicial.ListIndex = -1
        cmb_PerFinal.ListIndex = -1
        ipp_AnoInicial.Text = 0
        ipp_AnoFinal.Text = 0
    End If
        
    ipp_FecInicial.Text = ""
    ipp_FecFinal.Text = ""
    
    If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
        cmb_PerInicial.Enabled = True
        ipp_AnoInicial.Enabled = True
        cmb_PerFinal.Enabled = True
        ipp_AnoFinal.Enabled = True
        cmb_PerInicial.Enabled = True
        cmb_PerFinal.Enabled = True
        If (ipp_AnoInicial.Text = 0) Then
            ipp_AnoInicial.Text = Year(date)
        End If
        If (ipp_AnoFinal.Text = 0) Then
            ipp_AnoFinal.Text = Year(date)
        End If
    Else
        ipp_FecInicial.AllowNull = False
        ipp_FecFinal.AllowNull = False
        ipp_FecInicial.Enabled = True
        ipp_FecFinal.Enabled = True
        ipp_FecInicial.Text = date
        ipp_FecFinal.Text = date
    End If
    Call fs_GeneraProd
End Sub

Private Sub cmb_TipMoneda_Click()
    Call fs_GeneraProd
End Sub

Private Sub fs_GeneraProd()
    cmd_TipProducto.Clear
    If (cmb_TipReporte.ListIndex = 1 Or cmb_TipReporte.ListIndex = 2) Then 'REPORTE CxC
        If (cmb_TipMoneda.ListIndex = 0) Then 'soles
            cmd_TipProducto.AddItem "CME" '1
            cmd_TipProducto.AddItem "MICASITA" '3
            cmd_TipProducto.AddItem "MIVIVIENDA" '4
        ElseIf (cmb_TipMoneda.ListIndex = 1) Then 'dolares
            cmd_TipProducto.AddItem "CRC-PBP" '2
            cmd_TipProducto.AddItem "MICASITA" '3
        End If
        If (chk_Moneda.Value = 1) Then 'todos
            cmd_TipProducto.AddItem "CME" '1
            cmd_TipProducto.AddItem "CRC-PBP" '2
            cmd_TipProducto.AddItem "MICASITA" '3
            cmd_TipProducto.AddItem "MIVIVIENDA" '4
        End If
    ElseIf (cmb_TipReporte.ListIndex = 0) Then 'REPORTE CxP
        If (cmb_TipMoneda.ListIndex = 0) Then 'soles
            cmd_TipProducto.AddItem "CME" '1
            cmd_TipProducto.AddItem "MIVIVIENDA" '4
        End If
        If (chk_Moneda.Value = 1) Then 'todos
            cmd_TipProducto.AddItem "CME" '1
            cmd_TipProducto.AddItem "MIVIVIENDA" '4
        End If
    End If
End Sub


Private Sub cmb_TipReporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (chk_Moneda.Value = 1) Then
          If (chk_Product.Value = 1) Then
              If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
                  Call gs_SetFocus(cmb_PerInicial)
              Else
                  Call gs_SetFocus(ipp_FecInicial)
              End If
          Else
              Call gs_SetFocus(cmd_TipProducto)
          End If
      Else
          Call gs_SetFocus(cmb_TipMoneda)
      End If
   End If
End Sub

Private Sub cmb_TipMoneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (chk_Product.Value = 1) Then
          If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
              Call gs_SetFocus(cmb_PerInicial)
          Else
              Call gs_SetFocus(ipp_FecInicial)
          End If
      Else
          Call gs_SetFocus(cmd_TipProducto)
      End If
   End If
End Sub

Private Sub cmd_TipProducto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
          Call gs_SetFocus(cmb_PerInicial)
      Else
          Call gs_SetFocus(ipp_FecInicial)
      End If
   End If
End Sub

Private Sub cmb_PerFinal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AnoFinal)
   End If
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)

   Call gs_SetFocus(cmb_TipReporte)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerInicial.Clear
   cmb_PerFinal.Clear
   ipp_AnoInicial.Text = Year(date)
   ipp_AnoFinal.Text = Year(date)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerInicial, 1, "033")
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerFinal, 1, "033")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMoneda, 1, "204")
   
   cmb_TipReporte.Clear
   cmb_TipReporte.AddItem "REPORTE MENSUAL POR PAGAR COFIDE"
   cmb_TipReporte.AddItem "REPORTE MENSUAL CUENTAS POR COBRAR"
   cmb_TipReporte.AddItem "REPORTE DIARIO CUENTAS POR COBRAR"
   cmb_TipReporte.ListIndex = 0
End Sub

Private Sub ipp_AnoFinal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub

Private Sub cmb_PerInicial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AnoInicial)
   End If
End Sub

Private Sub ipp_AnoInicial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerFinal)
   End If
End Sub

Private Sub ipp_FecInicial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFinal)
   End If
End Sub

Private Sub ipp_FecFinal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub

Private Sub cmd_ExpExc_Click()
Dim r_str_Cadena   As String
   
   l_str_FecIni = ""
   r_str_Cadena = ""
   l_str_FecFin = ""
   l_int_TipRpt = 0
   l_int_TipMon = 0
   l_int_TipPrd = 0
   l_str_RptNom = ""
   
   If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
       If (cmb_PerInicial.ListIndex = -1) Then
           MsgBox "Debe seleccionar el periodo inicial.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_PerInicial)
           Exit Sub
       End If
       If (ipp_AnoInicial.Value = 0) Then
           MsgBox "Debe ingresar el año inicial.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_AnoInicial)
           Exit Sub
       End If
       If (cmb_PerFinal.ListIndex = -1) Then
           MsgBox "Debe seleccionar el periodo final.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_PerFinal)
           Exit Sub
       End If
       If (ipp_AnoFinal.Value = 0) Then
           MsgBox "Debe ingresar el año final.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_AnoFinal)
           Exit Sub
       End If
   Else
       If Not IsDate(ipp_FecInicial) Then
          MsgBox "La fecha inicial ingresada no es valida.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(ipp_FecInicial)
          Exit Sub
       End If
       If Not IsDate(ipp_FecFinal) Then
          MsgBox "La fecha final ingresada no es valida.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(ipp_FecFinal)
          Exit Sub
       End If
   End If
   
   If (cmb_TipReporte.ListIndex = 0 Or cmb_TipReporte.ListIndex = 1) Then
       l_str_FecIni = ipp_AnoInicial.Text & Format(cmb_PerInicial.ItemData(cmb_PerInicial.ListIndex), "00") & "01"
       r_str_Cadena = "01/" & Format(cmb_PerFinal.ItemData(cmb_PerFinal.ListIndex), "00") & "/" & ipp_AnoFinal.Text
       r_str_Cadena = Format(DateAdd("m", 1, CDate(r_str_Cadena)), "MM/YYYY")
       l_str_FecFin = Format(DateAdd("d", -1, CDate("01/" & r_str_Cadena)), "YYYYMMDD")
       
       If (CDbl(l_str_FecIni) > CDbl(l_str_FecFin)) Then
           MsgBox "El periodo inicial no puede ser mayor al periodo final.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_PerInicial)
           Exit Sub
       End If
   ElseIf (cmb_TipReporte.ListIndex = 2) Then
      l_str_FecIni = Format(CDate(ipp_FecInicial.Text), "YYYYMMDD")
      l_str_FecFin = Format(CDate(ipp_FecFinal.Text), "YYYYMMDD")
       If (CDbl(l_str_FecIni) > CDbl(l_str_FecFin)) Then
           MsgBox "La fecha inicial no puede ser mayor al fecha final.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FecInicial)
           Exit Sub
       End If
   End If
   
   If (chk_Moneda.Value = 1) Then
       l_int_TipMon = 0
   Else
       If (cmb_TipMoneda.ListIndex = -1) Then
           MsgBox "Seleccione un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_TipMoneda)
           Exit Sub
       End If
       l_int_TipMon = IIf(cmb_TipMoneda.ListIndex = 0, 1, 2)
   End If
   
   If (chk_Product.Value = 1) Then
       l_int_TipPrd = 0
   Else
       If (cmd_TipProducto.ListIndex = -1) Then
           MsgBox "Seleccione un tipo de producto.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmd_TipProducto)
           Exit Sub
       End If
       If (UCase(Trim(cmd_TipProducto.Text)) = "CME") Then  'CME
           l_int_TipPrd = 1
       ElseIf (UCase(Trim(cmd_TipProducto.Text)) = "CRC-PBP") Then   'CRC
           l_int_TipPrd = 2
       ElseIf (UCase(Trim(cmd_TipProducto.Text)) = "MICASITA") Then   'MICASITA
           l_int_TipPrd = 3
       ElseIf (UCase(Trim(cmd_TipProducto.Text)) = "MIVIVIENDA") Then   'MIVIVIENDA
           l_int_TipPrd = 4
       End If
   End If
            
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
            
   Screen.MousePointer = 11
   If (cmb_TipReporte.ListIndex = 0) Then
       If (l_int_TipMon = 2) Then
           Screen.MousePointer = 0
           Exit Sub
       End If
       If (l_int_TipPrd = 2 Or l_int_TipPrd = 3) Then
           Screen.MousePointer = 0
           Exit Sub
       End If
       l_str_RptNom = "REPORTE_PAGARCOFIDE"
       l_int_TipRpt = 0
       Call fs_GenExc_CtaxPagar
   ElseIf (cmb_TipReporte.ListIndex = 1) Then
       l_str_RptNom = "REPORTE_CTAXCOBRAR01"
       l_int_TipRpt = 1
       Call fs_GenExc_CtaxCobrar_01
   ElseIf (cmb_TipReporte.ListIndex = 2) Then
       l_str_RptNom = "REPORTE_CTAXCOBRAR02"
       l_int_TipRpt = 2
       Call fs_GenExc_CtaxCobrar_02
   End If
   Screen.MousePointer = 0
   
End Sub

Private Sub fs_GenExc_CtaxPagar()
Dim r_obj_Excel      As Excel.Application
Dim r_dbl_FilExl     As Double
Dim r_str_Fec01     As String
Dim r_str_Fec02     As String

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " USP_RPT_FLUJO_CTAXPAGAR( "
    g_str_Parame = g_str_Parame & l_str_FecIni & ", "
    g_str_Parame = g_str_Parame & l_str_FecFin & ", "
    g_str_Parame = g_str_Parame & l_int_TipPrd & ") "

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

    r_dbl_FilExl = 3
    With r_obj_Excel.Sheets(1)
       .Name = "SOLES"
       .Cells(r_dbl_FilExl, 1) = "ITEM"
       .Cells(r_dbl_FilExl, 2) = "MES"
       .Cells(r_dbl_FilExl, 3) = "AÑO"
       .Cells(r_dbl_FilExl, 4) = "CAPITAL"
       .Cells(r_dbl_FilExl, 5) = "INTERES"
       .Cells(r_dbl_FilExl, 6) = "COMISION"
       .Cells(r_dbl_FilExl, 7) = "SALDO"
       
       'ITEM
       .Columns("A").ColumnWidth = 9
       .Columns("A").HorizontalAlignment = xlHAlignCenter
       'MES
       .Columns("B").ColumnWidth = 9
       .Columns("B").HorizontalAlignment = xlHAlignCenter
       'AÑO
       .Columns("C").ColumnWidth = 9
       .Columns("C").HorizontalAlignment = xlHAlignCenter
       'CAPITAL
       .Columns("D").ColumnWidth = 16
       .Columns("D").NumberFormat = "###,###,##0.00"
       .Columns("D").HorizontalAlignment = xlHAlignRight
       'INTERES
       .Columns("E").ColumnWidth = 16
       .Columns("E").NumberFormat = "###,###,##0.00"
       .Columns("E").HorizontalAlignment = xlHAlignRight
       'COMISION
       .Columns("F").ColumnWidth = 16
       .Columns("F").NumberFormat = "###,###,##0.00"
       .Columns("F").HorizontalAlignment = xlHAlignRight
       'SALDO
       .Columns("G").ColumnWidth = 16
       .Columns("G").NumberFormat = "###,###,##0.00"
       .Columns("G").HorizontalAlignment = xlHAlignRight
       
       .Range(.Cells(1, 3), .Cells(1, 6)).Merge
       .Range(.Cells(2, 3), .Cells(2, 6)).Merge
       If (chk_Product.Value = 1) Then
           .Cells(1, 3) = "REPORTE MENSUAL POR PAGAR COFIDE"
       Else
           .Cells(1, 3) = "REPORTE MENSUAL POR PAGAR COFIDE - " & Trim(cmd_TipProducto.Text)
       End If
       r_str_Fec01 = ""
       r_str_Fec01 = Mid(l_str_FecIni, 7, 2) & "/" & Mid(l_str_FecIni, 5, 2) & "/" & Mid(l_str_FecIni, 1, 4)
       r_str_Fec01 = Format(CDate(r_str_Fec01), "mmmm yyyy")
       r_str_Fec02 = ""
       r_str_Fec02 = Mid(l_str_FecFin, 7, 2) & "/" & Mid(l_str_FecFin, 5, 2) & "/" & Mid(l_str_FecFin, 1, 4)
       r_str_Fec02 = Format(CDate(r_str_Fec02), "mmmm yyyy")
       .Cells(2, 3) = " De " & r_str_Fec01 _
                      & " a " & r_str_Fec02
       .Range(.Cells(1, 1), .Cells(3, 7)).Font.Bold = True
       .Range(.Cells(1, 1), .Cells(3, 7)).HorizontalAlignment = xlHAlignCenter
       
       g_rst_Princi.MoveFirst
       r_dbl_FilExl = 4
       Do While Not g_rst_Princi.EOF
          .Cells(r_dbl_FilExl, 1) = r_dbl_FilExl - 3
          .Cells(r_dbl_FilExl, 2) = g_rst_Princi!Mes
          .Cells(r_dbl_FilExl, 3) = g_rst_Princi!ANIO
          .Cells(r_dbl_FilExl, 4) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
          .Cells(r_dbl_FilExl, 5) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
          .Cells(r_dbl_FilExl, 6) = Format(g_rst_Princi!COMISION, "###,###,##0.00")
          .Cells(r_dbl_FilExl, 7) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
          
          r_dbl_FilExl = r_dbl_FilExl + 1
          g_rst_Princi.MoveNext
       Loop
       
    End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_CtaxCobrar_01()
Dim r_dbl_FilExl     As Double
Dim r_dbl_FilSol     As Double
Dim r_dbl_FilDol     As Double
Dim r_str_AuxAno     As String
Dim r_str_AuxMes     As String
Dim r_int_NumHoj     As Integer
Dim r_int_Contar     As Integer
                  
    r_str_AuxAno = Year(date)
    r_str_AuxMes = Month(date)
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " USP_RPT_FLUJO_CTAXCOBRAR ( "
    g_str_Parame = g_str_Parame & l_str_FecIni & ", "
    g_str_Parame = g_str_Parame & l_str_FecFin & ", "
    g_str_Parame = g_str_Parame & l_int_TipRpt & ", " 'TIPO REPORTE
    g_str_Parame = g_str_Parame & l_int_TipMon & ", " 'MONEDA
    g_str_Parame = g_str_Parame & l_int_TipPrd & ", " 'CODIGO PRODUCTO
    g_str_Parame = g_str_Parame & r_str_AuxAno & ", " 'ANO
    g_str_Parame = g_str_Parame & r_str_AuxMes & ", " 'MES
    g_str_Parame = g_str_Parame & "'" & l_str_RptNom & "', " 'NOMBRE REPORTE
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Exit Sub
    End If
    
    Set l_obj_Excel = New Excel.Application
            
    If (l_int_TipMon = 0) Then
        l_obj_Excel.SheetsInNewWorkbook = 2
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Men("REPORTE MENSUAL CUENTAS POR COBRAR SOLES", "SOLES", 1)
        Call fs_GenExc_Hoja_Men("REPORTE MENSUAL CUENTAS POR COBRAR DOLARES", "DOLARES", 2)
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 1) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 4) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           ElseIf (g_rst_Princi!Moneda = 2) Then
               With l_obj_Excel.Sheets(2)
                    .Cells(r_dbl_FilSol, 1) = r_dbl_FilSol - 3
                    .Cells(r_dbl_FilSol, 2) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilSol, 3) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilSol, 4) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 5) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 6) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 7) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 8) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 9) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilSol = r_dbl_FilSol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    ElseIf (l_int_TipMon = 1) Then
        l_obj_Excel.SheetsInNewWorkbook = 1
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Men("REPORTE MENSUAL CUENTAS POR COBRAR SOLES", "SOLES", 1)
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 1) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 4) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    ElseIf (l_int_TipMon = 2) Then
        l_obj_Excel.SheetsInNewWorkbook = 1
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Men("REPORTE MENSUAL CUENTAS POR COBRAR DOLARES", "DOLARES", 1)
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 2) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 4) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
   
   l_obj_Excel.Visible = True
   Set l_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Hoja_Men(p_Titulo As String, p_Hoja As String, p_NumHoj As Integer)
Dim r_dbl_FilExl As Double
Dim r_str_Fec01  As String
Dim r_str_Fec02  As String

    r_dbl_FilExl = 3
    With l_obj_Excel.Sheets(p_NumHoj)
       .Name = p_Hoja
       .Cells(r_dbl_FilExl, 1) = "ITEM"
       .Cells(r_dbl_FilExl, 2) = "MES"
       .Cells(r_dbl_FilExl, 3) = "AÑO"
       .Cells(r_dbl_FilExl, 4) = "CAPITAL"
       .Cells(r_dbl_FilExl, 5) = "INTERES"
       .Cells(r_dbl_FilExl, 6) = "SEG. DESG"
       .Cells(r_dbl_FilExl, 7) = "SEG. INM."
       .Cells(r_dbl_FilExl, 8) = "PORTES"
       .Cells(r_dbl_FilExl, 9) = "SALDO CAPITAL"
       
       'ITEM
       .Columns("A").ColumnWidth = 9
       .Columns("A").HorizontalAlignment = xlHAlignCenter
       'MES
       .Columns("B").ColumnWidth = 9
       .Columns("B").HorizontalAlignment = xlHAlignCenter
       'AÑO
       .Columns("C").ColumnWidth = 9
       .Columns("C").HorizontalAlignment = xlHAlignCenter
       'CAPITAL
       .Columns("D").ColumnWidth = 16
       .Columns("D").NumberFormat = "###,###,##0.00"
       .Columns("D").HorizontalAlignment = xlHAlignRight
       'INTERES
       .Columns("E").ColumnWidth = 16
       .Columns("E").NumberFormat = "###,###,##0.00"
       .Columns("E").HorizontalAlignment = xlHAlignRight
       'SEG. DESG.
       .Columns("F").ColumnWidth = 16
       .Columns("F").NumberFormat = "###,###,##0.00"
       .Columns("F").HorizontalAlignment = xlHAlignRight
       'SEG. INM
       .Columns("G").ColumnWidth = 16
       .Columns("G").NumberFormat = "###,###,##0.00"
       .Columns("G").HorizontalAlignment = xlHAlignRight
       'PORTES
       .Columns("H").ColumnWidth = 16
       .Columns("H").NumberFormat = "###,###,##0.00"
       .Columns("H").HorizontalAlignment = xlHAlignRight
       'SALDO CAPITAL
       .Columns("I").ColumnWidth = 16
       .Columns("I").NumberFormat = "###,###,##0.00"
       .Columns("I").HorizontalAlignment = xlHAlignRight
       
       .Range(.Cells(1, 4), .Cells(1, 8)).Merge
       .Range(.Cells(2, 4), .Cells(2, 8)).Merge
       If (chk_Product.Value = 1) Then
           .Cells(1, 4) = p_Titulo
       Else
           .Cells(1, 4) = p_Titulo & " - " & Trim(cmd_TipProducto.Text)
       End If
       r_str_Fec01 = ""
       r_str_Fec01 = Mid(l_str_FecIni, 7, 2) & "/" & Mid(l_str_FecIni, 5, 2) & "/" & Mid(l_str_FecIni, 1, 4)
       r_str_Fec01 = Format(CDate(r_str_Fec01), "mmmm yyyy")
       r_str_Fec02 = ""
       r_str_Fec02 = Mid(l_str_FecFin, 7, 2) & "/" & Mid(l_str_FecFin, 5, 2) & "/" & Mid(l_str_FecFin, 1, 4)
       r_str_Fec02 = Format(CDate(r_str_Fec02), "mmmm yyyy")
       .Cells(2, 4) = " De " & r_str_Fec01 _
                      & " a " & r_str_Fec02
       .Range(.Cells(1, 1), .Cells(3, 9)).Font.Bold = True
       .Range(.Cells(1, 1), .Cells(3, 9)).HorizontalAlignment = xlHAlignCenter
    End With
End Sub

Private Sub fs_GenExc_Hoja_Dia(p_Titulo As String, p_Hoja As String, p_NumHoj As Integer)
Dim r_dbl_FilExl As Double

    r_dbl_FilExl = 3
    With l_obj_Excel.Sheets(p_NumHoj)
       .Name = p_Hoja
       .Cells(r_dbl_FilExl, 1) = "ITEM"
       .Cells(r_dbl_FilExl, 2) = "DIA"
       .Cells(r_dbl_FilExl, 3) = "MES"
       .Cells(r_dbl_FilExl, 4) = "AÑO"
       .Cells(r_dbl_FilExl, 5) = "CAPITAL"
       .Cells(r_dbl_FilExl, 6) = "INTERES"
       .Cells(r_dbl_FilExl, 7) = "SEG. DESG"
       .Cells(r_dbl_FilExl, 8) = "SEG. INM."
       .Cells(r_dbl_FilExl, 9) = "PORTES"
       .Cells(r_dbl_FilExl, 10) = "SALDO CAPITAL"
       
       'ITEM
       .Columns("A").ColumnWidth = 9
       .Columns("A").HorizontalAlignment = xlHAlignCenter
       'DIA
       .Columns("B").ColumnWidth = 9
       .Columns("B").HorizontalAlignment = xlHAlignCenter
       'MES
       .Columns("C").ColumnWidth = 9
       .Columns("C").HorizontalAlignment = xlHAlignCenter
       'AÑO
       .Columns("D").ColumnWidth = 9
       .Columns("D").HorizontalAlignment = xlHAlignCenter
       'CAPITAL
       .Columns("E").ColumnWidth = 16
       .Columns("E").NumberFormat = "###,###,##0.00"
       .Columns("E").HorizontalAlignment = xlHAlignRight
       'INTERES
       .Columns("F").ColumnWidth = 16
       .Columns("F").NumberFormat = "###,###,##0.00"
       .Columns("F").HorizontalAlignment = xlHAlignRight
       'SEG. DESG.
       .Columns("G").ColumnWidth = 16
       .Columns("G").NumberFormat = "###,###,##0.00"
       .Columns("G").HorizontalAlignment = xlHAlignRight
       'SEG. INM
       .Columns("H").ColumnWidth = 16
       .Columns("H").NumberFormat = "###,###,##0.00"
       .Columns("H").HorizontalAlignment = xlHAlignRight
       'PORTES
       .Columns("I").ColumnWidth = 16
       .Columns("I").NumberFormat = "###,###,##0.00"
       .Columns("I").HorizontalAlignment = xlHAlignRight
       'SALDO CAPITAL
       .Columns("J").ColumnWidth = 16
       .Columns("J").NumberFormat = "###,###,##0.00"
       .Columns("J").HorizontalAlignment = xlHAlignRight
                          
       .Range(.Cells(1, 5), .Cells(1, 8)).Merge
       .Range(.Cells(2, 5), .Cells(2, 8)).Merge
       If (chk_Product.Value = 1) Then
           .Cells(1, 5) = p_Titulo
       Else
           .Cells(1, 5) = p_Titulo & " - " & Trim(cmd_TipProducto.Text)
       End If
       .Cells(2, 5) = " Del " & Mid(l_str_FecIni, 7, 2) & "/" & Mid(l_str_FecIni, 5, 2) & "/" & Mid(l_str_FecIni, 1, 4) _
                      & " al " & Mid(l_str_FecFin, 7, 2) & "/" & Mid(l_str_FecFin, 5, 2) & "/" & Mid(l_str_FecFin, 1, 4)
       .Range(.Cells(1, 1), .Cells(3, 10)).Font.Bold = True
       .Range(.Cells(1, 1), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
    End With
End Sub


Private Sub fs_GenExc_CtaxCobrar_02()
Dim r_obj_Excel      As Excel.Application
Dim r_dbl_FilExl     As Double
Dim r_dbl_FilSol     As Double
Dim r_dbl_FilDol     As Double
Dim r_str_AuxAno     As String
Dim r_str_AuxMes     As String
                  
    r_str_AuxAno = Year(date)
    r_str_AuxMes = Month(date)
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " USP_RPT_FLUJO_CTAXCOBRAR ( "
    g_str_Parame = g_str_Parame & l_str_FecIni & ", "
    g_str_Parame = g_str_Parame & l_str_FecFin & ", "
    g_str_Parame = g_str_Parame & l_int_TipRpt & ", " 'TIPO REPORTE
    g_str_Parame = g_str_Parame & l_int_TipMon & ", " 'MONEDA
    g_str_Parame = g_str_Parame & l_int_TipPrd & ", " 'CODIGO PRODUCTO
    g_str_Parame = g_str_Parame & r_str_AuxAno & ", " 'ANO
    g_str_Parame = g_str_Parame & r_str_AuxMes & ", " 'MES
    g_str_Parame = g_str_Parame & "'" & l_str_RptNom & "', " 'NOMBRE REPORTE
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Exit Sub
    End If
    
    Set l_obj_Excel = New Excel.Application
    
    If (l_int_TipMon = 0) Then
        l_obj_Excel.SheetsInNewWorkbook = 2
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Dia("REPORTE DIARIO CUENTAS POR COBRAR SOLES", "SOLES", 1)
        Call fs_GenExc_Hoja_Dia("REPORTE DIARIO CUENTAS POR COBRAR DOLARES", "DOLARES", 2)
        
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 1) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!DIA
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 4) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 10) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           ElseIf (g_rst_Princi!Moneda = 2) Then
               With l_obj_Excel.Sheets(2)
                    .Cells(r_dbl_FilSol, 1) = r_dbl_FilSol - 3
                    .Cells(r_dbl_FilSol, 2) = g_rst_Princi!DIA
                    .Cells(r_dbl_FilSol, 3) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilSol, 4) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilSol, 5) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 6) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 7) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 8) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 9) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilSol, 10) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilSol = r_dbl_FilSol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    ElseIf (l_int_TipMon = 1) Then
        l_obj_Excel.SheetsInNewWorkbook = 1
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Dia("REPORTE DIARIO CUENTAS POR COBRAR SOLES", "SOLES", 1)
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 1) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!DIA
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 4) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 10) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    ElseIf (l_int_TipMon = 2) Then
        l_obj_Excel.SheetsInNewWorkbook = 1
        l_obj_Excel.Workbooks.Add
        Call fs_GenExc_Hoja_Dia("REPORTE DIARIO CUENTAS POR COBRAR DOLARES", "DOLARES", 1)
        g_rst_Princi.MoveFirst
        r_dbl_FilSol = 4
        r_dbl_FilDol = 4
        Do While Not g_rst_Princi.EOF
           If (g_rst_Princi!Moneda = 2) Then
               With l_obj_Excel.Sheets(1)
                    .Cells(r_dbl_FilDol, 1) = r_dbl_FilDol - 3
                    .Cells(r_dbl_FilDol, 2) = g_rst_Princi!DIA
                    .Cells(r_dbl_FilDol, 3) = g_rst_Princi!Mes
                    .Cells(r_dbl_FilDol, 4) = g_rst_Princi!ANIO
                    .Cells(r_dbl_FilDol, 5) = Format(g_rst_Princi!CAPITAL, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 6) = Format(g_rst_Princi!INTERES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 7) = Format(g_rst_Princi!SEG_DESG, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 8) = Format(g_rst_Princi!SEG_INM, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 9) = Format(g_rst_Princi!PORTES, "###,###,##0.00")
                    .Cells(r_dbl_FilDol, 10) = Format(g_rst_Princi!SALDO_CAPITAL, "###,###,##0.00")
                    r_dbl_FilDol = r_dbl_FilDol + 1
               End With
           End If
           g_rst_Princi.MoveNext
        Loop
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
   
   l_obj_Excel.Visible = True
   Set l_obj_Excel = Nothing
End Sub

