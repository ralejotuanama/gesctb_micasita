VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_35 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   330
   ClientTop       =   2370
   ClientWidth     =   15120
   Icon            =   "GesCtb_frm_227.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8730
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   15399
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   17
         Top             =   810
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
            Left            =   1230
            Picture         =   "GesCtb_frm_227.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14430
            Picture         =   "GesCtb_frm_227.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_227.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_227.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   8
         Top             =   1500
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_SucAge 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   13425
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   390
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            Left            =   1560
            TabIndex        =   2
            Top             =   720
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label Label3 
            Caption         =   "Cuenta:"
            Height          =   225
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fin:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   690
            Width           =   1365
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   1244
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
            Height          =   585
            Left            =   600
            TabIndex        =   11
            Top             =   60
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Reporte de Conciliación"
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
            Left            =   14190
            Top             =   150
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
            Left            =   0
            Picture         =   "GesCtb_frm_227.frx":0D6C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6090
         Left            =   30
         TabIndex        =   12
         Top             =   2640
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   10751
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5675
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   14925
            _ExtentX        =   26326
            _ExtentY        =   10001
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumMov 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim."
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
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   5925
            TabIndex        =   14
            Top             =   60
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel pnl_Tit_FecMov 
            Height          =   285
            Left            =   1060
            TabIndex        =   15
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Movim."
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
         Begin Threed.SSPanel pnl_Tit_Glosa 
            Height          =   285
            Left            =   2250
            TabIndex        =   20
            Top             =   60
            Width           =   3670
            _Version        =   65536
            _ExtentX        =   6473
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
         Begin Threed.SSPanel pnl_Tit_TipoDoc 
            Height          =   285
            Left            =   12395
            TabIndex        =   22
            Top             =   60
            Width           =   895
            _Version        =   65536
            _ExtentX        =   1579
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Doc."
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
         Begin Threed.SSPanel pnl_Tit_NumDoc 
            Height          =   285
            Left            =   13290
            TabIndex        =   23
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Doc."
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
         Begin Threed.SSPanel pnl_Tit_NumOpAso 
            Height          =   285
            Left            =   8850
            TabIndex        =   24
            Top             =   60
            Width           =   3545
            _Version        =   65536
            _ExtentX        =   6253
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación Asociada"
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
         Begin Threed.SSPanel pnl_Tit_tipmov 
            Height          =   285
            Left            =   7515
            TabIndex        =   25
            Top             =   60
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Movimiento"
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
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   285
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "Importe"
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
Attribute VB_Name = "frm_RptCtb_35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_SucAge()      As moddat_tpo_Genera
Dim l_str_Existe        As String
Dim l_int_MsjErr        As Integer

Private Sub cmd_Buscar_Click()
   If cmb_SucAge.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SucAge)
      Exit Sub
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_SucAge)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   Dim cont As Integer
   Dim r_int_NroFil     As Integer
   Dim r_obj_Excel      As Excel.Application
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe ninguna operación financiera.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   l_int_MsjErr = 0
   Screen.MousePointer = 11
   cont = 0
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
   
   'IMAGEN
      On Local Error Resume Next
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("B").NumberFormat = "@"
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,##0.00"
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      'Unir celdas
      .Range("A" & r_int_NroFil & ":H" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & "") = "Conciliación Cuenta " & cmb_SucAge.Text
      .Range("A" & r_int_NroFil & "").Font.Underline = True
      .Range("A" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
         
      r_int_NroFil = r_int_NroFil + 1
      .Range("A" & r_int_NroFil) = "Fecha Inicio: " & Format(CDate(ipp_FecIni.Text), "DD/MM/YYYY")
      .Range("A" & r_int_NroFil).HorizontalAlignment = xlHAlignLeft
      
      r_int_NroFil = r_int_NroFil + 1
      
      .Range("A" & r_int_NroFil) = "Fecha Final: " & Format(CDate(ipp_FecFin.Text), "DD/MM/YYYY")
      .Range("A" & r_int_NroFil).HorizontalAlignment = xlHAlignLeft
      
      r_int_NroFil = r_int_NroFil + 2
      
      .Cells(r_int_NroFil, 1) = "NRO. MOVIM.":               .Columns("A").ColumnWidth = 13
      .Cells(r_int_NroFil, 2) = "F. MOVIM.":                 .Columns("B").ColumnWidth = 10
      .Cells(r_int_NroFil, 3) = "GLOSA":                     .Columns("C").ColumnWidth = 40
      .Cells(r_int_NroFil, 4) = "IMPORTE":                   .Columns("D").ColumnWidth = 12
      .Cells(r_int_NroFil, 5) = "TIPO MOVIMIENTO":           .Columns("E").ColumnWidth = 18
      .Cells(r_int_NroFil, 6) = "NRO. OPERACION ASOCIADA":   .Columns("F").ColumnWidth = 25
      .Cells(r_int_NroFil, 7) = "TIPO DOC.":                 .Columns("G").ColumnWidth = 13
      .Cells(r_int_NroFil, 8) = "NRO. DOC.":                 .Columns("H").ColumnWidth = 9
            
      .Range("A" & r_int_NroFil & ":H" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":H" & r_int_NroFil & "").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         
          Do While cont < grd_Listad.Rows
            grd_Listad.Row = cont
            grd_Listad.Col = 0
            .Cells(r_int_NroFil, 1) = grd_Listad.Text
            grd_Listad.Col = 1
            .Cells(r_int_NroFil, 2) = grd_Listad.Text
            grd_Listad.Col = 2
            .Cells(r_int_NroFil, 3) = grd_Listad.Text
            grd_Listad.Col = 3
            .Cells(r_int_NroFil, 4) = grd_Listad.Text
            grd_Listad.Col = 4
            .Cells(r_int_NroFil, 5) = grd_Listad.Text
            grd_Listad.Col = 5
            .Cells(r_int_NroFil, 6) = grd_Listad.Text
            grd_Listad.Col = 6
            .Cells(r_int_NroFil, 7) = grd_Listad.Text
            grd_Listad.Col = 7
            .Cells(r_int_NroFil, 8) = grd_Listad.Text
            r_int_NroFil = r_int_NroFil + 1
         
            cont = cont + 1
         Loop
      End If
      
      .Cells(1, 1).Select
   End With
     
   g_rst_Princi.Close
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
   Screen.MousePointer = 0
   'If l_int_MsjErr > 0 Then
   '   MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
   'End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1025
   grd_Listad.ColWidth(1) = 1175
   grd_Listad.ColWidth(2) = 3665
   grd_Listad.ColWidth(3) = 1575
   grd_Listad.ColWidth(4) = 1355
   grd_Listad.ColWidth(5) = 3565
   grd_Listad.ColWidth(6) = 885
   grd_Listad.ColWidth(7) = 1025
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   
   moddat_g_str_Codigo = "000001"
   Call moddat_gs_Carga_LisIte(cmb_SucAge, l_arr_SucAge, 1, "538")
   cmb_SucAge.ListIndex = -1
End Sub

Private Sub fs_Limpia()
   cmb_SucAge.ListIndex = -1
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_SucAge.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
End Sub

Private Sub cmb_SucAge_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_SucAge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SucAge_Click
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub fs_Buscar()
   Dim ctaBanco As String
   ctaBanco = Left(cmb_SucAge.Text, 12)
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CONBAN_NUMOPE AS CONBAN_NUMOPE ,A.CONBAN_FECMOV AS CONBAN_FECMOV ,A.CONBAN_CONCEP AS CONBAN_CONCEP ,A.CONBAN_IMPORT AS CONBAN_IMPORT ,B.CAJMOV_TIPMOV AS CAJMOV_TIPMOV ,B.CAJMOV_NUMOPE AS CAJMOV_NUMOPE ,B.CAJMOV_TIPDOC AS CAJMOV_TIPDOC ,B.CAJMOV_NUMDOC AS CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & " FROM CTB_CONBAN A LEFT JOIN OPE_CAJMOV B ON A.CONBAN_NUMOPE = RTRIM(B.CAJMOV_NUMCOM) "
   g_str_Parame = g_str_Parame & " WHERE A.CONBAN_CTABCO = '" & ctaBanco & "' "
   g_str_Parame = g_str_Parame & "   AND A.CONBAN_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND A.CONBAN_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & "   AND A.CONBAN_IMPORT > 0 "
   g_str_Parame = g_str_Parame & " ORDER BY A.CONBAN_NUMOPE ASC, A.CONBAN_FECMOV ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!CONBAN_NUMOPE)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!CONBAN_FECMOV))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!CONBAN_CONCEP)
      
      grd_Listad.Col = 3
      grd_Listad.Text = FormatNumber(CDbl(g_rst_Princi!CONBAN_IMPORT), 2)
      
      If (IsNull(g_rst_Princi!CAJMOV_TIPMOV)) Then
        grd_Listad.Col = 4
        grd_Listad.Text = ""
      Else
        grd_Listad.Col = 4
        grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPMOV)
      End If
      
      If (IsNull(g_rst_Princi!CAJMOV_NUMOPE)) Then
        grd_Listad.Col = 5
        grd_Listad.Text = ""
      Else
        grd_Listad.Col = 5
        grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_NUMOPE)
      End If
        
      If (IsNull(g_rst_Princi!CAJMOV_TIPDOC)) Then
        grd_Listad.Col = 6
        grd_Listad.Text = ""
      Else
        grd_Listad.Col = 6
        grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPDOC)
      End If
        
      If (IsNull(g_rst_Princi!CAJMOV_NUMDOC)) Then
        grd_Listad.Col = 7
        grd_Listad.Text = ""
      Else
        grd_Listad.Col = 7
        grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_NUMDOC)
      End If
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_FecMov_Click()
   If Len(Trim(pnl_Tit_FecMov.Tag)) = 0 Or pnl_Tit_FecMov.Tag = "D" Then
      pnl_Tit_FecMov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_FecMov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "N-")
   End If
End Sub

Private Sub pnl_Tit_Glosa_Click()
   If Len(Trim(pnl_Tit_Glosa.Tag)) = 0 Or pnl_Tit_Glosa.Tag = "D" Then
      pnl_Tit_Glosa.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_Glosa.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "N-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "N-")
   End If
End Sub

Private Sub pnl_Tit_NumDoc_Click()
   If Len(Trim(pnl_Tit_NumDoc.Tag)) = 0 Or pnl_Tit_NumDoc.Tag = "D" Then
      pnl_Tit_NumDoc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "C")
   Else
      pnl_Tit_NumDoc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "C-")
   End If
End Sub

Private Sub pnl_Tit_TipoDoc_Click()
   If Len(Trim(pnl_Tit_TipoDoc.Tag)) = 0 Or pnl_Tit_TipoDoc.Tag = "D" Then
      pnl_Tit_TipoDoc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_TipoDoc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumMov_Click()
   If Len(Trim(pnl_Tit_NumMov.Tag)) = 0 Or pnl_Tit_NumMov.Tag = "D" Then
      pnl_Tit_NumMov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumMov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpAso_Click()
   If Len(Trim(pnl_Tit_NumOpAso.Tag)) = 0 Or pnl_Tit_NumOpAso.Tag = "D" Then
      pnl_Tit_NumOpAso.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_NumOpAso.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_TipMov_Click()
   If Len(Trim(pnl_Tit_tipmov.Tag)) = 0 Or pnl_Tit_tipmov.Tag = "D" Then
      pnl_Tit_tipmov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_tipmov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub
