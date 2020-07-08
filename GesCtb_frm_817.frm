VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   1620
   ClientTop       =   1620
   ClientWidth     =   14940
   Icon            =   "GesCtb_frm_817.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9765
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14955
      _Version        =   65536
      _ExtentX        =   26379
      _ExtentY        =   17224
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   14850
         _Version        =   65536
         _ExtentX        =   26194
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   570
            TabIndex        =   12
            Top             =   180
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Estado de Ganancias y Pérdidas"
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
            Picture         =   "GesCtb_frm_817.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   14850
         _Version        =   65536
         _ExtentX        =   26194
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
         Begin VB.CommandButton cmd_ExpExcDet 
            Height          =   585
            Left            =   1245
            Picture         =   "GesCtb_frm_817.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel - Detallado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_817.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   645
            Picture         =   "GesCtb_frm_817.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel - Resumido"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14205
            Picture         =   "GesCtb_frm_817.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   60
         TabIndex        =   14
         Top             =   1470
         Width           =   14850
         _Version        =   65536
         _ExtentX        =   26194
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   9330
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   90
            Visible         =   0   'False
            Width           =   3795
         End
         Begin VB.ComboBox cmb_PerMesf 
            Height          =   315
            Left            =   5130
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   90
            Width           =   2500
         End
         Begin VB.ComboBox cmb_PerMesi 
            Height          =   315
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2500
         End
         Begin EditLib.fpLongInteger ipp_PerAnoi 
            Height          =   315
            Left            =   1290
            TabIndex        =   1
            Top             =   420
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
         Begin EditLib.fpLongInteger ipp_PerAnof 
            Height          =   315
            Left            =   5130
            TabIndex        =   3
            Top             =   420
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
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   9330
            TabIndex        =   9
            Top             =   420
            Visible         =   0   'False
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
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   8250
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   8250
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Año Final:"
            Height          =   195
            Left            =   4050
            TabIndex        =   18
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Periodo Final:"
            Height          =   255
            Left            =   4050
            TabIndex        =   17
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Periodo Inicial:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Año Inicial:"
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   480
            Width           =   780
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7365
         Left            =   60
         TabIndex        =   21
         Top             =   2340
         Width           =   14835
         _Version        =   65536
         _ExtentX        =   26167
         _ExtentY        =   12991
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEEFF 
            Height          =   7275
            Left            =   60
            TabIndex        =   22
            Top             =   45
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   12832
            _Version        =   393216
            Rows            =   21
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_int_PerMesi        As Integer
Dim r_int_PerAnoi        As Integer
Dim r_int_PerMesf        As Integer
Dim r_int_PerAnof        As Integer

Private Sub cmd_ExpExcDet_Click()
   
'   If cmb_PerMes.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_PerMes)
'      Exit Sub
'   End If
'   If ipp_PerAno.Text = "" Then
'      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_PerAno)
'      Exit Sub
'   End If
   
   If cmb_PerMesi.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If cmb_PerMesf.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesf)
      Exit Sub
   End If
   If cmb_PerMesi.ListIndex > cmb_PerMesf.ListIndex And ipp_PerAnoi = ipp_PerAnof Then
      MsgBox "El Mes inicial no puede ser mayor al Mes Final en el mismo año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If ipp_PerAnoi.Text = 0 Then
      MsgBox "Debe seleccionar el Año Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   If ipp_PerAnof.Text = 0 Then
      MsgBox "Debe seleccionar el Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnof)
      Exit Sub
   End If
   If ipp_PerAnoi.Text > ipp_PerAnof Then
      MsgBox "El Año inicial no puede ser mayor al Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcDet_NueVer
   'call fs_GenExcRes_AntVer
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExcRes_Click()
'   If cmb_PerMes.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_PerMes)
'      Exit Sub
'   End If
'   If ipp_PerAno.Text = "" Then
'      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_PerAno)
'      Exit Sub
'   End If

   If cmb_PerMesi.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If cmb_PerMesf.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesf)
      Exit Sub
   End If
   If cmb_PerMesi.ListIndex > cmb_PerMesf.ListIndex And ipp_PerAnoi = ipp_PerAnof Then
      MsgBox "El Mes inicial no puede ser mayor al Mes Final en el mismo año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If ipp_PerAnoi.Text = 0 Then
      MsgBox "Debe seleccionar el Año Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   If ipp_PerAnof.Text = 0 Then
      MsgBox "Debe seleccionar el Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnof)
      Exit Sub
   End If
   If ipp_PerAnoi.Text > ipp_PerAnof Then
      MsgBox "El Año inicial no puede ser mayor al Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes_NueVer
   'call fs_GenExcDet_AntVer
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Proces_Click()
  'Call fs_Proces_AntVer
   Call fs_Proces_NueVer
End Sub

Private Sub fs_Proces_AntVer()
Dim r_str_PerMes                        As String
Dim r_str_PerAno                        As String
Dim p, q                                As Integer

   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
        
   'If MsgBox("¿Está seguro que desea realizar el proceso ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
   '   Exit Sub
   'End If
   
   Screen.MousePointer = 11
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   grd_LisEEFF.Redraw = False
   Call gs_LimpiaGrid(grd_LisEEFF)
  ' Call fs_Recorset_nc_NueVer
   Call fs_Recorset_nc_AntVer
   
   'llama al SP de EEFF
   g_str_Parame = "USP_CUR_GEN_EEFF ("
   g_str_Parame = g_str_Parame & 12 & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) - 1 & ",1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "
    
   'EJECUTA CONSULTA
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEFF.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
    'ALMACENA RESULTADOS DEL AÑO ANTERIOR, EN UN RECORDSET NO CONECTADO
   Dim g_rst_Auxiliar As ADODB.Recordset
   g_str_Parame = "SELECT * FROM TT_EEFF"
   g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & " ORDER BY grupo, subgrp, item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Auxiliar, 3) Then
      MsgBox "Error al ejecutar la consulta para recorset no conectado.", vbCritical, modgen_g_str_NomPlt 'Exit Sub
   End If
   
   If Not (g_rst_Auxiliar.BOF And g_rst_Auxiliar.EOF) Then
      g_rst_Auxiliar.MoveFirst
      Do While Not g_rst_Auxiliar.EOF
         g_rst_GenAux.AddNew
         g_rst_GenAux.Fields(0).Value = g_rst_Auxiliar!GRUPO
         g_rst_GenAux.Fields(1).Value = g_rst_Auxiliar!NOMGRUPO
         g_rst_GenAux.Fields(2).Value = g_rst_Auxiliar!SUBGRP
         g_rst_GenAux.Fields(3).Value = g_rst_Auxiliar!NOMSUBGRP
         g_rst_GenAux.Fields(4).Value = g_rst_Auxiliar!CNTACTBLE
         g_rst_GenAux.Fields(5).Value = g_rst_Auxiliar!NOMCTA
         g_rst_GenAux.Fields(6).Value = g_rst_Auxiliar!MES01
         g_rst_GenAux.Fields(7).Value = g_rst_Auxiliar!MES02
         g_rst_GenAux.Fields(8).Value = g_rst_Auxiliar!MES03
         g_rst_GenAux.Fields(9).Value = g_rst_Auxiliar!MES04
         g_rst_GenAux.Fields(10).Value = g_rst_Auxiliar!MES05
         g_rst_GenAux.Fields(11).Value = g_rst_Auxiliar!MES06
         g_rst_GenAux.Fields(12).Value = g_rst_Auxiliar!MES07
         g_rst_GenAux.Fields(13).Value = g_rst_Auxiliar!MES08
         g_rst_GenAux.Fields(14).Value = g_rst_Auxiliar!MES09
         g_rst_GenAux.Fields(15).Value = g_rst_Auxiliar!MES10
         g_rst_GenAux.Fields(16).Value = g_rst_Auxiliar!MES11
         g_rst_GenAux.Fields(17).Value = g_rst_Auxiliar!MES12
         g_rst_GenAux.Fields(18).Value = g_rst_Auxiliar!ACUMU
         g_rst_GenAux.Fields(19).Value = g_rst_Auxiliar!INDTIPO
         g_rst_GenAux.Fields(20).Value = g_rst_Auxiliar!Item
         
         g_rst_GenAux.Update
         g_rst_Auxiliar.MoveNext
      Loop
   End If
   
   'CABECERA
   grd_LisEEFF.Rows = grd_LisEEFF.Rows + 2
   grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
   grd_LisEEFF.FixedRows = 1

   grd_LisEEFF.Row = 0
   grd_LisEEFF.Col = 3
   grd_LisEEFF.Text = "EJERCICIOS"
   
   If CInt(r_str_PerMes) = 12 Then
      GoTo SALTO1
      
S1:
      grd_LisEEFF.Col = q + 1:    grd_LisEEFF.Text = "FEB-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S2:
      grd_LisEEFF.Col = q + 2:    grd_LisEEFF.Text = "MAR-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S3:
      grd_LisEEFF.Col = q + 3:    grd_LisEEFF.Text = "ABR-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S4:
      grd_LisEEFF.Col = q + 4:    grd_LisEEFF.Text = "MAY-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S5:
      grd_LisEEFF.Col = q + 5:    grd_LisEEFF.Text = "JUN-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S6:
      grd_LisEEFF.Col = q + 6:    grd_LisEEFF.Text = "JUL-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S7:
      grd_LisEEFF.Col = q + 7:    grd_LisEEFF.Text = "AGO-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S8:
      grd_LisEEFF.Col = q + 8:    grd_LisEEFF.Text = "SET-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S9:
      grd_LisEEFF.Col = q + 9:    grd_LisEEFF.Text = "OCT-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S10:
      grd_LisEEFF.Col = q + 10:   grd_LisEEFF.Text = "NOV-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S11:
      grd_LisEEFF.Col = q + 11:   grd_LisEEFF.Text = "DIC-" & Right(CInt(r_str_PerAno) - 1, 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
  
   ElseIf CInt(r_str_PerMes) = 11 Then
      q = -7
      GoTo S11
   ElseIf CInt(r_str_PerMes) = 10 Then
      q = -6
      GoTo S10
   ElseIf CInt(r_str_PerMes) = 9 Then
      q = -5
      GoTo S9
   ElseIf CInt(r_str_PerMes) = 8 Then
      q = -4
      GoTo S8
   ElseIf CInt(r_str_PerMes) = 7 Then
      q = -3
      GoTo S7
   ElseIf CInt(r_str_PerMes) = 6 Then
      q = -2
      GoTo S6
   ElseIf CInt(r_str_PerMes) = 5 Then
      q = -1
      GoTo S5
   ElseIf CInt(r_str_PerMes) = 4 Then
      q = 0
      GoTo S4
   ElseIf CInt(r_str_PerMes) = 3 Then
      q = 1
      GoTo S3
   ElseIf CInt(r_str_PerMes) = 2 Then
      q = 2
      GoTo S2
   ElseIf CInt(r_str_PerMes) = 1 Then
      q = 3
      GoTo S1
   End If
   
SALTO1:
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         If Trim(g_rst_Princi!INDTIPO) <> "L" Then

            grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
            grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
            
            grd_LisEEFF.Col = 0
            grd_LisEEFF.Text = Trim(g_rst_Princi!GRUPO)
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
            grd_LisEEFF.Col = 1
            grd_LisEEFF.Text = Trim(g_rst_Princi!SUBGRP)
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
            grd_LisEEFF.Col = 2
            grd_LisEEFF.Text = Trim(g_rst_Princi!INDTIPO)
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
            If Trim(g_rst_Princi!INDTIPO) = "S" Then
               grd_LisEEFF.Col = 3
               grd_LisEEFF.Text = Space(5) & Trim(g_rst_Princi!NOMSUBGRP)
            ElseIf Trim(g_rst_Princi!INDTIPO) = "L" Then
               grd_LisEEFF.Col = 3
               grd_LisEEFF.Text = ""
            ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Then
               grd_LisEEFF.Col = 3
               grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
            Else
               grd_LisEEFF.Col = 3
               grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
            End If
            
            grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
            grd_LisEEFF.CellFontBold = True
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
            If CInt(r_str_PerMes) = 12 Then
               p = 4
               grd_LisEEFF.Col = p
               grd_LisEEFF.Text = g_rst_Princi!MES01 'Format(g_rst_Princi!MES01, "###,###,###,##0")
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S12:
               grd_LisEEFF.Col = p + 1
               grd_LisEEFF.Text = g_rst_Princi!MES02
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S13:
               grd_LisEEFF.Col = p + 2
               grd_LisEEFF.Text = g_rst_Princi!MES03
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S14:
               grd_LisEEFF.Col = p + 3
               grd_LisEEFF.Text = g_rst_Princi!MES04
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S15:
               grd_LisEEFF.Col = p + 4
               grd_LisEEFF.Text = g_rst_Princi!MES05
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S16:
               grd_LisEEFF.Col = p + 5
               grd_LisEEFF.Text = g_rst_Princi!MES06
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S17:
               grd_LisEEFF.Col = p + 6
               grd_LisEEFF.Text = g_rst_Princi!MES07
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S18:
               grd_LisEEFF.Col = p + 7
               grd_LisEEFF.Text = g_rst_Princi!MES08
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S19:
               grd_LisEEFF.Col = p + 8
               grd_LisEEFF.Text = g_rst_Princi!MES09
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S20:
               grd_LisEEFF.Col = p + 9
               grd_LisEEFF.Text = g_rst_Princi!MES10
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S21:
               grd_LisEEFF.Col = p + 10
               grd_LisEEFF.Text = g_rst_Princi!MES11
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
S22:
               grd_LisEEFF.Col = p + 11
               grd_LisEEFF.Text = g_rst_Princi!MES12
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               
            ElseIf CInt(r_str_PerMes) = 11 Then
               p = -7
               GoTo S22
            ElseIf CInt(r_str_PerMes) = 10 Then
               p = -6
               GoTo S21
            ElseIf CInt(r_str_PerMes) = 9 Then
               p = -5
               GoTo S20
            ElseIf CInt(r_str_PerMes) = 8 Then
               p = -4
               GoTo S19
            ElseIf CInt(r_str_PerMes) = 7 Then
               p = -3
               GoTo S18
            ElseIf CInt(r_str_PerMes) = 6 Then
               p = -2
               GoTo S17
            ElseIf CInt(r_str_PerMes) = 5 Then
               p = -1
               GoTo S16
            ElseIf CInt(r_str_PerMes) = 4 Then
               p = 0
               GoTo S15
            ElseIf CInt(r_str_PerMes) = 3 Then
               p = 1
               GoTo S14
            ElseIf CInt(r_str_PerMes) = 2 Then
               p = 2
               GoTo S13
            ElseIf CInt(r_str_PerMes) = 1 Then
               p = 3
               GoTo S12
            End If
            
'            grd_LisEEFF.Col = 16
'            grd_LisEEFF.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0")
'            grd_LisEEFF.CellFontName = "Arial"
'            grd_LisEEFF.CellFontSize = 8
        
            grd_LisEEFF.Col = 17
            grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
            grd_LisEEFF.Col = 18
            grd_LisEEFF.Text = Trim(g_rst_Princi!NOMSUBGRP & "")
            grd_LisEEFF.CellFontName = "Arial"
            grd_LisEEFF.CellFontSize = 8
            
         Else
            grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
            grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
         End If
         
SALTO:
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'llama al SP de EEFF del Año Anterior
   g_str_Parame = "USP_CUR_GEN_EEFF ("
   g_str_Parame = g_str_Parame & CInt(r_str_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) & ",1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "

   'EJECUTA CONSULTA
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEFF.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   Call fs_llenar_Grid_AntVer(CInt(r_str_PerMes), CInt(r_str_PerAno))
     
   grd_LisEEFF.Redraw = True
   If grd_LisEEFF.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEFF)
      Call fs_Activa(True)
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Proces_NueVer()
Dim r_str_PerMesi       As String
Dim r_str_PerAnoi       As String
Dim r_str_PerMesf       As String
Dim r_str_PerAnof       As String
Dim r_int_ConMes        As Integer
Dim r_int_ConAnn        As Integer
Dim r_int_MesIni        As Integer
Dim r_int_ColMes        As Integer
Dim r_int_VarAux1       As Integer
Dim r_int_VarAux2       As Integer
Dim r_int_VarAux3       As Integer

   If cmb_PerMesi.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If cmb_PerMesf.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesf)
      Exit Sub
   End If
   If cmb_PerMesi.ListIndex > cmb_PerMesf.ListIndex And ipp_PerAnoi = ipp_PerAnof Then
      MsgBox "El Mes inicial no puede ser mayor al Mes Final en el mismo año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMesi)
      Exit Sub
   End If
   If ipp_PerAnoi.Text = 0 Then
      MsgBox "Debe seleccionar el Año Inicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   If ipp_PerAnof.Text = 0 Then
      MsgBox "Debe seleccionar el Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnof)
      Exit Sub
   End If
   If ipp_PerAnoi.Text > ipp_PerAnof Then
      MsgBox "El Año inicial no puede ser mayor al Año Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAnoi)
      Exit Sub
   End If
   
   'If MsgBox("¿Está seguro que desea realizar el proceso ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
   '   Exit Sub
   'End If
   
   Screen.MousePointer = 11
   r_str_PerMesi = CInt(cmb_PerMesi.ItemData(cmb_PerMesi.ListIndex))
   r_str_PerAnoi = CInt(ipp_PerAnoi.Text)
   r_str_PerMesf = CInt(cmb_PerMesf.ItemData(cmb_PerMesf.ListIndex))
   r_str_PerAnof = CInt(ipp_PerAnof.Text)
   
   grd_LisEEFF.Redraw = False
   Call gs_LimpiaGrid(grd_LisEEFF)
   grd_LisEEFF.Cols = 5
   Call fs_Recorset_nc_NueVer
   'Call fs_Recorset_nc_AntVer
   
   If r_str_PerAnof >= r_str_PerAnoi Then
      For r_int_ConAnn = r_str_PerAnoi To r_str_PerAnof
      
         If r_int_ConAnn > r_str_PerAnoi And r_int_ConAnn <> r_str_PerAnof Then
            r_int_ConMes = 12

Consultar:
            'llama al SP de EEFF
            g_str_Parame = ""
            g_str_Parame = "USP_CUR_GEN_EEFF ("
            g_str_Parame = g_str_Parame & CInt(r_int_ConMes) & ", "
            g_str_Parame = g_str_Parame & CInt(r_int_ConAnn) & ",1, '" & modgen_g_str_CodUsu & "' ,'" & modgen_g_str_NombPC & "')  "
                      
         ElseIf r_int_ConAnn = r_str_PerAnof Then
            r_int_ConMes = r_str_PerMesf
            GoTo Consultar
         Else
            r_int_ConMes = 12
            GoTo Consultar
         End If
         
         'EJECUTA CONSULTA
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_EEFF.", vbCritical, modgen_g_str_NomPlt
            Exit Sub
         End If
            
         'ALMACENA RESULTADOS DEL AÑO ANTERIOR, EN UN RECORDSET NO CONECTADO
         Dim g_rst_Auxiliar As ADODB.Recordset
         
         g_str_Parame = "SELECT * FROM TT_EEFF"
         g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
         g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
         g_str_Parame = g_str_Parame & " ORDER BY grupo, subgrp, item, indtipo "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Auxiliar, 3) Then
            MsgBox "Error al ejecutar la consulta para recorset no conectado.", vbCritical, modgen_g_str_NomPlt 'Exit Sub
         End If
         
         If Not (g_rst_Auxiliar.BOF And g_rst_Auxiliar.EOF) Then
            g_rst_Auxiliar.MoveFirst
            Do While Not g_rst_Auxiliar.EOF
               g_rst_GenAux.AddNew
               g_rst_GenAux.Fields(0).Value = g_rst_Auxiliar!GRUPO
               g_rst_GenAux.Fields(1).Value = g_rst_Auxiliar!NOMGRUPO
               g_rst_GenAux.Fields(2).Value = g_rst_Auxiliar!SUBGRP
               g_rst_GenAux.Fields(3).Value = g_rst_Auxiliar!NOMSUBGRP
               g_rst_GenAux.Fields(4).Value = g_rst_Auxiliar!CNTACTBLE
               g_rst_GenAux.Fields(5).Value = g_rst_Auxiliar!NOMCTA
               g_rst_GenAux.Fields(6).Value = g_rst_Auxiliar!MES01
               g_rst_GenAux.Fields(7).Value = g_rst_Auxiliar!MES02
               g_rst_GenAux.Fields(8).Value = g_rst_Auxiliar!MES03
               g_rst_GenAux.Fields(9).Value = g_rst_Auxiliar!MES04
               g_rst_GenAux.Fields(10).Value = g_rst_Auxiliar!MES05
               g_rst_GenAux.Fields(11).Value = g_rst_Auxiliar!MES06
               g_rst_GenAux.Fields(12).Value = g_rst_Auxiliar!MES07
               g_rst_GenAux.Fields(13).Value = g_rst_Auxiliar!MES08
               g_rst_GenAux.Fields(14).Value = g_rst_Auxiliar!MES09
               g_rst_GenAux.Fields(15).Value = g_rst_Auxiliar!MES10
               g_rst_GenAux.Fields(16).Value = g_rst_Auxiliar!MES11
               g_rst_GenAux.Fields(17).Value = g_rst_Auxiliar!MES12
               g_rst_GenAux.Fields(18).Value = g_rst_Auxiliar!ACUMU
               g_rst_GenAux.Fields(19).Value = g_rst_Auxiliar!INDTIPO
               g_rst_GenAux.Fields(20).Value = g_rst_Auxiliar!Item
               g_rst_GenAux.Fields(21).Value = r_int_ConAnn
               
               g_rst_GenAux.Update
               g_rst_Auxiliar.MoveNext
            Loop
         End If
         'g_rst_Princi.Close
      Next r_int_ConAnn
   End If
   
   'CABECERA
   grd_LisEEFF.Rows = grd_LisEEFF.Rows + 2
   grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
   grd_LisEEFF.FixedRows = 1

   grd_LisEEFF.Row = 0
   grd_LisEEFF.Col = 3
   grd_LisEEFF.Text = "EJERCICIOS"
   grd_LisEEFF.FixedCols = 4
   
   '***************************************** CABECERA **************************************
    If r_str_PerAnof > r_str_PerAnoi Then
       
      For r_int_ConAnn = r_str_PerAnoi To r_str_PerAnof
      
         If r_int_ConAnn > r_str_PerAnoi And r_int_ConAnn <> r_str_PerAnof Then
             r_int_ConMes = 1
             grd_LisEEFF.Cols = grd_LisEEFF.Cols + 12
             r_int_VarAux1 = grd_LisEEFF.Col + 1
             
             If CInt(r_int_ConMes) = 12 Then
                r_int_VarAux1 = r_int_VarAux1 - 11: GoTo S11
S0:
                grd_LisEEFF.Col = r_int_VarAux1:        grd_LisEEFF.Text = "ENE-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S1:
                grd_LisEEFF.Col = r_int_VarAux1 + 1:    grd_LisEEFF.Text = "FEB-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S2:
                grd_LisEEFF.Col = r_int_VarAux1 + 2:    grd_LisEEFF.Text = "MAR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S3:
                grd_LisEEFF.Col = r_int_VarAux1 + 3:    grd_LisEEFF.Text = "ABR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S4:
                grd_LisEEFF.Col = r_int_VarAux1 + 4:    grd_LisEEFF.Text = "MAY-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S5:
                grd_LisEEFF.Col = r_int_VarAux1 + 5:    grd_LisEEFF.Text = "JUN-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S6:
                grd_LisEEFF.Col = r_int_VarAux1 + 6:    grd_LisEEFF.Text = "JUL-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S7:
                grd_LisEEFF.Col = r_int_VarAux1 + 7:    grd_LisEEFF.Text = "AGO-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S8:
                grd_LisEEFF.Col = r_int_VarAux1 + 8:    grd_LisEEFF.Text = "SET-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S9:
                grd_LisEEFF.Col = r_int_VarAux1 + 9:    grd_LisEEFF.Text = "OCT-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S10:
                grd_LisEEFF.Col = r_int_VarAux1 + 10:   grd_LisEEFF.Text = "NOV-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S11:
                grd_LisEEFF.Col = r_int_VarAux1 + 11:   grd_LisEEFF.Text = "DIC-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
            
             ElseIf CInt(r_int_ConMes) = 11 Then
                r_int_VarAux1 = r_int_VarAux1 - 10
                GoTo S10
             ElseIf CInt(r_int_ConMes) = 10 Then
                r_int_VarAux1 = r_int_VarAux1 - 9
                GoTo S9
             ElseIf CInt(r_int_ConMes) = 9 Then
                r_int_VarAux1 = r_int_VarAux1 - 8
                GoTo S8
             ElseIf CInt(r_int_ConMes) = 8 Then
                r_int_VarAux1 = r_int_VarAux1 - 7
                GoTo S7
             ElseIf CInt(r_int_ConMes) = 7 Then
                r_int_VarAux1 = r_int_VarAux1 - 6
                GoTo S6
             ElseIf CInt(r_int_ConMes) = 6 Then
                r_int_VarAux1 = r_int_VarAux1 - 5
                GoTo S5
             ElseIf CInt(r_int_ConMes) = 5 Then
                r_int_VarAux1 = r_int_VarAux1 - 4
                GoTo S4
             ElseIf CInt(r_int_ConMes) = 4 Then
                r_int_VarAux1 = r_int_VarAux1 - 3
                GoTo S3
             ElseIf CInt(r_int_ConMes) = 3 Then
                r_int_VarAux1 = r_int_VarAux1 - 2
                GoTo S2
             ElseIf CInt(r_int_ConMes) = 2 Then
                r_int_VarAux1 = r_int_VarAux1 - 1
                GoTo S1
             ElseIf CInt(r_int_ConMes) = 1 Then
                r_int_VarAux1 = r_int_VarAux1
                GoTo S0
             End If
             
         ElseIf r_int_ConAnn = r_str_PerAnof Then
            
            r_int_ConMes = r_str_PerMesf
            grd_LisEEFF.Cols = grd_LisEEFF.Cols + r_str_PerMesf + 1
            r_int_VarAux2 = grd_LisEEFF.Col + 1
            
            If CInt(r_int_ConMes) = 12 Then
S13:
                grd_LisEEFF.Col = r_int_VarAux2 + 11:   grd_LisEEFF.Text = "DIC-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 11) = 1350
S14:
                grd_LisEEFF.Col = r_int_VarAux2 + 10:   grd_LisEEFF.Text = "NOV-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 10) = 1350
S15:
                grd_LisEEFF.Col = r_int_VarAux2 + 9:    grd_LisEEFF.Text = "OCT-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 9) = 1350
S16:
                grd_LisEEFF.Col = r_int_VarAux2 + 8:    grd_LisEEFF.Text = "SET-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 8) = 1350
S17:
                grd_LisEEFF.Col = r_int_VarAux2 + 7:    grd_LisEEFF.Text = "AGO-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 7) = 1350
S18:
                grd_LisEEFF.Col = r_int_VarAux2 + 6:    grd_LisEEFF.Text = "JUL-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 6) = 1350
S19:
                grd_LisEEFF.Col = r_int_VarAux2 + 5:    grd_LisEEFF.Text = "JUN-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 5) = 1350
S20:
                grd_LisEEFF.Col = r_int_VarAux2 + 4:    grd_LisEEFF.Text = "MAY-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 4) = 1350
S21:
                grd_LisEEFF.Col = r_int_VarAux2 + 3:    grd_LisEEFF.Text = "ABR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 3) = 1350
S22:
                grd_LisEEFF.Col = r_int_VarAux2 + 2:    grd_LisEEFF.Text = "MAR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 2) = 1350
S23:
                grd_LisEEFF.Col = r_int_VarAux2 + 1:    grd_LisEEFF.Text = "FEB-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2 + 1) = 1350
S24:
                grd_LisEEFF.Col = r_int_VarAux2:        grd_LisEEFF.Text = "ENE-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2) = 1350
                
                grd_LisEEFF.Col = grd_LisEEFF.Cols - 3:  grd_LisEEFF.Text = "ACUM-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux2) = 1350

            ElseIf CInt(r_int_ConMes) = 11 Then
                GoTo S14
            ElseIf CInt(r_int_ConMes) = 10 Then
                GoTo S15
            ElseIf CInt(r_int_ConMes) = 9 Then
                GoTo S16
            ElseIf CInt(r_int_ConMes) = 8 Then
                GoTo S17
            ElseIf CInt(r_int_ConMes) = 7 Then
                GoTo S18
            ElseIf CInt(r_int_ConMes) = 6 Then
                GoTo S19
            ElseIf CInt(r_int_ConMes) = 5 Then
                GoTo S20
            ElseIf CInt(r_int_ConMes) = 4 Then
                GoTo S21
            ElseIf CInt(r_int_ConMes) = 3 Then
                GoTo S22
            ElseIf CInt(r_int_ConMes) = 2 Then
                GoTo S23
            ElseIf CInt(r_int_ConMes) = 1 Then
                GoTo S24
            End If
            
         Else
            r_int_ConMes = r_str_PerMesi
            grd_LisEEFF.Cols = grd_LisEEFF.Cols + (12 - r_str_PerMesi) + 1 + 1
            
            If CInt(r_int_ConMes) = 12 Then
            
                r_int_VarAux1 = -7: GoTo S36
S25:
                grd_LisEEFF.Col = r_int_VarAux1:        grd_LisEEFF.Text = "ENE-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S26:
                grd_LisEEFF.Col = r_int_VarAux1 + 1:    grd_LisEEFF.Text = "FEB-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S27:
                grd_LisEEFF.Col = r_int_VarAux1 + 2:    grd_LisEEFF.Text = "MAR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S28:
                grd_LisEEFF.Col = r_int_VarAux1 + 3:    grd_LisEEFF.Text = "ABR-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S29:
                grd_LisEEFF.Col = r_int_VarAux1 + 4:    grd_LisEEFF.Text = "MAY-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S30:
                grd_LisEEFF.Col = r_int_VarAux1 + 5:    grd_LisEEFF.Text = "JUN-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S31:
                grd_LisEEFF.Col = r_int_VarAux1 + 6:    grd_LisEEFF.Text = "JUL-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S32:
                grd_LisEEFF.Col = r_int_VarAux1 + 7:    grd_LisEEFF.Text = "AGO-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S33:
                grd_LisEEFF.Col = r_int_VarAux1 + 8:    grd_LisEEFF.Text = "SET-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S34:
                grd_LisEEFF.Col = r_int_VarAux1 + 9:    grd_LisEEFF.Text = "OCT-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S35:
                grd_LisEEFF.Col = r_int_VarAux1 + 10:   grd_LisEEFF.Text = "NOV-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
S36:
                grd_LisEEFF.Col = r_int_VarAux1 + 11:   grd_LisEEFF.Text = "DIC-" & Right(CInt(r_int_ConAnn), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter
                        
             ElseIf CInt(r_int_ConMes) = 11 Then
                r_int_VarAux1 = -6
                GoTo S35
             ElseIf CInt(r_int_ConMes) = 10 Then
                r_int_VarAux1 = -5
                GoTo S34
             ElseIf CInt(r_int_ConMes) = 9 Then
                r_int_VarAux1 = -4
                GoTo S33
             ElseIf CInt(r_int_ConMes) = 8 Then
                r_int_VarAux1 = -3
                GoTo S32
             ElseIf CInt(r_int_ConMes) = 7 Then
                r_int_VarAux1 = -2
                GoTo S31
             ElseIf CInt(r_int_ConMes) = 6 Then
                r_int_VarAux1 = -1
                GoTo S30
             ElseIf CInt(r_int_ConMes) = 5 Then
                r_int_VarAux1 = 0
                GoTo S29
             ElseIf CInt(r_int_ConMes) = 4 Then
                r_int_VarAux1 = 1
                GoTo S28
             ElseIf CInt(r_int_ConMes) = 3 Then
                r_int_VarAux1 = 2
                GoTo S27
             ElseIf CInt(r_int_ConMes) = 2 Then
                r_int_VarAux1 = 3
                GoTo S26
             ElseIf CInt(r_int_ConMes) = 1 Then
                r_int_VarAux1 = 4
                GoTo S25
             End If

         End If
      Next r_int_ConAnn
   End If

   '*********************************** FIN DE CABECERA *************************************
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM TT_EEFF"
   g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "   AND INDTIPO <> 'D' "
   g_str_Parame = g_str_Parame & " ORDER BY Grupo, Subgrp, Item, indtipo "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar la consulta para recorset no conectado.", vbCritical, modgen_g_str_NomPlt 'Exit Sub
   End If
        
   '******************************* DETALLE DE LA INFORMACIÓN *******************************
   If r_str_PerAnof = r_str_PerAnoi Then
         r_int_ConMes = r_str_PerMesi
         grd_LisEEFF.Cols = grd_LisEEFF.Cols + ((r_str_PerMesf - r_str_PerMesi) + 1) + 2 '1
               
         If CInt(r_int_ConMes) = 12 Then
             r_int_VarAux1 = -10
S37:
             If r_str_PerMesf + 1 = 1 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 3:   grd_LisEEFF.Text = "ENE-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 3) = 1350
S38:
             If r_str_PerMesf + 1 = 2 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 4:   grd_LisEEFF.Text = "FEB-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 4) = 1350
S39:
             If r_str_PerMesf + 1 = 3 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 5:   grd_LisEEFF.Text = "MAR-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 5) = 1350
S40:
             If r_str_PerMesf + 1 = 4 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 6:   grd_LisEEFF.Text = "ABR-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 6) = 1350
S41:
             If r_str_PerMesf + 1 = 5 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 7:   grd_LisEEFF.Text = "MAY-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 7) = 1350
S42:
             If r_str_PerMesf + 1 = 6 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 8:   grd_LisEEFF.Text = "JUN-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 8) = 1350
S43:
             If r_str_PerMesf + 1 = 7 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 9:   grd_LisEEFF.Text = "JUL-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 9) = 1350
S44:
             If r_str_PerMesf + 1 = 8 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 10:  grd_LisEEFF.Text = "AGO-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 10) = 1350
S45:
             If r_str_PerMesf + 1 = 9 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 11:  grd_LisEEFF.Text = "SET-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 11) = 1350
S46:
             If r_str_PerMesf + 1 = 10 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 12: grd_LisEEFF.Text = "OCT-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 12) = 1350
S47:
             If r_str_PerMesf + 1 = 11 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 13: grd_LisEEFF.Text = "NOV-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 13) = 1350
S48:
             If r_str_PerMesf + 1 = 12 Then GoTo SALTO3 Else grd_LisEEFF.Col = r_int_VarAux1 + 14: grd_LisEEFF.Text = "DIC-" & Right(CInt(r_str_PerAnof), 2): grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(r_int_VarAux1 + 14) = 1350
      
         ElseIf CInt(r_int_ConMes) = 11 Then
             r_int_VarAux1 = -9
             GoTo S47
         ElseIf CInt(r_int_ConMes) = 10 Then
             r_int_VarAux1 = -8
             GoTo S46
         ElseIf CInt(r_int_ConMes) = 9 Then
             r_int_VarAux1 = -7
             GoTo S45
         ElseIf CInt(r_int_ConMes) = 8 Then
             r_int_VarAux1 = -6
             GoTo S44
         ElseIf CInt(r_int_ConMes) = 7 Then
             r_int_VarAux1 = -5
             GoTo S43
         ElseIf CInt(r_int_ConMes) = 6 Then
             r_int_VarAux1 = -4
             GoTo S42
         ElseIf CInt(r_int_ConMes) = 5 Then
             r_int_VarAux1 = -3
             GoTo S41
         ElseIf CInt(r_int_ConMes) = 4 Then
             r_int_VarAux1 = -2
             GoTo S40
         ElseIf CInt(r_int_ConMes) = 3 Then
             r_int_VarAux1 = -1
             GoTo S39
         ElseIf CInt(r_int_ConMes) = 2 Then
             r_int_VarAux1 = 0
             GoTo S38
         ElseIf CInt(r_int_ConMes) = 1 Then
             r_int_VarAux1 = 1
             GoTo S37
         End If
   
SALTO3:
         grd_LisEEFF.Col = grd_LisEEFF.Cols - 3:   grd_LisEEFF.Text = "ACUM-" & Right(CInt(r_str_PerAnof), 2):   grd_LisEEFF.CellAlignment = flexAlignCenterCenter: grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 3) = 1350
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
               g_rst_Princi.MoveFirst
               Do While Not g_rst_Princi.EOF
                  If Trim(g_rst_Princi!INDTIPO) <> "L" Then
                        grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
                        grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
         
                        grd_LisEEFF.Col = 0
                        grd_LisEEFF.Text = Trim(g_rst_Princi!GRUPO)
                        grd_LisEEFF.Col = 1
                        grd_LisEEFF.Text = Trim(g_rst_Princi!SUBGRP)
                        grd_LisEEFF.Col = 2
                        grd_LisEEFF.Text = Trim(g_rst_Princi!INDTIPO)
         
                        If Trim(g_rst_Princi!INDTIPO) = "S" Then
                           grd_LisEEFF.Col = 3
                           grd_LisEEFF.Text = Space(5) & Trim(g_rst_Princi!NOMSUBGRP)
                        ElseIf Trim(g_rst_Princi!INDTIPO) = "L" Then
                           grd_LisEEFF.Col = 3
                           grd_LisEEFF.Text = ""
                        ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Then
                           grd_LisEEFF.Col = 3
                           grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
                        Else
                           grd_LisEEFF.Col = 3
                           grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
                        End If
         
                        grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
                        grd_LisEEFF.CellFontBold = True
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        
                        If CInt(r_str_PerMesi) = 12 Then
                           r_int_VarAux3 = -2
                           GoTo S60
S49:
                           If r_str_PerMesf + 1 = 1 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 - 5
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES01, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 - 5) = 1350
                           End If
S50:
                           If r_str_PerMesf + 1 = 2 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 - 4
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES02, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 - 4) = 1350
                           End If
S51:
                           If r_str_PerMesf + 1 = 3 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 - 3
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES03, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 - 3) = 1350
                           End If
S52:
                           If r_str_PerMesf + 1 = 4 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 - 2
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES04, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 - 2) = 1350
                           End If
S53:
                           If r_str_PerMesf + 1 = 5 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 - 1
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES05, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 - 1) = 1350
                           End If
S54:
                           If r_str_PerMesf + 1 = 6 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES06, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3) = 1350
                           End If
S55:
                           If r_str_PerMesf + 1 = 7 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 1
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES07, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 1) = 1350
                           End If
S56:
                           If r_str_PerMesf + 1 = 8 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 2
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES08, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 2) = 1350
                           End If
S57:
                           If r_str_PerMesf + 1 = 9 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 3
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES09, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 3) = 1350
                           End If
S58:
                           If r_str_PerMesf + 1 = 10 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 4
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES10, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 4) = 1350
                           End If
S59:
                           If r_str_PerMesf + 1 = 11 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 5
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES11, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 5) = 1350
                           End If
S60:
                           If r_str_PerMesf + 1 = 12 Then
                              GoTo SALTO4
                           Else
                              grd_LisEEFF.Col = r_int_VarAux3 + 6
                              grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES12, 2)
                              grd_LisEEFF.CellAlignment = flexAlignRightCenter
                              grd_LisEEFF.CellFontName = "Arial"
                              grd_LisEEFF.CellFontSize = 8
                              grd_LisEEFF.ColWidth(r_int_VarAux3 + 6) = 1350
                           End If
                        
                        ElseIf CInt(r_str_PerMesi) = 11 Then
                           r_int_VarAux3 = -1
                           GoTo S59
                        ElseIf CInt(r_str_PerMesi) = 10 Then
                           r_int_VarAux3 = 0
                           GoTo S58
                        ElseIf CInt(r_str_PerMesi) = 9 Then
                           r_int_VarAux3 = 1
                           GoTo S57
                        ElseIf CInt(r_str_PerMesi) = 8 Then
                           r_int_VarAux3 = 2
                           GoTo S56
                        ElseIf CInt(r_str_PerMesi) = 7 Then
                           r_int_VarAux3 = 3
                           GoTo S55
                        ElseIf CInt(r_str_PerMesi) = 6 Then
                           r_int_VarAux3 = 4
                           GoTo S54
                        ElseIf CInt(r_str_PerMesi) = 5 Then
                           r_int_VarAux3 = 5
                           GoTo S53
                        ElseIf CInt(r_str_PerMesi) = 4 Then
                           r_int_VarAux3 = 6
                           GoTo S52
                        ElseIf CInt(r_str_PerMesi) = 3 Then
                           r_int_VarAux3 = 7
                           GoTo S51
                        ElseIf CInt(r_str_PerMesi) = 2 Then
                           r_int_VarAux3 = 8
                           GoTo S50
                        ElseIf CInt(r_str_PerMesi) = 1 Then
                           r_int_VarAux3 = 9
                           GoTo S49
                        End If
      
SALTO4:
                        grd_LisEEFF.Col = grd_LisEEFF.Cols - 3
                        If CInt(r_str_PerMesi) > 1 Then
                           grd_LisEEFF.Text = FormatNumber(Sumar(grd_LisEEFF, grd_LisEEFF.Row), 2)
                        Else
                           grd_LisEEFF.Text = FormatNumber(g_rst_Princi!ACUMU, 2)
                        End If
                        
                        grd_LisEEFF.CellAlignment = flexAlignRightCenter
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 3) = 1350
                        
                        grd_LisEEFF.Col = grd_LisEEFF.Cols - 2
                        grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        
                        grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 2) = 0     ' NOMBRE GRUPO
         
                        grd_LisEEFF.Col = grd_LisEEFF.Cols - 1
                        grd_LisEEFF.Text = Trim(g_rst_Princi!NOMSUBGRP & "")
                        grd_LisEEFF.CellFontName = "Arial"
                        grd_LisEEFF.CellFontSize = 8
                        
                        grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 1) = 0     ' NOMBRE SUBGRUPO
                        
      
                  Else
                     grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
                     grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
                  End If
                  
SALTO1:
                  g_rst_Princi.MoveNext
                  
               Loop
       End If
   
   ElseIf CInt(r_str_PerAnoi) <> CInt(r_str_PerAnof) Then
      
      'ÚLTIMO AÑO CONSULTADO
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst

         Do While Not g_rst_Princi.EOF

            If Trim(g_rst_Princi!INDTIPO) <> "L" Then

               grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
               grd_LisEEFF.Row = grd_LisEEFF.Rows - 1

               grd_LisEEFF.Col = 0
               grd_LisEEFF.Text = Trim(g_rst_Princi!GRUPO)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(0) = 0

               grd_LisEEFF.Col = 1
               grd_LisEEFF.Text = Trim(g_rst_Princi!SUBGRP)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(1) = 0

               grd_LisEEFF.Col = 2
               grd_LisEEFF.Text = Trim(g_rst_Princi!INDTIPO)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(2) = 0

               If Trim(g_rst_Princi!INDTIPO) = "S" Then
                  grd_LisEEFF.Col = 3
                  grd_LisEEFF.Text = Space(5) & Trim(g_rst_Princi!NOMSUBGRP)
               ElseIf Trim(g_rst_Princi!INDTIPO) = "L" Then
                  grd_LisEEFF.Col = 3
                  grd_LisEEFF.Text = ""
               ElseIf Trim(g_rst_Princi!INDTIPO) = "G" Then
                  grd_LisEEFF.Col = 3
                  grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
               Else
                  grd_LisEEFF.Col = 3
                  grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
               End If
               grd_LisEEFF.ColWidth(3) = 3030

               grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
               grd_LisEEFF.CellFontBold = True
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8

               If CInt(r_str_PerMesf) = 12 Then

                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) - 4
                  grd_LisEEFF.Col = r_int_VarAux2 + 1
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES12, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 + 1) = 1350
S61:
                  grd_LisEEFF.Col = r_int_VarAux2
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES11, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2) = 1350
S62:
                  grd_LisEEFF.Col = r_int_VarAux2 - 1
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES10, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 1) = 1350
S63:
                  grd_LisEEFF.Col = r_int_VarAux2 - 2
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES09, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 2) = 1350
S64:
                  grd_LisEEFF.Col = r_int_VarAux2 - 3
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES08, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 3) = 1350
S65:
                  grd_LisEEFF.Col = r_int_VarAux2 - 4
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES07, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 4) = 1350
S66:
                  grd_LisEEFF.Col = r_int_VarAux2 - 5
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES06, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 5) = 1350
S67:
                  grd_LisEEFF.Col = r_int_VarAux2 - 6
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES05, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 6) = 1350
S68:
                  grd_LisEEFF.Col = r_int_VarAux2 - 7
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES04, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 7) = 1350
S69:
                  grd_LisEEFF.Col = r_int_VarAux2 - 8
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES03, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 8) = 1350
S70:
                  grd_LisEEFF.Col = r_int_VarAux2 - 9
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES02, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 9) = 1350
S71:
                  grd_LisEEFF.Col = r_int_VarAux2 - 10
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!MES01, 2)  ' Format(g_rst_Princi!MES01, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_VarAux2 - 10) = 1350
                  
                  
                  grd_LisEEFF.Col = grd_LisEEFF.Cols - 3
                  grd_LisEEFF.Text = FormatNumber(g_rst_Princi!ACUMU, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 3) = 1350

               ElseIf CInt(r_str_PerMesf) = 11 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) - 2
                  GoTo S61
               ElseIf CInt(r_str_PerMesf) = 10 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1)
                  GoTo S62
               ElseIf CInt(r_str_PerMesf) = 9 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 2
                  GoTo S63
               ElseIf CInt(r_str_PerMesf) = 8 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 4
                  GoTo S64
               ElseIf CInt(r_str_PerMesf) = 7 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 6
                  GoTo S65
               ElseIf CInt(r_str_PerMesf) = 6 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 8
                  GoTo S66
               ElseIf CInt(r_str_PerMesf) = 5 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 10
                  GoTo S67
               ElseIf CInt(r_str_PerMesf) = 4 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 12
                  GoTo S68
               ElseIf CInt(r_str_PerMesf) = 3 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 14
                  GoTo S69
               ElseIf CInt(r_str_PerMesf) = 2 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 16
                  GoTo S70
               ElseIf CInt(r_str_PerMesf) = 1 Then
                  r_int_VarAux2 = grd_LisEEFF.Cols - (12 - r_str_PerMesf + 1) + 18
                  GoTo S71
               End If

   '            grd_LisEEFF.Col = 16
   '            grd_LisEEFF.Text = Format(g_rst_Princi!ACUMU, "###,###,###,##0")
   '            grd_LisEEFF.CellFontName = "Arial"
   '            grd_LisEEFF.CellFontSize = 8

               grd_LisEEFF.Col = grd_LisEEFF.Cols - 2
               grd_LisEEFF.Text = Trim(g_rst_Princi!NOMGRUPO)
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 2) = 0

               grd_LisEEFF.Col = grd_LisEEFF.Cols - 1
               grd_LisEEFF.Text = Trim(g_rst_Princi!NOMSUBGRP & "")
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               grd_LisEEFF.ColWidth(grd_LisEEFF.Cols - 1) = 0

            Else
               grd_LisEEFF.Rows = grd_LisEEFF.Rows + 1
               grd_LisEEFF.Row = grd_LisEEFF.Rows - 1
            End If

            g_rst_Princi.MoveNext
         Loop
              
         'AÑOS ANTERIORES AL AÑO FINAL CONSULTADO
         For r_int_VarAux1 = r_str_PerAnoi To r_str_PerAnof
            If r_int_VarAux1 > r_str_PerAnoi And r_int_VarAux1 <> r_str_PerAnof Then
               r_int_MesIni = 1
               If r_int_ColMes = 4 Then
                  r_int_ColMes = (12 - r_str_PerMesi) + 1 + 4
               Else
                  r_int_ColMes = r_int_ColMes + 12
               End If
               Call fs_llenar_Grid_NueVer(CInt(r_int_MesIni), CInt(r_int_VarAux1), g_rst_GenAux, CInt(r_int_ColMes))
               
            ElseIf r_int_VarAux1 = r_str_PerAnof Then
              Exit For
            Else
              r_int_MesIni = r_str_PerMesi
              r_int_ColMes = 4
              Call fs_llenar_Grid_NueVer(CInt(r_int_MesIni), CInt(r_int_VarAux1), g_rst_GenAux, CInt(r_int_ColMes))
              
            End If
         Next r_int_VarAux1
      
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   grd_LisEEFF.Redraw = True
   If grd_LisEEFF.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisEEFF)
      Call fs_Activa(True)
   Else
      MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_llenar_Grid_NueVer(ByVal r_str_PerMesi As Integer, ByVal r_int_Anno As Integer, ByVal g_rst_Aux As ADODB.Recordset, ByVal r_int_Columna As Integer)
Dim r_int_NumCol     As Integer
Dim Fila             As Integer
Dim p                As Integer

   Fila = 1
   If Not (g_rst_Aux.BOF And g_rst_Aux.EOF) Then
   
      g_rst_Aux.MoveFirst
      Do While Not g_rst_Aux.EOF
      
         If r_int_Anno = g_rst_Aux!anno Then
            If Trim(g_rst_Aux!INDTIPO) <> "L" Then
               'Fila = Fila + 1
               If Trim(g_rst_Aux!INDTIPO) = "D" Then GoTo SALTO
                  Fila = Fila + 1
                
               If CInt(r_str_PerMesi) = 12 Then                                  'DICIEMBRE
                  r_int_NumCol = r_int_Columna - 11
                  GoTo L
A:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol) = FormatNumber(g_rst_Aux!MES01, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol) = 1350
   
B:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 1) = FormatNumber(g_rst_Aux!MES02, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 1) = 1350
   
C:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 2) = FormatNumber(g_rst_Aux!MES03, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 2) = 1350
   
D:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 3) = FormatNumber(g_rst_Aux!MES04, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 3) = 1350
   
E:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 4) = FormatNumber(g_rst_Aux!MES05, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 4) = 1350
   
F:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 5) = FormatNumber(g_rst_Aux!MES06, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 5) = 1350
   
G:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 6) = FormatNumber(g_rst_Aux!MES07, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 6) = 1350
   
H:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 7) = FormatNumber(g_rst_Aux!MES08, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 7) = 1350
   
i:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 8) = FormatNumber(g_rst_Aux!MES09, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 8) = 1350
   
j:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 9) = FormatNumber(g_rst_Aux!MES10, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 9) = 1350
   
k:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 10) = FormatNumber(g_rst_Aux!MES11, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 10) = 1350
   
L:
                  grd_LisEEFF.TextMatrix(Fila, r_int_NumCol + 11) = FormatNumber(g_rst_Aux!MES12, 2)
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                  grd_LisEEFF.ColWidth(r_int_NumCol + 11) = 1350
                  
               ElseIf CInt(r_str_PerMesi) = 11 Then         'NOVIEMBRE
                  r_int_NumCol = r_int_Columna - 10
                  GoTo k
               ElseIf CInt(r_str_PerMesi) = 10 Then         'OCTUBRE
                  r_int_NumCol = r_int_Columna - 9
                  GoTo j
               ElseIf CInt(r_str_PerMesi) = 9 Then          'SEPTIEMBRE
                  r_int_NumCol = r_int_Columna - 8
                  GoTo i
               ElseIf CInt(r_str_PerMesi) = 8 Then          'AGOSTO
                  r_int_NumCol = r_int_Columna - 7
                  GoTo H
               ElseIf CInt(r_str_PerMesi) = 7 Then          'JULIO
                  r_int_NumCol = r_int_Columna - 6
                  GoTo G
               ElseIf CInt(r_str_PerMesi) = 6 Then          'JUNIO
                  r_int_NumCol = r_int_Columna - 5
                  GoTo F
               ElseIf CInt(r_str_PerMesi) = 5 Then          'MAYO
                  r_int_NumCol = r_int_Columna - 4
                  GoTo E
               ElseIf CInt(r_str_PerMesi) = 4 Then          'ABRIL
                  r_int_NumCol = r_int_Columna - 3
                  GoTo D
               ElseIf CInt(r_str_PerMesi) = 3 Then          'MARZO
                  r_int_NumCol = r_int_Columna - 2
                  GoTo C
               ElseIf CInt(r_str_PerMesi) = 2 Then          'FEBRERO
                  r_int_NumCol = r_int_Columna - 1
                  GoTo B
               ElseIf CInt(r_str_PerMesi) = 1 Then          'ENERO
                  r_int_NumCol = r_int_Columna
                  GoTo A
               End If
               
            Else
               Fila = Fila + 1
            End If
   
         End If

SALTO:
            g_rst_Aux.MoveNext
         
      Loop
   End If
End Sub

Private Sub fs_llenar_Grid_AntVer(ByVal Columna As Integer, ByVal anno As Integer)
Dim Fila    As Integer
Dim j       As Integer
Dim k       As Integer
    
   grd_LisEEFF.Row = 0
   
   If Columna = 12 Then
      j = -12
      grd_LisEEFF.Col = 4 + Columna + j + 11: grd_LisEEFF.Text = "DIC-" & Right(anno, 2)
S1:
      grd_LisEEFF.Col = 4 + Columna + j + 10: grd_LisEEFF.Text = "NOV-" & Right(anno, 2)
S2:
      grd_LisEEFF.Col = 4 + Columna + j + 9:  grd_LisEEFF.Text = "OCT-" & Right(anno, 2)
S3:
      grd_LisEEFF.Col = 4 + Columna + j + 8:  grd_LisEEFF.Text = "SET-" & Right(anno, 2)
S4:
      grd_LisEEFF.Col = 4 + Columna + j + 7:  grd_LisEEFF.Text = "AGO-" & Right(anno, 2)
S5:
      grd_LisEEFF.Col = 4 + Columna + j + 6:  grd_LisEEFF.Text = "JUL-" & Right(anno, 2)
S6:
      grd_LisEEFF.Col = 4 + Columna + j + 5:  grd_LisEEFF.Text = "JUN-" & Right(anno, 2)
S7:
      grd_LisEEFF.Col = 4 + Columna + j + 4:  grd_LisEEFF.Text = "MAY-" & Right(anno, 2)
S8:
      grd_LisEEFF.Col = 4 + Columna + j + 3:  grd_LisEEFF.Text = "ABR-" & Right(anno, 2)
S9:
      grd_LisEEFF.Col = 4 + Columna + j + 2:  grd_LisEEFF.Text = "MAR-" & Right(anno, 2)
S10:
      grd_LisEEFF.Col = 4 + Columna + j + 1: grd_LisEEFF.Text = "FEB-" & Right(anno, 2)
S11:
      grd_LisEEFF.Col = 4 + Columna + j: grd_LisEEFF.Text = "ENE-" & Right(anno, 2)
        
   ElseIf Columna = 11 Then
      j = -10
      GoTo S1
   ElseIf Columna = 10 Then
      j = -8
      GoTo S2
   ElseIf Columna = 9 Then
      j = -6
      GoTo S3
   ElseIf Columna = 8 Then
      j = -4
      GoTo S4
   ElseIf Columna = 7 Then
      j = -2
      GoTo S5
   ElseIf Columna = 6 Then
      j = 0
      GoTo S6
   ElseIf Columna = 5 Then
      j = 2
      GoTo S7
   ElseIf Columna = 4 Then
      j = 4
      GoTo S8
   ElseIf Columna = 3 Then
      j = 6
      GoTo S9
   ElseIf Columna = 2 Then
      j = 8
      GoTo S10
   ElseIf Columna = 1 Then
      j = 10
      GoTo S11
   End If
        
   grd_LisEEFF.Col = 16:   grd_LisEEFF.Text = "ACUM-" & Right(anno, 2): grd_LisEEFF.Font.Bold = False
     
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
            
         For Fila = 2 To grd_LisEEFF.Rows - 1
            If Trim(g_rst_Princi!INDTIPO) <> "L" Then
               grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
               'grd_LisEEFF.CellFontBold = True
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
               
               If Columna = 12 Then
                  k = 13
                  grd_LisEEFF.TextMatrix(Fila, k + 2) = Format(g_rst_Princi!MES12, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S12:
                  grd_LisEEFF.TextMatrix(Fila, k + 1) = Format(g_rst_Princi!MES11, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S13:
                  grd_LisEEFF.TextMatrix(Fila, k) = Format(g_rst_Princi!MES10, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S14:
                  grd_LisEEFF.TextMatrix(Fila, k - 1) = Format(g_rst_Princi!MES09, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S15:
                  grd_LisEEFF.TextMatrix(Fila, k - 2) = Format(g_rst_Princi!MES08, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S16:
                  grd_LisEEFF.TextMatrix(Fila, k - 3) = Format(g_rst_Princi!MES07, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S17:
                  grd_LisEEFF.TextMatrix(Fila, k - 4) = Format(g_rst_Princi!MES06, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S18:
                  grd_LisEEFF.TextMatrix(Fila, k - 5) = Format(g_rst_Princi!MES05, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S19:
                  grd_LisEEFF.TextMatrix(Fila, k - 6) = Format(g_rst_Princi!MES04, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S20:
                  grd_LisEEFF.TextMatrix(Fila, k - 7) = Format(g_rst_Princi!MES03, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S21:
                  grd_LisEEFF.TextMatrix(Fila, k - 8) = Format(g_rst_Princi!MES02, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
S22:
                  grd_LisEEFF.TextMatrix(Fila, k - 9) = Format(g_rst_Princi!MES01, "###,###,###,##0")
                  grd_LisEEFF.CellFontName = "Arial"
                  grd_LisEEFF.CellFontSize = 8
                           
               ElseIf Columna = 11 Then
                   k = 14
                   GoTo S12
               ElseIf Columna = 10 Then
                   k = 15
                   GoTo S13
               ElseIf Columna = 9 Then
                   k = 16
                   GoTo S14
               ElseIf Columna = 8 Then
                   k = 17
                   GoTo S15
               ElseIf Columna = 7 Then
                   k = 18
                   GoTo S16
               ElseIf Columna = 6 Then
                   k = 19
                   GoTo S17
               ElseIf Columna = 5 Then
                   k = 20
                   GoTo S18
               ElseIf Columna = 4 Then
                   k = 21
                   GoTo S19
               ElseIf Columna = 3 Then
                   k = 22
                   GoTo S20
               ElseIf Columna = 2 Then
                   k = 23
                   GoTo S21
               ElseIf Columna = 1 Then
                   k = 24
                   GoTo S22
               End If
                        
               grd_LisEEFF.TextMatrix(Fila, 16) = Format(g_rst_Princi!ACUMU, "###,###,###,##0")
               grd_LisEEFF.CellFontName = "Arial"
               grd_LisEEFF.CellFontSize = 8
            End If
                    
            g_rst_Princi.MoveNext
         Next
      Loop
   End If
    
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Centra fila (meses) del encabezado
   For j = 4 To 18
      grd_LisEEFF.Row = 0
      grd_LisEEFF.Col = j
      grd_LisEEFF.CellFontName = "Arial"
      grd_LisEEFF.Font.Bold = False
      grd_LisEEFF.CellForeColor = modgen_g_con_ColVer
      grd_LisEEFF.CellFontSize = 8
      grd_LisEEFF.CellAlignment = flexAlignCenterCenter
   Next j
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia_NueVer
   'Call fs_Inicia_AntVer
  
   Call gs_CentraForm(Me)
   Call fs_Activa(False)
   Call fs_Recorset_nc_NueVer
   'Call fs_Recorset_nc_AntVer
   
   Call gs_SetFocus(cmb_PerMesi)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Recorset_nc_NueVer()
    Set g_rst_GenAux = New ADODB.Recordset
    
    g_rst_GenAux.Fields.Append "GRUPO", adBigInt, 2, adFldFixed
    g_rst_GenAux.Fields.Append "NOMGRUPO", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "SUBGRP", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "NOMSUBGRP", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "CNTACTBLE", adChar, 30, adFldIsNullable
    g_rst_GenAux.Fields.Append "NOMCTA", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "MES01", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES02", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES03", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES04", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES05", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES06", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES07", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES08", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES09", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES10", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES11", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES12", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "ACUMU", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "INDTIPO", adChar, 5, adFldFixed
    g_rst_GenAux.Fields.Append "ITEM", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "ANNO", adChar, 4, adFldFixed
    g_rst_GenAux.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub fs_Recorset_nc_AntVer()
    Set g_rst_GenAux = New ADODB.Recordset
    
    g_rst_GenAux.Fields.Append "GRUPO", adBigInt, 2, adFldFixed
    g_rst_GenAux.Fields.Append "NOMGRUPO", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "SUBGRP", adBigInt, 3, adFldFixed
    g_rst_GenAux.Fields.Append "NOMSUBGRP", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "CNTACTBLE", adChar, 30, adFldIsNullable
    g_rst_GenAux.Fields.Append "NOMCTA", adChar, 150, adFldIsNullable
    g_rst_GenAux.Fields.Append "MES01", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES02", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES03", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES04", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES05", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES06", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES07", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES08", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES09", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES10", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES11", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "MES12", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "ACUMU", adDouble, , adFldFixed
    g_rst_GenAux.Fields.Append "INDTIPO", adChar, 5, adFldFixed
    g_rst_GenAux.Fields.Append "ITEM", adBigInt, 3, adFldFixed
    g_rst_GenAux.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub fs_Inicia_AntVer()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
       
   'LISTADO INGRESOS MANUALES
   grd_LisEEFF.ColWidth(0) = 0      ' GRUPO
   grd_LisEEFF.ColWidth(1) = 0      ' COD SUBGRUPO
   grd_LisEEFF.ColWidth(2) = 0      ' INDICA TIPO
   grd_LisEEFF.ColWidth(3) = 3030   ' DESCRIPCION
   grd_LisEEFF.ColWidth(4) = 870    ' MES 1
   grd_LisEEFF.ColWidth(5) = 870    ' MES 2
   grd_LisEEFF.ColWidth(6) = 870    ' MES 3
   grd_LisEEFF.ColWidth(7) = 870    ' MES 4
   grd_LisEEFF.ColWidth(8) = 870    ' MES 5
   grd_LisEEFF.ColWidth(9) = 870    ' MES 6
   grd_LisEEFF.ColWidth(10) = 870   ' MES 7
   grd_LisEEFF.ColWidth(11) = 870   ' MES 8
   grd_LisEEFF.ColWidth(12) = 870   ' MES 9
   grd_LisEEFF.ColWidth(13) = 870   ' MES 10
   grd_LisEEFF.ColWidth(14) = 870   ' MES 11
   grd_LisEEFF.ColWidth(15) = 870   ' MES 12
   grd_LisEEFF.ColWidth(16) = 930   ' ACUMULADO
   grd_LisEEFF.ColWidth(17) = 0     ' NOMBRE GRUPO
   grd_LisEEFF.ColWidth(18) = 0     ' NOMBRE SUBGRUPO
   grd_LisEEFF.ColAlignment(3) = flexAlignLeftCenter
   grd_LisEEFF.ColAlignment(4) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(5) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(6) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(7) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(8) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(9) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(10) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(11) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(12) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(13) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(14) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(15) = flexAlignRightCenter
   grd_LisEEFF.ColAlignment(16) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisEEFF)
End Sub

Private Sub fs_Inicia_NueVer()
   cmb_PerMesi.Clear
   cmb_PerMesf.Clear
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMesi, 1, "033")
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMesf, 1, "033")
   
   r_int_PerMesf = Month(date)
   r_int_PerAnof = Year(date)
   r_int_PerMesi = Month(date) + 1
   r_int_PerMesf = Month(date)
   r_int_PerAnoi = Year(date) - 1
   r_int_PerAnof = Year(date)
 
   Call gs_BuscarCombo_Item(cmb_PerMesi, r_int_PerMesi)
   Call gs_BuscarCombo_Item(cmb_PerMesf, r_int_PerMesf)
   ipp_PerAnoi.Text = Format(r_int_PerAnoi, "0000")
   ipp_PerAnof.Text = Format(r_int_PerAnof, "0000")
   
   'LISTADO INGRESOS MANUALES
   grd_LisEEFF.ColWidth(0) = 0      ' GRUPO
   grd_LisEEFF.ColWidth(1) = 0      ' COD SUBGRUPO
   grd_LisEEFF.ColWidth(2) = 0      ' INDICA TIPO
   grd_LisEEFF.ColWidth(3) = 3030   ' DESCRIPCION
   Call gs_LimpiaGrid(grd_LisEEFF)
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExcRes.Enabled = estado
    cmd_ExpExcDet.Enabled = estado
End Sub

Private Sub grd_LisEEFF_DblClick()
   fs_grd_LisEEFF_NueVer
End Sub

Private Sub fs_grd_LisEEFF_NueVer()
   Dim r_str_FecRpt        As String

   If grd_LisEEFF.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_FecRpt = "01/" & Format(r_int_PerMesi, "00") & "/" & r_int_PerAnoi & " AL " & Format(ff_Ultimo_Dia_Mes(r_int_PerMesf, r_int_PerAnof), "00") & "/" & Format(r_int_PerMesf, "00") & "/" & r_int_PerAnof
   r_int_PerMesi = CInt(cmb_PerMesi.ItemData(cmb_PerMesi.ListIndex))
   r_int_PerAnoi = CInt(ipp_PerAnoi.Text)
   
   moddat_g_str_FecIng = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMesf.Text, 1) & LCase(Mid(cmb_PerMesf.Text, 2, Len(cmb_PerMesf.Text))) & " del " & Format(r_int_PerAnof, "0000")
   
   grd_LisEEFF.Col = 0
   moddat_g_str_CodPrd = Trim(grd_LisEEFF & "")
   
   grd_LisEEFF.Col = 1
   moddat_g_str_CodSub = Trim(grd_LisEEFF)
   
   grd_LisEEFF.Col = 2
   moddat_g_str_TipCre = Trim(grd_LisEEFF)
         
   grd_LisEEFF.Col = grd_LisEEFF.Cols - 2
   moddat_g_str_NomPrd = UCase(Trim(grd_LisEEFF))
   
   grd_LisEEFF.Col = grd_LisEEFF.Cols - 1
   moddat_g_str_NomPrd = moddat_g_str_NomPrd & " " & IIf(Len(Trim(grd_LisEEFF)) > 0, " - " & UCase(Trim(grd_LisEEFF)), "")
   
   Call gs_RefrescaGrid(grd_LisEEFF)
   
   If moddat_g_str_TipCre <> "F" And moddat_g_str_TipCre <> "L" And moddat_g_str_TipCre <> "" Then
        frm_RptCtb_20.Show 1
   End If
End Sub

Private Sub fs_grd_LisEEFF_AntVer()
   Dim r_str_FecRpt        As String

   If grd_LisEEFF.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_FecRpt = "01/" & Format(r_int_PerMesi, "00") & "/" & r_int_PerAnoi & " AL " & Format(ff_Ultimo_Dia_Mes(r_int_PerMesf, r_int_PerAnof), "00") & "/" & Format(r_int_PerMesf, "00") & "/" & r_int_PerAnof
   r_int_PerMesi = CInt(cmb_PerMesi.ItemData(cmb_PerMesi.ListIndex))
   r_int_PerAnoi = CInt(ipp_PerAnoi.Text)
   
   moddat_g_str_FecIng = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMesf.Text, 1) & LCase(Mid(cmb_PerMesf.Text, 2, Len(cmb_PerMesf.Text))) & " del " & Format(r_int_PerAnof, "0000")
   
   grd_LisEEFF.Col = 0
   moddat_g_str_CodPrd = Trim(grd_LisEEFF & "")
   
   grd_LisEEFF.Col = 1
   moddat_g_str_CodSub = Trim(grd_LisEEFF)
   
   grd_LisEEFF.Col = 2
   moddat_g_str_TipCre = Trim(grd_LisEEFF)
         
   grd_LisEEFF.Col = 17
   moddat_g_str_NomPrd = UCase(Trim(grd_LisEEFF))
   
   grd_LisEEFF.Col = 18
   moddat_g_str_NomPrd = moddat_g_str_NomPrd & " " & IIf(Len(Trim(grd_LisEEFF)) > 0, " - " & UCase(Trim(grd_LisEEFF)), "")
   
   Call gs_RefrescaGrid(grd_LisEEFF)
   
   If moddat_g_str_TipCre <> "F" And moddat_g_str_TipCre <> "L" And moddat_g_str_TipCre <> "" Then
        frm_RptCtb_20.Show 1
   End If
End Sub

Private Sub fs_GenExcRes_AntVer()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim q                   As Integer

   r_int_nrofil = 4
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(4, 2) = "EJERCICIOS"

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4), "-") = 0 Then
          .Cells(r_int_nrofil, 4) = "'" & "ENE " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 4) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5), "-") = 0 Then
          .Cells(r_int_nrofil, 5) = "'" & "FEB " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 5) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6), "-") = 0 Then
         .Cells(r_int_nrofil, 6) = "'" & "MAR " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_nrofil, 6) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7), "-") = 0 Then
         .Cells(r_int_nrofil, 7) = "'" & "ABR " & Right(r_int_PerAno, 2)
      Else
         .Cells(r_int_nrofil, 7) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8), "-") = 0 Then
          .Cells(r_int_nrofil, 8) = "'" & "MAY " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 8) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9), "-") = 0 Then
          .Cells(r_int_nrofil, 9) = "'" & "JUN " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 9) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10), "-") = 0 Then
          .Cells(r_int_nrofil, 10) = "'" & "JUL " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 10) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11), "-") = 0 Then
          .Cells(r_int_nrofil, 11) = "'" & "AGO " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 11) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12), "-") = 0 Then
          .Cells(r_int_nrofil, 12) = "'" & "SET " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 12) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12) & ""
      End If
        
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13), "-") = 0 Then
          .Cells(r_int_nrofil, 13) = "'" & "OCT " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 13) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14), "-") = 0 Then
          .Cells(r_int_nrofil, 14) = "'" & "NOV " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 14) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15), "-") = 0 Then
          .Cells(r_int_nrofil, 15) = "'" & "DIC " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 15) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15) & ""
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 16), "-") = 0 Then
          .Cells(r_int_nrofil, 16) = "'" & "ACUM " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_nrofil, 16) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 16) & ""
      End If
      
      .Range(.Cells(r_int_nrofil, 2), .Cells(4, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 2), .Cells(4, 16)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 3), .Cells(4, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_nrofil + 1, 2), .Cells(200, 3)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 5
      .Columns("C").ColumnWidth = 37
      .Columns("D").ColumnWidth = 11
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0"
      .Columns("E").ColumnWidth = 11
      .Columns("E").NumberFormat = "###,###,###,##0"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 11
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 11
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 11
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 11
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 11
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 11
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 11
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 11
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 11
      .Columns("O").NumberFormat = "###,###,###,##0"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 12
      .Columns("P").NumberFormat = "###,###,###,##0"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11

      g_str_Parame = "SELECT * "
      g_str_Parame = g_str_Parame & "FROM TT_EEFF WHERE "
      g_str_Parame = g_str_Parame & "INDTIPO <> 'D' "
      g_str_Parame = g_str_Parame & "  AND USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "  AND TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If

      r_int_Contad = 4

      Do While Not g_rst_Princi.EOF
         r_int_Contad = r_int_Contad + 1
         If Trim(g_rst_Princi!INDTIPO) = "L" Then
            g_rst_Princi.MoveNext
            r_int_Contad = r_int_Contad + 1
         End If
          
         If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
            .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_Contad, 4), .Cells(r_int_Contad, 16)).Font.Bold = True
         End If
         If Trim(g_rst_Princi!INDTIPO) = "S" Then
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
         End If
        
          If frm_RptCtb_19.cmb_PerMes.ListIndex = 11 Then 'DICIEMBRE
                q = 6
                     
                .Cells(r_int_Contad, q + 9) = g_rst_Princi!MES12
S1:
                .Cells(r_int_Contad, q + 8) = g_rst_Princi!MES11
S2:
                .Cells(r_int_Contad, q + 7) = g_rst_Princi!MES10
S3:
                .Cells(r_int_Contad, q + 6) = g_rst_Princi!MES09
S4:
                .Cells(r_int_Contad, q + 5) = g_rst_Princi!MES08
S5:
                .Cells(r_int_Contad, q + 4) = g_rst_Princi!MES07
S6:
                .Cells(r_int_Contad, q + 3) = g_rst_Princi!MES06
S7:
                .Cells(r_int_Contad, q + 2) = g_rst_Princi!MES05
S8:
                .Cells(r_int_Contad, q + 1) = g_rst_Princi!MES04
S9:
                .Cells(r_int_Contad, q) = g_rst_Princi!MES03
S10:
                .Cells(r_int_Contad, q - 1) = g_rst_Princi!MES02
S11:
                .Cells(r_int_Contad, q - 2) = g_rst_Princi!MES01
          
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 10 Then   'NOVIEMBRE
            q = 7
            GoTo S1
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 9 Then    'OCTUBRE
            q = 8
            GoTo S2
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 8 Then    'SETIEMBRE
            q = 9
            GoTo S3
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 7 Then    'AGOSTO
            q = 10
            GoTo S4
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 6 Then    'JULIO
            q = 11
            GoTo S5
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 5 Then    'JUNIO
            q = 12
            GoTo S6
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 4 Then    'MAYO
            q = 13
            GoTo S7
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 3 Then    'ABRIL
            q = 14
            GoTo S8
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 2 Then    'MARZO
            q = 15
            GoTo S9
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 1 Then    'FEBRERO
            q = 16
            GoTo S10
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 0 Then    'ENERO
            q = 17
            GoTo S11
          End If
          
          .Cells(r_int_Contad, 16) = g_rst_Princi!ACUMU
          g_rst_Princi.MoveNext
          DoEvents
       Loop

       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       
      'PARA AÑO ANTERIOR
      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If

      r_int_Contad = 4
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
      
          If Trim(g_rst_GenAux!INDTIPO) <> "D" Then
                  
               r_int_Contad = r_int_Contad + 1
               If Trim(g_rst_GenAux!INDTIPO) = "L" Then
                 g_rst_GenAux.MoveNext
                 r_int_Contad = r_int_Contad + 1
               End If
                
               If frm_RptCtb_19.cmb_PerMes.ListIndex = 11 Then       'DICIEMBRE
                     GoTo SALTO1
                             
                     .Cells(r_int_Contad, q + 7) = g_rst_GenAux!MES01
S12:
                     .Cells(r_int_Contad, q + 8) = g_rst_GenAux!MES02
S13:
                     .Cells(r_int_Contad, q + 9) = g_rst_GenAux!MES03
S14:
                     .Cells(r_int_Contad, q + 10) = g_rst_GenAux!MES04
S15:
                     .Cells(r_int_Contad, q + 11) = g_rst_GenAux!MES05
S16:
                     .Cells(r_int_Contad, q + 12) = g_rst_GenAux!MES06
S17:
                     .Cells(r_int_Contad, q + 13) = g_rst_GenAux!MES07
S18:
                     .Cells(r_int_Contad, q + 14) = g_rst_GenAux!MES08
S19:
                     .Cells(r_int_Contad, q + 15) = g_rst_GenAux!MES09
S20:
                     .Cells(r_int_Contad, q + 16) = g_rst_GenAux!MES10
S21:
                     .Cells(r_int_Contad, q + 17) = g_rst_GenAux!MES11
S22:
                     .Cells(r_int_Contad, q + 18) = g_rst_GenAux!MES12
                  
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 10 Then   'NOVIEMBRE
                 q = -14
                 GoTo S22
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 9 Then    'OCTUBRE
                 q = -13
                 GoTo S21
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 8 Then    'SETIEMBRE
                 q = -12
                 GoTo S20
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 7 Then    'AGOSTO
                 q = -11
                 GoTo S19
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 6 Then    'JULIO
                 q = -10
                 GoTo S18
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 5 Then    'JUNIO
                 q = -9
                 GoTo S17
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 4 Then    'MAYO
                 q = -8
                 GoTo S16
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 3 Then    'ABRIL
                 q = -7
                 GoTo S15
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 2 Then    'MARZO
                 q = -6
                 GoTo S14
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 1 Then    'FEBRERO
                 q = -5
                 GoTo S13
               ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 0 Then    'ENERO
                 q = -4
                 GoTo S12
               End If
          End If
          g_rst_GenAux.MoveNext
          DoEvents
       Loop

SALTO1:
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcRes_NueVer()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_nrofil = 5
    r_int_NoFlLi = 2
    r_int_PerMesi = CInt(cmb_PerMesi.ItemData(Me.cmb_PerMesi.ListIndex))
    r_int_PerAnoi = CInt(ipp_PerAnoi.Text)
    r_int_PerMesf = CInt(cmb_PerMesf.ItemData(Me.cmb_PerMesf.ListIndex))
    r_int_PerAnof = CInt(ipp_PerAnof.Text)
    'r_str_FecRpt = "01/" & Format(r_int_PerMesi, "00") & "/" & r_int_PerAnoi & " AL " & Format(ff_Ultimo_Dia_Mes(r_int_PerMesf, r_int_PerAnof), "00") & "/" & Format(r_int_PerMesf, "00") & "/" & r_int_PerAnof
    r_str_FecRpt = Format(ff_Ultimo_Dia_Mes(r_int_PerMesf, r_int_PerAnof), "00") & "/" & Format(r_int_PerMesf, "00") & "/" & r_int_PerAnof
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
        .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
        .Range(.Cells(1, 2), .Cells(1, 3)).Merge
        .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
        .Cells(2, 2) = "Del " & "01 de " & Left(Me.cmb_PerMesi.Text, 1) & LCase(Mid(Me.cmb_PerMesi.Text, 2, Len(cmb_PerMesi.Text))) & " del " & Me.ipp_PerAnoi.Text & " Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(Me.cmb_PerMesf.Text, 1) & LCase(Mid(Me.cmb_PerMesf.Text, 2, Len(cmb_PerMesf.Text))) & " del " & Format(r_int_PerAnof, "0000")
        .Range(.Cells(2, 2), .Cells(2, 3)).Merge
        .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
        .Cells(3, 2) = "( En Soles )"
        .Cells(5, 2) = "EJERCICIOS"
        
         For r_int_Contad = 4 To grd_LisEEFF.Cols - 3
            .Cells(r_int_nrofil, r_int_Contad) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, r_int_Contad) & ""
         Next r_int_Contad
        
        .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_Contad - 1)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_Contad - 1)).Font.Bold = True
                
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 43 '37
        
        For r_int_Contad = 4 To grd_LisEEFF.Cols - 3
           .Columns(r_int_Contad).HorizontalAlignment = xlHAlignRight
           .Columns(r_int_Contad).NumberFormat = "###,###,###,##0.00"
           .Columns(r_int_Contad).ColumnWidth = 14.5
        Next r_int_Contad
                
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
         
        r_int_nrofil = r_int_nrofil + 2
        .Range(.Cells(5, 4), .Cells(5, r_int_Contad - 1)).HorizontalAlignment = xlHAlignCenter
         
        For r_int_NoFlLi = 2 To grd_LisEEFF.Rows - 1
            If Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 2)) = "G" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 2)) = "F" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 2)) = "A" _
                 Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 2)) = "T" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 2)) = "X" Or Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 3)) = "PASIVO" Then
                'TITULO
                .Cells(r_int_nrofil, 2) = grd_LisEEFF.TextMatrix(r_int_NoFlLi, 3)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_Contad - 1)).Interior.Color = RGB(146, 208, 80)
                .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, r_int_Contad - 1)).Font.Bold = True
            Else
                .Cells(r_int_nrofil, 3) = Trim(grd_LisEEFF.TextMatrix(r_int_NoFlLi, 3))
            End If
             
             For r_int_Contad = 4 To grd_LisEEFF.Cols - 3
               .Cells(r_int_nrofil, r_int_Contad) = grd_LisEEFF.TextMatrix(r_int_NoFlLi, r_int_Contad)
             Next r_int_Contad
            
            r_int_nrofil = r_int_nrofil + 1
        Next r_int_NoFlLi
        
       .Columns("C:C").EntireColumn.AutoFit
   End With
   
   r_obj_Excel.Cells(6, 4).Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   r_obj_Excel.Visible = True
End Sub

Private Sub fs_GenExcDet_AntVer()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim A                   As Integer
Dim B                   As Integer
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_Contad = 4
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = "Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Format(r_int_PerAno, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      
      .Cells(4, 2) = "EJERCICIOS"

      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4), "-") = 0 Then
          .Cells(r_int_Contad, 4) = "'" & "ENE " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 4) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 4)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5), "-") = 0 Then
          .Cells(r_int_Contad, 5) = "'" & "FEB " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 5) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 5)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6), "-") = 0 Then
          .Cells(r_int_Contad, 6) = "'" & "MAR " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 6) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 6)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7), "-") = 0 Then
          .Cells(r_int_Contad, 7) = "'" & "ABR " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 7) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 7)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8), "-") = 0 Then
          .Cells(r_int_Contad, 8) = "'" & "MAY " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 8) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 8)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9), "-") = 0 Then
          .Cells(r_int_Contad, 9) = "'" & "JUN " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 9) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 9)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10), "-") = 0 Then
          .Cells(r_int_Contad, 10) = "'" & "JUL " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 10) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 10)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11), "-") = 0 Then
          .Cells(r_int_Contad, 11) = "'" & "AGO " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 11) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 11)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12), "-") = 0 Then
          .Cells(r_int_Contad, 12) = "'" & "SEP " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 12) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 12)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13), "-") = 0 Then
          .Cells(r_int_Contad, 13) = "'" & "OCT " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 13) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 13)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14), "-") = 0 Then
          .Cells(r_int_Contad, 14) = "'" & "NOV " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 14) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 14)
      End If
      
      If InStr(frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15), "-") = 0 Then
          .Cells(r_int_Contad, 15) = "'" & "DIC " & Right(r_int_PerAno, 2)
      Else
          .Cells(r_int_Contad, 15) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 15)
      End If
        
      .Cells(r_int_Contad, 16) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, 16)
      .Range(.Cells(4, 2), .Cells(4, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 16)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 37
      .Columns("D").ColumnWidth = 11
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,###,##0"
      .Columns("E").ColumnWidth = 11
      .Columns("E").NumberFormat = "###,###,###,##0"
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("F").ColumnWidth = 11
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 11
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 11
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 11
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 11
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 11
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 11
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 11
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 11
      .Columns("O").NumberFormat = "###,###,###,##0"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 12
      .Columns("P").NumberFormat = "###,###,###,##0"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      g_str_Parame = "SELECT * "
      g_str_Parame = g_str_Parame & "FROM TT_EEFF  "
      g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
      'g_str_Parame = g_str_Parame & "INDTIPO <> 'D' "
      g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      'AÑO ACTUAL
      r_int_Contad = 4
      Do While Not g_rst_Princi.EOF
         r_int_Contad = r_int_Contad + 1
          
         If Trim(g_rst_Princi!INDTIPO) = "L" Then
            g_rst_Princi.MoveNext
            r_int_Contad = r_int_Contad + 1
         End If
         If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
            .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Font.Bold = True
         End If
         If Trim(g_rst_Princi!INDTIPO) = "S" Then
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 16)).Font.Bold = True
         End If
         If Trim(g_rst_Princi!INDTIPO) = "D" Then
            .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
            .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
         End If
          
         If frm_RptCtb_19.cmb_PerMes.ListIndex = 11 Then
            A = 5
            .Cells(r_int_Contad, A + 10) = g_rst_Princi!MES12
S1:
            .Cells(r_int_Contad, A + 9) = g_rst_Princi!MES11
S2:
            .Cells(r_int_Contad, A + 8) = g_rst_Princi!MES10
S3:
            .Cells(r_int_Contad, A + 7) = g_rst_Princi!MES09
S4:
            .Cells(r_int_Contad, A + 6) = g_rst_Princi!MES08
S5:
            .Cells(r_int_Contad, A + 5) = g_rst_Princi!MES07
S6:
            .Cells(r_int_Contad, A + 4) = g_rst_Princi!MES06
S7:
            .Cells(r_int_Contad, A + 3) = g_rst_Princi!MES05
S8:
            .Cells(r_int_Contad, A + 2) = g_rst_Princi!MES04
S9:
            .Cells(r_int_Contad, A + 1) = g_rst_Princi!MES03
S10:
            .Cells(r_int_Contad, A) = g_rst_Princi!MES02
S11:
            .Cells(r_int_Contad, A - 1) = g_rst_Princi!MES01

          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 10 Then
            A = 6
            GoTo S1
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 9 Then
            A = 7
            GoTo S2
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 8 Then
            A = 8
            GoTo S3
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 7 Then
            A = 9
            GoTo S4
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 6 Then
            A = 10
            GoTo S5
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 5 Then
            A = 11
            GoTo S6
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 4 Then
            A = 12
            GoTo S7
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 3 Then
            A = 13
            GoTo S8
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 2 Then
            A = 14
            GoTo S9
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 1 Then
            A = 15
            GoTo S10
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 0 Then
            A = 16
            GoTo S11
          End If
          
          Dim total As Currency
          
          If g_rst_Princi!ACUMU = 0 Then
             total = g_rst_Princi!MES01 + g_rst_Princi!MES02 + g_rst_Princi!MES03 + g_rst_Princi!MES04 + _
                      g_rst_Princi!MES05 + g_rst_Princi!MES06 + g_rst_Princi!MES07 + g_rst_Princi!MES08 + _
                      g_rst_Princi!MES09 + g_rst_Princi!MES10 + g_rst_Princi!MES11 + g_rst_Princi!MES12
          
             .Cells(r_int_Contad, 16) = total
            
          Else
             .Cells(r_int_Contad, 16) = g_rst_Princi!ACUMU
          End If
                      
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
       
      'AÑO ANTERIOR
      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_Contad = 4
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
         r_int_Contad = r_int_Contad + 1
            
         If Trim(g_rst_GenAux!INDTIPO) = "L" Then
            g_rst_GenAux.MoveNext
            r_int_Contad = r_int_Contad + 1
         End If
          
         If frm_RptCtb_19.cmb_PerMes.ListIndex = 11 Then
            GoTo SALTO1
S12:
            .Cells(r_int_Contad, B) = g_rst_GenAux!MES02
S13:
            .Cells(r_int_Contad, B + 1) = g_rst_GenAux!MES03
S14:
            .Cells(r_int_Contad, B + 2) = g_rst_GenAux!MES04
S15:
            .Cells(r_int_Contad, B + 3) = g_rst_GenAux!MES05
S16:
            .Cells(r_int_Contad, B + 4) = g_rst_GenAux!MES06
S17:
            .Cells(r_int_Contad, B + 5) = g_rst_GenAux!MES07
S18:
            .Cells(r_int_Contad, B + 6) = g_rst_GenAux!MES08
S19:
            .Cells(r_int_Contad, B + 7) = g_rst_GenAux!MES09
S20:
            .Cells(r_int_Contad, B + 8) = g_rst_GenAux!MES10
S21:
            .Cells(r_int_Contad, B + 9) = g_rst_GenAux!MES11
S22:
            .Cells(r_int_Contad, B + 10) = g_rst_GenAux!MES12

          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 10 Then
            B = -6
            GoTo S22
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 9 Then
            B = -5
            GoTo S21
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 8 Then
            B = -4
            GoTo S20
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 7 Then
            B = -3
            GoTo S19
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 6 Then
            B = -2
            GoTo S18
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 5 Then
            B = -1
            GoTo S17
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 4 Then
            B = 0
            GoTo S16
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 3 Then
            B = 1
            GoTo S15
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 2 Then
            B = 2
            GoTo S14
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 1 Then
            B = 3
            GoTo S13
          ElseIf frm_RptCtb_19.cmb_PerMes.ListIndex = 0 Then
            B = 4
            GoTo S12
          End If
           
          g_rst_GenAux.MoveNext
          
          DoEvents
       Loop

SALTO1:
   End With
         
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcDet_NueVer()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_ConAux        As Integer
Dim r_str_PerMesi       As String
Dim r_str_PerAnoi       As String
Dim r_str_PerMesf       As String
Dim r_str_PerAnof       As String
Dim r_int_ConMes        As Integer
Dim r_int_ConAnn        As Integer
Dim r_int_VarAux1       As Integer
Dim r_int_VarAux2       As Integer
    
   r_int_Contad = 5
   r_int_PerMesi = CInt(cmb_PerMesi.ItemData(Me.cmb_PerMesi.ListIndex))
   r_int_PerAnoi = CInt(ipp_PerAnoi.Text)
   r_int_PerMesf = CInt(cmb_PerMesf.ItemData(Me.cmb_PerMesf.ListIndex))
   r_int_PerAnof = CInt(ipp_PerAnof.Text)
   r_str_FecRpt = Format(ff_Ultimo_Dia_Mes(r_int_PerMesf, r_int_PerAnof), "00") & "/" & Format(r_int_PerMesf, "00") & "/" & r_int_PerAnof
  
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE"
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE GESTION FINANCIERA"
      .Range(.Cells(1, 2), .Cells(1, 3)).Merge
      .Range(.Cells(1, 2), .Cells(1, 3)).Font.Bold = True
      .Cells(2, 2) = "Del " & "01 de " & Left(Me.cmb_PerMesi.Text, 1) & LCase(Mid(Me.cmb_PerMesi.Text, 2, Len(cmb_PerMesi.Text))) & " del " & Me.ipp_PerAnoi.Text & " Al " & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & " de " & Left(Me.cmb_PerMesf.Text, 1) & LCase(Mid(Me.cmb_PerMesf.Text, 2, Len(cmb_PerMesf.Text))) & " del " & Format(r_int_PerAnof, "0000")
      .Range(.Cells(2, 2), .Cells(2, 3)).Merge
      .Range(.Cells(2, 2), .Cells(2, 3)).Font.Bold = True
      .Cells(3, 2) = "( En Soles )"
      .Cells(5, 2) = "EJERCICIOS"
      
       For r_int_ConAux = 4 To grd_LisEEFF.Cols - 3
            .Cells(r_int_Contad, r_int_ConAux) = "'" & frm_RptCtb_19.grd_LisEEFF.TextMatrix(0, r_int_ConAux) & ""
       Next r_int_ConAux
      
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, r_int_ConAux - 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, r_int_ConAux - 1)).Font.Bold = True
      .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, r_int_ConAux - 1)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13
      .Columns("C").ColumnWidth = 56
      
      For r_int_ConAux = 4 To grd_LisEEFF.Cols - 3
           .Columns(r_int_ConAux).HorizontalAlignment = xlHAlignRight
           .Columns(r_int_ConAux).NumberFormat = "###,###,###,##0.00"
           .Columns(r_int_ConAux).ColumnWidth = 14.5
      Next r_int_ConAux

      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
      
      'r_int_NroFil = r_int_NroFil + 2
      .Range(.Cells(5, 4), .Cells(5, grd_LisEEFF.Cols - 3)).HorizontalAlignment = xlHAlignCenter

      'PARA ÚLTIMO AÑO CONSULTADO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TT_EEFF  "
      g_str_Parame = g_str_Parame & " WHERE USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame & "   AND TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame & " ORDER BY GRUPO, SUBGRP, ITEM, INDTIPO "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
        
      r_int_Contad = r_int_Contad + 1
        
      Do While Not g_rst_Princi.EOF
            r_int_Contad = r_int_Contad + 1
            If Trim(g_rst_Princi!INDTIPO) = "L" Then
              g_rst_Princi.MoveNext
              r_int_Contad = r_int_Contad + 1
            End If
            If Trim(g_rst_Princi!INDTIPO) = "G" Or Trim(g_rst_Princi!INDTIPO) = "F" Then
              .Cells(r_int_Contad, 2) = Trim(g_rst_Princi!NOMGRUPO)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, r_int_ConAux - 1)).Interior.Color = RGB(146, 208, 80)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, r_int_ConAux - 1)).Font.Bold = True
            End If
            If Trim(g_rst_Princi!INDTIPO) = "S" Then
              .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMSUBGRP)
              .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, r_int_ConAux - 1)).Font.Bold = True
            End If
            If Trim(g_rst_Princi!INDTIPO) = "D" Then
              .Cells(r_int_Contad, 2) = "'" & Trim(g_rst_Princi!CNTACTBLE & "")
              .Cells(r_int_Contad, 3) = Trim(g_rst_Princi!NOMCTA & "")
            End If
            
            If Me.ipp_PerAnoi.Text = Me.ipp_PerAnof.Text Then
               r_int_ConMes = Me.cmb_PerMesi.ListIndex + 1
               
               If CInt(r_int_ConMes) = 12 Then
                  r_int_VarAux1 = -2
                  GoTo L
A:
                  If r_int_PerMesf + 1 = 1 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES01, 2)
                  End If
   
B:
                  If r_int_PerMesf + 1 = 2 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES02, 2)
                  End If
C:
                  If r_int_PerMesf + 1 = 3 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES03, 2)
                  End If
D:
                  If r_int_PerMesf + 1 = 4 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES04, 2)
                  End If
E:
                  If r_int_PerMesf + 1 = 5 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES05, 2)
                  End If
F:
                  If r_int_PerMesf + 1 = 6 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES06, 2)
                  End If
G:
                  If r_int_PerMesf + 1 = 7 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_Princi!MES07, 2)
                  End If
H:
                  If r_int_PerMesf + 1 = 8 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_Princi!MES08, 2)
                  End If
i:
                  If r_int_PerMesf + 1 = 9 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_Princi!MES09, 2)
                  End If
j:
                  If r_int_PerMesf + 1 = 10 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_Princi!MES10, 2)
                  End If
k:
                  If r_int_PerMesf + 1 = 11 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_Princi!MES11, 2)
                  End If
L:
                  If r_int_PerMesf + 1 = 12 Then
                     GoTo Salir
                  Else
                     .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_Princi!MES12, 2)
                  End If
               
               ElseIf CInt(r_int_ConMes) = 11 Then
                       r_int_VarAux1 = -1
                       GoTo k
               ElseIf CInt(r_int_ConMes) = 10 Then
                       r_int_VarAux1 = 0
                       GoTo j
               ElseIf CInt(r_int_ConMes) = 9 Then
                       r_int_VarAux1 = 1
                       GoTo i
               ElseIf CInt(r_int_ConMes) = 8 Then
                       r_int_VarAux1 = 2
                       GoTo H
               ElseIf CInt(r_int_ConMes) = 7 Then
                       r_int_VarAux1 = 3
                       GoTo G
               ElseIf CInt(r_int_ConMes) = 6 Then
                       r_int_VarAux1 = 4
                       GoTo F
               ElseIf CInt(r_int_ConMes) = 5 Then
                       r_int_VarAux1 = 5
                       GoTo E
               ElseIf CInt(r_int_ConMes) = 4 Then
                       r_int_VarAux1 = 6
                       GoTo D
               ElseIf CInt(r_int_ConMes) = 3 Then
                       r_int_VarAux1 = 7
                       GoTo C
               ElseIf CInt(r_int_ConMes) = 2 Then
                       r_int_VarAux1 = 8
                       GoTo B
               ElseIf CInt(r_int_ConMes) = 1 Then
                       r_int_VarAux1 = 9
                       GoTo A
               End If
              
            Else
               r_int_VarAux1 = r_int_ConAux - 2
               r_int_ConMes = Me.cmb_PerMesf.ListIndex
         
               If r_int_ConMes = 11 Then                                  'DICIEMBRE
                 .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES12, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES11, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES10, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES09, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES08, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES07, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES06, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 7) = FormatNumber(g_rst_Princi!MES05, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 8) = FormatNumber(g_rst_Princi!MES04, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 9) = FormatNumber(g_rst_Princi!MES03, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 10) = FormatNumber(g_rst_Princi!MES02, 2)
                 .Cells(r_int_Contad, r_int_VarAux1 - 11) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 10 Then                               'NOVIEMBRE
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES11, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES10, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES09, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES08, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES07, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 7) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 8) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 9) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 10) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 9 Then                                   'OCTUBRE
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES10, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES09, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES08, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES07, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 7) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 8) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 9) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 8 Then                                   'SETIEMBRE
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES09, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES08, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES07, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 7) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 8) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 7 Then                                   'AGOSTO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES08, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES07, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 7) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 6 Then                                   'JULIO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES07, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 6) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 5 Then                                  'JUNIO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES06, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 5) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 4 Then                                   'MAYO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES05, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 4) = FormatNumber(g_rst_Princi!MES01, 2)
                   
               ElseIf r_int_ConMes = 3 Then                                  'ABRIL
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES04, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 3) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 2 Then                                   'MARZO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES03, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 2) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 1 Then                                   'FEBRERO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES02, 2)
                   .Cells(r_int_Contad, r_int_VarAux1 - 1) = FormatNumber(g_rst_Princi!MES01, 2)
   
               ElseIf r_int_ConMes = 0 Then                                   'ENERO
                   .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_Princi!MES01, 2)
   
               End If
            End If
                      
Salir:
          'MOSTRAR ACUMU
          If CInt(r_int_ConMes) > 0 Then
             If r_int_PerAnoi = r_int_PerAnof Then
                .Cells(r_int_Contad, (r_int_PerMesf - r_int_PerMesi) + 1 + 4).FormulaR1C1 = "=SUM(RC[-" & (r_int_PerMesf - r_int_PerMesi) + 1 & "]:RC[-1])"
             Else
                .Cells(r_int_Contad, r_int_VarAux1 + 1).FormulaR1C1 = "=SUM(RC[-" & (r_int_PerMesf) & "]:RC[-1])"
             End If
          Else
             .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_Princi!ACUMU, 2)
          End If
          
          g_rst_Princi.MoveNext
          DoEvents
       Loop
       
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       
       'INGRESO DE AÑOS ANTERIORES AL ÚLTIMO AÑO CONSULTADO
       If g_rst_GenAux.RecordCount > 0 Then g_rst_GenAux.MoveFirst
       
       For r_int_ConAnn = r_int_PerAnoi To r_int_PerAnof
      
         If r_int_ConAnn > r_int_PerAnoi And r_int_ConAnn <> r_int_PerAnof Then
            g_rst_GenAux.MoveFirst
            g_rst_GenAux.Find " anno = '" & r_int_ConAnn & "'"
            r_int_ConMes = 0
            
            If r_int_VarAux1 = 4 Then
               r_int_VarAux1 = (12 - r_int_PerMesi + 1) + 4
            Else
               r_int_VarAux1 = r_int_VarAux1 + 12
            End If
         
            r_int_Contad = 6 '8
            GoTo Ingresar
            
         ElseIf r_int_ConAnn = r_int_PerAnof Then
            Exit For
            
         Else
            g_rst_GenAux.MoveFirst
            g_rst_GenAux.Find " anno = '" & r_int_ConAnn & "'"
            r_int_ConMes = Me.cmb_PerMesi.ListIndex
            r_int_VarAux1 = 4
            r_int_Contad = 6
   
Ingresar:
            Do While Not g_rst_GenAux.EOF
                  r_int_Contad = r_int_Contad + 1
                  
                  If g_rst_GenAux!anno <> r_int_ConAnn Then GoTo Saltar
                     
                  If Trim(g_rst_GenAux!INDTIPO) = "L" Then
                    g_rst_GenAux.MoveNext
                    r_int_Contad = r_int_Contad + 1
                  End If
                  
                  If r_int_ConMes = 11 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                     'DICIEMBRE
                  
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES12, 2)
      
                  ElseIf r_int_ConMes = 10 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                 'NOVIEMBRE
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES12, 2)
      
                  ElseIf r_int_ConMes = 9 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'OCTUBRE
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 8 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'SETIEMBRE
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 7 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'AGOSTO
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 6 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'JULIO
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 5 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'JUNIO
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 4 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'MAYO
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES05, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 3 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                 'ABRIL
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES04, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES05, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES12, 2)
                      
                  ElseIf r_int_ConMes = 2 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                 'MARZO
      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES03, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES04, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES05, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES12, 2)
      
                  ElseIf r_int_ConMes = 1 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                  'FEBRERO
                      
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES02, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES03, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES04, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES05, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 10) = FormatNumber(g_rst_GenAux!MES12, 2)
      
                  ElseIf r_int_ConMes = 0 And g_rst_GenAux!anno <> Me.ipp_PerAnof Then                 'ENERO
                  
                      .Cells(r_int_Contad, r_int_VarAux1) = FormatNumber(g_rst_GenAux!MES01, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 1) = FormatNumber(g_rst_GenAux!MES02, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 2) = FormatNumber(g_rst_GenAux!MES03, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 3) = FormatNumber(g_rst_GenAux!MES04, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 4) = FormatNumber(g_rst_GenAux!MES05, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 5) = FormatNumber(g_rst_GenAux!MES06, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 6) = FormatNumber(g_rst_GenAux!MES07, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 7) = FormatNumber(g_rst_GenAux!MES08, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 8) = FormatNumber(g_rst_GenAux!MES09, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 9) = FormatNumber(g_rst_GenAux!MES10, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 10) = FormatNumber(g_rst_GenAux!MES11, 2)
                      .Cells(r_int_Contad, r_int_VarAux1 + 11) = FormatNumber(g_rst_GenAux!MES12, 2)
                  End If
Saltar:
               g_rst_GenAux.MoveNext
               
               DoEvents
            Loop
         End If
         
       Next r_int_ConAnn
   
   End With
   
   r_obj_Excel.Cells(6, 4).Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMesi_Click()
   If cmb_PerMesi.ListIndex > -1 Then
      Call gs_SetFocus(ipp_PerAnoi)
   End If
End Sub

Private Sub cmb_PerMesf_Click()
   If cmb_PerMesf.ListIndex > -1 Then
      Call gs_SetFocus(ipp_PerAnof)
   End If
End Sub

Private Sub cmb_PerMesi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMesi.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAnoi)
      End If
   End If
End Sub

Private Sub ipp_PerAnoi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMesf)
   End If
End Sub

Private Sub cmb_PerMesf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMesf.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAnof)
      End If
   End If
End Sub

Private Sub ipp_PerAnof_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

Function Sumar(MSHFlexGrid As Object, Fila As Integer) As Currency
   On Error GoTo error_function
  
   With MSHFlexGrid
        Dim total As Currency
        Dim i As Long
        For i = 4 To .Cols - 1 '.Rows - 1
            If IsNumeric(.TextMatrix(Fila, i)) Then
                total = total + .TextMatrix(Fila, i)
            End If
        Next
        Sumar = total
    End With
    Exit Function
   
error_function:
   MsgBox Err.Description, vbCritical, "error al sumar"
End Function

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub
